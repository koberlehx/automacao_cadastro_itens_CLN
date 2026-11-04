"""
Processa TODOS os arquivos .xlsx que estiverem na pasta de entrada manual de
uma vez, gerando sempre os dois modelos (itens + classificação) para cada um.
O filtro por Lote 01/02 pode ser calculado automaticamente, item por item,
dentro da conversão, ou forçado na hora de rodar.

Uso:
    python converter_manual.py

Estrutura de pastas esperada (relativa ao diretório de execução; pode ser
alterada pela variável de ambiente CONVERSOR_PASTA_DADOS):

    ./config/modelo_itens.xlsx
    ./config/modelo_classificacao.xlsx
    ./dados/manual/            <- coloque aqui as planilhas de origem
    ./dados/prontos/itens/
    ./dados/prontos/classificacao/
    ./dados/processados/
    ./dados/erros/
"""

import os
import shutil
import time
import traceback
from datetime import datetime

import pandas as pd

from converter_itens import (
    converter,
    detectar_tipo_origem,
    encontrar_aba_de_dados,
    escrever_planilha_consolidada,
    localizar_aba_e_cabecalho,
)

# ----------------------------- CONFIG -----------------------------
PASTA_CONFIG = os.path.join(".", "config")
PASTA_BASE = os.environ.get("CONVERSOR_PASTA_DADOS", os.path.join(".", "dados"))

CAMINHO_MODELO_ITENS = os.path.join(PASTA_CONFIG, "modelo_itens.xlsx")
CAMINHO_MODELO_CLASSIFICACAO = os.path.join(PASTA_CONFIG, "modelo_classificacao.xlsx")
PASTA_ORIGEM_MANUAL = os.path.join(PASTA_BASE, "manual")
PASTA_SAIDA_ITENS = os.path.join(PASTA_BASE, "prontos", "itens")
PASTA_SAIDA_CLASSIFICACAO = os.path.join(PASTA_BASE, "prontos", "classificacao")
PASTA_PROCESSADOS = os.path.join(PASTA_BASE, "processados")
PASTA_ERROS = os.path.join(PASTA_BASE, "erros")
# ------------------------------------------------------------------


def detectar_tipo(caminho_origem):
    aba, linha_cabecalho = localizar_aba_e_cabecalho(caminho_origem)
    df_cabecalho = pd.read_excel(caminho_origem, sheet_name=aba, header=linha_cabecalho, nrows=0)
    colunas = [str(c).strip() for c in df_cabecalho.columns]
    return detectar_tipo_origem(colunas)


def copiar_com_retry(origem, destino, tentativas=6, espera=3):
    for tentativa in range(1, tentativas + 1):
        try:
            shutil.copy2(origem, destino)
            return
        except PermissionError:
            if tentativa == tentativas:
                raise
            time.sleep(espera)


def apagar_com_retry(caminho, tentativas=6, espera=3):
    for tentativa in range(1, tentativas + 1):
        try:
            os.remove(caminho)
            return True
        except PermissionError:
            if tentativa == tentativas:
                return False
            time.sleep(espera)


def ja_foi_convertido(nome_sem_ext):
    """Evita reconverter (e duplicar) um arquivo que já foi processado
    antes e só ficou preso na pasta de entrada por não ter conseguido apagar."""
    prefixo = f"{nome_sem_ext}_ITEM_"
    return any(f.startswith(prefixo) for f in os.listdir(PASTA_SAIDA_ITENS))


def processar_arquivo(caminho_origem):
    nome_arquivo = os.path.basename(caminho_origem)
    nome_sem_ext = os.path.splitext(nome_arquivo)[0]
    carimbo = datetime.now().strftime("%Y%m%d_%H%M%S")

    if ja_foi_convertido(nome_sem_ext):
        if apagar_com_retry(caminho_origem):
            print(f"LIMPEZA: '{nome_arquivo}' já tinha sido convertido antes; removido da pasta de entrada agora.\n")
        else:
            print(f"AVISO: '{nome_arquivo}' já foi convertido, mas ainda não consegui apagar da pasta de entrada.\n")
        return

    print(f"Arquivo: {nome_arquivo}")

    try:
        tipo = detectar_tipo(caminho_origem)
    except Exception:
        tipo = None  # deixa o erro real aparecer na conversão, com detalhe salvo em erros

    ticket = None
    lote = None
    if tipo in ("cliente_padrao", "cliente_multimarca"):
        ticket = input("  Número do ticket (Referência) — deixe em branco se não souber: ").strip()
        ticket = ticket or None

        while True:
            resposta_lote = input(
                "  Lote — 1 (Atributos) / 2 (Classificação completa) / Enter (calcular automaticamente por item): "
            ).strip()
            if resposta_lote == "":
                lote = None
                break
            if resposta_lote == "1":
                lote = "01"
                break
            if resposta_lote == "2":
                lote = "02"
                break
            print("  Opção inválida, digite 1, 2 ou deixe em branco.")
    elif tipo == "erp_export":
        print("  (exportação do ERP detectada — Referência e Lote são lidos direto da planilha)")

    try:
        encontrar_aba_de_dados(caminho_origem)  # valida que é um arquivo reconhecível
        caminho_saida = os.path.join(PASTA_SAIDA_ITENS, f"{nome_sem_ext}_ITEM_{carimbo}.xlsx")
        caminho_saida_classificacao = os.path.join(
            PASTA_SAIDA_CLASSIFICACAO, f"{nome_sem_ext}_CLASSIFICACAO_{carimbo}.xlsx"
        )
        resultado = converter(
            caminho_origem, CAMINHO_MODELO_ITENS, caminho_saida,
            CAMINHO_MODELO_CLASSIFICACAO, caminho_saida_classificacao,
            referencia_forcada=ticket, lote_forcado=lote, retornar_linhas=True,
        )
    except Exception as e:
        os.makedirs(PASTA_ERROS, exist_ok=True)
        try:
            copiar_com_retry(caminho_origem, os.path.join(PASTA_ERROS, nome_arquivo))
            apagar_com_retry(caminho_origem)
        except Exception:
            pass
        detalhe = os.path.join(PASTA_ERROS, f"{nome_sem_ext}_ERRO_{carimbo}.txt")
        with open(detalhe, "w", encoding="utf-8") as f:
            f.write(traceback.format_exc())
        print(f"ERRO ao converter '{nome_arquivo}': {e} (detalhes em {detalhe})\n")
        return None

    # Chegou aqui: a conversão JÁ deu certo. Falha na limpeza não é erro de
    # conversão — só falta remover o original (tentaremos de novo depois).
    try:
        copiar_com_retry(caminho_origem, os.path.join(PASTA_PROCESSADOS, nome_arquivo))
        if apagar_com_retry(caminho_origem):
            print(f"OK: '{nome_arquivo}' convertido.\n")
        else:
            print(f"OK (com aviso): '{nome_arquivo}' convertido, mas ainda preso na pasta de entrada "
                  f"(tentaremos limpar na próxima execução).\n")
    except Exception:
        print(f"OK (com aviso): '{nome_arquivo}' convertido, mas não consegui copiar "
              f"para a pasta de processados (tentaremos de novo na próxima execução).\n")

    return resultado


def main():
    if not os.path.exists(CAMINHO_MODELO_ITENS):
        print(f"ERRO: modelo de itens não encontrado em {CAMINHO_MODELO_ITENS}")
        return
    if not os.path.exists(CAMINHO_MODELO_CLASSIFICACAO):
        print(f"ERRO: modelo de classificação não encontrado em {CAMINHO_MODELO_CLASSIFICACAO}")
        return

    os.makedirs(PASTA_ORIGEM_MANUAL, exist_ok=True)
    os.makedirs(PASTA_SAIDA_ITENS, exist_ok=True)
    os.makedirs(PASTA_SAIDA_CLASSIFICACAO, exist_ok=True)
    os.makedirs(PASTA_PROCESSADOS, exist_ok=True)

    arquivos = sorted(
        f for f in os.listdir(PASTA_ORIGEM_MANUAL)
        if f.lower().endswith((".xlsx", ".xls")) and not f.startswith("~$")
    )

    if not arquivos:
        print(f"Não achei nenhum arquivo .xlsx na pasta:\n  {PASTA_ORIGEM_MANUAL}")
        return

    print(f"{len(arquivos)} arquivo(s) encontrado(s). Processando...\n")

    carimbo_rodada = datetime.now().strftime("%Y%m%d_%H%M%S")
    cabecalho_itens = None
    todas_linhas_itens = []
    cabecalho_classificacao = None
    todas_linhas_classificacao = []

    for nome_arquivo in arquivos:
        resultado = processar_arquivo(os.path.join(PASTA_ORIGEM_MANUAL, nome_arquivo))
        if resultado:
            cabecalho_itens = resultado["cabecalho_itens"]
            todas_linhas_itens.extend(resultado["linhas_itens"])
            if resultado["linhas_classificacao"]:
                cabecalho_classificacao = resultado["cabecalho_classificacao"]
                todas_linhas_classificacao.extend(resultado["linhas_classificacao"])

    # Além dos arquivos individuais (já gerados por arquivo), gera também
    # 2 arquivos consolidados juntando tudo que foi processado nesta rodada.
    if todas_linhas_itens:
        caminho_consolidado_itens = os.path.join(
            PASTA_SAIDA_ITENS, f"CONSOLIDADO_ITEM_{carimbo_rodada}.xlsx"
        )
        escrever_planilha_consolidada(
            cabecalho_itens, todas_linhas_itens, CAMINHO_MODELO_ITENS, caminho_consolidado_itens
        )
        print(f"Consolidado de itens ({len(todas_linhas_itens)} linha(s)): {caminho_consolidado_itens}")

    if todas_linhas_classificacao:
        caminho_consolidado_classificacao = os.path.join(
            PASTA_SAIDA_CLASSIFICACAO, f"CONSOLIDADO_CLASSIFICACAO_{carimbo_rodada}.xlsx"
        )
        escrever_planilha_consolidada(
            cabecalho_classificacao, todas_linhas_classificacao,
            CAMINHO_MODELO_CLASSIFICACAO, caminho_consolidado_classificacao,
        )
        print(f"Consolidado de classificação ({len(todas_linhas_classificacao)} linha(s)): {caminho_consolidado_classificacao}")

    print(f"\nConcluído! Itens em: {PASTA_SAIDA_ITENS}\nClassificação em: {PASTA_SAIDA_CLASSIFICACAO}")
    input("\nAperte Enter para fechar...")


if __name__ == "__main__":
    main()

"""
Converte planilhas de origem para modelos usados em carga de itens numa
plataforma corporativa. Reconhece automaticamente TRÊS formatos de origem,
pelo cabeçalho da planilha:

  1) "cliente_padrao"      -> planilha de solicitação padrão do cliente
  2) "cliente_multimarca"  -> planilha com Part Numbers por marca
  3) "erp_export"          -> planilha de exportação do ERP_Corporativo

Em todos os casos, gera dois arquivos a partir da mesma leitura:
  1) modelo_itens.xlsx         -> todos os itens (cadastro geral)
  2) modelo_classificacao.xlsx -> todos os itens, com os campos de
     classificação já preenchidos quando a origem já os fornece

Uso:
    python converter_itens.py origem.xlsx modelo_itens.xlsx saida_itens.xlsx
    python converter_itens.py origem.xlsx modelo_itens.xlsx saida_itens.xlsx modelo_classificacao.xlsx saida_classificacao.xlsx

- origem.xlsx               : planilha de origem (cliente OU exportação do ERP)
- modelo_itens.xlsx         : planilha modelo em branco (define cabeçalho/ordem final)
- saida_itens.xlsx          : arquivo de itens gerado, pronto para carga
- modelo_classificacao.xlsx : (opcional) modelo em branco de classificação
- saida_classificacao.xlsx  : (opcional) arquivo de classificação gerado
"""

import os
import re
import sys

import pandas as pd
from openpyxl import load_workbook

# Caminho padrão do modelo em branco de classificação, usado quando o script
# é chamado com apenas 3 argumentos. Pode ser sobrescrito pela variável de
# ambiente CONVERSOR_MODELO_CLASSIFICACAO.
CAMINHO_MODELO_CLASSIFICACAO_PADRAO = os.environ.get(
    "CONVERSOR_MODELO_CLASSIFICACAO",
    os.path.join(".", "config", "modelo_classificacao.xlsx"),
)

# ============================================================
# MAPEAMENTOS — um conjunto por tipo de planilha de origem
# Formato: {coluna_na_origem: coluna_no_modelo_de_destino}
# ============================================================

# ----- Tipo 1: planilha de solicitação do cliente (formato padrão) -----
MAPA_ITENS_CLIENTE_PADRAO = {
    "Empresa (código interno)": "Empresa (código interno)",
    "Part Number": "Part Number",
    "Modalidade": "Modalidade",
    "Descrição Original": "Descrição Original",
    "Código Tarifário (NCM)": "Código Tarifário",
    "ERP": "ERP",
    "Grupo Produto": "Grupo Produto",
    # "Solicitante" e "Área Negócio" ficam sempre em branco nos dois tipos
    # de origem — por isso nunca entram nesses mapas.
    "Peso": "Peso",
    "Referência": "Referência",
}

MAPA_CLASSIFICACAO_CLIENTE_PADRAO = {
    "Empresa (código interno)": "Empresa (código interno)",
    "Part Number": "Part Number",
    "Modalidade": "Modalidade",
    "Descrição Longa": "Descrição Longa",
    "Descrição Curta": "Descrição Curta",
    "Observação Classificação": "Observações Classificação",
    "NCM": "NCM",
    "Destaque": "Destaque",
    "Naladi": "Naladi",
    "EX II": "Ex II",
    "EX IPI": "Ex IPI",
    "NVEs": "NVEs",
    "Necessidade Li": "Necessita LI",
}

# ----- Tipo 2: exportação do ERP_Corporativo -----
# Aqui os nomes já batem quase todos com o destino, então o mapa é bem direto.
MAPA_ITENS_ERP_EXPORT = {
    "Código interno": "Empresa (código interno)",
    "Part Number": "Part Number",
    "Modalidade": "Modalidade",
    "Descrição Original": "Descrição Original",
    "Código Tarifário": "Código Tarifário",
    "ERP": "ERP",
    "Grupo Produto": "Grupo Produto",
    "Peso": "Peso",
    "Observações": "Observações",
    "Referência": "Referência",
}

MAPA_CLASSIFICACAO_ERP_EXPORT = {
    "Código interno": "Empresa (código interno)",
    "Part Number": "Part Number",
    "Modalidade": "Modalidade",
    "Descrição Longa": "Descrição Longa",
    "Descrição Curta": "Descrição Curta",
    "Observações Classificação": "Observações Classificação",
    "NCM": "NCM",
    "Destaque": "Destaque",
    "Naladi": "Naladi",
    "Ex II": "Ex II",
    "Ex IPI": "Ex IPI",
    "NVEs": "NVEs",
    "Necessita LI": "Necessita LI",
    "Código Suframa": "Código Suframa",
    "Tipo Suframa": "Tipo Suframa",
    "Destaque Suframa": "Destaque Suframa",
    "Modelo Suframa": "Modelo Suframa",
    "ICMS SP": "ICMS SP",
}

# ----- Tipo 3: planilha multimarca -----
# Formato com 2 colunas de Part Number diferentes (usa-se a da marca do
# produto, não a interna do grupo) e sem alguns campos que os outros
# formatos têm (ERP, Grupo Produto, Peso, Código Tarifário próprio).
MAPA_ITENS_CLIENTE_MULTIMARCA = {
    "Empresa (código grupo)": "Empresa (código interno)",
    "PartNumber (Marca)": "Part Number",
    "Modalidade": "Modalidade",
    "Descrição Longa": "Descrição Original",
    "NCM": "Código Tarifário",
}

MAPA_CLASSIFICACAO_CLIENTE_MULTIMARCA = {
    "Empresa (código grupo)": "Empresa (código interno)",
    "PartNumber (Marca)": "Part Number",
    "Modalidade": "Modalidade",
    "Descrição Longa": "Descrição Longa",
    "Descrição Curta": "Descrição Curta",
    "Observação Classificação": "Observações Classificação",
    "NCM": "NCM",
}

# Campos que NUNCA podem ficar em branco, mesmo se vierem "Não"/vazios na origem
CAMPOS_OBRIGATORIOS = [
    "Empresa (código interno)",
    "Part Number",
    "Modalidade",
    "Descrição Original",
]

MODALIDADE_PADRAO = "IMPORTAÇÃO"
LOTE_ATRIBUTOS = "01"
LOTE_CLASSIFICACAO_FISCAL = "02"

# Esse campo é booleano de verdade (Sim/Não são respostas válidas, não
# devem ser tratadas como "sem informação").
CAMPO_BOOLEANO_CLASSIFICACAO = "Necessita LI"

# Esse campo precisa sempre ter 3 dígitos (ex: 81 -> "081", 999 -> "999")
CAMPO_DESTAQUE_CLASSIFICACAO = "Destaque"

# Esse campo vem em várias linhas (uma por código NVE) e precisa virar um
# único texto no formato "AA=XXXX;AB=XXXX;AC=XXXX" (códigos em ordem alfabética).
CAMPO_NVE_CLASSIFICACAO = "NVEs"

# Esse campo sempre precisa ir sem pontos (ex: "7318.15.00" -> "73181500")
CAMPO_NCM_CLASSIFICACAO = "NCM"
CAMPO_NCM_ITENS = "Código Tarifário"


def remover_pontos(valor):
    if valor is None:
        return None
    return valor.replace(".", "")


def valor_valido(valor):
    """Retorna True se o valor existir de verdade (não é vazio, NaN, nem uma
    string placeholder ('Não', 'Não informado') usada na planilha de origem
    quando o campo não se aplica)."""
    if pd.isna(valor):
        return False
    texto = str(valor).strip().lower()
    if texto in ("", "não", "nao", "não informado", "nao informado"):
        return False
    return True


def eh_numero_de_registro(valor):
    """Alguns clientes usam a coluna 'Registro Anvisa' como indicador
    Sim/Não (só avisa se precisa de registro), enquanto outros colocam o
    número de registro de verdade ali. Só tratamos como conteúdo válido para
    a Observação quando parece ser o número em si, não uma resposta Sim/Não."""
    if not valor_valido(valor):
        return False
    texto = str(valor).strip().lower()
    if texto in ("sim", "não", "nao", "n/a", "na"):
        return False
    return True


def limpar(valor):
    """Devolve o valor tratado (string sem espaço nas pontas) ou None se
    não for um valor válido (vazio/NaN/'Não'/'Não informado')."""
    if valor_valido(valor):
        return str(valor).strip()
    return None


def transformar_nve(valor):
    """Converte o texto de várias linhas (uma por código NVE) no formato
    'U - AC - TEOR DE COMONÔMERO: 0006 - Superior a 30 até 50' para o
    formato exigido pela plataforma: 'AA=0002;AB=0002;AC=0006' (códigos
    em ordem alfabética)."""
    if not valor_valido(valor):
        return None
    pares = []
    for linha in str(valor).splitlines():
        m = re.search(r"([A-Z]{2,3})\s*-\s*[^:]+:\s*(\S+)", linha)
        if m:
            pares.append((m.group(1), m.group(2)))
    if not pares:
        return None
    pares.sort(key=lambda p: p[0])
    return ";".join(f"{codigo}={numero}" for codigo, numero in pares)


def limpar_classificacao(col_modelo, valor):
    """Igual a limpar(), com tratamentos especiais para alguns campos:
    - Necessita LI: 'Sim'/'Não' são respostas de verdade, não devem virar None.
    - Destaque: sempre precisa ter 3 dígitos (completa com zero à esquerda).
    - NVEs: reformatado de várias linhas para 'AA=XXXX;AB=XXXX;AC=XXXX'.
    - NCM: sempre sem pontos."""
    if col_modelo == CAMPO_BOOLEANO_CLASSIFICACAO:
        if pd.isna(valor) or str(valor).strip() == "":
            return None
        return str(valor).strip().capitalize()

    if col_modelo == CAMPO_NVE_CLASSIFICACAO:
        return transformar_nve(valor)

    valor_limpo = limpar(valor)

    if col_modelo == CAMPO_DESTAQUE_CLASSIFICACAO and valor_limpo is not None:
        if valor_limpo.isdigit():
            valor_limpo = valor_limpo.zfill(3)

    if col_modelo == CAMPO_NCM_CLASSIFICACAO:
        valor_limpo = remover_pontos(valor_limpo)

    return valor_limpo


def encontrar_aba_de_dados(caminho_arquivo):
    """Procura, entre as abas do arquivo, a que contém a tabela de itens
    (mantido por compatibilidade — usa localizar_aba_e_cabecalho por trás)."""
    aba, _ = localizar_aba_e_cabecalho(caminho_arquivo)
    return aba


# Colunas usadas como "marcador": se uma linha contém pelo menos uma
# delas, consideramos que essa linha é o cabeçalho de verdade da tabela
# (o cabeçalho nem sempre fica na linha 1 — em alguns formatos vem
# depois de uma linha em branco ou de um título).
MARCADORES_CABECALHO = {"Part Number", "Código interno", "PartNumber (Marca)"}


def localizar_aba_e_cabecalho(caminho_arquivo, linhas_max=15):
    """Procura, em todas as abas do arquivo, a linha que contém algum dos
    MARCADORES_CABECALHO. Retorna (nome_aba, indice_da_linha_cabecalho)
    — o índice é 0-based, pronto para usar como 'header=' no pandas."""
    xl = pd.ExcelFile(caminho_arquivo)
    for nome_aba in xl.sheet_names:
        df_bruto = xl.parse(nome_aba, header=None, nrows=linhas_max)
        for i, linha in df_bruto.iterrows():
            valores = {str(v).strip() for v in linha if pd.notna(v)}
            if valores & MARCADORES_CABECALHO:
                return nome_aba, i
    raise ValueError(
        "Não encontrei nenhuma aba com um cabeçalho reconhecível "
        "(procurei por 'Part Number', 'Código interno' ou 'PartNumber (Marca)') "
        "no arquivo de origem."
    )


def detectar_tipo_origem(colunas):
    """Reconhece o tipo de planilha de origem pelo cabeçalho: solicitação
    de cliente (formato padrão), planilha multimarca ou exportação do
    ERP_Corporativo."""
    colunas = set(colunas)
    if "Código interno" in colunas and "Razão Social" in colunas:
        return "erp_export"
    if "PartNumber (Marca)" in colunas and "Empresa (código grupo)" in colunas:
        return "cliente_multimarca"
    if "Empresa (código interno)" in colunas:
        return "cliente_padrao"
    raise ValueError(
        "Não reconheço o formato desta planilha de origem — não encontrei "
        "'Empresa (código interno)' (formato padrão do cliente), "
        "'PartNumber (Marca)' + 'Empresa (código grupo)' (multimarca) nem "
        "'Código interno' + 'Razão Social' (exportação do ERP) no cabeçalho."
    )


def calcular_lote_cliente_padrao(codigo_tarifario):
    """Regra de negócio para planilhas de cliente:
    - Se o cliente já informou o Código Tarifário (NCM) -> pedido é só de
      atributos -> Lote 01.
    - Se o Código Tarifário veio em branco -> pedido precisa de
      classificação fiscal -> Lote 02.
    """
    if valor_valido(codigo_tarifario):
        return LOTE_ATRIBUTOS
    return LOTE_CLASSIFICACAO_FISCAL


def calcular_lote_erp_export(valor_lote_bruto):
    """Na exportação do ERP, a coluna 'Lote' vem com um texto tipo
    'S1_2024 - Itens Primeira semana 2024'. Só a parte antes do ' - '
    interessa (ex: 'S1_2024')."""
    valor_limpo = limpar(valor_lote_bruto)
    if valor_limpo is None:
        return None
    if " - " in valor_limpo:
        return valor_limpo.split(" - ")[0].strip()
    if "-" in valor_limpo:
        return valor_limpo.split("-")[0].strip()
    return valor_limpo


def calcular_observacao(observacoes_item, registro_anvisa, composicao):
    """(Só usado para planilhas de cliente.) Monta o campo Observações a
    partir de duas fontes independentes (cada uma entra se tiver conteúdo):
      1) Observações do Item, se preenchida.
      2) Registro Anvisa e/ou Composição, concatenados, se algum dos dois
         estiver preenchido.
    Se nada estiver preenchido, o resultado é None (fica em branco)."""
    partes_observacao = []

    if valor_valido(observacoes_item):
        partes_observacao.append(str(observacoes_item).strip())

    if eh_numero_de_registro(registro_anvisa) or valor_valido(composicao):
        sub_partes = []
        if eh_numero_de_registro(registro_anvisa):
            sub_partes.append(f"Registro anvisa: {str(registro_anvisa).strip()}")
        if valor_valido(composicao):
            sub_partes.append(f"Composição: {str(composicao).strip()}")
        partes_observacao.append(" | ".join(sub_partes))

    if not partes_observacao:
        return None
    return " | ".join(partes_observacao)


def gerar_saida_classificacao(df_origem, mapa_classificacao, caminho_modelo_classificacao, caminho_saida_classificacao):
    """Gera o modelo_classificacao.xlsx com TODOS os itens. Itens que já
    vêm com os campos de classificação preenchidos na origem saem
    preenchidos; os demais saem com esses campos em branco."""
    wb_modelo = load_workbook(caminho_modelo_classificacao)
    ws_modelo = wb_modelo.active
    cabecalho_modelo = [c.value for c in ws_modelo[1]]

    linhas_convertidas = []
    for idx, linha in df_origem.iterrows():
        linha_final = {}
        for col_modelo in cabecalho_modelo:
            col_origem = next(
                (k for k, v in mapa_classificacao.items() if v == col_modelo), None
            )
            valor_origem = linha.get(col_origem) if col_origem else None
            linha_final[col_modelo] = limpar_classificacao(col_modelo, valor_origem)

        linhas_convertidas.append(linha_final)

    if not linhas_convertidas:
        print("Nenhum item para colocar no arquivo de classificação.")
        return cabecalho_modelo, []

    for i, linha_final in enumerate(linhas_convertidas, start=2):
        for j, col_modelo in enumerate(cabecalho_modelo, start=1):
            ws_modelo.cell(row=i, column=j, value=linha_final[col_modelo])

    wb_modelo.save(caminho_saida_classificacao)
    print(f"{len(linhas_convertidas)} item(ns) gerados em: {caminho_saida_classificacao}")
    return cabecalho_modelo, linhas_convertidas


def escrever_planilha_consolidada(cabecalho, linhas, caminho_modelo, caminho_saida):
    """Escreve um arquivo novo a partir de linhas já processadas (usado
    para juntar vários arquivos de origem num único resultado consolidado)."""
    wb_modelo = load_workbook(caminho_modelo)
    ws_modelo = wb_modelo.active
    for i, linha_final in enumerate(linhas, start=2):
        for j, col_modelo in enumerate(cabecalho, start=1):
            ws_modelo.cell(row=i, column=j, value=linha_final.get(col_modelo))
    wb_modelo.save(caminho_saida)


def converter(caminho_origem, caminho_modelo, caminho_saida,
              caminho_modelo_classificacao=None, caminho_saida_classificacao=None,
              referencia_forcada=None, lote_forcado=None, retornar_linhas=False):
    aba, linha_cabecalho = localizar_aba_e_cabecalho(caminho_origem)
    df_origem = pd.read_excel(caminho_origem, sheet_name=aba, header=linha_cabecalho, dtype=str)
    df_origem.columns = [str(c).strip() for c in df_origem.columns]

    # Remove linhas totalmente vazias
    df_origem = df_origem.dropna(how="all")

    tipo_origem = detectar_tipo_origem(df_origem.columns)
    if tipo_origem == "cliente_padrao":
        mapa_itens = MAPA_ITENS_CLIENTE_PADRAO
        mapa_classificacao = MAPA_CLASSIFICACAO_CLIENTE_PADRAO
    elif tipo_origem == "cliente_multimarca":
        mapa_itens = MAPA_ITENS_CLIENTE_MULTIMARCA
        mapa_classificacao = MAPA_CLASSIFICACAO_CLIENTE_MULTIMARCA
    else:
        mapa_itens = MAPA_ITENS_ERP_EXPORT
        mapa_classificacao = MAPA_CLASSIFICACAO_ERP_EXPORT

    wb_modelo = load_workbook(caminho_modelo)
    ws_modelo = wb_modelo.active
    cabecalho_modelo = [c.value for c in ws_modelo[1]]

    linhas_convertidas = []
    avisos = []
    lotes_por_linha = {}

    for idx, linha in df_origem.iterrows():
        linha_final = {}

        # Mapeamento direto coluna a coluna, tratando "Não"/vazio como None
        for col_modelo in cabecalho_modelo:
            col_origem = next(
                (k for k, v in mapa_itens.items() if v == col_modelo), None
            )
            valor_origem = linha.get(col_origem) if col_origem else None
            linha_final[col_modelo] = limpar(valor_origem)

        # Regra: Código Tarifário sempre sem pontos (ex: "7318.15.00" -> "73181500")
        if CAMPO_NCM_ITENS in linha_final:
            linha_final[CAMPO_NCM_ITENS] = remover_pontos(linha_final[CAMPO_NCM_ITENS])

        # Regra: Modalidade em branco -> assume "IMPORTAÇÃO"
        if not linha_final.get("Modalidade"):
            linha_final["Modalidade"] = MODALIDADE_PADRAO

        # Regra: Lote — calculado de um jeito diferente conforme o tipo de origem
        if tipo_origem == "cliente_padrao":
            lote_calculado = lote_forcado or calcular_lote_cliente_padrao(
                linha.get("Código Tarifário (NCM)")
            )
        elif tipo_origem == "cliente_multimarca":
            lote_calculado = lote_forcado or calcular_lote_cliente_padrao(linha.get("NCM"))
        else:
            lote_calculado = lote_forcado or calcular_lote_erp_export(linha.get("Lote"))
        lotes_por_linha[idx] = lote_calculado
        if "Lote" in linha_final:
            linha_final["Lote"] = lote_calculado

        # Regra: Observações — só planilha de cliente (formato padrão) precisa
        # concatenar (Registro Anvisa + Composição); os outros formatos ou
        # já trazem pronto, ou nem têm esse campo na origem.
        if tipo_origem == "cliente_padrao" and "Observações" in linha_final:
            linha_final["Observações"] = calcular_observacao(
                linha.get("Observações do Item"),
                linha.get("Registro Anvisa"),
                linha.get("COMPOSIÇÃO"),
            )

        # Regra: se um número de ticket foi informado manualmente, ele
        # sempre substitui o que viria da planilha de origem.
        if referencia_forcada and "Referência" in linha_final:
            linha_final["Referência"] = referencia_forcada

        # Aviso se faltar algum campo obrigatório
        faltando = [c for c in CAMPOS_OBRIGATORIOS if not linha_final.get(c)]
        if faltando:
            avisos.append(
                f"Linha {idx + 2} da origem: faltando {', '.join(faltando)}"
            )

        linhas_convertidas.append(linha_final)

    # Aviso de Part Numbers duplicados (a plataforma não aceita 2 itens com
    # o mesmo Part Number na mesma planilha, mesmo que os dados sejam
    # diferentes). Não decide sozinho qual manter — só avisa pra revisão manual.
    part_numbers_linhas = {}
    for idx_linha, linha_final in enumerate(linhas_convertidas):
        pn = linha_final.get("Part Number")
        if pn:
            part_numbers_linhas.setdefault(pn, []).append(idx_linha + 2)
    for pn, linhas_pn in part_numbers_linhas.items():
        if len(linhas_pn) > 1:
            avisos.append(
                f"Part Number '{pn}' duplicado nas linhas: {', '.join(map(str, linhas_pn))} — revisar manualmente antes da carga"
            )

    # Escreve os dados a partir da linha 2, respeitando a ordem do cabeçalho do modelo
    for i, linha_final in enumerate(linhas_convertidas, start=2):
        for j, col_modelo in enumerate(cabecalho_modelo, start=1):
            ws_modelo.cell(row=i, column=j, value=linha_final[col_modelo])

    wb_modelo.save(caminho_saida)

    print(f"Tipo de origem detectado: '{tipo_origem}'")
    print(f"Aba de origem detectada: '{aba}'")
    print(f"{len(linhas_convertidas)} itens convertidos.")
    print(f"Arquivo gerado: {caminho_saida}")
    if avisos:
        print("\nAvisos:")
        for a in avisos:
            print(" -", a)

    # Gera também o modelo de classificação, se os caminhos foram informados.
    cabecalho_classificacao, linhas_classificacao = None, []
    if caminho_modelo_classificacao and caminho_saida_classificacao:
        cabecalho_classificacao, linhas_classificacao = gerar_saida_classificacao(
            df_origem, mapa_classificacao, caminho_modelo_classificacao, caminho_saida_classificacao
        )

    if retornar_linhas:
        return {
            "cabecalho_itens": cabecalho_modelo,
            "linhas_itens": linhas_convertidas,
            "cabecalho_classificacao": cabecalho_classificacao,
            "linhas_classificacao": linhas_classificacao,
        }


if __name__ == "__main__":
    if len(sys.argv) not in (4, 6):
        print(
            "Uso: python converter_itens.py origem.xlsx modelo_itens.xlsx saida_itens.xlsx "
            "[modelo_classificacao.xlsx saida_classificacao.xlsx]"
        )
        sys.exit(1)

    if len(sys.argv) == 6:
        converter(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5])
    else:
        # Com 3 argumentos, usa o modelo de classificação padrão
        # automaticamente (sem perguntar nada).
        origem, modelo_itens, saida_itens = sys.argv[1], sys.argv[2], sys.argv[3]
        if not os.path.exists(CAMINHO_MODELO_CLASSIFICACAO_PADRAO):
            print(
                f"\nAVISO: não encontrei o modelo de classificação em "
                f"'{CAMINHO_MODELO_CLASSIFICACAO_PADRAO}'. Ajuste a constante "
                f"CAMINHO_MODELO_CLASSIFICACAO_PADRAO no topo deste script, "
                f"defina a variável de ambiente CONVERSOR_MODELO_CLASSIFICACAO "
                f"ou rode de novo passando os 5 argumentos manualmente."
            )
            sys.exit(1)
        nome_base, _ = os.path.splitext(saida_itens)
        saida_classificacao = f"{nome_base}_classificacao.xlsx"
        converter(
            origem, modelo_itens, saida_itens,
            CAMINHO_MODELO_CLASSIFICACAO_PADRAO, saida_classificacao,
        )

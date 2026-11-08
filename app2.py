import io
import os
from copy import copy
from datetime import date, datetime
from typing import Dict, List, Tuple

import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import pandas as pd
import streamlit as st

# ==============================================================================
# 1. DICIONÁRIO DE MAPEAMENTO DO TEMPLATE EXCEL (CELL_MAPPING)
# Calibrado para o layout exato de "Relatorio Orcamento.xlsx" (Aba: "Orçamento")
# ==============================================================================
CELL_MAPPING = {
    "SHEET_NAME": "Orçamento",
    "PROPOSTA_NUM": "I7",
    "PROPOSTA_DATA": "L7",
    "PROPOSTA_DATA_P2": "L55",
    "DESTINATARIO": "B8",
    "TELEFONE": "F8",
    "CONTATO": "B9",
    "EMAIL": "F9",
    "PROJETO_BASICO_TEXTO": "B11",
    "SERVICO_TITULO": "B59",
    "START_ROW": 60,
    "END_TEMPLATE_ROW": 131,
    "COL_WBS": 2,
    "COL_DESC_START": 3,
    "COL_DESC_END": 6,
    "COL_UNIDADE": 7,
    "COL_QUANTIDADE": 8,
    "COL_VALOR": 9,
    "COL_MARGIN_START": 10,
    "COL_MARGIN_END": 13,
    "ROW_VALOR_GLOBAL": 133,
    "COL_RODAPE_VALORES": 4,       # Coluna D (Intervalo mesclado D133:E133 para o Valor Global numérico)
    "COL_VALOR_EXTENSO": 6,        # Coluna F (Intervalo mesclado F133:M133 ao lado do Valor Global)
    "ROW_PRAZO_EXECUCAO": 137,
    "ROW_IMPOSTOS": 138,
    "ROW_COND_COMERCIAIS": 139,
    "ROW_VALIDADE": 140,
    "MAX_FOOTER_ROW": 141,
}

SINAPI_CSV_PATH = "sinapi_ce.csv"
TEMPLATE_CANDIDATES = ["template_base.xlsx", "Relatorio Orcamento.xlsx"]


# ==============================================================================
# 2. CARREGAMENTO E NORMALIZAÇÃO DA BASE SINAPI-CE
# ==============================================================================
def normalize_numeric_series(series: pd.Series) -> pd.Series:
    """Converte strings numéricas no padrão BR (1.234,56) ou US (1234.56) para float."""
    def parse_val(val):
        if pd.isna(val) or str(val).strip() in ("", "-", "None", "nan"):
            return 0.0
        if isinstance(val, (int, float, np.number)):
            return float(val)
        s = str(val).strip().replace("R$", "").replace(" ", "")
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return 0.0

    return series.apply(parse_val)


@st.cache_data(show_spinner="Carregando base de dados SINAPI-CE...")
def load_sinapi_data(filepath: str) -> pd.DataFrame:
    """Carrega o arquivo sinapi_ce.csv (tratando BOM UTF-8) e normaliza tipos de dados."""
    df = None
    for enc in ("utf-8-sig", "utf-8", "latin1"):
        try:
            df = pd.read_csv(filepath, dtype=str, encoding=enc)
            if len(df.columns) == 1:
                df = pd.read_csv(filepath, dtype=str, encoding=enc, sep=None, engine="python")
            break
        except Exception:
            continue

    if df is None:
        df = pd.read_csv(filepath, dtype=str, encoding="latin1", sep=None, engine="python")

    # Remove BOM residual (\ufeff) e espaços extras nos nomes das colunas
    df.columns = [str(c).replace("\ufeff", "").strip() for c in df.columns]

    # Mapear variações de nomes caso necessário
    if "Codigo_SINAPI" not in df.columns:
        col_map = {}
        for col in df.columns:
            c_low = col.lower()
            if "cod" in c_low or "cód" in c_low:
                col_map[col] = "Codigo_SINAPI"
            elif "desc" in c_low:
                col_map[col] = "Descricao_SINAPI"
            elif "unid" in c_low or "medida" in c_low:
                col_map[col] = "Unidade"
            elif "sem" in c_low and "deson" in c_low:
                col_map[col] = "Preco_Sem_Desoneracao"
            elif "com" in c_low and "deson" in c_low:
                col_map[col] = "Preco_Com_Desoneracao"
            elif "prec" in c_low or "preç" in c_low or "unit" in c_low:
                col_map[col] = "Preco_Unitario"
        df = df.rename(columns=col_map)

    df["Codigo_SINAPI"] = (
        df["Codigo_SINAPI"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )
    df["Descricao_SINAPI"] = df["Descricao_SINAPI"].fillna("").astype(str).str.strip()
    df["Unidade"] = df["Unidade"].fillna("").astype(str).str.strip()
    df["Preco_Unitario"] = normalize_numeric_series(df["Preco_Unitario"])

    if "Preco_Sem_Desoneracao" in df.columns:
        df["Preco_Sem_Desoneracao"] = normalize_numeric_series(df["Preco_Sem_Desoneracao"])
    else:
        df["Preco_Sem_Desoneracao"] = df["Preco_Unitario"]

    if "Preco_Com_Desoneracao" in df.columns:
        df["Preco_Com_Desoneracao"] = normalize_numeric_series(df["Preco_Com_Desoneracao"])
    else:
        df["Preco_Com_Desoneracao"] = df["Preco_Unitario"]

    df = df[df["Codigo_SINAPI"] != ""].drop_duplicates(subset=["Codigo_SINAPI"], keep="first")
    return df.reset_index(drop=True)


# ==============================================================================
# 3. MOTOR DE REGRAS DE NEGÓCIO, ACRÉSCIMO PERCENTUAL E ROLL-UP WBS / EAP
# ==============================================================================
def process_budget_wbs(
    df_input: pd.DataFrame,
    df_sinapi: pd.DataFrame,
    acrescimo_percentual: float = 0.0,
    regime_preco: str = "Preco_Com_Desoneracao",
) -> Tuple[pd.DataFrame, float, List[str], List[str]]:
    """
    Realiza o LEFT JOIN com o SINAPI, aplica o acréscimo percentual (caso > 0%)
    diretamente sobre os valores unitários/totais e calcula o Roll-up WBS sem dupla contagem.
    """
    df = df_input.copy()
    df["Item_WBS"] = df["Item_WBS"].fillna("").astype(str).str.strip()
    df = df[df["Item_WBS"] != ""].reset_index(drop=True)

    if df.empty:
        return df, 0.0, [], []

    invalid_wbs = [
        wbs for wbs in df["Item_WBS"] if any(not p.isdigit() for p in wbs.split("."))
    ]

    df["Codigo_SINAPI"] = (
        df["Codigo_SINAPI"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .replace({"None": "", "nan": ""})
    )

    df["Quantidade"] = normalize_numeric_series(df["Quantidade"])
    df["Preco_Unitario_Manual"] = normalize_numeric_series(df["Preco_Unitario_Manual"])

    # Seleciona a coluna de preço conforme o regime (Com ou Sem Desoneração)
    col_preco_sinapi = regime_preco if regime_preco in df_sinapi.columns else "Preco_Unitario"
    sinapi_lookup = df_sinapi[["Codigo_SINAPI", "Descricao_SINAPI", "Unidade", col_preco_sinapi]].rename(
        columns={col_preco_sinapi: "Preco_SINAPI_Base"}
    )

    # 1. LEFT JOIN com o banco SINAPI-CE
    df = pd.merge(
        df,
        sinapi_lookup,
        on="Codigo_SINAPI",
        how="left",
        suffixes=("", "_SINAPI"),
    )

    # 2. Identificar Nós Pais vs. Nós Folha (comparação estrita com ponto delimitador)
    all_wbs = df["Item_WBS"].tolist()
    df["Is_Parent"] = [
        any(other.startswith(wbs + ".") for other in all_wbs if other != wbs)
        for wbs in all_wbs
    ]

    # 3. Fator multiplicador da porcentagem (Ex: 40% -> fator = 1.40 | 0% -> fator = 1.00)
    fator_acrescimo = 1.0 + (float(acrescimo_percentual) / 100.0) if acrescimo_percentual else 1.0

    desc_final = []
    unid_final = []
    preco_final = []
    origem_item = []
    missing_codes = []

    for _, row in df.iterrows():
        cod = row["Codigo_SINAPI"]
        user_desc = str(row["Descricao"]).strip() if pd.notna(row["Descricao"]) and str(row["Descricao"]).strip() not in ("None", "nan") else ""
        user_unid = str(row["Unidade"]).strip() if pd.notna(row["Unidade"]) and str(row["Unidade"]).strip() not in ("None", "nan") else ""

        if row["Is_Parent"]:
            desc_final.append(user_desc if user_desc != "" else f"GRUPO {row['Item_WBS']}")
            unid_final.append("")
            preco_final.append(0.0)
            origem_item.append("Grupo (EAP)")
        elif cod != "":
            if pd.notna(row["Preco_SINAPI_Base"]):
                desc_final.append(user_desc if user_desc != "" else str(row["Descricao_SINAPI"]))
                unid_final.append(user_unid if user_unid != "" else str(row["Unidade_SINAPI"]))
                preco_base = float(row["Preco_SINAPI_Base"])
                preco_final.append(round(preco_base * fator_acrescimo, 2))
                origem_item.append("SINAPI-CE")
            else:
                missing_codes.append(cod)
                desc_final.append(user_desc if user_desc != "" else f"[SINAPI {cod} NÃO ENCONTRADO]")
                unid_final.append(user_unid)
                preco_base = float(row["Preco_Unitario_Manual"])
                preco_final.append(round(preco_base * fator_acrescimo, 2))
                origem_item.append("Código Inválido")
        else:
            desc_final.append(user_desc)
            unid_final.append(user_unid)
            preco_base = float(row["Preco_Unitario_Manual"])
            preco_final.append(round(preco_base * fator_acrescimo, 2))
            origem_item.append("Manual / Cotado")

    df["Descricao_Final"] = desc_final
    df["Unidade_Final"] = unid_final
    df["Preco_Unitario_Final"] = preco_final
    df["Origem"] = origem_item
    df["Quantidade_Final"] = np.where(df["Is_Parent"], np.nan, df["Quantidade"])

    # 4. Cálculo do Valor Total dos Nós Folha (já com a porcentagem embutida se preenchida)
    df["Valor_Total"] = np.where(
        ~df["Is_Parent"],
        np.round(df["Quantidade"] * df["Preco_Unitario_Final"], 2),
        0.0,
    )

    # 5. Roll-up WBS/EAP: Soma exclusiva dos Nós Folha abaixo de cada Nó Pai
    leaf_df = df[~df["Is_Parent"]]
    for idx, row in df[df["Is_Parent"]].iterrows():
        parent_prefix = row["Item_WBS"] + "."
        descendant_leaves = leaf_df[leaf_df["Item_WBS"].str.startswith(parent_prefix)]
        df.at[idx, "Valor_Total"] = round(float(descendant_leaves["Valor_Total"].sum()), 2)

    # 6. Valor Global da Obra
    valor_global = round(float(leaf_df["Valor_Total"].sum()), 2)

    return df, valor_global, sorted(list(set(missing_codes))), invalid_wbs


# ==============================================================================
# 4. CONVERSOR DE VALOR MONETÁRIO POR EXTENSO (PT-BR) E MOTOR OPENPYXL
# ==============================================================================
def _centena_por_extenso(n: int) -> str:
    """Converte um inteiro entre 1 e 999 para extenso em português."""
    unidades = [
        "", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove",
        "dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis",
        "dezessete", "dezoito", "dezenove",
    ]
    dezenas = [
        "", "", "vinte", "trinta", "quarenta", "cinquenta",
        "sessenta", "setenta", "oitenta", "noventa",
    ]
    centenas = [
        "", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos",
        "seiscentos", "setecentos", "oitocentos", "novecentos",
    ]
    if n <= 0:
        return ""
    if n == 100:
        return "cem"

    c = n // 100
    d = (n % 100) // 10
    u = n % 10
    resto = n % 100

    partes = []
    if c > 0:
        partes.append(centenas[c])
    if resto > 0:
        if resto < 20:
            partes.append(unidades[resto])
        else:
            partes.append(dezenas[d])
            if u > 0:
                partes.append(unidades[u])
    return " e ".join(partes)


def _inteiro_por_extenso(n: int) -> str:
    """Converte um número inteiro não-negativo para extenso em português."""
    if n == 0:
        return "zero"

    escalas = [
        (1_000_000_000, "bilhão", "bilhões"),
        (1_000_000, "milhão", "milhões"),
        (1_000, "mil", "mil"),
        (1, "", ""),
    ]

    grupos = []
    resto = n
    for divisor, sing, plur in escalas:
        qtd = resto // divisor
        resto %= divisor
        if qtd > 0:
            if divisor == 1_000 and qtd == 1:
                texto_grupo = "mil"
            else:
                ext = _centena_por_extenso(qtd)
                sufixo = sing if qtd == 1 else plur
                texto_grupo = f"{ext} {sufixo}".strip()
            grupos.append((qtd, divisor, texto_grupo))

    if len(grupos) == 1:
        return grupos[0][2]

    partes_texto = [g[2] for g in grupos]
    ultimo_val = grupos[-1][0]
    ultimo_div = grupos[-1][1]
    # Regra gramatical PT-BR: usa "e" antes do último grupo se for < 100 ou múltiplo exato de 100
    usar_e = (ultimo_div == 1 and (ultimo_val < 100 or ultimo_val % 100 == 0)) or (ultimo_div > 1 and ultimo_val <= 100)

    if usar_e:
        return ", ".join(partes_texto[:-1]) + " e " + partes_texto[-1]
    return ", ".join(partes_texto)


def valor_por_extenso(valor: float) -> str:
    """
    Converte um valor float em Reais (R$) para texto por extenso em maiúsculas.
    Ex: 4981.32 -> 'QUATRO MIL, NOVECENTOS E OITENTA E UM REAIS E TRINTA E DOIS CENTAVOS'
    """
    total_centavos = int(round(abs(float(valor)) * 100))
    if total_centavos == 0:
        return "ZERO REAIS"

    inteiro, centavos = divmod(total_centavos, 100)
    partes = []

    if inteiro > 0:
        ext_int = _inteiro_por_extenso(inteiro)
        if inteiro == 1:
            moeda = "real"
        elif inteiro % 1_000_000 == 0:
            moeda = "de reais"
        else:
            moeda = "reais"
        partes.append(f"{ext_int} {moeda}")

    if centavos > 0:
        ext_cent = _centena_por_extenso(centavos)
        unid_cent = "centavo" if centavos == 1 else "centavos"
        partes.append(f"{ext_cent} {unid_cent}")

    return " e ".join(partes).upper()


def write_merged_safe(ws, coord: str, value):
    """Escreve sempre na célula superior-esquerda caso a coordenada seja mesclada."""
    cell = ws[coord]
    for m_range in ws.merged_cells.ranges:
        if cell.coordinate in m_range:
            ws.cell(m_range.min_row, m_range.min_col).value = value
            return
    cell.value = value


def generate_excel_budget(
    template_path: str,
    header_data: Dict[str, str],
    df_calc: pd.DataFrame,
    valor_global: float,
) -> io.BytesIO:
    """Injeta os dados no template Excel mantendo logomarca, bordas e rodapé."""
    wb = openpyxl.load_workbook(template_path)
    sheet_name = CELL_MAPPING["SHEET_NAME"]
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active

    proposta_str = (
        f"Nº {header_data['proposta_num']}"
        if not str(header_data["proposta_num"]).upper().startswith("Nº")
        else str(header_data["proposta_num"])
    )
    write_merged_safe(ws, CELL_MAPPING["PROPOSTA_NUM"], proposta_str)
    write_merged_safe(ws, CELL_MAPPING["PROPOSTA_DATA"], header_data["data_proposta"])
    write_merged_safe(ws, CELL_MAPPING["PROPOSTA_DATA_P2"], header_data["data_proposta"])
    write_merged_safe(ws, CELL_MAPPING["DESTINATARIO"], f"DESTINATÁRIO: {header_data['destinatario']}")
    write_merged_safe(ws, CELL_MAPPING["TELEFONE"], f"TELEFONE: {header_data['telefone']}")
    write_merged_safe(ws, CELL_MAPPING["CONTATO"], f"CONTATO: {header_data['contato']}")
    write_merged_safe(ws, CELL_MAPPING["EMAIL"], f"E-MAIL: {header_data['email']}")
    write_merged_safe(ws, CELL_MAPPING["SERVICO_TITULO"], f"SERVIÇO: {header_data['servico_titulo']}")

    texto_projeto = (
        f"\n\n1. OBJETO\n"
        f"1.1. {header_data['servico_titulo']}.\n\n\n"
        f"2. PRAZO DE ENTREGA, DO LOCAL DE ENTREGA E DO RECEBIMENTO.\n"
        f"2.1. O prazo de execução dos serviços {header_data['prazo_execucao']}."
    )
    write_merged_safe(ws, CELL_MAPPING["PROJETO_BASICO_TEXTO"], texto_projeto)

    start_row = CELL_MAPPING["START_ROW"]
    end_template_row = CELL_MAPPING["END_TEMPLATE_ROW"]
    max_footer_row = CELL_MAPPING["MAX_FOOTER_ROW"]
    available_rows = end_template_row - start_row + 1
    num_items = max(len(df_calc), 1)
    diff = num_items - available_rows

    footer_heights = {
        r: ws.row_dimensions[r].height
        for r in range(end_template_row + 1, max_footer_row + 1)
    }

    parent_style_snapshot = {}
    leaf_style_snapshot = {}
    default_item_height = ws.row_dimensions[start_row + 1].height or 18.0
    for col in range(1, CELL_MAPPING["COL_MARGIN_END"] + 1):
        c_parent = ws.cell(start_row, col)
        c_leaf = ws.cell(start_row + 1, col)
        parent_style_snapshot[col] = {
            "font": copy(c_parent.font),
            "border": copy(c_parent.border),
            "fill": copy(c_parent.fill),
            "number_format": c_parent.number_format,
            "alignment": copy(c_parent.alignment),
        }
        leaf_style_snapshot[col] = {
            "font": copy(c_leaf.font),
            "border": copy(c_leaf.border),
            "fill": copy(c_leaf.fill),
            "number_format": c_leaf.number_format,
            "alignment": copy(c_leaf.alignment),
        }

    footer_merges = []
    for m_range in list(ws.merged_cells.ranges):
        if m_range.max_row >= start_row:
            if m_range.min_row > end_template_row:
                footer_merges.append((m_range.min_row, m_range.min_col, m_range.max_row, m_range.max_col))
            ws.unmerge_cells(str(m_range))

    if diff > 0:
        ws.insert_rows(idx=end_template_row + 1, amount=diff)
    elif diff < 0:
        ws.delete_rows(idx=start_row + num_items, amount=abs(diff))

    parent_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    currency_format = '[$R$ -416]#,##0.00'

    for i in range(num_items):
        current_row = start_row + i
        ws.row_dimensions[current_row].height = default_item_height

        row_data = df_calc.iloc[i]
        is_parent = bool(row_data["Is_Parent"])
        style_source = parent_style_snapshot if is_parent else leaf_style_snapshot

        for col, st_props in style_source.items():
            cell = ws.cell(current_row, col)
            cell.font = copy(st_props["font"])
            cell.border = copy(st_props["border"])
            cell.fill = (
                copy(parent_fill)
                if is_parent and CELL_MAPPING["COL_WBS"] <= col <= CELL_MAPPING["COL_VALOR"]
                else copy(st_props["fill"])
            )
            cell.number_format = st_props["number_format"]
            cell.alignment = copy(st_props["alignment"])

        c_wbs = ws.cell(current_row, CELL_MAPPING["COL_WBS"], str(row_data["Item_WBS"]))
        c_desc = ws.cell(current_row, CELL_MAPPING["COL_DESC_START"], str(row_data["Descricao_Final"]))
        c_unid = ws.cell(current_row, CELL_MAPPING["COL_UNIDADE"], "" if is_parent else str(row_data["Unidade_Final"]))
        c_qtd = ws.cell(current_row, CELL_MAPPING["COL_QUANTIDADE"], None if is_parent else float(row_data["Quantidade_Final"]))
        c_val = ws.cell(current_row, CELL_MAPPING["COL_VALOR"], float(row_data["Valor_Total"]))

        c_wbs.font = Font(name=c_wbs.font.name or "Calibri", size=c_wbs.font.size or 10, bold=True)
        c_desc.font = Font(name=c_desc.font.name or "Calibri", size=c_desc.font.size or 10, bold=is_parent)
        c_unid.font = Font(name=c_unid.font.name or "Calibri", size=c_unid.font.size or 10, bold=is_parent)
        c_qtd.font = Font(name=c_qtd.font.name or "Calibri", size=c_qtd.font.size or 10, bold=is_parent)
        c_qtd.number_format = "#,##0.00"
        c_val.font = Font(name=c_val.font.name or "Calibri", size=c_val.font.size or 10, bold=is_parent)
        c_val.number_format = currency_format

        ws.merge_cells(
            start_row=current_row,
            start_column=CELL_MAPPING["COL_DESC_START"],
            end_row=current_row,
            end_column=CELL_MAPPING["COL_DESC_END"],
        )

    new_end_row = start_row + num_items - 1
    ws.merge_cells(
        start_row=start_row,
        start_column=CELL_MAPPING["COL_MARGIN_START"],
        end_row=new_end_row,
        end_column=CELL_MAPPING["COL_MARGIN_END"],
    )

    for min_r, min_c, max_r, max_c in footer_merges:
        ws.merge_cells(
            start_row=min_r + diff,
            start_column=min_c,
            end_row=max_r + diff,
            end_column=max_c,
        )

    for orig_r, h_val in footer_heights.items():
        if h_val is not None:
            ws.row_dimensions[orig_r + diff].height = h_val

    col_rodape = CELL_MAPPING["COL_RODAPE_VALORES"]
    row_global = CELL_MAPPING["ROW_VALOR_GLOBAL"] + diff

    # Valor Global numérico na célula D (mesclada D:E)
    cell_global = ws.cell(row_global, col_rodape)
    cell_global.value = float(valor_global)
    cell_global.number_format = currency_format

    # Valor Global por extenso na célula ao lado F (mesclada F:M)
    col_extenso = CELL_MAPPING["COL_VALOR_EXTENSO"]
    cell_extenso = ws.cell(row_global, col_extenso)
    cell_extenso.value = f"({valor_por_extenso(valor_global)})"
    cell_extenso.font = Font(
        name=cell_extenso.font.name or "Calibri",
        size=cell_extenso.font.size or 10,
        bold=True,
    )
    cell_extenso.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws.cell(CELL_MAPPING["ROW_PRAZO_EXECUCAO"] + diff, col_rodape, header_data["prazo_execucao"])
    ws.cell(CELL_MAPPING["ROW_IMPOSTOS"] + diff, col_rodape, header_data["impostos"])
    ws.cell(CELL_MAPPING["ROW_COND_COMERCIAIS"] + diff, col_rodape, header_data["cond_comerciais"])
    ws.cell(CELL_MAPPING["ROW_VALIDADE"] + diff, col_rodape, header_data["validade"])

    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer


# ==============================================================================
# 5. INTERFACE DO USUÁRIO NO STREAMLIT
# ==============================================================================
def main():
    st.set_page_config(
        page_title="Orçamentador SINAPI-CE | Mais Engenharia",
        page_icon="🏗️",
        layout="wide",
    )

    st.title("🏗️ Sistema de Orçamentação Dinâmica (SINAPI-CE)")
    st.caption("Mais Engenharia & Consultoria — Geração Automatizada de Propostas Comerciais em EAP/WBS")

    if not os.path.exists(SINAPI_CSV_PATH):
        st.error(f"Arquivo `{SINAPI_CSV_PATH}` não encontrado. Coloque o CSV gerado na mesma pasta do `app.py`.")
        st.stop()

    try:
        df_sinapi = load_sinapi_data(SINAPI_CSV_PATH)
    except Exception as e:
        st.error(f"Erro ao carregar `{SINAPI_CSV_PATH}`: {e}")
        st.stop()

    template_path = next((p for p in TEMPLATE_CANDIDATES if os.path.exists(p)), None)

    # --------------------------------------------------------------------------
    # BARRA LATERAL: CONSULTA RÁPIDA SINAPI-CE
    # --------------------------------------------------------------------------
    with st.sidebar:
        st.header("🔍 Consulta Rápida SINAPI-CE")
        termo_busca = st.text_input(
            "Pesquisar por código ou descrição:",
            placeholder="Ex: 103328, alvenaria, chapisco...",
        )

        if termo_busca:
            mask = (
                df_sinapi["Codigo_SINAPI"].str.contains(termo_busca, case=False, na=False)
                | df_sinapi["Descricao_SINAPI"].str.contains(termo_busca, case=False, na=False)
            )
            df_filtrado = df_sinapi[mask]
        else:
            df_filtrado = df_sinapi.head(25)

        st.dataframe(
            df_filtrado[["Codigo_SINAPI", "Descricao_SINAPI", "Unidade", "Preco_Com_Desoneracao", "Preco_Sem_Desoneracao"]],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Codigo_SINAPI": st.column_config.TextColumn("Código"),
                "Descricao_SINAPI": st.column_config.TextColumn("Descrição"),
                "Unidade": st.column_config.TextColumn("Unid."),
                "Preco_Com_Desoneracao": st.column_config.NumberColumn("Com Deson.", format="R$ %.2f"),
                "Preco_Sem_Desoneracao": st.column_config.NumberColumn("Sem Deson.", format="R$ %.2f"),
            },
        )
        st.caption(f"Total de composições ativas (Fortaleza-CE): **{len(df_sinapi)}**")

        st.divider()
        uploaded_template = st.file_uploader("Substituir Template (.xlsx)", type=["xlsx"])

    # --------------------------------------------------------------------------
    # SEÇÃO 1: DADOS DO CABEÇALHO, ACRÉSCIMO PERCENTUAL (%) E RODAPÉ
    # --------------------------------------------------------------------------
    with st.expander("📋 1. Dados da Proposta, Acréscimo Percentual (%) e Condições", expanded=True):
        col1, col2, col3, col4 = st.columns([1, 1, 2, 2])
        with col1:
            proposta_num = st.text_input("Nº da Proposta", value="26123")
        with col2:
            data_proposta = st.date_input("Data da Proposta", value=date.today())
        with col3:
            destinatario = st.text_input("Destinatário (Cliente)", value="Solar")
        with col4:
            servico_titulo = st.text_input(
                "Objeto / Serviço",
                value="Reforma ambiente corporativo Iguatemi Empresarial",
            )

        col5, col6, col7, col8, col9 = st.columns([1.5, 1.2, 1.8, 1.5, 1.5])
        with col5:
            contato = st.text_input("Contato", value="Jaqueline Azevedo")
        with col6:
            telefone = st.text_input("Telefone", value="(85) 9.9957-1480")
        with col7:
            email = st.text_input("E-mail", value="jaqueline.silva@solarbr.com.br")
        with col8:
            regime_label = st.selectbox(
                "Tabela SINAPI-CE",
                options=["Com Desoneração", "Sem Desoneração"],
                index=0,
            )
            regime_col = "Preco_Com_Desoneracao" if regime_label == "Com Desoneração" else "Preco_Sem_Desoneracao"
        with col9:
            acrescimo_percentual = st.number_input(
                "Acréscimo Percentual (%)",
                min_value=0.0,
                max_value=500.0,
                value=0.0,
                step=1.0,
                format="%.2f",
                help="Ex: Digite 40 para somar 40% a todos os valores (100 vira 140). Deixe 0 para imprimir os valores in natura.",
            )

        st.markdown("**Condições de Fornecimento (Rodapé)**")
        c_rod1, c_rod2, c_rod3, c_rod4 = st.columns(4)
        with c_rod1:
            prazo_execucao = st.text_input("Prazo de Execução", value="30 dias")
        with c_rod2:
            impostos = st.text_input("Impostos", value="INCLUSOS")
        with c_rod3:
            cond_comerciais = st.text_input("Cond. Comerciais", value="120 dias")
        with c_rod4:
            validade = st.text_input("Validade da Proposta", value="30 dias")

    # --------------------------------------------------------------------------
    # SEÇÃO 2: TABELA INTERATIVA DE ORÇAMENTO (WBS / EAP)
    # --------------------------------------------------------------------------
    st.subheader("📐 2. Estrutura Analítica do Orçamento (WBS / EAP)")

    if "budget_df" not in st.session_state:
        st.session_state.budget_df = pd.DataFrame(
            [
                {"Item_WBS": "1", "Codigo_SINAPI": "", "Descricao": "CUSTOS INDIRETOS", "Unidade": "", "Quantidade": 0.0, "Preco_Unitario_Manual": 0.0},
                {"Item_WBS": "1.1", "Codigo_SINAPI": "", "Descricao": "CUSTO COM EQUIPE DE APOIO E LIMPEZA", "Unidade": "MÊS", "Quantidade": 1.0, "Preco_Unitario_Manual": 3995.00},
                {"Item_WBS": "2", "Codigo_SINAPI": "", "Descricao": "CUSTOS DIRETOS", "Unidade": "", "Quantidade": 0.0, "Preco_Unitario_Manual": 0.0},
                {"Item_WBS": "2.1", "Codigo_SINAPI": "", "Descricao": "ALVENARIA E REVESTIMENTOS", "Unidade": "", "Quantidade": 0.0, "Preco_Unitario_Manual": 0.0},
                {"Item_WBS": "2.1.1", "Codigo_SINAPI": "103328", "Descricao": "", "Unidade": "", "Quantidade": 4.0, "Preco_Unitario_Manual": 0.0},
                {"Item_WBS": "2.1.2", "Codigo_SINAPI": "87879", "Descricao": "", "Unidade": "", "Quantidade": 4.0, "Preco_Unitario_Manual": 0.0},
            ]
        )

    edited_df = st.data_editor(
        st.session_state.budget_df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "Item_WBS": st.column_config.TextColumn("Item (WBS)", required=True, width="small"),
            "Codigo_SINAPI": st.column_config.TextColumn("Código SINAPI", width="small"),
            "Descricao": st.column_config.TextColumn("Descrição (Opcional p/ SINAPI)", width="large"),
            "Unidade": st.column_config.TextColumn("Medida", width="small"),
            "Quantidade": st.column_config.NumberColumn("Quantidade", min_value=0.0, step=0.01, format="%.2f"),
            "Preco_Unitario_Manual": st.column_config.NumberColumn("Preço Unit. Manual (R$)", min_value=0.0, step=0.01, format="R$ %.2f"),
        },
    )

    # --------------------------------------------------------------------------
    # SEÇÃO 3: PROCESSAMENTO, PRÉVIA E EXPORTAÇÃO EXCEL
    # --------------------------------------------------------------------------
    df_calc, valor_global, missing_codes, invalid_wbs = process_budget_wbs(
        edited_df,
        df_sinapi,
        acrescimo_percentual=acrescimo_percentual,
        regime_preco=regime_col,
    )

    if invalid_wbs:
        st.warning(f"⚠️ Códigos WBS fora do padrão numérico (ex: 1.1.2): {', '.join(invalid_wbs)}")

    if missing_codes:
        st.warning(f"⚠️ Códigos SINAPI não encontrados na base: **{', '.join(missing_codes)}**")

    st.divider()
    col_prev, col_metric = st.columns([3, 1])
    with col_prev:
        status_pct = f"(Com acréscimo de +{acrescimo_percentual:.2f}%)" if acrescimo_percentual > 0 else "(Valores In Natura)"
        st.subheader(f"📊 3. Prévia do Orçamento {status_pct}")
        st.caption(f"**Valor por extenso (Célula F do Rodapé):** ({valor_por_extenso(valor_global)})")
    with col_metric:
        valor_br = f"R$ {valor_global:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        st.metric(label="VALOR GLOBAL DA OBRA", value=valor_br)

    if not df_calc.empty:
        preview_df = df_calc[
            [
                "Item_WBS",
                "Origem",
                "Codigo_SINAPI",
                "Descricao_Final",
                "Unidade_Final",
                "Quantidade_Final",
                "Preco_Unitario_Final",
                "Valor_Total",
            ]
        ].copy()

        st.dataframe(
            preview_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Item_WBS": st.column_config.TextColumn("ITEM"),
                "Origem": st.column_config.TextColumn("TIPO"),
                "Codigo_SINAPI": st.column_config.TextColumn("SINAPI"),
                "Descricao_Final": st.column_config.TextColumn("SERVIÇO / DESCRIÇÃO", width="large"),
                "Unidade_Final": st.column_config.TextColumn("MEDIDA"),
                "Quantidade_Final": st.column_config.NumberColumn("QUANTIDADE", format="%.2f"),
                "Preco_Unitario_Final": st.column_config.NumberColumn("UNITÁRIO FINAL (R$)", format="R$ %.2f"),
                "Valor_Total": st.column_config.NumberColumn("VALOR TOTAL (R$)", format="R$ %.2f"),
            },
        )

        active_template = uploaded_template if uploaded_template is not None else template_path
        if active_template is not None:
            header_payload = {
                "proposta_num": proposta_num,
                "data_proposta": data_proposta.strftime("%d/%m/%Y") if isinstance(data_proposta, (date, datetime)) else str(data_proposta),
                "destinatario": destinatario,
                "telefone": telefone,
                "contato": contato,
                "email": email,
                "servico_titulo": servico_titulo,
                "prazo_execucao": prazo_execucao,
                "impostos": impostos,
                "cond_comerciais": cond_comerciais,
                "validade": validade,
            }

            try:
                excel_buffer = generate_excel_budget(
                    template_path=active_template,
                    header_data=header_payload,
                    df_calc=df_calc,
                    valor_global=valor_global,
                )
                nome_arquivo = f"Proposta_{proposta_num}_{destinatario.replace(' ', '_')}.xlsx"
                st.download_button(
                    label="📥 Baixar Proposta Oficial em Excel (.xlsx)",
                    data=excel_buffer,
                    file_name=nome_arquivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Erro ao gerar a planilha Excel: {e}")


if __name__ == "__main__":
    main()
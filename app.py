import io
import os
from copy import copy
from datetime import date, datetime
from typing import Dict, List, Tuple

import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill
import pandas as pd
import streamlit as st

# ==============================================================================
# 1. DICIONÁRIO DE MAPEAMENTO DO TEMPLATE EXCEL (CELL_MAPPING)
# Calibrado para o layout exato de "Relatorio Orcamento.xlsx" (Aba: "Orçamento")
# ==============================================================================
CELL_MAPPING = {
    "SHEET_NAME": "Orçamento",
    # Cabeçalho Página 1 (A Página 2 nas linhas 55-57 já possui fórmulas =I7, =B8, etc.)
    "PROPOSTA_NUM": "I7",          # Intervalo mesclado I7:K7
    "PROPOSTA_DATA": "L7",         # Intervalo mesclado L7:M7
    "PROPOSTA_DATA_P2": "L55",     # Intervalo mesclado L55:M55
    "DESTINATARIO": "B8",          # Intervalo mesclado B8:E8 (Prefixo "DESTINATÁRIO: ")
    "TELEFONE": "F8",              # Intervalo mesclado F8:M8 (Prefixo "TELEFONE: ")
    "CONTATO": "B9",               # Intervalo mesclado B9:E9 (Prefixo "CONTATO: ")
    "EMAIL": "F9",                 # Intervalo mesclado F9:M9 (Prefixo "E-MAIL: ")
    "PROJETO_BASICO_TEXTO": "B11", # Intervalo mesclado B11:M26 (Texto do Objeto/Prazo)
    "SERVICO_TITULO": "B59",       # Intervalo mesclado B59:F59 (Prefixo "SERVIÇO: ")
    # Limites originais da tabela de itens no template_base.xlsx
    "START_ROW": 60,               # Primeira linha de itens (logo abaixo do cabeçalho da linha 59)
    "END_TEMPLATE_ROW": 131,       # Última linha em branco antes do separador do rodapé (linha 132)
    # Colunas da tabela de itens (índices numéricos base 1: B=2, C=3, G=7, H=8, I=9)
    "COL_WBS": 2,                  # Coluna B: Índice WBS (ex: 1, 1.1, 2.1.1)
    "COL_DESC_START": 3,           # Coluna C: Início da Descrição (mesclada até Coluna F)
    "COL_DESC_END": 6,             # Coluna F: Fim da mesclagem da Descrição
    "COL_UNIDADE": 7,              # Coluna G: MEDIDA
    "COL_QUANTIDADE": 8,           # Coluna H: QUANTIDADE
    "COL_VALOR": 9,                # Coluna I: VALOR
    "COL_MARGIN_START": 10,        # Coluna J: Margem lateral direita mesclada (J:M)
    "COL_MARGIN_END": 13,          # Coluna M: Fim da margem lateral direita
    # Linhas originais do Rodapé (serão deslocadas dinamicamente por `offset`)
    "ROW_VALOR_GLOBAL": 133,       # Linha do "VALOR GLOBAL DA OBRA" (Valor numérico na Col D)
    "COL_RODAPE_VALORES": 4,       # Coluna D (D133, D137, D138, D139, D140)
    "ROW_PRAZO_EXECUCAO": 137,     # Linha "PRAZO DE EXECUÇÃO:"
    "ROW_IMPOSTOS": 138,           # Linha "IMPOSTOS:"
    "ROW_COND_COMERCIAIS": 139,    # Linha "COND. COMERCIAIS:"
    "ROW_VALIDADE": 140,           # Linha "VALIDADE:"
    "MAX_FOOTER_ROW": 141,         # Última linha formatada do rodapé
}

SINAPI_CSV_PATH = "sinapi_ce.csv"
TEMPLATE_CANDIDATES = ["template_base.xlsx", "Relatorio Orcamento.xlsx"]


# ==============================================================================
# 2. CARREGAMENTO E NORMALIZAÇÃO DA BASE SINAPI-CE
# ==============================================================================
def ensure_sample_sinapi_csv(filepath: str) -> bool:
    """
    Cria um arquivo sinapi_ce.csv de exemplo caso não exista no diretório,
    garantindo que a aplicação possa ser testada imediatamente.
    Retorna True se criou o arquivo de exemplo, False se o arquivo já existia.
    """
    if os.path.exists(filepath):
        return False

    sample_data = [
        {"Codigo_SINAPI": "87298", "Descricao_SINAPI": "ARGAMASSA TRAÇO 1:3 (EM VOLUME DE CIMENTO E AREIA MÉDIA ÚMIDA), PREPARO MANUAL. AF_08/2019", "Unidade": "M³", "Preco_Unitario": 615.40},
        {"Codigo_SINAPI": "87313", "Descricao_SINAPI": "ALVENARIA DE VEDAÇÃO DE BLOCOS CERÂMICOS FURADOS NA HORIZONTAL DE 9X19X19 CM (ESPESSURA 9 CM)", "Unidade": "M²", "Preco_Unitario": 85.20},
        {"Codigo_SINAPI": "87879", "Descricao_SINAPI": "CHAPISCO APLICADO EM ALVENARIAS E ESTRUTURAS DE CONCRETO INTERNAS, COM COLHER DE PEDREIRO. ARGAMASSA TRAÇO 1:3", "Unidade": "M²", "Preco_Unitario": 6.45},
        {"Codigo_SINAPI": "87529", "Descricao_SINAPI": "MASSA ÚNICA, PARA RECEBIMENTO DE PINTURA, EM ARGAMASSA TRAÇO 1:2:8, PREPARO MECÂNICO COM BETONEIRA", "Unidade": "M²", "Preco_Unitario": 34.10},
        {"Codigo_SINAPI": "88485", "Descricao_SINAPI": "APLICAÇÃO DE FUNDO SELADOR ACRÍLICO EM PAREDES, UMA DEMÃO. AF_04/2023", "Unidade": "M²", "Preco_Unitario": 3.85},
        {"Codigo_SINAPI": "88489", "Descricao_SINAPI": "APLICAÇÃO MANUAL DE PINTURA COM TINTA LÁTEX ACRÍLICA EM PAREDES, DUAS DEMÃOS. AF_04/2023", "Unidade": "M²", "Preco_Unitario": 16.90},
        {"Codigo_SINAPI": "96116", "Descricao_SINAPI": "FORRO EM PLACAS DE GESSO, PARA AMBIENTES COMERCIAIS. AF_05/2017_PS", "Unidade": "M²", "Preco_Unitario": 58.70},
        {"Codigo_SINAPI": "91854", "Descricao_SINAPI": "ELETRODUTO FLEXÍVEL CORRUGADO, PVC, DN 25 MM (3/4''), PARA CIRCUITOS TERMINAIS, INSTALADO EM PAREDE", "Unidade": "M", "Preco_Unitario": 11.25},
        {"Codigo_SINAPI": "91926", "Descricao_SINAPI": "CABO DE COBRE FLEXÍVEL ISOLADO, 2,5 MM², ANTI-CHAMA 450/750V, PARA CIRCUITOS TERMINAIS - FORNECIMENTO E INSTALAÇÃO", "Unidade": "M", "Preco_Unitario": 6.80},
        {"Codigo_SINAPI": "92000", "Descricao_SINAPI": "TOMADA MÉDIA DE EMBUTIR (1 MÓDULO), 2P+T 10 A, INCLUINDO SUPORTE E PLACA - FORNECIMENTO E INSTALAÇÃO", "Unidade": "UND", "Preco_Unitario": 31.50},
        {"Codigo_SINAPI": "9537", "Descricao_SINAPI": "LIMPEZA FINAL DA OBRA", "Unidade": "M²", "Preco_Unitario": 6.80},
    ]
    pd.DataFrame(sample_data).to_csv(filepath, index=False, encoding="utf-8-sig")
    return True


def normalize_numeric_series(series: pd.Series) -> pd.Series:
    """Converte strings numéricas no padrão BR (1.234,56) ou US (1234.56) para float."""
    def parse_val(val):
        if pd.isna(val) or str(val).strip() == "":
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
    """
    Carrega e normaliza o arquivo sinapi_ce.csv, tratando Codigo_SINAPI como string
    limpa (sem espaços ou sufixo .0) e Preco_Unitario como float.
    """
    try:
        df = pd.read_csv(filepath, dtype=str, sep=None, engine="python")
    except Exception:
        df = pd.read_csv(filepath, dtype=str, encoding="latin1", sep=None, engine="python")

    # Padronizar nomes das colunas caso haja variações de acentuação
    col_map = {}
    for col in df.columns:
        c_clean = col.strip().lower()
        if "cod" in c_clean or "cód" in c_clean:
            col_map[col] = "Codigo_SINAPI"
        elif "desc" in c_clean:
            col_map[col] = "Descricao_SINAPI"
        elif "unid" in c_clean or "medida" in c_clean:
            col_map[col] = "Unidade"
        elif "prec" in c_clean or "preç" in c_clean or "valor" in c_clean or "custo" in c_clean:
            col_map[col] = "Preco_Unitario"
    df = df.rename(columns=col_map)

    required_cols = {"Codigo_SINAPI", "Descricao_SINAPI", "Unidade", "Preco_Unitario"}
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"O arquivo CSV não possui as colunas obrigatórias: {missing}")

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

    # Remover duplicidades de código mantendo a primeira ocorrência
    df = df[df["Codigo_SINAPI"] != ""].drop_duplicates(subset=["Codigo_SINAPI"], keep="first")
    return df.reset_index(drop=True)


# ==============================================================================
# 3. MOTOR DE REGRAS DE NEGÓCIO E ROLL-UP HIERÁRQUICO WBS / EAP
# ==============================================================================
def process_budget_wbs(
    df_input: pd.DataFrame, df_sinapi: pd.DataFrame
) -> Tuple[pd.DataFrame, float, List[str], List[str]]:
    """
    Executa o cruzamento LEFT JOIN com o SINAPI, identifica Nós Pais vs. Nós Folha
    por prefixo estrito (X + ".") e realiza o Roll-up matemático sem dupla contagem.
    """
    df = df_input.copy()

    # Limpar linhas completamente vazias ou sem Item_WBS
    df["Item_WBS"] = df["Item_WBS"].fillna("").astype(str).str.strip()
    df = df[df["Item_WBS"] != ""].reset_index(drop=True)

    if df.empty:
        return df, 0.0, [], []

    # Validar formato dos códigos WBS
    invalid_wbs = []
    for wbs in df["Item_WBS"]:
        parts = wbs.split(".")
        if any(not p.isdigit() for p in parts):
            invalid_wbs.append(wbs)

    # Normalizar Codigo_SINAPI digitado pelo usuário
    df["Codigo_SINAPI"] = (
        df["Codigo_SINAPI"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .replace("None", "")
        .replace("nan", "")
    )

    # Normalizar Quantidade e Preco_Unitario_Manual
    df["Quantidade"] = normalize_numeric_series(df["Quantidade"])
    df["Preco_Unitario_Manual"] = normalize_numeric_series(df["Preco_Unitario_Manual"])

    # 1. LEFT JOIN com o banco SINAPI-CE
    df = pd.merge(
        df,
        df_sinapi[["Codigo_SINAPI", "Descricao_SINAPI", "Unidade", "Preco_Unitario"]],
        on="Codigo_SINAPI",
        how="left",
        suffixes=("", "_SINAPI"),
    )

    # 2. Identificar Nós Pais vs. Nós Folha usando comparação estrita com ponto delimitador
    # Isso garante que "1.1" NÃO seja considerado pai de "1.10" (pois "1.10" não começa com "1.1.")
    all_wbs = df["Item_WBS"].tolist()
    is_parent_flags = []
    for wbs in all_wbs:
        strict_prefix = wbs + "."
        has_children = any(other.startswith(strict_prefix) for other in all_wbs if other != wbs)
        is_parent_flags.append(has_children)
    df["Is_Parent"] = is_parent_flags

    # 3. Resolver Descrição, Unidade e Preço Unitário de cada linha
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
            # Nó Pai (Título de Grupo ou Subgrupo)
            desc_final.append(user_desc if user_desc != "" else f"GRUPO {row['Item_WBS']}")
            unid_final.append("")
            preco_final.append(0.0)
            origem_item.append("Grupo (EAP)")
        elif cod != "":
            # Nó Folha com Código SINAPI informado
            if pd.notna(row["Preco_Unitario"]):
                desc_final.append(user_desc if user_desc != "" else str(row["Descricao_SINAPI"]))
                unid_final.append(user_unid if user_unid != "" else str(row["Unidade_SINAPI"]))
                preco_final.append(float(row["Preco_Unitario"]))
                origem_item.append("SINAPI-CE")
            else:
                missing_codes.append(cod)
                desc_final.append(user_desc if user_desc != "" else f"[SINAPI {cod} NÃO ENCONTRADO]")
                unid_final.append(user_unid)
                preco_final.append(float(row["Preco_Unitario_Manual"]))
                origem_item.append("Código Inválido")
        else:
            # Nó Folha Manual / Próprio (sem Código SINAPI)
            desc_final.append(user_desc)
            unid_final.append(user_unid)
            preco_final.append(float(row["Preco_Unitario_Manual"]))
            origem_item.append("Manual / Cotado")

    df["Descricao_Final"] = desc_final
    df["Unidade_Final"] = unid_final
    df["Preco_Unitario_Final"] = preco_final
    df["Origem"] = origem_item

    # Zerar quantidade em Nós Pais para manter consistência visual
    df["Quantidade_Final"] = np.where(df["Is_Parent"], np.nan, df["Quantidade"])

    # 4. Cálculo do Valor Total dos Nós Folha
    df["Valor_Total"] = np.where(
        ~df["Is_Parent"],
        np.round(df["Quantidade"] * df["Preco_Unitario_Final"], 2),
        0.0,
    )

    # 5. Roll-up WBS/EAP:
    # O valor de qualquer Nó Pai é a soma EXCLUSIVA dos Nós Folha cujo Item_WBS
    # começa com (Item_WBS_Pai + "."). Isso evita 100% a dupla contagem entre subgrupos.
    leaf_df = df[~df["Is_Parent"]]
    for idx, row in df[df["Is_Parent"]].iterrows():
        parent_prefix = row["Item_WBS"] + "."
        descendant_leaves = leaf_df[leaf_df["Item_WBS"].str.startswith(parent_prefix)]
        df.at[idx, "Valor_Total"] = round(float(descendant_leaves["Valor_Total"].sum()), 2)

    # 6. Valor Global da Obra = Soma exclusiva de todos os Nós Folha
    valor_global = round(float(leaf_df["Valor_Total"].sum()), 2)

    return df, valor_global, sorted(list(set(missing_codes))), invalid_wbs


# ==============================================================================
# 4. MOTOR OPENPYXL: INJEÇÃO DINÂMICA E PRESERVAÇÃO DE MERGED CELLS
# ==============================================================================
def write_merged_safe(ws, coord: str, value):
    """
    Escreve com segurança em uma coordenada do Excel, localizando a célula
    superior-esquerda caso a coordenada pertença a um intervalo mesclado.
    """
    cell = ws[coord]
    for m_range in ws.merged_cells.ranges:
        if cell.coordinate in m_range:
            ws.cell(m_range.min_row, m_range.min_col).value = value
            return
    cell.value = value


def copy_row_style(ws, source_row: int, target_row: int, max_col: int = 13):
    """Copia fontes, bordas, preenchimentos, máscaras numéricas, alinhamentos e altura."""
    if ws.row_dimensions[source_row].height is not None:
        ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height

    for col in range(1, max_col + 1):
        src = ws.cell(source_row, col)
        dst = ws.cell(target_row, col)
        if src.has_style:
            dst.font = copy(src.font)
            dst.border = copy(src.border)
            dst.fill = copy(src.fill)
            dst.number_format = src.number_format
            dst.alignment = copy(src.alignment)


def generate_excel_budget(
    template_path: str,
    header_data: Dict[str, str],
    df_calc: pd.DataFrame,
    valor_global: float,
) -> io.BytesIO:
    """
    Abre o template Excel, injeta cabeçalho, ajusta dinamicamente a quantidade de linhas
    preservando células mescladas e alturas do rodapé, e retorna o arquivo em memória (BytesIO).
    """
    wb = openpyxl.load_workbook(template_path)
    sheet_name = CELL_MAPPING["SHEET_NAME"]
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active

    # 1. Injetar Dados do Cabeçalho Superior
    proposta_str = f"Nº {header_data['proposta_num']}" if not str(header_data["proposta_num"]).upper().startswith("Nº") else str(header_data["proposta_num"])
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

    # 2. Preparar Ajuste Dinâmico de Linhas da Tabela e Preservação de Mesclagens
    start_row = CELL_MAPPING["START_ROW"]
    end_template_row = CELL_MAPPING["END_TEMPLATE_ROW"]
    max_footer_row = CELL_MAPPING["MAX_FOOTER_ROW"]
    available_rows = end_template_row - start_row + 1
    num_items = max(len(df_calc), 1)
    diff = num_items - available_rows  # > 0: inserir linhas | < 0: remover linhas excedentes

    # Salvar alturas originais do rodapé antes de inserir/deletar linhas
    footer_heights = {
        r: ws.row_dimensions[r].height
        for r in range(end_template_row + 1, max_footer_row + 1)
    }

    # Capturar o estilo de uma linha Pai (linha 60) e de uma linha Folha (linha 61)
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

    # Desfazer temporariamente todas as células mescladas a partir de start_row
    # (Isso evita que o Openpyxl corrompa as mesclagens do rodapé durante insert_rows/delete_rows)
    footer_merges = []
    for m_range in list(ws.merged_cells.ranges):
        if m_range.max_row >= start_row:
            if m_range.min_row > end_template_row:
                footer_merges.append((m_range.min_row, m_range.min_col, m_range.max_row, m_range.max_col))
            ws.unmerge_cells(str(m_range))

    # Executar inserção ou exclusão física de linhas
    if diff > 0:
        ws.insert_rows(idx=end_template_row + 1, amount=diff)
    elif diff < 0:
        ws.delete_rows(idx=start_row + num_items, amount=abs(diff))

    # 3. Preencher os Itens do Orçamento e Aplicar Estilos (Pai vs. Folha)
    parent_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    currency_format = '[$R$ -416]#,##0.00'

    for i in range(num_items):
        current_row = start_row + i
        ws.row_dimensions[current_row].height = default_item_height

        row_data = df_calc.iloc[i]
        is_parent = bool(row_data["Is_Parent"])
        style_source = parent_style_snapshot if is_parent else leaf_style_snapshot

        # Aplicar estilos em todas as colunas da linha
        for col, st_props in style_source.items():
            cell = ws.cell(current_row, col)
            cell.font = copy(st_props["font"])
            cell.border = copy(st_props["border"])
            cell.fill = copy(parent_fill) if is_parent and col <= CELL_MAPPING["COL_VALOR"] and col >= CELL_MAPPING["COL_WBS"] else copy(st_props["fill"])
            cell.number_format = st_props["number_format"]
            cell.alignment = copy(st_props["alignment"])

        # Escrever valores nas colunas mapeadas
        c_wbs = ws.cell(current_row, CELL_MAPPING["COL_WBS"], str(row_data["Item_WBS"]))
        c_desc = ws.cell(current_row, CELL_MAPPING["COL_DESC_START"], str(row_data["Descricao_Final"]))
        c_unid = ws.cell(current_row, CELL_MAPPING["COL_UNIDADE"], "" if is_parent else str(row_data["Unidade_Final"]))
        c_qtd = ws.cell(current_row, CELL_MAPPING["COL_QUANTIDADE"], None if is_parent else float(row_data["Quantidade_Final"]))
        c_val = ws.cell(current_row, CELL_MAPPING["COL_VALOR"], float(row_data["Valor_Total"]))

        # Ajustar negrito e formatos numéricos
        c_wbs.font = Font(name=c_wbs.font.name or "Calibri", size=c_wbs.font.size or 10, bold=True)
        c_desc.font = Font(name=c_desc.font.name or "Calibri", size=c_desc.font.size or 10, bold=is_parent)
        c_unid.font = Font(name=c_unid.font.name or "Calibri", size=c_unid.font.size or 10, bold=is_parent)
        c_qtd.font = Font(name=c_qtd.font.name or "Calibri", size=c_qtd.font.size or 10, bold=is_parent)
        c_qtd.number_format = "#,##0.00"
        c_val.font = Font(name=c_val.font.name or "Calibri", size=c_val.font.size or 10, bold=is_parent)
        c_val.number_format = currency_format

        # Recriar mesclagem horizontal da Descrição (Colunas C até F)
        ws.merge_cells(
            start_row=current_row,
            start_column=CELL_MAPPING["COL_DESC_START"],
            end_row=current_row,
            end_column=CELL_MAPPING["COL_DESC_END"],
        )

    # Recriar mesclagem da margem lateral direita (Colunas J até M ao longo de toda a tabela)
    new_end_row = start_row + num_items - 1
    ws.merge_cells(
        start_row=start_row,
        start_column=CELL_MAPPING["COL_MARGIN_START"],
        end_row=new_end_row,
        end_column=CELL_MAPPING["COL_MARGIN_END"],
    )

    # 4. Recriar Mesclagens e Alturas do Rodapé com o Deslocamento (`diff`)
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

    # 5. Injetar Valor Global da Obra e Condições Comerciais no Rodapé Deslocado
    col_rodape = CELL_MAPPING["COL_RODAPE_VALORES"]

    cell_global = ws.cell(CELL_MAPPING["ROW_VALOR_GLOBAL"] + diff, col_rodape)
    cell_global.value = float(valor_global)
    cell_global.number_format = currency_format

    ws.cell(CELL_MAPPING["ROW_PRAZO_EXECUCAO"] + diff, col_rodape, header_data["prazo_execucao"])
    ws.cell(CELL_MAPPING["ROW_IMPOSTOS"] + diff, col_rodape, header_data["impostos"])
    ws.cell(CELL_MAPPING["ROW_COND_COMERCIAIS"] + diff, col_rodape, header_data["cond_comerciais"])
    ws.cell(CELL_MAPPING["ROW_VALIDADE"] + diff, col_rodape, header_data["validade"])

    # 6. Exportar para buffer em memória (BytesIO) - Zero escrita em disco
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

    # Verificar e carregar base SINAPI
    created_sample = ensure_sample_sinapi_csv(SINAPI_CSV_PATH)
    if created_sample:
        st.info(
            "ℹ️ Arquivo `sinapi_ce.csv` não encontrado no diretório. Um banco de dados SINAPI-CE de demonstração "
            "foi gerado automaticamente para teste imediato. Você pode substituí-lo pela tabela completa da Caixa."
        )

    try:
        df_sinapi = load_sinapi_data(SINAPI_CSV_PATH)
    except Exception as e:
        st.error(f"Erro ao carregar `{SINAPI_CSV_PATH}`: {e}")
        st.stop()

    # Localizar arquivo de template Excel
    template_path = next((p for p in TEMPLATE_CANDIDATES if os.path.exists(p)), None)
    if not template_path:
        st.warning(
            "⚠️ Arquivo `template_base.xlsx` (ou `Relatorio Orcamento.xlsx`) não encontrado na pasta local. "
            "Faça o upload do template na barra lateral ou adicione-o ao repositório."
        )

    # --------------------------------------------------------------------------
    # BARRA LATERAL: CONSULTA RÁPIDA SINAPI-CE E UPLOAD OPCIONAL DE ARQUIVOS
    # --------------------------------------------------------------------------
    with st.sidebar:
        st.header("🔍 Consulta Rápida SINAPI-CE")
        termo_busca = st.text_input(
            "Pesquisar por código ou descrição:",
            placeholder="Ex: 87298, alvenaria, pintura, cabo...",
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
            df_filtrado,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Codigo_SINAPI": st.column_config.TextColumn("Código"),
                "Descricao_SINAPI": st.column_config.TextColumn("Descrição"),
                "Unidade": st.column_config.TextColumn("Unid."),
                "Preco_Unitario": st.column_config.NumberColumn("Preço (R$)", format="R$ %.2f"),
            },
        )
        st.caption(f"Total de itens na base SINAPI-CE: **{len(df_sinapi)}**")

        st.divider()
        st.subheader("📂 Arquivos Base (Opcional)")
        uploaded_template = st.file_uploader("Substituir Template (.xlsx)", type=["xlsx"])
        uploaded_sinapi = st.file_uploader("Atualizar Base SINAPI (.csv)", type=["csv"])

        if uploaded_sinapi is not None:
            with open(SINAPI_CSV_PATH, "wb") as f:
                f.write(uploaded_sinapi.getbuffer())
            st.cache_data.clear()
            st.rerun()

    # --------------------------------------------------------------------------
    # SEÇÃO 1: DADOS DO CABEÇALHO E RODAPÉ DA PROPOSTA
    # --------------------------------------------------------------------------
    with st.expander("📋 1. Dados da Proposta Comercial e Condições de Fornecimento", expanded=True):
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

        col5, col6, col7 = st.columns(3)
        with col5:
            contato = st.text_input("Contato", value="Jaqueline Azevedo")
        with col6:
            telefone = st.text_input("Telefone", value="(85) 9.9957-1480")
        with col7:
            email = st.text_input("E-mail", value="jaqueline.silva@solarbr.com.br")

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
    st.markdown(
        "💡 **Como preencher:** "
        "**(1) Títulos de Grupo/Subgrupo:** preencha apenas `Item_WBS` (ex: `1`, `2.1`) e `Descricao` (o sistema somará os filhos automaticamente). "
        "**(2) Itens SINAPI:** informe `Item_WBS`, `Codigo_SINAPI` e `Quantidade`. "
        "**(3) Itens Próprios/Cotados:** deixe `Codigo_SINAPI` em branco e informe `Descricao`, `Unidade`, `Quantidade` e `Preco_Unitario_Manual`."
    )

    if "budget_df" not in st.session_state:
        st.session_state.budget_df = pd.DataFrame(
            [
                {"Item_WBS": "1", "Codigo_SINAPI": "", "Descricao": "CUSTOS INDIRETOS", "Unidade": "", "Quantidade": 0.0, "Preco_Unitario_Manual": 0.0},
                {"Item_WBS": "1.1", "Codigo_SINAPI": "", "Descricao": "CUSTO COM EQUIPE DE APOIO E LIMPEZA", "Unidade": "MÊS", "Quantidade": 1.0, "Preco_Unitario_Manual": 3995.00},
                {"Item_WBS": "2", "Codigo_SINAPI": "", "Descricao": "SERVIÇOS DE ENGENHARIA CIVIL", "Unidade": "", "Quantidade": 0.0, "Preco_Unitario_Manual": 0.0},
                {"Item_WBS": "2.1", "Codigo_SINAPI": "", "Descricao": "ALVENARIA E REVESTIMENTOS", "Unidade": "", "Quantidade": 0.0, "Preco_Unitario_Manual": 0.0},
                {"Item_WBS": "2.1.1", "Codigo_SINAPI": "87313", "Descricao": "", "Unidade": "", "Quantidade": 27.98, "Preco_Unitario_Manual": 0.0},
                {"Item_WBS": "2.1.2", "Codigo_SINAPI": "87529", "Descricao": "", "Unidade": "", "Quantidade": 27.98, "Preco_Unitario_Manual": 0.0},
                {"Item_WBS": "2.2", "Codigo_SINAPI": "", "Descricao": "PINTURA", "Unidade": "", "Quantidade": 0.0, "Preco_Unitario_Manual": 0.0},
                {"Item_WBS": "2.2.1", "Codigo_SINAPI": "88489", "Descricao": "", "Unidade": "", "Quantidade": 55.96, "Preco_Unitario_Manual": 0.0},
                {"Item_WBS": "3", "Codigo_SINAPI": "", "Descricao": "DIVERSOS", "Unidade": "", "Quantidade": 0.0, "Preco_Unitario_Manual": 0.0},
                {"Item_WBS": "3.1", "Codigo_SINAPI": "9537", "Descricao": "", "Unidade": "", "Quantidade": 19.43, "Preco_Unitario_Manual": 0.0},
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
    df_calc, valor_global, missing_codes, invalid_wbs = process_budget_wbs(edited_df, df_sinapi)

    if invalid_wbs:
        st.warning(f"⚠️ Atenção: Os seguintes códigos WBS estão fora do padrão numérico separado por ponto (ex: 1.1.2): {', '.join(invalid_wbs)}")

    if missing_codes:
        st.warning(f"⚠️ Os seguintes Códigos SINAPI não foram encontrados em `{SINAPI_CSV_PATH}`: **{', '.join(missing_codes)}**")

    st.divider()
    col_prev, col_metric = st.columns([3, 1])
    with col_prev:
        st.subheader("📊 3. Prévia Calculada do Orçamento (Roll-up WBS)")
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
                "Preco_Unitario_Final": st.column_config.NumberColumn("UNITÁRIO (R$)", format="R$ %.2f"),
                "Valor_Total": st.column_config.NumberColumn("VALOR TOTAL (R$)", format="R$ %.2f"),
            },
        )

        # Preparar exportação em memória
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
                st.error(f"Erro ao gerar a planilha Excel a partir do template: {e}")


if __name__ == "__main__":
    main()
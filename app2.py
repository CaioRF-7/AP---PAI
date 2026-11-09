import io
import math
import os
import re
from copy import copy
from datetime import date, datetime
from decimal import Decimal
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import pandas as pd
from rapidfuzz import fuzz, process as rf_process
import streamlit as st

from converter_orcamento import (
    auto_detect_sheet_and_columns,
    export_clean_budget_xlsx,
    extract_budget_from_sheet,
    load_raw_sheets_from_file,
    parse_br_number,
    parse_seinfra_file,
    quantize_money,
    strip_accents,
)

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
SEINFRA_CSV_PATH = "seinfra_ce.csv"
TEMPLATE_CANDIDATES = ["template_base.xlsx", "Relatorio Orcamento.xlsx"]

MODO_PCT_TODOS = "Todos os itens"
MODO_PCT_MARCADOS = "Apenas itens marcados na tabela (coluna '%?')"
MODO_PCT_BASES = "Apenas itens SINAPI / SEINFRA"
MODO_PCT_MANUAIS = "Apenas itens sem código (Importados / Manuais)"


# ==============================================================================
# 2. CARREGAMENTO E UNIFICAÇÃO DAS BASES SINAPI-CE E SEINFRA-CE + BUSCA GOOGLE
# ==============================================================================
def normalize_numeric_series(series: pd.Series) -> pd.Series:
    """Converte strings numéricas no padrão BR (1.234,56) ou US (1234.56) para float."""
    return series.apply(lambda v: float(p) if (p := parse_br_number(v)) is not None else 0.0)


def normalize_code_str(val: Any) -> str:
    """Padroniza códigos SINAPI ('103328') ou SEINFRA ('C2820', 'I0001') em caixa alta."""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return ""
    s = str(val).strip().upper()
    if s in ("NONE", "NAN", "-"):
        return ""
    s = re.sub(r"\.0$", "", s)
    return s


def _normalize_search_text(text: str) -> str:
    """Normaliza texto sem acentos e unifica dimensões ('9 x 19' -> '9x19') para busca estilo Google."""
    s = strip_accents(text).lower()
    s = re.sub(r"(\d)\s*[x×]\s*(\d)", r"\1x\2", s)
    s = re.sub(r"[^a-z0-9x.,/%]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _read_csv_flexible(filepath: str) -> pd.DataFrame:
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
    df.columns = [str(c).replace("\ufeff", "").strip() for c in df.columns]
    return df


@st.cache_data(show_spinner="Carregando bases oficiais SINAPI-CE e SEINFRA-CE...")
def load_unified_catalog(sinapi_path: str, seinfra_path: str) -> pd.DataFrame:
    """
    Carrega e unifica SINAPI-CE e SEINFRA-CE em um único catálogo indexável com colunas:
    [Fonte, Codigo, Descricao, Unidade, Preco_Unitario, Preco_Sem_Desoneracao, Preco_Com_Desoneracao, Search_Norm]
    """
    frames: List[pd.DataFrame] = []

    if os.path.exists(sinapi_path):
        df_sin = _read_csv_flexible(sinapi_path)
        col_map = {}
        for col in df_sin.columns:
            c_low = strip_accents(col).lower()
            if "cod" in c_low:
                col_map[col] = "Codigo"
            elif "desc" in c_low:
                col_map[col] = "Descricao"
            elif "unid" in c_low or "medida" in c_low:
                col_map[col] = "Unidade"
            elif "sem" in c_low and "deson" in c_low:
                col_map[col] = "Preco_Sem_Desoneracao"
            elif "com" in c_low and "deson" in c_low:
                col_map[col] = "Preco_Com_Desoneracao"
            elif "prec" in c_low or "unit" in c_low:
                col_map[col] = "Preco_Unitario"
        df_sin = df_sin.rename(columns=col_map)
        df_sin["Fonte"] = "SINAPI"
        frames.append(df_sin)

    if os.path.exists(seinfra_path):
        df_sei = _read_csv_flexible(seinfra_path)
        col_map = {}
        for col in df_sei.columns:
            c_low = strip_accents(col).lower()
            if "cod" in c_low:
                col_map[col] = "Codigo"
            elif "desc" in c_low:
                col_map[col] = "Descricao"
            elif "unid" in c_low or "medida" in c_low:
                col_map[col] = "Unidade"
            elif "sem" in c_low and "deson" in c_low:
                col_map[col] = "Preco_Sem_Desoneracao"
            elif "com" in c_low and "deson" in c_low:
                col_map[col] = "Preco_Com_Desoneracao"
            elif "prec" in c_low or "unit" in c_low:
                col_map[col] = "Preco_Unitario"
        df_sei = df_sei.rename(columns=col_map)
        df_sei["Fonte"] = "SEINFRA"
        frames.append(df_sei)

    if not frames:
        return pd.DataFrame(
            columns=[
                "Fonte",
                "Codigo",
                "Descricao",
                "Unidade",
                "Preco_Unitario",
                "Preco_Sem_Desoneracao",
                "Preco_Com_Desoneracao",
                "Search_Norm",
            ]
        )

    df_all = pd.concat(frames, ignore_index=True)
    df_all["Codigo"] = df_all["Codigo"].apply(normalize_code_str)
    df_all["Descricao"] = df_all["Descricao"].fillna("").astype(str).str.strip()
    df_all["Unidade"] = df_all["Unidade"].fillna("").astype(str).str.strip()
    df_all["Preco_Unitario"] = normalize_numeric_series(df_all["Preco_Unitario"])

    if "Preco_Sem_Desoneracao" in df_all.columns:
        df_all["Preco_Sem_Desoneracao"] = normalize_numeric_series(df_all["Preco_Sem_Desoneracao"])
        mask_zero_sem = df_all["Preco_Sem_Desoneracao"] == 0.0
        df_all.loc[mask_zero_sem, "Preco_Sem_Desoneracao"] = df_all.loc[mask_zero_sem, "Preco_Unitario"]
    else:
        df_all["Preco_Sem_Desoneracao"] = df_all["Preco_Unitario"]

    if "Preco_Com_Desoneracao" in df_all.columns:
        df_all["Preco_Com_Desoneracao"] = normalize_numeric_series(df_all["Preco_Com_Desoneracao"])
        mask_zero_com = df_all["Preco_Com_Desoneracao"] == 0.0
        df_all.loc[mask_zero_com, "Preco_Com_Desoneracao"] = df_all.loc[mask_zero_com, "Preco_Unitario"]
    else:
        df_all["Preco_Com_Desoneracao"] = df_all["Preco_Unitario"]

    df_all = df_all[df_all["Codigo"] != ""].drop_duplicates(subset=["Fonte", "Codigo"], keep="first").reset_index(drop=True)
    df_all["Search_Norm"] = [
        _normalize_search_text(f"{f} {c} {d} {u}")
        for f, c, d, u in zip(df_all["Fonte"], df_all["Codigo"], df_all["Descricao"], df_all["Unidade"])
    ]
    return df_all


def build_catalog_lookup(df_catalog: pd.DataFrame) -> Dict[str, Dict[str, Any]]:
    """Cria índice O(1) por código (aceitando tanto '103328' / 'C2820' quanto 'SINAPI:103328')."""
    lookup: Dict[str, Dict[str, Any]] = {}
    for rec in df_catalog.to_dict("records"):
        cod = rec["Codigo"]
        fonte = rec["Fonte"]
        lookup[f"{fonte}:{cod}"] = rec
        if cod not in lookup:
            lookup[cod] = rec
    return lookup


def search_catalog_google_style(
    query: str,
    df_catalog: pd.DataFrame,
    fonte_filter: str = "Todas",
    limit: int = 40,
) -> pd.DataFrame:
    """
    Busca inteligente estilo Google sobre SINAPI + SEINFRA:
    - Ignora acentos e maiúsculas/minúsculas ('ceramico' acha 'CERÂMICO')
    - Aceita palavras fora de ordem ('alvenaria 9x19', 'forro gesso', 'piso vinilico')
    - Prioriza código exato, início da descrição e correspondência de todos os termos
    - Usa tolerância a erros de digitação (fuzzy matching via rapidfuzz) se necessário
    """
    if df_catalog.empty:
        return df_catalog

    df_sub = df_catalog
    if fonte_filter in ("SINAPI", "SEINFRA"):
        df_sub = df_catalog[df_catalog["Fonte"] == fonte_filter]

    q_norm = _normalize_search_text(query)
    if not q_norm:
        return df_sub.head(limit)

    tokens = [t for t in q_norm.split() if t]
    if not tokens:
        return df_sub.head(limit)

    q_upper = query.strip().upper()
    cod_series = df_sub["Codigo"]
    search_series = df_sub["Search_Norm"]

    # 1. Filtro por presença de todos os tokens
    mask_all = pd.Series(True, index=df_sub.index)
    for tok in tokens:
        mask_all &= search_series.str.contains(re.escape(tok), regex=True, na=False)

    exact_code_mask = cod_series == q_upper
    prefix_code_mask = cod_series.str.startswith(q_upper, na=False)
    combined_mask = exact_code_mask | prefix_code_mask | mask_all

    candidates = df_sub[combined_mask].copy()
    if not candidates.empty:
        desc_norm_first = candidates["Descricao"].apply(lambda d: _normalize_search_text(d[:40]))
        scores = np.zeros(len(candidates), dtype=float)
        scores += np.where(candidates["Codigo"] == q_upper, 1000.0, 0.0)
        scores += np.where(candidates["Codigo"].str.startswith(q_upper, na=False), 500.0, 0.0)
        scores += np.where(desc_norm_first.str.startswith(tokens[0], na=False), 120.0, 0.0)
        # Descrições mais diretas/concisas recebem leve desempate positivo
        scores += np.maximum(0.0, 60.0 - candidates["Descricao"].str.len() * 0.15)
        candidates["_score"] = scores
        candidates = candidates.sort_values("_score", ascending=False).drop(columns=["_score"])

    # 2. Fallback parcial / fuzzy caso a busca exata retorne poucos itens
    if len(candidates) < min(10, limit) and len(df_sub) > 0:
        already_idx = set(candidates.index.tolist())
        remaining = df_sub.loc[~df_sub.index.isin(already_idx)]

        # Se tiver múltiplos tokens, tenta correspondência parcial (N-1 termos)
        partial_df = pd.DataFrame()
        if len(tokens) >= 2 and not remaining.empty:
            hit_counts = sum(
                remaining["Search_Norm"].str.contains(re.escape(tok), regex=True, na=False).astype(int)
                for tok in tokens
            )
            min_hits = max(1, len(tokens) - 1)
            partial_df = remaining[hit_counts >= min_hits].copy()
            if not partial_df.empty:
                partial_df["_hits"] = hit_counts[hit_counts >= min_hits]
                partial_df = partial_df.sort_values("_hits", ascending=False).drop(columns=["_hits"]).head(limit)
                already_idx.update(partial_df.index.tolist())
                remaining = df_sub.loc[~df_sub.index.isin(already_idx)]

        # Fuzzy matching com rapidfuzz para erros de digitação
        fuzzy_df = pd.DataFrame()
        if len(candidates) + len(partial_df) < 8 and not remaining.empty:
            choices = remaining["Search_Norm"].to_dict()
            matches = rf_process.extract(
                q_norm,
                choices,
                scorer=fuzz.WRatio,
                limit=15,
                score_cutoff=58,
            )
            if matches:
                matched_indices = [m[2] for m in matches]
                fuzzy_df = remaining.loc[matched_indices]

        parts = [p for p in (candidates, partial_df, fuzzy_df) if not p.empty]
        if parts:
            return pd.concat(parts, ignore_index=False).head(limit)

    return candidates.head(limit)


# ==============================================================================
# 3. GERENCIAMENTO DE ESTADO DE EDIÇÃO E SOBRESCRITA DE PREÇOS
# ==============================================================================
def build_initial_records(df: pd.DataFrame, catalog_lookup: Dict[str, Dict[str, Any]], regime_col: str) -> List[Dict[str, Any]]:
    """Constrói metadados por linha para rastrear códigos da base vs. edições manuais."""
    records: List[Dict[str, Any]] = []
    for _, r in df.iterrows():
        wbs = str(r.get("Item_WBS", "")).strip()
        cod = normalize_code_str(r.get("Codigo", r.get("Codigo_SINAPI", "")))
        desc = str(r.get("Descricao", "") or "").strip()
        unid = str(r.get("Unidade", "") or "").strip()
        qty = float(parse_br_number(r.get("Quantidade", 0.0)) or 0.0)
        price = float(parse_br_number(r.get("Preco_Unitario", r.get("Preco_Unitario_Manual", 0.0))) or 0.0)
        ap_pct = bool(r.get("Aplicar_Pct", True))

        cat = catalog_lookup.get(cod) if cod else None
        desc_from_base = False
        unid_from_base = False
        price_overridden = False

        if cat is not None:
            base_price = float(cat[regime_col])
            if desc == "":
                desc = cat["Descricao"]
                desc_from_base = True
            elif desc == cat["Descricao"]:
                desc_from_base = True

            if unid == "":
                unid = cat["Unidade"]
                unid_from_base = True
            elif unid == cat["Unidade"]:
                unid_from_base = True

            if price == 0.0:
                price = base_price
                price_overridden = False
            elif abs(price - base_price) > 1e-4:
                price_overridden = True

        records.append(
            {
                "Item_WBS": wbs,
                "Codigo": cod,
                "Descricao": desc,
                "Unidade": unid,
                "Quantidade": qty,
                "Preco_Unitario": price,
                "Aplicar_Pct": ap_pct,
                "price_overridden": price_overridden,
                "desc_from_base": desc_from_base,
                "unid_from_base": unid_from_base,
            }
        )
    return records


def records_to_editor_df(records: List[Dict[str, Any]]) -> pd.DataFrame:
    cols = ["Item_WBS", "Codigo", "Descricao", "Unidade", "Quantidade", "Preco_Unitario", "Aplicar_Pct"]
    if not records:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame([{k: rec[k] for k in cols} for rec in records], columns=cols)


def reconcile_edited_df(
    edited_df: pd.DataFrame,
    prev_records: List[Dict[str, Any]],
    catalog_lookup: Dict[str, Dict[str, Any]],
    regime_col: str,
    regime_changed: bool,
) -> Tuple[List[Dict[str, Any]], bool, bool]:
    """
    Reconcilia a edição do usuário no `st.data_editor` aplicando as regras exatas:
    1. Se o Código (SINAPI ou SEINFRA) foi inserido/alterado na linha -> puxa o Preço da base
       (e preenche Descrição/Medida caso estejam em branco ou viessem do código anterior).
    2. Se o Código NÃO mudou, mas o usuário editou manualmente o Preço Unitário -> mantém o
       valor digitado pelo usuário sobrescrevendo o valor puxado da base (`price_overridden=True`).
    3. Retorna `(new_records, needs_editor_refresh, pct_checkbox_toggled)`.
    """
    new_records: List[Dict[str, Any]] = []
    needs_refresh = False
    pct_checkbox_toggled = False

    same_len = len(edited_df) == len(prev_records)
    # Pool auxiliar caso linhas tenham sido adicionadas/excluídas
    prev_pool: List[Optional[Dict[str, Any]]] = list(prev_records)

    for idx, (_, row) in enumerate(edited_df.iterrows()):
        wbs = str(row.get("Item_WBS", "") or "").strip()
        if wbs in ("None", "nan"):
            wbs = ""
        cod = normalize_code_str(row.get("Codigo", ""))
        desc = str(row.get("Descricao", "") or "").strip()
        if desc in ("None", "nan"):
            desc = ""
        unid = str(row.get("Unidade", "") or "").strip()
        if unid in ("None", "nan"):
            unid = ""
        qty = float(parse_br_number(row.get("Quantidade", 0.0)) or 0.0)
        price = float(parse_br_number(row.get("Preco_Unitario", 0.0)) or 0.0)
        ap_pct_raw = row.get("Aplicar_Pct", True)
        ap_pct = True if pd.isna(ap_pct_raw) else bool(ap_pct_raw)

        prev_rec: Optional[Dict[str, Any]] = None
        if same_len:
            prev_rec = prev_records[idx]
        else:
            # Tenta casar pelo Item_WBS primeiro
            for p_i, cand in enumerate(prev_pool):
                if cand is not None and cand["Item_WBS"] == wbs and wbs != "":
                    prev_rec = cand
                    prev_pool[p_i] = None
                    break

        prev_cod = prev_rec["Codigo"] if prev_rec else ""
        cat = catalog_lookup.get(cod) if cod else None

        price_overridden = prev_rec["price_overridden"] if prev_rec else False
        desc_from_base = prev_rec["desc_from_base"] if prev_rec else False
        unid_from_base = prev_rec["unid_from_base"] if prev_rec else False

        if prev_rec is not None and ap_pct != prev_rec["Aplicar_Pct"]:
            pct_checkbox_toggled = True

        # Caso A: Código SINAPI/SEINFRA foi adicionado ou trocado nesta linha
        if cod != prev_cod:
            if cod != "" and cat is not None:
                price = float(cat[regime_col])
                price_overridden = False
                if desc == "" or desc_from_base:
                    desc = cat["Descricao"]
                    desc_from_base = True
                else:
                    desc_from_base = False
                if unid == "" or unid_from_base:
                    unid = cat["Unidade"]
                    unid_from_base = True
                else:
                    unid_from_base = False
                needs_refresh = True
            elif cod == "":
                price_overridden = False
                desc_from_base = False
                unid_from_base = False

        # Caso B: Código permaneceu o mesmo
        else:
            if prev_rec is not None:
                # Verifica se o usuário alterou manualmente o Preço Unitário
                if abs(price - float(prev_rec["Preco_Unitario"])) > 1e-5:
                    if cod != "" and cat is not None:
                        price_overridden = True
                elif regime_changed and cod != "" and cat is not None and not price_overridden:
                    price = float(cat[regime_col])
                    needs_refresh = True

                # Verifica se o usuário alterou ou limpou a Descrição
                if desc != prev_rec["Descricao"]:
                    if desc == "" and cat is not None:
                        desc = cat["Descricao"]
                        desc_from_base = True
                        needs_refresh = True
                    else:
                        desc_from_base = False

                if unid != prev_rec["Unidade"]:
                    if unid == "" and cat is not None:
                        unid = cat["Unidade"]
                        unid_from_base = True
                        needs_refresh = True
                    else:
                        unid_from_base = False

        new_records.append(
            {
                "Item_WBS": wbs,
                "Codigo": cod,
                "Descricao": desc,
                "Unidade": unid,
                "Quantidade": qty,
                "Preco_Unitario": price,
                "Aplicar_Pct": ap_pct,
                "price_overridden": price_overridden,
                "desc_from_base": desc_from_base,
                "unid_from_base": unid_from_base,
            }
        )

    return new_records, needs_refresh, pct_checkbox_toggled


# ==============================================================================
# 4. MOTOR DE REGRAS DE NEGÓCIO, ACRÉSCIMO SELETIVO (%), TRUNC E ROLL-UP WBS
# ==============================================================================
def process_budget_wbs(
    records: List[Dict[str, Any]],
    catalog_lookup: Dict[str, Dict[str, Any]],
    acrescimo_percentual: float = 0.0,
    modo_aplicacao_pct: str = MODO_PCT_TODOS,
    regime_preco: str = "Preco_Com_Desoneracao",
    calc_mode: str = "trunc",
) -> Tuple[pd.DataFrame, float, List[str], List[str]]:
    """
    Processa a lista de registros do orçamento aplicando:
    - Identificação de Grupos EAP vs. Itens Folha
    - Resolução SINAPI-CE / SEINFRA-CE com suporte a sobrescrita manual de preço
    - Aplicação seletiva de porcentagem (Todos, Marcados, Só Base, Só Manuais)
    - Truncamento exato em 2 casas (padrão TRUNC do Excel) ou Arredondamento ABNT
    - Roll-up WBS de subtotais sem dupla contagem
    """
    valid_recs = [r.copy() for r in records if str(r.get("Item_WBS", "")).strip() != ""]
    if not valid_recs:
        return pd.DataFrame(), 0.0, [], []

    df = pd.DataFrame(valid_recs)
    df["Item_WBS"] = df["Item_WBS"].astype(str).str.strip()

    invalid_wbs = [
        wbs for wbs in df["Item_WBS"] if any(not p.isdigit() for p in wbs.split("."))
    ]

    all_wbs = df["Item_WBS"].tolist()
    df["Is_Parent"] = [
        any(other.startswith(wbs + ".") for other in all_wbs if other != wbs)
        for wbs in all_wbs
    ]

    pct_dec = Decimal(str(acrescimo_percentual)) if acrescimo_percentual else Decimal("0")
    fator_dec = Decimal("1") + (pct_dec / Decimal("100"))

    desc_final: List[str] = []
    unid_final: List[str] = []
    preco_base_list: List[float] = []
    preco_final_list: List[float] = []
    valor_total_list: List[Decimal] = []
    origem_item: List[str] = []
    pct_aplicado_list: List[str] = []
    missing_codes: List[str] = []

    for _, row in df.iterrows():
        is_parent = bool(row["Is_Parent"])
        cod = normalize_code_str(row["Codigo"])
        user_desc = str(row["Descricao"] or "").strip()
        user_unid = str(row["Unidade"] or "").strip()
        qty = float(row["Quantidade"] or 0.0)
        row_price = float(row["Preco_Unitario"] or 0.0)
        overridden = bool(row.get("price_overridden", False))
        marked_pct = bool(row.get("Aplicar_Pct", True))

        if is_parent:
            desc_final.append(user_desc if user_desc != "" else f"GRUPO {row['Item_WBS']}")
            unid_final.append("")
            preco_base_list.append(0.0)
            preco_final_list.append(0.0)
            valor_total_list.append(Decimal("0.00"))
            origem_item.append("Grupo (EAP)")
            pct_aplicado_list.append("-")
            continue

        cat = catalog_lookup.get(cod) if cod else None
        has_valid_base = cat is not None

        if cod != "" and has_valid_base:
            fonte_nome = f"{cat['Fonte']}-CE"
            d_fin = user_desc if user_desc != "" else str(cat["Descricao"])
            u_fin = user_unid if user_unid != "" else str(cat["Unidade"])
            if overridden:
                p_base = row_price
                orig = f"{fonte_nome} (Editado)"
            else:
                p_base = float(cat[regime_preco]) if row_price == 0.0 else row_price
                orig = fonte_nome
        elif cod != "" and not has_valid_base:
            missing_codes.append(cod)
            d_fin = user_desc if user_desc != "" else f"[CÓDIGO {cod} NÃO ENCONTRADO]"
            u_fin = user_unid
            p_base = row_price
            orig = "Código Inválido"
        else:
            d_fin = user_desc
            u_fin = user_unid
            p_base = row_price
            orig = "Manual / Importado"

        # Decide se o percentual incide neste item conforme o seletor escolhido
        if pct_dec == Decimal("0"):
            recebe_pct = False
        elif modo_aplicacao_pct == MODO_PCT_TODOS:
            recebe_pct = True
        elif modo_aplicacao_pct == MODO_PCT_MARCADOS:
            recebe_pct = marked_pct
        elif modo_aplicacao_pct == MODO_PCT_BASES:
            recebe_pct = has_valid_base
        elif modo_aplicacao_pct == MODO_PCT_MANUAIS:
            recebe_pct = not has_valid_base
        else:
            recebe_pct = True

        p_base_dec = Decimal(str(p_base))
        qty_dec = Decimal(str(qty))

        if recebe_pct and pct_dec != Decimal("0"):
            p_final_val = quantize_money(p_base_dec * fator_dec, mode=calc_mode)
            p_final_dec = Decimal(str(p_final_val))
            pct_str = f"{acrescimo_percentual:+.2f}%"
        else:
            # Preserva precisão integral do unitário (ex: 35789.947) quando não há %
            p_final_dec = p_base_dec
            p_final_val = float(p_base_dec)
            pct_str = "0%"

        v_tot_val = quantize_money(qty_dec * p_final_dec, mode=calc_mode)

        desc_final.append(d_fin)
        unid_final.append(u_fin)
        preco_base_list.append(p_base)
        preco_final_list.append(p_final_val)
        valor_total_list.append(Decimal(f"{v_tot_val:.2f}"))
        origem_item.append(orig)
        pct_aplicado_list.append(pct_str)

    df["Descricao_Final"] = desc_final
    df["Unidade_Final"] = unid_final
    df["Preco_Unitario_Base"] = preco_base_list
    df["Preco_Unitario_Final"] = preco_final_list
    df["Pct_Aplicado"] = pct_aplicado_list
    df["Origem"] = origem_item
    df["Quantidade_Final"] = np.where(df["Is_Parent"], np.nan, df["Quantidade"])
    df["_Valor_Dec"] = valor_total_list

    # Roll-up exato em Decimal dos Nós Folha abaixo de cada Nó Pai
    leaf_df = df[~df["Is_Parent"]]
    for idx, row in df[df["Is_Parent"]].iterrows():
        parent_prefix = row["Item_WBS"] + "."
        descendant_leaves = leaf_df[leaf_df["Item_WBS"].str.startswith(parent_prefix)]
        sub_sum = sum(descendant_leaves["_Valor_Dec"].tolist(), Decimal("0.00"))
        df.at[idx, "_Valor_Dec"] = sub_sum

    df["Valor_Total"] = df["_Valor_Dec"].apply(float)
    valor_global_dec = sum(leaf_df["_Valor_Dec"].tolist(), Decimal("0.00"))
    df = df.drop(columns=["_Valor_Dec"])

    return df, float(valor_global_dec), sorted(list(set(missing_codes))), invalid_wbs


# ==============================================================================
# 5. CONVERSOR DE VALOR MONETÁRIO POR EXTENSO (PT-BR) E MOTOR OPENPYXL
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
    template_path: Any,
    header_data: Dict[str, str],
    df_calc: pd.DataFrame,
    valor_global: float,
) -> io.BytesIO:
    """Injeta os dados no template Excel mantendo logomarca, bordas e rodapé."""
    wb = openpyxl.load_workbook(template_path)
    sheet_name = CELL_MAPPING["SHEET_NAME"]
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.worksheets[0]

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
        row_data = df_calc.iloc[i]
        is_parent = bool(row_data["Is_Parent"])
        desc_text = str(row_data["Descricao_Final"])

        # Ajusta altura da linha automaticamente para descrições longas não cortarem no Excel
        est_lines = max(1, math.ceil(len(desc_text) / 68))
        ws.row_dimensions[current_row].height = max(default_item_height, est_lines * 15.5)

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
        c_desc = ws.cell(current_row, CELL_MAPPING["COL_DESC_START"], desc_text)
        c_unid = ws.cell(current_row, CELL_MAPPING["COL_UNIDADE"], "" if is_parent else str(row_data["Unidade_Final"]))
        c_qtd = ws.cell(current_row, CELL_MAPPING["COL_QUANTIDADE"], None if is_parent else float(row_data["Quantidade_Final"]))
        c_val = ws.cell(current_row, CELL_MAPPING["COL_VALOR"], float(row_data["Valor_Total"]))

        c_wbs.font = Font(name=c_wbs.font.name or "Calibri", size=c_wbs.font.size or 10, bold=True)
        c_desc.font = Font(name=c_desc.font.name or "Calibri", size=c_desc.font.size or 10, bold=is_parent)
        c_desc.alignment = Alignment(
            horizontal=c_desc.alignment.horizontal or "left",
            vertical="center",
            wrap_text=True,
        )
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


def format_brl(val: float) -> str:
    return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def suggest_next_wbs(records: List[Dict[str, Any]]) -> str:
    """Sugere o próximo código WBS sequencial com base na última linha da tabela."""
    wbs_list = [str(r.get("Item_WBS", "")).strip() for r in records if str(r.get("Item_WBS", "")).strip()]
    if not wbs_list:
        return "1.1"
    last = wbs_list[-1]
    parts = last.split(".")
    if all(p.isdigit() for p in parts):
        parts[-1] = str(int(parts[-1]) + 1)
        return ".".join(parts)
    return "1.1"


# ==============================================================================
# 6. INTERFACE DO USUÁRIO NO STREAMLIT
# ==============================================================================
def main():
    st.set_page_config(
        page_title="Orçamentador SINAPI + SEINFRA | Mais Engenharia",
        page_icon="🏗️",
        layout="wide",
    )

    st.title("🏗️ Sistema de Orçamentação Dinâmica (SINAPI-CE + SEINFRA-CE)")
    st.caption(
        "Mais Engenharia & Consultoria — Importação de Orçamentos, Busca Inteligente e Propostas Oficiais em EAP/WBS"
    )

    if not os.path.exists(SINAPI_CSV_PATH) and not os.path.exists(SEINFRA_CSV_PATH):
        st.error(
            f"Nenhuma base encontrada (`{SINAPI_CSV_PATH}` ou `{SEINFRA_CSV_PATH}`). "
            "Coloque os arquivos na pasta do projeto ou envie pela barra lateral."
        )
        st.stop()

    try:
        df_catalog = load_unified_catalog(SINAPI_CSV_PATH, SEINFRA_CSV_PATH)
        catalog_lookup = build_catalog_lookup(df_catalog)
    except Exception as e:
        st.error(f"Erro ao carregar catálogo SINAPI/SEINFRA: {e}")
        st.stop()

    template_path = next((p for p in TEMPLATE_CANDIDATES if os.path.exists(p)), None)

    # Inicialização de chaves de sessão
    if "editor_version" not in st.session_state:
        st.session_state.editor_version = 0
    if "modo_aplicacao_pct" not in st.session_state:
        st.session_state.modo_aplicacao_pct = MODO_PCT_TODOS

    # --------------------------------------------------------------------------
    # BARRA LATERAL: CONSULTA RÁPIDA (SINAPI + SEINFRA) E GESTÃO DE BASES
    # --------------------------------------------------------------------------
    with st.sidebar:
        st.header("🔍 Consulta Rápida (Bases)")
        filtro_sidebar = st.radio(
            "Filtrar Base:",
            options=["Todas", "SINAPI", "SEINFRA"],
            horizontal=True,
        )
        termo_sidebar = st.text_input(
            "Pesquisar por código ou descrição:",
            placeholder="Ex: 103328, C2820, alvenaria 9x19...",
        )

        df_side = search_catalog_google_style(
            termo_sidebar,
            df_catalog,
            fonte_filter=filtro_sidebar,
            limit=50 if termo_sidebar else 25,
        )

        st.dataframe(
            df_side[["Fonte", "Codigo", "Descricao", "Unidade", "Preco_Com_Desoneracao", "Preco_Sem_Desoneracao"]],
            width="stretch",
            hide_index=True,
            column_config={
                "Fonte": st.column_config.TextColumn("Base"),
                "Codigo": st.column_config.TextColumn("Código"),
                "Descricao": st.column_config.TextColumn("Descrição"),
                "Unidade": st.column_config.TextColumn("Unid."),
                "Preco_Com_Desoneracao": st.column_config.NumberColumn("Com Deson.", format="R$ %.2f"),
                "Preco_Sem_Desoneracao": st.column_config.NumberColumn("Sem Deson.", format="R$ %.2f"),
            },
        )

        qtd_sinapi = int((df_catalog["Fonte"] == "SINAPI").sum())
        qtd_seinfra = int((df_catalog["Fonte"] == "SEINFRA").sum())
        st.caption(
            f"**SINAPI-CE:** {qtd_sinapi:,} | **SEINFRA-CE:** {qtd_seinfra:,} | **Total:** {len(df_catalog):,}".replace(",", ".")
        )

        st.divider()
        with st.expander("⚙️ Substituir Template ou Atualizar SEINFRA"):
            uploaded_template = st.file_uploader("Template Oficial (.xlsx)", type=["xlsx"])
            uploaded_seinfra = st.file_uploader(
                "Atualizar Tabela SEINFRA (.xls, .xlsx, .csv)",
                type=["xls", "xlsx", "csv"],
                key="up_seinfra_base",
            )
            if uploaded_seinfra is not None:
                if st.button("💾 Processar e Atualizar Base SEINFRA", width="stretch"):
                    try:
                        df_new_sei = parse_seinfra_file(uploaded_seinfra.getvalue(), filename=uploaded_seinfra.name)
                        df_new_sei.to_csv(SEINFRA_CSV_PATH, index=False, encoding="utf-8-sig")
                        load_unified_catalog.clear()
                        st.success(f"SEINFRA atualizada ({len(df_new_sei)} itens)!")
                        st.rerun()
                    except Exception as ex:
                        st.error(f"Erro ao converter SEINFRA: {ex}")

    # --------------------------------------------------------------------------
    # SEÇÃO 0: IMPORTADOR / CONVERSOR DE ORÇAMENTOS DE TERCEIROS
    # --------------------------------------------------------------------------
    with st.expander("📥 Importar / Converter Orçamento de Outra Empresa (.xlsx, .xls, .csv, .pdf)", expanded=False):
        st.markdown(
            "Envie a planilha ou PDF recebido do cliente/terceiro. O sistema detecta as colunas automaticamente, "
            "mantém todos os itens e valores originais, confere o total e carrega direto na área de edição."
        )
        imp_file = st.file_uploader(
            "Selecione o arquivo de orçamento (.xlsx, .xls, .csv ou .pdf)",
            type=["xlsx", "xls", "csv", "pdf"],
            key="budget_importer_file",
        )

        if imp_file is not None:
            try:
                raw_sheets = load_raw_sheets_from_file(imp_file.getvalue(), filename=imp_file.name)
                if not raw_sheets:
                    st.warning("Nenhuma tabela pôde ser identificada no arquivo enviado.")
                else:
                    c_sh1, c_sh2, c_sh3 = st.columns([1.5, 1, 2.5])
                    sheet_names = list(raw_sheets.keys())
                    auto_init = auto_detect_sheet_and_columns(raw_sheets)
                    default_sh_idx = sheet_names.index(auto_init["sheet_name"]) if auto_init["sheet_name"] in sheet_names else 0

                    with c_sh1:
                        sel_sheet = st.selectbox("Aba / Origem detectada", options=sheet_names, index=default_sh_idx)

                    det = auto_detect_sheet_and_columns(raw_sheets, preferred_sheet=sel_sheet)
                    sheet_rows = raw_sheets[sel_sheet]

                    with c_sh2:
                        hdr_row_1based = st.number_input(
                            "Linha do Cabeçalho",
                            min_value=1,
                            max_value=max(1, len(sheet_rows)),
                            value=int(det["header_row_idx"]) + 1,
                            step=1,
                        )
                    hdr_idx = int(hdr_row_1based) - 1

                    # Se o usuário alterou manualmente a linha de cabeçalho, recalcula rótulos daquela linha
                    if hdr_idx != det["header_row_idx"]:
                        col_labels = [
                            f"[{openpyxl.utils.get_column_letter(i + 1)}] {str(c or f'Coluna {i + 1}').strip()}"
                            for i, c in enumerate(sheet_rows[hdr_idx])
                        ]
                    else:
                        col_labels = det["col_labels"]

                    with c_sh3:
                        opt_skip_empty = st.checkbox(
                            "Ignorar linhas em branco, grupos vazios (ex: OMISSOS) e observações de rodapé",
                            value=True,
                        )
                        opt_autofix_wbs = st.checkbox(
                            "Corrigir automaticamente numeração WBS duplicada ou fora de hierarquia",
                            value=True,
                        )

                    # Ajuste fino opcional de mapeamento de colunas
                    none_label = "<Nenhuma>"
                    options_with_none = [none_label] + col_labels

                    def _idx_for(col_i: Optional[int]) -> int:
                        if col_i is None or col_i < 0 or col_i >= len(col_labels):
                            return 0
                        return col_i + 1

                    mp = det["mapping"]
                    with st.expander("🔧 Conferir / Ajustar Colunas Mapeadas (Opcional)", expanded=False):
                        mc1, mc2, mc3, mc4, mc5, mc6, mc7 = st.columns(7)
                        sel_item = mc1.selectbox("Item (WBS)", options_with_none, index=_idx_for(mp["item_col"]))
                        sel_cod = mc2.selectbox("Código (se houver)", options_with_none, index=_idx_for(mp["codigo_col"]))
                        sel_desc = mc3.selectbox("Descrição", options_with_none, index=_idx_for(mp["desc_col"]))
                        sel_unid = mc4.selectbox("Unidade", options_with_none, index=_idx_for(mp["unid_col"]))
                        sel_qty = mc5.selectbox("Quantidade", options_with_none, index=_idx_for(mp["qty_col"]))
                        sel_prc = mc6.selectbox("Valor Unitário", options_with_none, index=_idx_for(mp["price_col"]))
                        sel_tot = mc7.selectbox("Total (Conferência)", options_with_none, index=_idx_for(mp["total_col"]))

                    def _to_col_idx(chosen: str) -> Optional[int]:
                        return None if chosen == none_label else col_labels.index(chosen)

                    final_map = {
                        "item_col": _to_col_idx(sel_item),
                        "codigo_col": _to_col_idx(sel_cod),
                        "desc_col": _to_col_idx(sel_desc),
                        "unid_col": _to_col_idx(sel_unid),
                        "qty_col": _to_col_idx(sel_qty),
                        "price_col": _to_col_idx(sel_prc),
                        "total_col": _to_col_idx(sel_tot),
                    }

                    imp_res = extract_budget_from_sheet(
                        sheet_rows,
                        header_row_idx=hdr_idx,
                        mapping=final_map,
                        skip_empty_and_omissos=opt_skip_empty,
                        auto_fix_wbs=opt_autofix_wbs,
                        calc_mode="trunc",
                    )

                    df_imp = imp_res["df"]
                    calc_tot = imp_res["calculated_total"]
                    rep_tot = imp_res["reported_total"]

                    km1, km2, km3, km4 = st.columns(4)
                    km1.metric("Linhas Estruturadas", len(df_imp))
                    km2.metric(
                        "Total na Planilha Original",
                        format_brl(rep_tot) if rep_tot is not None else "Não informado",
                    )
                    km3.metric("Total Calculado (Truncado)", format_brl(calc_tot))
                    if rep_tot is not None:
                        diff_val = round(calc_tot - rep_tot, 2)
                        status_txt = "✅ 100% Idêntico" if abs(diff_val) < 0.01 else f"Dif: {format_brl(diff_val)}"
                    else:
                        status_txt = "Pronto p/ importar"
                    km4.metric("Conferência", status_txt)

                    if imp_res["anomalies"]:
                        if opt_autofix_wbs and imp_res["fixes_applied"]:
                            st.info(
                                "🛠️ **Inconsistências de numeração detectadas no arquivo original e corrigidas automaticamente:**\n"
                                + "\n".join(f"- {fx}" for fx in imp_res["fixes_applied"])
                            )
                        else:
                            st.warning(
                                "⚠️ **Atenção — Inconsistências de numeração WBS encontradas no arquivo original:**\n"
                                + "\n".join(f"- {an}" for an in imp_res["anomalies"])
                            )

                    if imp_res["warnings"]:
                        st.warning("\n".join(f"- {w}" for w in imp_res["warnings"]))

                    btn_imp_col1, btn_imp_col2 = st.columns([2, 1.5])
                    with btn_imp_col1:
                        if st.button(
                            "🚀 Carregar Orçamento na Tabela de Edição",
                            type="primary",
                            width="stretch",
                            disabled=df_imp.empty,
                        ):
                            active_reg = st.session_state.get("last_regime_col", "Preco_Com_Desoneracao")
                            st.session_state.budget_records = build_initial_records(df_imp, catalog_lookup, active_reg)
                            st.session_state.budget_df = records_to_editor_df(st.session_state.budget_records)
                            st.session_state.editor_version += 1
                            st.toast("Orçamento importado com sucesso para a tabela de edição!", icon="✅")
                            st.rerun()

                    with btn_imp_col2:
                        if not df_imp.empty:
                            temp_recs = build_initial_records(df_imp, catalog_lookup, "Preco_Com_Desoneracao")
                            df_conv_calc, v_conv_glob, _, _ = process_budget_wbs(temp_recs, catalog_lookup, calc_mode="trunc")
                            xlsx_clean_buf = export_clean_budget_xlsx(
                                df_conv_calc,
                                v_conv_glob,
                                titulo=f"Orçamento Convertido — {imp_file.name}",
                            )
                            base_no_ext = os.path.splitext(imp_file.name)[0]
                            st.download_button(
                                label="📊 Converter e Baixar Excel Limpo (.xlsx)",
                                data=xlsx_clean_buf,
                                file_name=f"{base_no_ext}_convertido.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                width="stretch",
                            )
            except Exception as ex:
                st.error(f"Erro ao ler arquivo para importação: {ex}")

    # --------------------------------------------------------------------------
    # SEÇÃO 1: DADOS DO CABEÇALHO, ACRÉSCIMO SELETIVO (%), CÁLCULO E RODAPÉ
    # --------------------------------------------------------------------------
    with st.expander("📋 1. Dados da Proposta, Acréscimo Seletivo (%) e Condições", expanded=True):
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

        col5, col6, col7 = st.columns([1.5, 1.2, 2.0])
        with col5:
            contato = st.text_input("Contato", value="Jaqueline Azevedo")
        with col6:
            telefone = st.text_input("Telefone", value="(85) 9.9957-1480")
        with col7:
            email = st.text_input("E-mail", value="jaqueline.silva@solarbr.com.br")

        st.markdown("**Configurações de Precificação, Porcentagem (%) e Arredondamento**")
        cp1, cp2, cp3, cp4 = st.columns([1.3, 1.1, 2.1, 1.8])
        with cp1:
            regime_label = st.selectbox(
                "Regime das Bases",
                options=["Com Desoneração", "Sem Desoneração"],
                index=0,
                help="Aplica-se às composições puxadas do SINAPI-CE / SEINFRA-CE que não foram sobrescritas manualmente.",
            )
            regime_col = "Preco_Com_Desoneracao" if regime_label == "Com Desoneração" else "Preco_Sem_Desoneracao"
        with cp2:
            acrescimo_percentual = st.number_input(
                "Acréscimo / Ajuste (%)",
                min_value=-99.0,
                max_value=500.0,
                value=0.0,
                step=1.0,
                format="%.2f",
                help="Ex: Digite 25 para somar +25% nos itens escolhidos ao lado. Deixe 0 para valores originais.",
            )
        with cp3:
            opcoes_modo_pct = [MODO_PCT_TODOS, MODO_PCT_MARCADOS, MODO_PCT_BASES, MODO_PCT_MANUAIS]
            idx_modo = (
                opcoes_modo_pct.index(st.session_state.modo_aplicacao_pct)
                if st.session_state.modo_aplicacao_pct in opcoes_modo_pct
                else 0
            )
            modo_aplicacao_pct = st.selectbox(
                "Onde aplicar a % (Seletor de Incidência)",
                options=opcoes_modo_pct,
                index=idx_modo,
                help="Permite aplicar a porcentagem em cima de todos os valores ou apenas em um/alguns itens selecionados na tabela.",
            )
            st.session_state.modo_aplicacao_pct = modo_aplicacao_pct
        with cp4:
            regra_calculo_label = st.selectbox(
                "Regra de Casas Decimais",
                options=[
                    "Truncar em 2 casas (igual TRUNC do Excel)",
                    "Arredondar em 2 casas (padrão ABNT)",
                ],
                index=0,
                help="Truncar corta na 2ª casa decimal sem arredondar para cima, mantendo igualdade centavo a centavo com orçamentos importados.",
            )
            calc_mode = "trunc" if regra_calculo_label.startswith("Truncar") else "round"

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
    # INICIALIZAÇÃO DA TABELA PADRÃO CASO AINDA NÃO EXISTA NA SESSÃO
    # --------------------------------------------------------------------------
    if "budget_records" not in st.session_state:
        default_seed = pd.DataFrame(
            [
                {"Item_WBS": "1", "Codigo": "", "Descricao": "CUSTOS INDIRETOS", "Unidade": "", "Quantidade": 0.0, "Preco_Unitario": 0.0, "Aplicar_Pct": False},
                {"Item_WBS": "1.1", "Codigo": "", "Descricao": "CUSTO COM EQUIPE DE APOIO E LIMPEZA", "Unidade": "MÊS", "Quantidade": 1.0, "Preco_Unitario": 3995.00, "Aplicar_Pct": True},
                {"Item_WBS": "2", "Codigo": "", "Descricao": "CUSTOS DIRETOS", "Unidade": "", "Quantidade": 0.0, "Preco_Unitario": 0.0, "Aplicar_Pct": False},
                {"Item_WBS": "2.1", "Codigo": "", "Descricao": "ALVENARIA E REVESTIMENTOS", "Unidade": "", "Quantidade": 0.0, "Preco_Unitario": 0.0, "Aplicar_Pct": False},
                {"Item_WBS": "2.1.1", "Codigo": "103328", "Descricao": "", "Unidade": "", "Quantidade": 4.0, "Preco_Unitario": 0.0, "Aplicar_Pct": True},
                {"Item_WBS": "2.1.2", "Codigo": "87879", "Descricao": "", "Unidade": "", "Quantidade": 4.0, "Preco_Unitario": 0.0, "Aplicar_Pct": True},
            ]
        )
        st.session_state.budget_records = build_initial_records(default_seed, catalog_lookup, regime_col)
        st.session_state.budget_df = records_to_editor_df(st.session_state.budget_records)
        st.session_state.last_regime_col = regime_col

    regime_changed = st.session_state.get("last_regime_col", regime_col) != regime_col
    st.session_state.last_regime_col = regime_col

    # --------------------------------------------------------------------------
    # SEÇÃO 2: BUSCA ESTILO GOOGLE (SINAPI + SEINFRA) + TABELA INTERATIVA WBS
    # --------------------------------------------------------------------------
    st.subheader("📐 2. Estrutura Analítica do Orçamento (WBS / EAP)")

    with st.container(border=True):
        st.markdown("##### 🔎 Assistente de Busca estilo Google (SINAPI + SEINFRA) — Pesquisar Serviço e Inserir Código")
        g_col1, g_col2, g_col3 = st.columns([2.8, 1.0, 2.2])

        # Monta lista de itens atuais para permitir puxar palavras ou aplicar o código encontrado
        current_recs = st.session_state.budget_records
        all_curr_wbs = [r["Item_WBS"] for r in current_recs if r["Item_WBS"]]
        leaf_targets: List[Tuple[int, str]] = []
        for idx_r, r in enumerate(current_recs):
            w = r["Item_WBS"]
            is_par = any(o.startswith(w + ".") for o in all_curr_wbs if o != w) if w else False
            if not is_par:
                label_t = f"{w or f'Linha {idx_r + 1}'} — {r['Descricao'][:55] or '(Sem descrição)'}"
                leaf_targets.append((idx_r, label_t))

        with g_col3:
            target_options = ["➕ Adicionar como nova linha no final"] + [lbl for _, lbl in leaf_targets]
            selected_target_lbl = st.selectbox(
                "Item alvo na tabela:",
                options=target_options,
                index=1 if len(target_options) > 1 else 0,
            )

        with g_col1:
            google_query = st.text_input(
                "Digite palavras-chave do serviço (sem se preocupar com acentos ou ordem) ou código:",
                placeholder="Ex: alvenaria 9x19, chapisco, piso vinilico, forro gesso, limpeza piso, C1628...",
                key="google_search_input",
            )
        with g_col2:
            google_fonte = st.selectbox("Filtrar Base:", options=["Todas", "SINAPI", "SEINFRA"], key="google_fonte_sel")

        if google_query.strip():
            df_matches = search_catalog_google_style(google_query, df_catalog, fonte_filter=google_fonte, limit=35)
            if df_matches.empty:
                st.warning("Nenhuma composição encontrada para os termos digitados.")
            else:
                match_records = df_matches.to_dict("records")
                match_labels = [
                    f"[{m['Fonte']} {m['Codigo']}] {m['Descricao'][:95]} ({m['Unidade']}) — {format_brl(float(m[regime_col]))}"
                    for m in match_records
                ]
                chosen_match_lbl = st.selectbox(
                    f"Sugestões encontradas ({len(match_records)} resultados mais relevantes):",
                    options=match_labels,
                )
                chosen_rec = match_records[match_labels.index(chosen_match_lbl)]

                if selected_target_lbl == "➕ Adicionar como nova linha no final":
                    add_c1, add_c2, add_c3 = st.columns([1, 1, 2.5])
                    new_wbs_in = add_c1.text_input("Novo Item (WBS)", value=suggest_next_wbs(current_recs))
                    new_qty_in = add_c2.number_input("Quantidade", min_value=0.0, value=1.0, step=1.0, format="%.2f")
                    if add_c3.button("➕ Inserir Novo Serviço na Tabela", type="primary", width="stretch"):
                        st.session_state.budget_records.append(
                            {
                                "Item_WBS": new_wbs_in.strip(),
                                "Codigo": chosen_rec["Codigo"],
                                "Descricao": chosen_rec["Descricao"],
                                "Unidade": chosen_rec["Unidade"],
                                "Quantidade": float(new_qty_in),
                                "Preco_Unitario": float(chosen_rec[regime_col]),
                                "Aplicar_Pct": True,
                                "price_overridden": False,
                                "desc_from_base": True,
                                "unid_from_base": True,
                            }
                        )
                        st.session_state.budget_df = records_to_editor_df(st.session_state.budget_records)
                        st.session_state.editor_version += 1
                        st.rerun()
                else:
                    target_pos = [lbl for _, lbl in leaf_targets].index(selected_target_lbl)
                    target_row_idx = leaf_targets[target_pos][0]

                    act_c1, act_c2 = st.columns(2)
                    if act_c1.button(
                        f"💰 Aplicar Código [{chosen_rec['Codigo']}] + Preço (manter descrição atual)",
                        type="primary",
                        width="stretch",
                    ):
                        rec_t = st.session_state.budget_records[target_row_idx]
                        rec_t["Codigo"] = chosen_rec["Codigo"]
                        rec_t["Preco_Unitario"] = float(chosen_rec[regime_col])
                        rec_t["price_overridden"] = False
                        if not rec_t["Descricao"]:
                            rec_t["Descricao"] = chosen_rec["Descricao"]
                            rec_t["desc_from_base"] = True
                        if not rec_t["Unidade"]:
                            rec_t["Unidade"] = chosen_rec["Unidade"]
                            rec_t["unid_from_base"] = True
                        st.session_state.budget_df = records_to_editor_df(st.session_state.budget_records)
                        st.session_state.editor_version += 1
                        st.rerun()

                    if act_c2.button(
                        f"📋 Aplicar Completo (Código + Descrição Oficial + Medida + Preço)",
                        width="stretch",
                    ):
                        rec_t = st.session_state.budget_records[target_row_idx]
                        rec_t["Codigo"] = chosen_rec["Codigo"]
                        rec_t["Descricao"] = chosen_rec["Descricao"]
                        rec_t["Unidade"] = chosen_rec["Unidade"]
                        rec_t["Preco_Unitario"] = float(chosen_rec[regime_col])
                        rec_t["price_overridden"] = False
                        rec_t["desc_from_base"] = True
                        rec_t["unid_from_base"] = True
                        st.session_state.budget_df = records_to_editor_df(st.session_state.budget_records)
                        st.session_state.editor_version += 1
                        st.rerun()

    # Barra rápida de ações em lote sobre a tabela
    tb1, tb2, tb3, tb4 = st.columns([1.2, 1.2, 1.5, 2.1])
    if tb1.button("☑️ Marcar todos (%?)", width="stretch"):
        for r in st.session_state.budget_records:
            r["Aplicar_Pct"] = True
        st.session_state.modo_aplicacao_pct = MODO_PCT_MARCADOS
        st.session_state.budget_df = records_to_editor_df(st.session_state.budget_records)
        st.session_state.editor_version += 1
        st.rerun()

    if tb2.button("⬜ Desmarcar todos (%?)", width="stretch"):
        for r in st.session_state.budget_records:
            r["Aplicar_Pct"] = False
        st.session_state.modo_aplicacao_pct = MODO_PCT_MARCADOS
        st.session_state.budget_df = records_to_editor_df(st.session_state.budget_records)
        st.session_state.editor_version += 1
        st.rerun()

    if tb3.button("🏛️ Marcar só SINAPI/SEINFRA (%?)", width="stretch"):
        for r in st.session_state.budget_records:
            r["Aplicar_Pct"] = bool(r["Codigo"] and r["Codigo"] in catalog_lookup)
        st.session_state.modo_aplicacao_pct = MODO_PCT_MARCADOS
        st.session_state.budget_df = records_to_editor_df(st.session_state.budget_records)
        st.session_state.editor_version += 1
        st.rerun()

    if tb4.button("🔄 Restaurar Preços da Base (desfazer sobrescrita manual)", width="stretch"):
        for r in st.session_state.budget_records:
            cat_r = catalog_lookup.get(r["Codigo"]) if r["Codigo"] else None
            if cat_r is not None:
                r["Preco_Unitario"] = float(cat_r[regime_col])
                r["price_overridden"] = False
        st.session_state.budget_df = records_to_editor_df(st.session_state.budget_records)
        st.session_state.editor_version += 1
        st.rerun()

    edited_df = st.data_editor(
        st.session_state.budget_df,
        key=f"budget_editor_{st.session_state.editor_version}",
        num_rows="dynamic",
        width="stretch",
        hide_index=True,
        column_config={
            "Item_WBS": st.column_config.TextColumn("Item (WBS)", required=True, width="small"),
            "Codigo": st.column_config.TextColumn(
                "Código (SINAPI/SEINFRA)",
                width="small",
                help="Digite um código SINAPI (ex: 103328) ou SEINFRA (ex: C2820) para puxar o valor da base automaticamente.",
            ),
            "Descricao": st.column_config.TextColumn(
                "Descrição do Serviço",
                width="large",
                help="Se deixada em branco ao digitar um código, será preenchida automaticamente pela base.",
            ),
            "Unidade": st.column_config.TextColumn("Medida", width="small"),
            "Quantidade": st.column_config.NumberColumn("Quantidade", min_value=0.0, format="%.2f"),
            "Preco_Unitario": st.column_config.NumberColumn(
                "Preço Unit. Base (R$)",
                min_value=0.0,
                format="R$ %.2f",
                help="Puxado automaticamente ao inserir código SINAPI/SEINFRA. Se você editar este valor manualmente, sua edição sobrescreve o preço da base!",
            ),
            "Aplicar_Pct": st.column_config.CheckboxColumn(
                "%?",
                width="small",
                help="Marque/desmarque os itens que devem receber o Acréscimo Percentual (%).",
            ),
        },
    )

    # Reconcilia alterações de código vs. sobrescritas manuais de preço
    updated_records, needs_refresh, pct_toggled = reconcile_edited_df(
        edited_df,
        st.session_state.budget_records,
        catalog_lookup,
        regime_col=regime_col,
        regime_changed=regime_changed,
    )
    st.session_state.budget_records = updated_records
    st.session_state.budget_df = records_to_editor_df(updated_records)

    if pct_toggled and st.session_state.modo_aplicacao_pct == MODO_PCT_TODOS:
        st.session_state.modo_aplicacao_pct = MODO_PCT_MARCADOS
        modo_aplicacao_pct = MODO_PCT_MARCADOS
        st.toast("Seletor de % ajustado para 'Apenas itens marcados na tabela (%?)'!", icon="☑️")
        st.rerun()

    if needs_refresh:
        st.session_state.editor_version += 1
        st.rerun()

    # --------------------------------------------------------------------------
    # SEÇÃO 3: PROCESSAMENTO, PRÉVIA E EXPORTAÇÃO EXCEL
    # --------------------------------------------------------------------------
    df_calc, valor_global, missing_codes, invalid_wbs = process_budget_wbs(
        st.session_state.budget_records,
        catalog_lookup,
        acrescimo_percentual=acrescimo_percentual,
        modo_aplicacao_pct=modo_aplicacao_pct,
        regime_preco=regime_col,
        calc_mode=calc_mode,
    )

    if invalid_wbs:
        st.warning(f"⚠️ Códigos WBS fora do padrão numérico (ex: 1.1.2): {', '.join(invalid_wbs)}")

    if missing_codes:
        st.warning(f"⚠️ Códigos não encontrados nas bases SINAPI-CE / SEINFRA-CE: **{', '.join(missing_codes)}**")

    st.divider()
    col_prev, col_metric = st.columns([3, 1])
    with col_prev:
        if acrescimo_percentual != 0:
            status_pct = f"(Ajuste {acrescimo_percentual:+.2f}% — {modo_aplicacao_pct})"
        else:
            status_pct = "(Valores In Natura)"
        modo_tag = "TRUNC" if calc_mode == "trunc" else "ROUND"
        st.subheader(f"📊 3. Prévia do Orçamento {status_pct} [{modo_tag}]")
        st.caption(f"**Valor por extenso (Célula F do Rodapé):** ({valor_por_extenso(valor_global)})")
    with col_metric:
        st.metric(label="VALOR GLOBAL DA OBRA", value=format_brl(valor_global))

    if not df_calc.empty:
        preview_df = df_calc[
            [
                "Item_WBS",
                "Origem",
                "Codigo",
                "Descricao_Final",
                "Unidade_Final",
                "Quantidade_Final",
                "Pct_Aplicado",
                "Preco_Unitario_Final",
                "Valor_Total",
            ]
        ].copy()

        st.dataframe(
            preview_df,
            width="stretch",
            hide_index=True,
            column_config={
                "Item_WBS": st.column_config.TextColumn("ITEM"),
                "Origem": st.column_config.TextColumn("ORIGEM"),
                "Codigo": st.column_config.TextColumn("CÓDIGO"),
                "Descricao_Final": st.column_config.TextColumn("SERVIÇO / DESCRIÇÃO", width="large"),
                "Unidade_Final": st.column_config.TextColumn("MEDIDA"),
                "Quantidade_Final": st.column_config.NumberColumn("QUANTIDADE", format="%.2f"),
                "Pct_Aplicado": st.column_config.TextColumn("% APLIC."),
                "Preco_Unitario_Final": st.column_config.NumberColumn("UNITÁRIO FINAL (R$)", format="R$ %.2f"),
                "Valor_Total": st.column_config.NumberColumn("VALOR TOTAL (R$)", format="R$ %.2f"),
            },
        )

        dl_col1, dl_col2 = st.columns([2, 1.3])

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
                with dl_col1:
                    st.download_button(
                        label="📥 Baixar Proposta Oficial Mais Engenharia (.xlsx)",
                        data=excel_buffer,
                        file_name=nome_arquivo,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        width="stretch",
                    )
            except Exception as e:
                st.error(f"Erro ao gerar a planilha oficial Excel: {e}")

        with dl_col2:
            mem_buf = export_clean_budget_xlsx(
                df_calc,
                valor_global,
                titulo=f"Memória de Cálculo — Proposta {proposta_num} ({destinatario})",
            )
            st.download_button(
                label="📄 Baixar Memória de Cálculo Detalhada (.xlsx)",
                data=mem_buf,
                file_name=f"Memoria_Calculo_{proposta_num}_{destinatario.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
            )


if __name__ == "__main__":
    main()

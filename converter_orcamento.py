"""
Utilitário de Conversão e Importação de Bases (SEINFRA / SINAPI) e Orçamentos Externos
(.xlsx, .xls, .csv, .pdf -> DataFrame padronizado e/ou planilha Excel .xlsx limpa).

Pode ser utilizado tanto importado pelo `app2.py` quanto via linha de comando (CLI):
  python converter_orcamento.py orcamento_terceiro.pdf -o orcamento_convertido.xlsx
  python converter_orcamento.py orcamento_antigo.xls -o orcamento_convertido.xlsx
  python converter_orcamento.py Planos-de-Servicos-028.xls --seinfra -o seinfra_ce.csv
"""

from __future__ import annotations

import argparse
import io
import os
import re
import unicodedata
from decimal import Decimal, InvalidOperation, ROUND_DOWN, ROUND_HALF_UP
from typing import Any, Dict, List, Optional, Tuple, Union

import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd


def strip_accents(text: str) -> str:
    """Remove acentos e normaliza texto para comparação insensível a acentuação."""
    if not text:
        return ""
    nfkd = unicodedata.normalize("NFKD", str(text))
    return "".join(ch for ch in nfkd if not unicodedata.combining(ch))


def parse_br_number(val: Any) -> Optional[float]:
    """
    Converte número em formato brasileiro ('1.234,56', 'R$ 45,20') ou americano/float
    para float Python. Retorna None caso não seja numérico (ex: 'DIMENSIONAR', '').
    """
    if val is None:
        return None
    if isinstance(val, bool):
        return None
    if isinstance(val, (int, float, np.number)):
        if pd.isna(val):
            return None
        return float(val)

    s = str(val).strip()
    if s == "" or s.lower() in ("none", "nan", "-", "--", "n/a"):
        return None

    s = re.sub(r"[R$\s\u00a0]", "", s, flags=re.IGNORECASE)
    # Verifica se contém letras (ex: 'DIMENSIONAR')
    if re.search(r"[A-Za-zÀ-ÿ]", s):
        return None

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")

    try:
        return float(s)
    except ValueError:
        return None


def quantize_money(val: Union[float, Decimal, str], mode: str = "trunc") -> float:
    """
    Trunca ('trunc', igual a =TRUNC(x,2) no Excel) ou arredonda ('round', ABNT/padrão)
    um valor monetário para exatamente 2 casas decimais usando aritmética Decimal exata.
    """
    try:
        d = val if isinstance(val, Decimal) else Decimal(str(val))
    except (InvalidOperation, ValueError):
        return 0.0
    rounding_rule = ROUND_DOWN if mode == "trunc" else ROUND_HALF_UP
    return float(d.quantize(Decimal("0.01"), rounding=rounding_rule))


# ==============================================================================
# 1. CONVERSOR DA TABELA SEINFRA (.xls / .xlsx / .csv -> CSV PADRONIZADO)
# ==============================================================================
def parse_seinfra_file(file_source: Union[str, bytes, io.BytesIO], filename: str = "") -> pd.DataFrame:
    """
    Lê uma planilha oficial da SEINFRA-CE (.xls, .xlsx ou .csv) — ex: Tabela Unificada 028 —
    e extrai todas as composições/insumos válidos (ex: C2820, I0001) em formato padronizado.
    """
    ext = os.path.splitext(filename.lower())[1] if filename else ""
    if not ext and isinstance(file_source, str):
        ext = os.path.splitext(file_source.lower())[1]

    if ext == ".csv":
        if isinstance(file_source, bytes):
            buf = io.BytesIO(file_source)
        else:
            buf = file_source
        df_csv = pd.read_csv(buf, dtype=str, encoding="utf-8-sig")
        df_csv.columns = [str(c).replace("\ufeff", "").strip() for c in df_csv.columns]
        return _normalize_seinfra_df(df_csv)

    # Leitura de .xls (xlrd) ou .xlsx (openpyxl)
    raw_rows: List[List[Any]] = []
    header_meta_text = ""

    if ext == ".xls":
        import xlrd

        if isinstance(file_source, str):
            wb = xlrd.open_workbook(file_source)
        else:
            content = file_source.getvalue() if isinstance(file_source, io.BytesIO) else file_source
            wb = xlrd.open_workbook(file_contents=content)

        for sh_name in wb.sheet_names():
            sh = wb.sheet_by_name(sh_name)
            for r in range(min(5, sh.nrows)):
                header_meta_text += " " + " ".join(str(sh.cell_value(r, c)) for c in range(sh.ncols))
            for r in range(sh.nrows):
                raw_rows.append([sh.cell_value(r, c) for c in range(sh.ncols)])
    else:
        wb = openpyxl.load_workbook(file_source, data_only=True)
        for ws in wb.worksheets:
            for r in range(1, min(6, ws.max_row + 1)):
                header_meta_text += " " + " ".join(str(ws.cell(r, c).value or "") for c in range(1, ws.max_column + 1))
            for row in ws.iter_rows(values_only=True):
                raw_rows.append(list(row))

    # Localizar cabeçalho da SEINFRA (ITEM | CÓDIGO | DESCRIÇÃO | ... | UNIDADE | PREÇO UNITÁRIO)
    col_codigo = 1
    col_desc = 2
    col_unid = 5
    col_preco = 6
    start_idx = 0

    for idx, row in enumerate(raw_rows[:30]):
        norm_cells = [strip_accents(str(c or "")).upper().strip() for c in row]
        if any("CODIGO" in c for c in norm_cells) and any("DESCRICAO" in c for c in norm_cells):
            for c_i, val in enumerate(norm_cells):
                if "CODIGO" in val:
                    col_codigo = c_i
                elif "DESCRICAO" in val:
                    col_desc = c_i
                elif "UNID" in val:
                    col_unid = c_i
                elif "PRECO" in val or "UNIT" in val or "VALOR" in val:
                    col_preco = c_i
            start_idx = idx + 1
            break

    records = []
    for row in raw_rows[start_idx:]:
        if len(row) <= max(col_codigo, col_desc, col_unid, col_preco):
            continue
        cod = str(row[col_codigo] or "").strip().upper()
        if cod.endswith(".0"):
            cod = cod[:-2]
        desc = str(row[col_desc] or "").strip()
        unid = str(row[col_unid] or "").strip().upper()
        preco = parse_br_number(row[col_preco])

        # Na tabela SEINFRA, linhas de grupo têm título na coluna CÓDIGO e DESCRIÇÃO/UNIDADE vazias
        if not cod or not desc or not unid or preco is None:
            continue
        if cod.upper() in ("CODIGO", "CÓDIGO"):
            continue

        records.append(
            {
                "Codigo_SEINFRA": cod,
                "Descricao_SEINFRA": re.sub(r"\s+", " ", desc),
                "Unidade": unid,
                "Preco_Unitario": round(float(preco), 2),
                "Preco_Sem_Desoneracao": round(float(preco), 2),
                "Preco_Com_Desoneracao": round(float(preco), 2),
            }
        )

    df_out = pd.DataFrame(records)
    if not df_out.empty:
        df_out = df_out.drop_duplicates(subset=["Codigo_SEINFRA"], keep="first").reset_index(drop=True)
    return df_out


def _normalize_seinfra_df(df: pd.DataFrame) -> pd.DataFrame:
    col_map = {}
    for col in df.columns:
        c_norm = strip_accents(col).lower()
        if "cod" in c_norm:
            col_map[col] = "Codigo_SEINFRA"
        elif "desc" in c_norm:
            col_map[col] = "Descricao_SEINFRA"
        elif "unid" in c_norm or "medida" in c_norm:
            col_map[col] = "Unidade"
        elif "sem" in c_norm and "deson" in c_norm:
            col_map[col] = "Preco_Sem_Desoneracao"
        elif "com" in c_norm and "deson" in c_norm:
            col_map[col] = "Preco_Com_Desoneracao"
        elif "prec" in c_norm or "unit" in c_norm or "valor" in c_norm:
            col_map[col] = "Preco_Unitario"
    df = df.rename(columns=col_map)
    df["Codigo_SEINFRA"] = df["Codigo_SEINFRA"].fillna("").astype(str).str.strip().str.upper()
    df["Descricao_SEINFRA"] = df["Descricao_SEINFRA"].fillna("").astype(str).str.strip()
    df["Unidade"] = df["Unidade"].fillna("").astype(str).str.strip().str.upper()
    df["Preco_Unitario"] = df["Preco_Unitario"].apply(lambda v: parse_br_number(v) or 0.0)
    if "Preco_Sem_Desoneracao" in df.columns:
        df["Preco_Sem_Desoneracao"] = df["Preco_Sem_Desoneracao"].apply(lambda v: parse_br_number(v) or 0.0)
    else:
        df["Preco_Sem_Desoneracao"] = df["Preco_Unitario"]
    if "Preco_Com_Desoneracao" in df.columns:
        df["Preco_Com_Desoneracao"] = df["Preco_Com_Desoneracao"].apply(lambda v: parse_br_number(v) or 0.0)
    else:
        df["Preco_Com_Desoneracao"] = df["Preco_Unitario"]
    return df[df["Codigo_SEINFRA"] != ""].drop_duplicates(subset=["Codigo_SEINFRA"], keep="first").reset_index(drop=True)


# ==============================================================================
# 2. LEITURA E DETECÇÃO INTELIGENTE DE ORÇAMENTOS DE TERCEIROS (.xlsx, .xls, .csv, .pdf)
# ==============================================================================
def _score_header_row(cells: List[str]) -> int:
    """Pontua uma linha candidata a cabeçalho de orçamento."""
    joined = [strip_accents(c).lower().strip() for c in cells if c and str(c).strip()]
    if len(joined) < 3:
        return 0
    score = 0
    if any(re.search(r"\b(item|it\.|eap|wbs|seq)\b", c) for c in joined):
        score += 3
    if any(re.search(r"(descri|servic|discrimina|especifica|escopo)", c) for c in joined):
        score += 4
    if any(re.search(r"\b(unid|und|un\.|medida)\b", c) for c in joined):
        score += 3
    if any(re.search(r"(quant|qtd|qtde)", c) for c in joined):
        score += 3
    if any(re.search(r"(unit|p\.?\s*u\.?|preco|valor)", c) for c in joined):
        score += 3
    if any("total" in c for c in joined):
        score += 2
    return score


def load_raw_sheets_from_file(
    file_source: Union[str, bytes, io.BytesIO],
    filename: str,
) -> Dict[str, List[List[Any]]]:
    """
    Carrega qualquer arquivo de orçamento (.xlsx, .xls, .csv, .pdf) retornando um
    dicionário {nome_da_aba: matriz_de_linhas}.
    """
    ext = os.path.splitext(filename.lower())[1]
    content_bytes: Optional[bytes] = None
    if isinstance(file_source, bytes):
        content_bytes = file_source
    elif isinstance(file_source, io.BytesIO):
        content_bytes = file_source.getvalue()

    sheets: Dict[str, List[List[Any]]] = {}

    if ext == ".xlsx":
        src = io.BytesIO(content_bytes) if content_bytes is not None else file_source
        wb = openpyxl.load_workbook(src, data_only=True)
        for ws in wb.worksheets:
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            if rows:
                sheets[ws.title] = rows

    elif ext == ".xls":
        import xlrd

        if content_bytes is not None:
            wb = xlrd.open_workbook(file_contents=content_bytes)
        else:
            wb = xlrd.open_workbook(str(file_source))
        for name in wb.sheet_names():
            sh = wb.sheet_by_name(name)
            rows = [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)]
            if rows:
                sheets[name] = rows

    elif ext == ".csv":
        raw_data = content_bytes if content_bytes is not None else open(str(file_source), "rb").read()
        df_csv = None
        for enc in ("utf-8-sig", "utf-8", "latin1", "cp1252"):
            try:
                df_csv = pd.read_csv(io.BytesIO(raw_data), header=None, dtype=object, encoding=enc, sep=None, engine="python")
                break
            except Exception:
                continue
        if df_csv is not None:
            sheets["CSV"] = df_csv.where(pd.notna(df_csv), None).values.tolist()

    elif ext == ".pdf":
        import pdfplumber

        src = io.BytesIO(content_bytes) if content_bytes is not None else str(file_source)
        table_rows: List[List[Any]] = []
        text_rows: List[List[Any]] = []

        with pdfplumber.open(src) as pdf:
            for page in pdf.pages:
                extracted = page.extract_tables()
                for tb in extracted:
                    for r in tb:
                        if r and any(cell and str(cell).strip() for cell in r):
                            table_rows.append([str(cell).replace("\n", " ").strip() if cell else None for cell in r])

                # Fallback para linhas em texto corrido caso o PDF não tenha grade de tabela
                txt = page.extract_text() or ""
                for line in txt.splitlines():
                    line_s = line.strip()
                    if not line_s:
                        continue
                    # Tenta capturar linha folha: ITEM | DESC | UNID | QTD | UNIT | TOTAL
                    m_leaf = re.match(
                        r"^(\d+(?:\.\d+)*)\s+(.+?)\s+([A-Za-z²³%/]+)\s+([\d.,]+)\s+(?:R\$\s*)?([\d.,]+)\s+(?:R\$\s*)?([\d.,]+)$",
                        line_s,
                    )
                    if m_leaf:
                        text_rows.append(list(m_leaf.groups()))
                        continue
                    # Tenta capturar linha de grupo: ITEM | DESC | TOTAL
                    m_grp = re.match(r"^(\d+(?:\.\d+)*)\s+(.+?)\s+(?:R\$\s*)?([\d.,]+)$", line_s)
                    if m_grp:
                        text_rows.append([m_grp.group(1), m_grp.group(2), None, None, None, m_grp.group(3)])

        if table_rows:
            sheets["Tabelas PDF"] = table_rows
        if text_rows:
            header = [["ITEM", "DESCRIÇÃO", "UNIDADE", "QUANTIDADE", "VALOR UNITÁRIO", "VALOR TOTAL"]]
            sheets["Texto PDF"] = header + text_rows

    return sheets


def auto_detect_sheet_and_columns(
    sheets: Dict[str, List[List[Any]]],
    preferred_sheet: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Identifica automaticamente a melhor aba, a linha de cabeçalho e o mapeamento ideal
    de colunas (priorizando QUANTIDADE PROPONENTE quando houver múltiplas quantidades).
    """
    best_sheet = preferred_sheet if preferred_sheet in sheets else None
    best_header_idx = 0
    best_score = -1

    search_sheets = [best_sheet] if best_sheet else list(sheets.keys())
    for sh_name in search_sheets:
        rows = sheets[sh_name]
        for r_idx in range(min(45, len(rows))):
            row_strs = [str(c or "").strip() for c in rows[r_idx]]
            sc = _score_header_row(row_strs)
            if sc > best_score:
                best_score = sc
                best_sheet = sh_name
                best_header_idx = r_idx

    if not best_sheet:
        best_sheet = next(iter(sheets.keys()))
        best_header_idx = 0

    rows = sheets[best_sheet]
    header_row = rows[best_header_idx] if best_header_idx < len(rows) else []
    prev_row = rows[best_header_idx - 1] if best_header_idx > 0 else []
    max_cols = max((len(r) for r in rows[: min(len(rows), best_header_idx + 50)]), default=len(header_row))

    # Constrói rótulos de colunas únicos combinando cabeçalho mesclado superior se necessário
    col_labels: List[str] = []
    carry_prev = ""
    for c_idx in range(max_cols):
        p_val = str(prev_row[c_idx]).strip() if c_idx < len(prev_row) and prev_row[c_idx] is not None else ""
        if p_val:
            carry_prev = p_val
        h_val = str(header_row[c_idx]).strip() if c_idx < len(header_row) and header_row[c_idx] is not None else ""
        h_val = re.sub(r"\s+", " ", h_val)

        if h_val:
            # Se a linha de cima for ex: 'QUANTIDADE' e a de baixo 'PROPONENTE', combina
            if carry_prev and strip_accents(carry_prev).lower() in ("quantidade", "preco", "valor") and strip_accents(carry_prev).lower() not in strip_accents(h_val).lower():
                label = f"{carry_prev} {h_val}"
            else:
                label = h_val
        else:
            label = f"Coluna {c_idx + 1}"

        col_letter = openpyxl.utils.get_column_letter(c_idx + 1)
        col_labels.append(f"[{col_letter}] {label}")

    # Mapeamento inteligente de colunas
    item_col = None
    codigo_col = None
    desc_col = None
    unid_col = None
    qty_candidates: List[Tuple[int, int, str]] = []
    unit_price_candidates: List[Tuple[int, int, str]] = []
    total_col = None

    data_sample = rows[best_header_idx + 1 : best_header_idx + 80]

    for c_idx, lbl in enumerate(col_labels):
        norm = strip_accents(lbl).lower()
        # remove prefixo '[a] '
        norm_clean = re.sub(r"^\[[a-z]+\]\s*", "", norm)

        if item_col is None and re.search(r"\b(item|it\.|eap|wbs|nivel)\b", norm_clean):
            item_col = c_idx
        elif codigo_col is None and re.search(r"\b(codigo|cod\.|sinapi|seinfra)\b", norm_clean):
            codigo_col = c_idx
        elif desc_col is None and re.search(r"(descri|servic|discrimina|especifica|escopo)", norm_clean):
            desc_col = c_idx
        elif unid_col is None and re.search(r"\b(unid|und|un\.|medida)\b", norm_clean):
            unid_col = c_idx
        elif re.search(r"(quant|qtd|qtde)", norm_clean):
            # Priorizar QUANTIDADE PROPONENTE / ORÇADA e com mais números válidos
            num_valid = sum(1 for r in data_sample if c_idx < len(r) and parse_br_number(r[c_idx]) is not None)
            prio = 10
            if any(k in norm_clean for k in ("proponente", "orcad", "contratad", "final", "execut")):
                prio += 50
            if any(k in norm_clean for k in ("solar", "referencia", "cliente", "projeto")):
                prio -= 20
            qty_candidates.append((prio, num_valid, lbl))
        elif "total" in norm_clean and "unit" not in norm_clean:
            if total_col is None:
                total_col = c_idx
        elif any(k in norm_clean for k in ("unit", "preco", "valor")) and "total" not in norm_clean:
            prio = 10
            if "unit" in norm_clean:
                prio += 20
            if "com bdi" in norm_clean or "c/ bdi" in norm_clean or "final" in norm_clean:
                prio += 15
            unit_price_candidates.append((prio, c_idx, lbl))

    qty_col = None
    if qty_candidates:
        qty_candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
        qty_col = col_labels.index(qty_candidates[0][2])

    price_col = None
    if unit_price_candidates:
        unit_price_candidates.sort(key=lambda x: x[0], reverse=True)
        price_col = unit_price_candidates[0][1]

    # Fallback posicional caso alguma coluna chave não tenha título claro
    if item_col is None and max_cols >= 1:
        item_col = 0
    if desc_col is None and max_cols >= 2:
        desc_col = 1
    if unid_col is None and max_cols >= 3:
        unid_col = 2
    if qty_col is None and max_cols >= 4:
        qty_col = 3
    if price_col is None and max_cols >= 5:
        price_col = 4

    return {
        "sheet_name": best_sheet,
        "header_row_idx": best_header_idx,
        "col_labels": col_labels,
        "mapping": {
            "item_col": item_col,
            "codigo_col": codigo_col,
            "desc_col": desc_col,
            "unid_col": unid_col,
            "qty_col": qty_col,
            "price_col": price_col,
            "total_col": total_col,
        },
    }


def _normalize_wbs_cell(raw_val: Any) -> str:
    """
    Converte valor de célula WBS preservando strings como '1.10' e convertendo floats
    inteiros do Excel (ex: 1.0 -> '1', 2.0 -> '2').
    """
    if raw_val is None:
        return ""
    if isinstance(raw_val, int):
        return str(raw_val)
    if isinstance(raw_val, float):
        if np.isnan(raw_val):
            return ""
        if raw_val.is_integer():
            return str(int(round(raw_val)))
        return str(raw_val).strip()

    s = str(raw_val).strip()
    # Remove ponto final solto (ex: '1.1.' -> '1.1'), mantendo apenas dígitos e pontos se for WBS
    s = re.sub(r"\.+$", "", s)
    return s


def _is_valid_wbs_code(wbs: str) -> bool:
    if not wbs:
        return False
    parts = wbs.split(".")
    return all(p.isdigit() for p in parts)


def extract_budget_from_sheet(
    rows: List[List[Any]],
    header_row_idx: int,
    mapping: Dict[str, Optional[int]],
    skip_empty_and_omissos: bool = True,
    auto_fix_wbs: bool = True,
    calc_mode: str = "trunc",
) -> Dict[str, Any]:
    """
    Extrai os itens de orçamento a partir das linhas brutas, identifica anomalias reais
    (como itens fora de hierarquia, códigos WBS duplicados, células 'DIMENSIONAR',
    grupos vazios 'OMISSOS' e linha final de TOTAL), e retorna o DataFrame pronto para o editor.
    """
    item_c = mapping.get("item_col")
    cod_c = mapping.get("codigo_col")
    desc_c = mapping.get("desc_col")
    unid_c = mapping.get("unid_col")
    qty_c = mapping.get("qty_col")
    price_c = mapping.get("price_col")
    tot_c = mapping.get("total_col")

    def get_cell(r: List[Any], idx: Optional[int]) -> Any:
        if idx is None or idx < 0 or idx >= len(r):
            return None
        return r[idx]

    extracted: List[Dict[str, Any]] = []
    warnings_list: List[str] = []
    fixes_applied: List[str] = []
    skipped_rows_info: List[str] = []
    reported_global_total: Optional[float] = None

    data_rows = rows[header_row_idx + 1 :]

    # Verifica se strings '1.0', '2.0' vindas de CSV/PDF são na verdade grupos inteiros '1', '2'
    raw_wbs_preview = [_normalize_wbs_cell(get_cell(r, item_c)) for r in data_rows]
    has_dot_zero_children = any(w.startswith(tuple(f"{k}.0." for k in range(1, 30))) for w in raw_wbs_preview if w)

    for offset, r in enumerate(data_rows):
        excel_row_num = header_row_idx + 2 + offset
        if not any(c is not None and str(c).strip() != "" for c in r):
            continue

        raw_item = get_cell(r, item_c)
        wbs = _normalize_wbs_cell(raw_item)
        if not has_dot_zero_children and re.match(r"^\d+\.0$", wbs) and isinstance(raw_item, str):
            wbs = wbs[:-2]

        raw_desc = get_cell(r, desc_c)
        desc = re.sub(r"\s+", " ", str(raw_desc).strip()) if raw_desc is not None and str(raw_desc).strip() not in ("None", "nan") else ""

        raw_unid = get_cell(r, unid_c)
        unid = str(raw_unid).strip() if raw_unid is not None and str(raw_unid).strip() not in ("None", "nan") else ""

        raw_cod = get_cell(r, cod_c)
        cod = str(raw_cod).strip().upper() if raw_cod is not None and str(raw_cod).strip() not in ("None", "nan") else ""
        if cod.endswith(".0"):
            cod = cod[:-2]

        raw_qty = get_cell(r, qty_c)
        qty_parsed = parse_br_number(raw_qty)

        raw_price = get_cell(r, price_c)
        price_parsed = parse_br_number(raw_price)

        raw_tot = get_cell(r, tot_c)
        tot_parsed = parse_br_number(raw_tot)

        # Detecta linha de Rodapé / TOTAL GERAL / OBS:
        row_text_join = " ".join(str(c or "").strip().upper() for c in r if c is not None)
        if not _is_valid_wbs_code(wbs):
            if any(k in row_text_join for k in ("TOTAL", "GLOBAL", "VALOR GERAL")) and tot_parsed is not None and tot_parsed > 0:
                reported_global_total = float(tot_parsed)
            if wbs or desc:
                skipped_rows_info.append(f"Linha {excel_row_num} ignorada (rodapé/observação): `{wbs or desc[:50]}`")
            continue

        # Se a célula de quantidade selecionada continha texto como 'DIMENSIONAR'
        if raw_qty is not None and str(raw_qty).strip() != "" and qty_parsed is None:
            warnings_list.append(
                f"Linha {excel_row_num} (Item `{wbs}`): quantidade não numérica (`{raw_qty}`) convertida para `0,00`."
            )

        qty_val = qty_parsed if qty_parsed is not None else 0.0
        price_val = price_parsed if price_parsed is not None else 0.0

        # Classificação estrutural na planilha de origem: Grupo (sem unid/qtd/preço) vs Item Folha
        is_structural_group = (unid == "" and qty_parsed is None and price_parsed is None)

        # Filtrar linhas folha completamente vazias (ex: 5.1, 5.2, 5.3, 5.4 em OMISSOS)
        if skip_empty_and_omissos and desc == "" and unid == "" and qty_val == 0.0 and price_val == 0.0:
            skipped_rows_info.append(f"Linha {excel_row_num} ignorada (item em branco): `{wbs}`")
            continue

        extracted.append(
            {
                "excel_row": excel_row_num,
                "Item_WBS": wbs,
                "Codigo": cod,
                "Descricao": desc,
                "Unidade": unid,
                "Quantidade": qty_val,
                "Preco_Unitario": price_val,
                "Sheet_Total": tot_parsed,
                "Is_Structural_Group": is_structural_group,
            }
        )

    # Remover grupos vazios sem filhos (ex: '5 OMISSOS' após remover 5.1..5.4 vazios)
    if skip_empty_and_omissos and extracted:
        cleaned: List[Dict[str, Any]] = []
        for i, item in enumerate(extracted):
            if item["Is_Structural_Group"]:
                # Verifica se há algum item folha ou subgrupo válido antes do próximo grupo de mesmo nível ou superior
                grp_depth = len(item["Item_WBS"].split("."))
                has_leaf = False
                for nxt in extracted[i + 1 :]:
                    nxt_depth = len(nxt["Item_WBS"].split("."))
                    if nxt["Is_Structural_Group"] and nxt_depth <= grp_depth:
                        break
                    if not nxt["Is_Structural_Group"]:
                        has_leaf = True
                        break
                if not has_leaf and (item["Sheet_Total"] in (None, 0.0)):
                    skipped_rows_info.append(
                        f"Linha {item['excel_row']} ignorada (grupo sem serviços/zerado): `{item['Item_WBS']} - {item['Descricao']}`"
                    )
                    continue
            cleaned.append(item)
        extracted = cleaned

    # --------------------------------------------------------------------------
    # Detecção e Correção Opcional de Anomalias de Hierarquia / Duplicidade WBS
    # --------------------------------------------------------------------------
    anomalies_detected: List[str] = []
    has_any_group = any(it["Is_Structural_Group"] for it in extracted)

    if has_any_group:
        # Agrupa itens folha sob o grupo estrutural imediatamente aberto acima deles
        current_group_idx: Optional[int] = None
        group_children: Dict[int, List[int]] = {}

        for idx, it in enumerate(extracted):
            if it["Is_Structural_Group"]:
                current_group_idx = idx
                group_children[idx] = []
            elif current_group_idx is not None:
                group_children[current_group_idx].append(idx)

        for g_idx, child_indices in group_children.items():
            if not child_indices:
                continue
            g_wbs = extracted[g_idx]["Item_WBS"]
            g_desc = extracted[g_idx]["Descricao"]
            expected_prefix = g_wbs + "."

            wrong_prefix_children = [
                c_i for c_i in child_indices if not extracted[c_i]["Item_WBS"].startswith(expected_prefix)
            ]
            seen_in_group: Dict[str, int] = {}
            dup_children: List[int] = []
            for c_i in child_indices:
                cw = extracted[c_i]["Item_WBS"]
                seen_in_group[cw] = seen_in_group.get(cw, 0) + 1
                if seen_in_group[cw] > 1:
                    dup_children.append(c_i)

            for c_i in wrong_prefix_children:
                cw = extracted[c_i]["Item_WBS"]
                cdesc = extracted[c_i]["Descricao"][:55]
                anomalies_detected.append(
                    f"Linha {extracted[c_i]['excel_row']}: item `{cw}` ({cdesc}) está dentro do grupo `{g_wbs} - {g_desc}`, fora do prefixo `{expected_prefix}`."
                )

            if dup_children:
                dups_str = ", ".join(sorted({extracted[c_i]["Item_WBS"] for c_i in dup_children}))
                anomalies_detected.append(
                    f"Grupo `{g_wbs} - {g_desc}` possui código(s) WBS duplicado(s): `{dups_str}`."
                )

            if auto_fix_wbs and (wrong_prefix_children or dup_children):
                for seq_num, c_i in enumerate(child_indices, start=1):
                    old_wbs = extracted[c_i]["Item_WBS"]
                    new_wbs = f"{g_wbs}.{seq_num}"
                    if old_wbs != new_wbs:
                        extracted[c_i]["Item_WBS"] = new_wbs
                        fixes_applied.append(
                            f"Linha {extracted[c_i]['excel_row']}: `{old_wbs}` ➔ `{new_wbs}` ({extracted[c_i]['Descricao'][:50]})"
                        )
    else:
        # Lista plana sem grupos: apenas checa duplicidades diretas
        seen_flat: Dict[str, int] = {}
        for idx, it in enumerate(extracted):
            cw = it["Item_WBS"]
            seen_flat[cw] = seen_flat.get(cw, 0) + 1
            if seen_flat[cw] > 1:
                anomalies_detected.append(f"Linha {it['excel_row']}: código WBS duplicado `{cw}`.")

    # Monta DataFrame final para o editor do app2.py
    editor_rows = []
    calc_leaf_sum = Decimal("0.00")
    for it in extracted:
        is_grp = it["Is_Structural_Group"]
        if not is_grp:
            item_tot = Decimal(str(quantize_money(Decimal(str(it["Quantidade"])) * Decimal(str(it["Preco_Unitario"])), mode=calc_mode)))
            calc_leaf_sum += item_tot

        editor_rows.append(
            {
                "Item_WBS": it["Item_WBS"],
                "Codigo": "" if is_grp else it["Codigo"],
                "Descricao": it["Descricao"],
                "Unidade": "" if is_grp else it["Unidade"],
                "Quantidade": 0.0 if is_grp else float(it["Quantidade"]),
                "Preco_Unitario": 0.0 if is_grp else float(it["Preco_Unitario"]),
                "Aplicar_Pct": False if is_grp else True,
            }
        )

    df_editor = pd.DataFrame(
        editor_rows,
        columns=["Item_WBS", "Codigo", "Descricao", "Unidade", "Quantidade", "Preco_Unitario", "Aplicar_Pct"],
    )

    # Caso não haja linha explícita de TOTAL no rodapé, tenta somar grupos raiz da planilha original
    if reported_global_total is None and extracted:
        top_groups = [it for it in extracted if it["Is_Structural_Group"] and "." not in it["Item_WBS"] and it["Sheet_Total"] is not None]
        if top_groups:
            reported_global_total = round(sum(float(it["Sheet_Total"]) for it in top_groups), 2)

    return {
        "df": df_editor,
        "calculated_total": float(calc_leaf_sum),
        "reported_total": reported_global_total,
        "anomalies": anomalies_detected,
        "fixes_applied": fixes_applied,
        "warnings": warnings_list,
        "skipped_rows": skipped_rows_info,
    }


def export_clean_budget_xlsx(df_calc: pd.DataFrame, valor_global: float, titulo: str = "Orçamento Convertido") -> io.BytesIO:
    """
    Exporta qualquer orçamento (vindo de PDF, XLS, CSV ou da tabela de edição) para uma
    planilha Excel (.xlsx) limpa, formatada e com fórmulas/valores conferidos.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Orçamento"

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    group_fill = PatternFill(start_color="EAEAEA", end_color="EAEAEA", fill_type="solid")
    total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    thin_side = Side(border_style="thin", color="BFBFBF")
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    ws.cell(1, 1, titulo).font = Font(name="Calibri", size=13, bold=True)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)

    headers = ["ITEM", "FONTE", "CÓDIGO", "DESCRIÇÃO DO SERVIÇO", "UNID.", "QUANTIDADE", "VALOR UNITÁRIO (R$)", "VALOR TOTAL (R$)"]
    for col_i, h_text in enumerate(headers, start=1):
        c = ws.cell(3, col_i, h_text)
        c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_all

    currency_fmt = '[$R$ -416]#,##0.00'
    for idx, (_, row) in enumerate(df_calc.iterrows(), start=4):
        is_parent = bool(row.get("Is_Parent", False))
        vals = [
            str(row.get("Item_WBS", "")),
            "" if is_parent else str(row.get("Origem", "")),
            "" if is_parent else str(row.get("Codigo", row.get("Codigo_SINAPI", ""))),
            str(row.get("Descricao_Final", row.get("Descricao", ""))),
            "" if is_parent else str(row.get("Unidade_Final", row.get("Unidade", ""))),
            None if is_parent else float(row.get("Quantidade_Final", row.get("Quantidade", 0.0))),
            None if is_parent else float(row.get("Preco_Unitario_Final", row.get("Preco_Unitario", 0.0))),
            float(row.get("Valor_Total", 0.0)),
        ]
        for col_i, val in enumerate(vals, start=1):
            c = ws.cell(idx, col_i, val)
            c.font = Font(name="Calibri", size=10, bold=is_parent)
            c.border = border_all
            if is_parent:
                c.fill = group_fill
            if col_i == 6 and val is not None:
                c.number_format = "#,##0.00"
            elif col_i in (7, 8) and val is not None:
                c.number_format = currency_fmt

    tot_row = 4 + len(df_calc)
    ws.cell(tot_row, 7, "VALOR GLOBAL:").font = Font(name="Calibri", size=11, bold=True)
    ws.cell(tot_row, 7).fill = total_fill
    ws.cell(tot_row, 7).border = border_all
    c_tot = ws.cell(tot_row, 8, float(valor_global))
    c_tot.font = Font(name="Calibri", size=11, bold=True)
    c_tot.number_format = currency_fmt
    c_tot.fill = total_fill
    c_tot.border = border_all

    col_widths = {1: 10, 2: 16, 3: 12, 4: 65, 5: 10, 6: 14, 7: 20, 8: 20}
    for col_i, width in col_widths.items():
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_i)].width = width

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ==============================================================================
# 3. INTERFACE DE LINHA DE COMANDO (CLI)
# ==============================================================================
def _cli_main():
    parser = argparse.ArgumentParser(
        description="Converte tabelas SEINFRA (.xls) para CSV ou orçamentos (.pdf, .xls, .xlsx, .csv) para Excel (.xlsx)."
    )
    parser.add_argument("input_file", help="Caminho do arquivo de entrada (.xls, .xlsx, .csv, .pdf)")
    parser.add_argument("-o", "--output", help="Caminho do arquivo de saída (.xlsx ou .csv)", default=None)
    parser.add_argument("--seinfra", action="store_true", help="Tratar entrada como Tabela Unificada SEINFRA e gerar CSV")
    args = parser.parse_args()

    in_path = args.input_file
    if not os.path.exists(in_path):
        raise SystemExit(f"Arquivo não encontrado: {in_path}")

    if args.seinfra:
        out_path = args.output or "seinfra_ce.csv"
        df_seinfra = parse_seinfra_file(in_path, filename=in_path)
        df_seinfra.to_csv(out_path, index=False, encoding="utf-8-sig")
        print(f"Base SEINFRA salva com sucesso em '{out_path}' ({len(df_seinfra)} composições/insumos).")
        return

    sheets = load_raw_sheets_from_file(in_path, filename=in_path)
    if not sheets:
        raise SystemExit("Nenhuma tabela encontrada no arquivo informado.")

    detected = auto_detect_sheet_and_columns(sheets)
    res = extract_budget_from_sheet(
        sheets[detected["sheet_name"]],
        detected["header_row_idx"],
        detected["mapping"],
        skip_empty_and_omissos=True,
        auto_fix_wbs=True,
        calc_mode="trunc",
    )
    df_ed = res["df"].copy()
    all_wbs = df_ed["Item_WBS"].tolist()
    df_ed["Is_Parent"] = [any(o.startswith(w + ".") for o in all_wbs if o != w) for w in all_wbs]
    df_ed["Origem"] = np.where(df_ed["Is_Parent"], "Grupo (EAP)", "Importado")
    df_ed["Descricao_Final"] = df_ed["Descricao"]
    df_ed["Unidade_Final"] = df_ed["Unidade"]
    df_ed["Quantidade_Final"] = np.where(df_ed["Is_Parent"], np.nan, df_ed["Quantidade"])
    df_ed["Preco_Unitario_Final"] = df_ed["Preco_Unitario"]
    df_ed["Valor_Total"] = [
        0.0 if row["Is_Parent"] else quantize_money(Decimal(str(row["Quantidade"])) * Decimal(str(row["Preco_Unitario"])), "trunc")
        for _, row in df_ed.iterrows()
    ]
    leaf_df = df_ed[~df_ed["Is_Parent"]]
    for idx, row in df_ed[df_ed["Is_Parent"]].iterrows():
        prefix = row["Item_WBS"] + "."
        df_ed.at[idx, "Valor_Total"] = round(float(leaf_df[leaf_df["Item_WBS"].str.startswith(prefix)]["Valor_Total"].sum()), 2)

    out_path = args.output or (os.path.splitext(in_path)[0] + "_convertido.xlsx")
    buf = export_clean_budget_xlsx(df_ed, res["calculated_total"], titulo=os.path.basename(in_path))
    with open(out_path, "wb") as f:
        f.write(buf.getvalue())
    print(f"Orçamento convertido com sucesso para '{out_path}' | Total Calculado: R$ {res['calculated_total']:,.2f}")


if __name__ == "__main__":
    _cli_main()

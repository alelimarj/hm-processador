# HM.py
# ------------------------------------------------------------
# HM — Processador automático de TXT (TAB) → Tabela Virtual
# Pasta monitorada: C:\base_de_dados
#
# AJUSTE (novo, homologado agora):
# - Formatar CÓD. TUSS e CÓD. PRODUTO como valores (numéricos) no Excel exportado
# - CBHPM com exatamente 2 casas decimais (já em BR) e no Excel number_format "#.##0,00"
#
# Mantido todo o processamento homologado anteriormente.
# ------------------------------------------------------------

from __future__ import annotations

from pathlib import Path
import io
import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter


# =========================
# CONFIG
# =========================
PASTA_BASE = r"C:\base_de_dados"

CABECALHOS = [
    "REGISTRO",
    "NOME DO PACIENTE",
    "ENTRADA",
    "SAÍDA",
    "TIPO DE PRODUTO",
    "CÓD. PRODUTO",
    "DESC. PRODUTO",
    "QUANTIDADE",
    "COMANDA",
    "CÓD. TUSS",
    "DESTINO",
    "DATA E HORA DO PROC.",
    "UNIDADE",
    "MÉDICO",
    "SETOR",
    "NUM. FATURA",
    "CONVÊNIO",
    "NUM. REMESSA",
    "DATA DA REMESSA",
    "VALOR DO PROC.",
    "ATO",
    "VIA DE ACESSO",
    "PORTE CIRÚRGICO",
    "ACOMODAÇÃO",
    "PORTA DE ENTRADA",
]

COLS_NUMERICAS = {"REGISTRO", "QUANTIDADE", "VALOR DO PROC."}
EXCEL_MAX_LINHAS = 1_048_576  # limite por aba


# =========================
# HELPERS
# =========================
def list_txt_files(folder: str) -> list[Path]:
    p = Path(folder)
    if not p.exists():
        return []
    return sorted(p.glob("*.txt"), key=lambda x: x.stat().st_mtime, reverse=True)


def find_tabela_excel(folder: str) -> Path | None:
    p = Path(folder)
    for ext in (".xlsx", ".xlsm", ".xls"):
        candidate = p / f"TABELA{ext}"
        if candidate.exists():
            return candidate
    return None


def read_txt_tab(filepath: Path) -> pd.DataFrame:
    df = None
    for enc in ("utf-8", "utf-8-sig", "latin-1"):
        try:
            df = pd.read_csv(
                filepath,
                sep="\t",
                dtype=str,
                encoding=enc,
                engine="python",
                header=None,
                on_bad_lines="skip",
            )
            break
        except Exception:
            df = None

    if df is None:
        raise RuntimeError("Não foi possível ler o TXT (UTF-8/Latin-1).")

    if df.shape[1] == len(CABECALHOS):
        df.columns = CABECALHOS

    df = df.fillna("").applymap(lambda x: x.strip() if isinstance(x, str) else x)
    return df


def br_to_float(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    s = s.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
    s = s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce")


def aplicar_tipos(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in COLS_NUMERICAS:
        if col in out.columns:
            out[col] = br_to_float(out[col])

    if "REGISTRO" in out.columns:
        out["REGISTRO"] = pd.to_numeric(out["REGISTRO"], errors="coerce").astype("Int64")
    return out


def _money_any_to_float(val) -> float | None:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip()
    if s == "" or s.lower() in ("nan", "none"):
        return None
    s = s.replace(" ", "")

    if "," in s and "." in s:
        last_comma = s.rfind(",")
        last_dot = s.rfind(".")
        if last_comma > last_dot:
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        if "," in s and "." not in s:
            s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None


def _fmt_br_money(x: float) -> str:
    return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _to_float_series_any(s: pd.Series) -> pd.Series:
    return s.apply(_money_any_to_float).astype("Float64")


def formatar_exibicao(df: pd.DataFrame) -> pd.DataFrame:
    disp = df.copy()

    if "QUANTIDADE" in disp.columns and pd.api.types.is_numeric_dtype(disp["QUANTIDADE"]):
        disp["QUANTIDADE"] = disp["QUANTIDADE"].apply(
            lambda x: "" if pd.isna(x) else f"{int(x):,}".replace(",", ".")
        )

    if "VALOR DO PROC." in disp.columns and pd.api.types.is_numeric_dtype(disp["VALOR DO PROC."]):
        disp["VALOR DO PROC."] = disp["VALOR DO PROC."].apply(
            lambda x: "" if pd.isna(x) else _fmt_br_money(float(x))
        )

    if "CBHPM" in disp.columns:
        disp["CBHPM"] = disp["CBHPM"].apply(
            lambda v: "" if _money_any_to_float(v) is None else _fmt_br_money(_money_any_to_float(v))
        )

    money_cols = [
        "CIRURGIÃO",
        "1º AUXILIAR",
        "2º AUXILIAR",
        "3º AUXILIAR",
        "DEFLATOR",
        "VALOR DE REP. REGRA",
        "VALOR REP. SISHOP",
    ]
    for c in money_cols:
        if c in disp.columns:
            disp[c] = disp[c].apply(
                lambda v: "" if _money_any_to_float(v) is None else _fmt_br_money(_money_any_to_float(v))
            )
    return disp


def parse_entrada_dates(df: pd.DataFrame) -> pd.Series:
    if "ENTRADA" not in df.columns:
        return pd.Series([pd.NaT] * len(df))

    s = df["ENTRADA"].astype(str).str.strip()
    s = s.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})

    dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
    if dt.isna().all():
        m = s.str.extract(r"(?P<data>\b\d{1,2}/\d{1,2}/\d{2,4}\b)")["data"]
        dt = pd.to_datetime(m, dayfirst=True, errors="coerce")
    return dt


def safe_filename_component_date(dt: pd.Timestamp) -> str:
    if pd.isna(dt):
        return "data_invalida"
    return dt.strftime("%d-%m-%y")


def build_hm_filename(min_dt: pd.Timestamp, max_dt: pd.Timestamp) -> str:
    return f"{safe_filename_component_date(min_dt)}_a_{safe_filename_component_date(max_dt)}_HM.txt"


def ensure_unique_path(target: Path) -> Path:
    if not target.exists():
        return target
    stem = target.stem
    suffix = target.suffix
    parent = target.parent
    i = 1
    while True:
        candidate = parent / f"{stem}_{i}{suffix}"
        if not candidate.exists():
            return candidate
        i += 1


def maybe_rename_txt_by_entrada(filepath: Path, df: pd.DataFrame) -> tuple[Path, str]:
    if "ENTRADA" not in df.columns:
        return filepath, "Sem coluna ENTRADA — não renomeado."

    dt = parse_entrada_dates(df)
    min_dt = dt.min()
    max_dt = dt.max()

    if pd.isna(min_dt) or pd.isna(max_dt):
        return filepath, "ENTRADA inválida/ausente — não renomeado."

    novo_nome = build_hm_filename(min_dt, max_dt)

    if filepath.name == novo_nome:
        return filepath, "Já está no padrão HM — mantido."

    alvo = ensure_unique_path(filepath.with_name(novo_nome))
    try:
        novo_path = filepath.rename(alvo)
        return novo_path, f"Renomeado para: {novo_path.name}"
    except Exception as e:
        return filepath, f"Falha ao renomear: {e}"


def compute_minmax_for_sheet(df: pd.DataFrame) -> tuple[pd.Timestamp, pd.Timestamp]:
    dt = parse_entrada_dates(df)
    return dt.min(), dt.max()


def sheet_name_from_minmax(min_dt: pd.Timestamp, max_dt: pd.Timestamp) -> str:
    return f"{safe_filename_component_date(min_dt)}_a_{safe_filename_component_date(max_dt)}"


def chunk_df_for_excel_by_max_rows(df: pd.DataFrame, max_rows: int = EXCEL_MAX_LINHAS) -> list[tuple[str, pd.DataFrame]]:
    if df.empty:
        return []
    chunks: list[tuple[str, pd.DataFrame]] = []
    total = len(df)
    start = 0
    while start < total:
        end = min(start + max_rows, total)
        chunk = df.iloc[start:end].copy()
        mn, mx = compute_minmax_for_sheet(chunk)
        sname = sheet_name_from_minmax(mn, mx)
        chunks.append((sname, chunk))
        start = end
    return chunks


def _norm_key_series(s: pd.Series) -> pd.Series:
    out = s.astype(str).str.strip()
    out = out.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
    out = out.str.replace(r"\.0$", "", regex=True)
    return out


def load_tabela_mapping(excel_path: Path) -> tuple[pd.DataFrame, str]:
    try:
        df_tab = pd.read_excel(excel_path, sheet_name="TABELA", dtype=object)
    except Exception as e:
        return pd.DataFrame(), f"Falha ao ler Excel/aba 'TABELA': {e}"

    if df_tab.empty:
        return pd.DataFrame(), "Aba 'TABELA' está vazia."

    cols_lower = {str(c).strip().lower(): c for c in df_tab.columns}
    key_col = None
    for candidate in ("id do procedimento", "id procedimento", "id_procedimento", "procedimento id"):
        if candidate in cols_lower:
            key_col = cols_lower[candidate]
            break
    if key_col is None:
        key_col = df_tab.columns[0]

    if df_tab.shape[1] < 16:
        return pd.DataFrame(), (
            f"Aba 'TABELA' tem {df_tab.shape[1]} colunas — preciso de pelo menos 16 "
            "(para pegar colunas i, K e p)."
        )

    col_i = df_tab.columns[8]    # i
    col_k = df_tab.columns[10]   # K
    col_p = df_tab.columns[15]   # p

    mapping = df_tab[[key_col, col_i, col_k, col_p]].copy()
    mapping.columns = ["ID do Procedimento", "PORTE", "QTD_AUX_TABELA", "CBHPM"]

    mapping["ID do Procedimento"] = _norm_key_series(mapping["ID do Procedimento"])
    mapping["PORTE"] = mapping["PORTE"].astype(str).str.strip().replace({"nan": "", "None": ""})
    mapping["CBHPM"] = mapping["CBHPM"].astype(str).str.strip().replace({"nan": "", "None": ""})
    mapping["QTD_AUX_TABELA"] = pd.to_numeric(mapping["QTD_AUX_TABELA"], errors="coerce").fillna(0).astype("Int64")

    mapping = mapping.dropna(subset=["ID do Procedimento"])
    mapping = mapping[mapping["ID do Procedimento"].astype(str).str.len() > 0]
    mapping = mapping.drop_duplicates(subset=["ID do Procedimento"], keep="last")

    return mapping, f"OK: mapeamento carregado ({len(mapping):,} chaves)".replace(",", ".")


def load_tabela_sheet_df(excel_path: Path) -> pd.DataFrame:
    return pd.read_excel(excel_path, sheet_name="TABELA", dtype=object)


def enrich_with_tabela(df_final: pd.DataFrame, mapping: pd.DataFrame) -> pd.DataFrame:
    out = df_final.copy()

    if "PORTE" not in out.columns:
        out["PORTE"] = ""
    if "CBHPM" not in out.columns:
        out["CBHPM"] = ""
    if "QTD_AUX_TABELA" not in out.columns:
        out["QTD_AUX_TABELA"] = pd.Series([0] * len(out), dtype="Int64")

    if out.empty or mapping.empty:
        return out
    if "CÓD. TUSS" not in out.columns:
        return out

    out["_KEY_TUSS"] = _norm_key_series(out["CÓD. TUSS"])
    map2 = mapping.rename(columns={"ID do Procedimento": "_KEY_TUSS"}).copy()

    out = out.merge(
        map2[["_KEY_TUSS", "PORTE", "QTD_AUX_TABELA", "CBHPM"]],
        on="_KEY_TUSS",
        how="left",
        suffixes=("", "_m"),
    )

    if "PORTE_m" in out.columns:
        out["PORTE"] = out["PORTE"].astype(str).str.strip()
        out["PORTE_m"] = out["PORTE_m"].fillna("").astype(str).str.strip()
        out["PORTE"] = out["PORTE"].where(out["PORTE"] != "", out["PORTE_m"])
        out.drop(columns=["PORTE_m"], inplace=True, errors="ignore")

    if "CBHPM_m" in out.columns:
        out["CBHPM"] = out["CBHPM"].astype(str).str.strip()
        out["CBHPM_m"] = out["CBHPM_m"].fillna("").astype(str).str.strip()
        out["CBHPM"] = out["CBHPM"].where(out["CBHPM"] != "", out["CBHPM_m"])
        out.drop(columns=["CBHPM_m"], inplace=True, errors="ignore")

    if "QTD_AUX_TABELA_m" in out.columns:
        out["QTD_AUX_TABELA_m"] = pd.to_numeric(out["QTD_AUX_TABELA_m"], errors="coerce")
        out["QTD_AUX_TABELA"] = pd.to_numeric(out["QTD_AUX_TABELA"], errors="coerce")
        out["QTD_AUX_TABELA"] = out["QTD_AUX_TABELA"].where(out["QTD_AUX_TABELA"].notna(), out["QTD_AUX_TABELA_m"])
        out["QTD_AUX_TABELA"] = out["QTD_AUX_TABELA"].fillna(out["QTD_AUX_TABELA_m"])
        out.drop(columns=["QTD_AUX_TABELA_m"], inplace=True, errors="ignore")

    out["PORTE"] = out["PORTE"].fillna("")
    out["CBHPM"] = out["CBHPM"].fillna("")
    out["QTD_AUX_TABELA"] = pd.to_numeric(out["QTD_AUX_TABELA"], errors="coerce").fillna(0).astype("Int64")

    out.drop(columns=["_KEY_TUSS"], inplace=True, errors="ignore")
    return out


def ensure_col_ab_cirurgiao(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "CIRURGIÃO" not in out.columns:
        out["CIRURGIÃO"] = ""
    if "CBHPM" in out.columns:
        cols = list(out.columns)
        cols = [c for c in cols if c != "CIRURGIÃO"]
        idx = cols.index("CBHPM") + 1
        cols.insert(idx, "CIRURGIÃO")
        out = out[cols]
    return out


def ensure_cols_ac_to_ak(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    novas = [
        "1º AUXILIAR",
        "2º AUXILIAR",
        "3º AUXILIAR",
        "DEFLATOR",
        "VALOR DE REP. REGRA",
        "VALOR REP. SISHOP",
        "COMPLEMENTE",
        "DEDUÇÃO",
        "PROFISSIONAL",
    ]
    for c in novas:
        if c not in out.columns:
            out[c] = ""

    if "CIRURGIÃO" in out.columns:
        cols = list(out.columns)
        cols_sem_novas = [c for c in cols if c not in novas]
        try:
            idx = cols_sem_novas.index("CIRURGIÃO") + 1
            for offset, c in enumerate(novas):
                cols_sem_novas.insert(idx + offset, c)
            out = out[cols_sem_novas]
        except ValueError:
            pass

    return out


# =========================
# PIPELINE (homologado)
# =========================
def normalizar_cbhpm_para_numerico(df: pd.DataFrame) -> pd.DataFrame:
    """
    CBHPM numérico Float64 e arredondado em 2 casas (para manter 2 decimais).
    """
    out = df.copy()
    if "CBHPM" in out.columns:
        out["CBHPM"] = _to_float_series_any(out["CBHPM"].astype(str)).round(2)
    return out


def preencher_cirurgiao_por_formula(df: pd.DataFrame) -> pd.DataFrame:
    """
    CIRURGIÃO = CBHPM * VIA DE ACESSO * QUANTIDADE
    """
    out = df.copy()
    if "CIRURGIÃO" not in out.columns:
        out["CIRURGIÃO"] = ""
    if "CBHPM" not in out.columns or "VIA DE ACESSO" not in out.columns or "QUANTIDADE" not in out.columns:
        return out

    cbhpm_f = pd.to_numeric(out["CBHPM"], errors="coerce").fillna(0).astype("Float64")
    via_f = _to_float_series_any(out["VIA DE ACESSO"].astype(str)).fillna(0)
    qtd_f = pd.to_numeric(out["QUANTIDADE"], errors="coerce").fillna(0).astype("Float64")

    out["CIRURGIÃO"] = (cbhpm_f * via_f * qtd_f).astype("Float64")
    return out


def preencher_auxiliares_por_regra(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for c in ("1º AUXILIAR", "2º AUXILIAR", "3º AUXILIAR"):
        if c not in out.columns:
            out[c] = ""
    if "QTD_AUX_TABELA" not in out.columns or "CIRURGIÃO" not in out.columns:
        return out

    qtd = pd.to_numeric(out["QTD_AUX_TABELA"], errors="coerce").fillna(0).astype(int)
    cir = _to_float_series_any(out["CIRURGIÃO"].astype(str)).fillna(0)

    aux1 = (qtd >= 1).astype(int) * (cir * 0.3)
    aux2 = (qtd >= 2).astype(int) * (cir * 0.2)
    aux3 = (qtd >= 3).astype(int) * (aux1 * 0.1)

    out["1º AUXILIAR"] = aux1.astype("Float64")
    out["2º AUXILIAR"] = aux2.astype("Float64")
    out["3º AUXILIAR"] = aux3.astype("Float64")
    return out


def preencher_deflator_e_rep_regra(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "DEFLATOR" not in out.columns:
        out["DEFLATOR"] = ""
    if "VALOR DE REP. REGRA" not in out.columns:
        out["VALOR DE REP. REGRA"] = ""

    needed = ["CIRURGIÃO", "1º AUXILIAR", "2º AUXILIAR", "3º AUXILIAR"]
    for c in needed:
        if c not in out.columns:
            return out

    cir = _to_float_series_any(out["CIRURGIÃO"].astype(str)).fillna(0)
    a1 = _to_float_series_any(out["1º AUXILIAR"].astype(str)).fillna(0)
    a2 = _to_float_series_any(out["2º AUXILIAR"].astype(str)).fillna(0)
    a3 = _to_float_series_any(out["3º AUXILIAR"].astype(str)).fillna(0)

    deflator = (a1 + a2 + a3) * 0.20
    rep_regra = (cir + a1 + a2 + a3) - deflator

    out["DEFLATOR"] = deflator.astype("Float64")
    out["VALOR DE REP. REGRA"] = rep_regra.astype("Float64")
    return out


# =========================
# EXPORTAÇÃO EXCEL COM FÓRMULAS + FORMATAÇÕES
# =========================
def build_excel_with_formulas(
    chunks: list[tuple[str, pd.DataFrame]],
    tabela_sheet_df: pd.DataFrame | None,
) -> bytes:
    """
    Excel exportado:
      - abas de dados (chunks)
      - aba TABELA (para VLOOKUP)
      - fórmulas:
        CIRURGIÃO, 1º AUXILIAR, 2º AUXILIAR, 3º AUXILIAR, DEFLATOR, VALOR DE REP. REGRA
      - FORMATAÇÕES:
        * CBHPM: "#.##0,00" (milhar "." / decimal "," / 2 casas)
        * CÓD. TUSS e CÓD. PRODUTO: numéricos (cell.value = número) e formato "0"
    """
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, df_chunk in chunks:
            safe_name = sheet_name[:31]
            df_chunk.to_excel(writer, index=False, sheet_name=safe_name)

        if tabela_sheet_df is not None and not tabela_sheet_df.empty:
            tabela_sheet_df.to_excel(writer, index=False, sheet_name="TABELA")

        wb = writer.book

        for sheet_name, _df_chunk in chunks:
            ws = wb[sheet_name[:31]]
            max_row = ws.max_row
            if max_row < 2:
                continue

            header_to_col = {}
            for idx, cell in enumerate(ws[1], start=1):
                header_to_col[str(cell.value).strip()] = idx

            required = [
                "CBHPM",
                "VIA DE ACESSO",
                "QUANTIDADE",
                "CÓD. TUSS",
                "CÓD. PRODUTO",
                "CIRURGIÃO",
                "1º AUXILIAR",
                "2º AUXILIAR",
                "3º AUXILIAR",
                "DEFLATOR",
                "VALOR DE REP. REGRA",
            ]
            if any(c not in header_to_col for c in required):
                continue

            col_cbhpm = get_column_letter(header_to_col["CBHPM"])
            col_via = get_column_letter(header_to_col["VIA DE ACESSO"])
            col_qtd = get_column_letter(header_to_col["QUANTIDADE"])
            col_tuss = get_column_letter(header_to_col["CÓD. TUSS"])
            col_prod = get_column_letter(header_to_col["CÓD. PRODUTO"])

            col_cir = get_column_letter(header_to_col["CIRURGIÃO"])
            col_a1 = get_column_letter(header_to_col["1º AUXILIAR"])
            col_a2 = get_column_letter(header_to_col["2º AUXILIAR"])
            col_a3 = get_column_letter(header_to_col["3º AUXILIAR"])
            col_def = get_column_letter(header_to_col["DEFLATOR"])
            col_rep = get_column_letter(header_to_col["VALOR DE REP. REGRA"])

            # FORMATAÇÕES por coluna (e conversão para número em TUSS/PROD)
            for r in range(2, max_row + 1):
                # CBHPM: 2 casas e BR
                ws[f"{col_cbhpm}{r}"].number_format = "#.##0,00"

                # CÓD. TUSS e CÓD. PRODUTO: converter para número quando possível
                # (mantém vazio se não converter)
                for col in (col_tuss, col_prod):
                    cell = ws[f"{col}{r}"]
                    v = cell.value
                    if v is None:
                        continue
                    s = str(v).strip()
                    if s == "" or s.lower() in ("nan", "none"):
                        cell.value = None
                        continue
                    # remove separadores comuns e .0
                    s = s.replace(".", "").replace(",", "").strip()
                    if s.endswith(".0"):
                        s = s[:-2]
                    # tenta int
                    try:
                        num = int(float(s))
                        cell.value = num
                        cell.number_format = "0"
                    except Exception:
                        # se não der, mantém como está
                        pass

            # fórmulas por linha
            for r in range(2, max_row + 1):
                cbhpm_cell = f"{col_cbhpm}{r}"
                via_cell = f"{col_via}{r}"
                qtd_cell = f"{col_qtd}{r}"
                tuss_cell = f"{col_tuss}{r}"

                cir_cell = f"{col_cir}{r}"
                a1_cell = f"{col_a1}{r}"
                a2_cell = f"{col_a2}{r}"
                a3_cell = f"{col_a3}{r}"
                def_cell = f"{col_def}{r}"
                rep_cell = f"{col_rep}{r}"

                ws[cir_cell].value = f"={cbhpm_cell}*{via_cell}*{qtd_cell}"

                ws[a1_cell].value = f"=IF(1<=VLOOKUP({tuss_cell},TABELA!$E:$K,7,FALSE),{cir_cell}*0.3,0)"
                ws[a2_cell].value = f"=IF(2<=VLOOKUP({tuss_cell},TABELA!$E:$K,7,FALSE),{cir_cell}*0.2,0)"
                ws[a3_cell].value = f"=IF(3<=VLOOKUP({tuss_cell},TABELA!$E:$K,7,FALSE),{a1_cell}*0.1,0)"

                ws[def_cell].value = f"=({a1_cell}+{a2_cell}+{a3_cell})*0.2"
                ws[rep_cell].value = f"=({cir_cell}+{a1_cell}+{a2_cell}+{a3_cell})-{def_cell}"

    buffer.seek(0)
    return buffer.getvalue()


# =========================
# STREAMLIT UI
# =========================
st.set_page_config(page_title="HM — Processador TXT", layout="wide")
st.title("🏥 HM — Processador de TXT (TAB)")
st.caption(f"Pasta monitorada: {PASTA_BASE}")

arquivos = list_txt_files(PASTA_BASE)
if not arquivos:
    st.warning("Nenhum arquivo .txt encontrado na pasta C:\\base_de_dados")
    st.stop()

st.markdown("### 1) Normalização + Acumulado + Enriquecimento (TABELA)")

c1, c2, c3, c4 = st.columns([2, 1, 1, 2])

with c1:
    renomear_automatico = st.checkbox("Renomear automaticamente os TXT por ENTRADA", value=True)
with c2:
    mostrar_detalhes_renomeio = st.checkbox("Mostrar log", value=True)
with c3:
    usar_tabela = st.checkbox("Aplicar TABELA (PORTE/CBHPM)", value=True)

mapping = pd.DataFrame()
status_tabela = "TABELA não aplicada."
excel_tabela_path = None
tabela_sheet_df = None

if usar_tabela:
    excel_tabela_path = find_tabela_excel(PASTA_BASE)
    if excel_tabela_path is None:
        status_tabela = "Arquivo Excel 'TABELA' não encontrado em C:\\base_de_dados (TABELA.xlsx/.xlsm/.xls)."
    else:
        mapping, status_tabela = load_tabela_mapping(excel_tabela_path)
        try:
            tabela_sheet_df = load_tabela_sheet_df(excel_tabela_path)
        except Exception:
            tabela_sheet_df = None

st.caption(f"📌 Status TABELA: {status_tabela}")

dfs = []
log_execucao = []

for fpath in arquivos:
    try:
        df_raw = read_txt_tab(fpath)
    except Exception as e:
        log_execucao.append((fpath.name, f"Falha ao ler TXT: {e}"))
        continue

    if df_raw.shape[1] != len(CABECALHOS):
        log_execucao.append((fpath.name, f"⚠️ Colunas: {df_raw.shape[1]} (esperado {len(CABECALHOS)})"))

    if renomear_automatico and fpath.exists():
        novo_path, status = maybe_rename_txt_by_entrada(fpath, df_raw)
        log_execucao.append((fpath.name, status))
        fpath = novo_path
    else:
        log_execucao.append((fpath.name, "Leitura OK (sem renomeio)"))

    df_proc = aplicar_tipos(df_raw)
    df_proc["ARQUIVO_ORIGEM"] = fpath.name
    dfs.append(df_proc)

if mostrar_detalhes_renomeio:
    with st.expander("🧾 Log de execução"):
        for nome_arq, status in log_execucao:
            st.write(f"- **{nome_arq}** → {status}")

if not dfs:
    st.error("Não foi possível processar nenhum arquivo TXT (verifique o log acima).")
    st.stop()

df_final = pd.concat(dfs, ignore_index=True)

# pipeline homologado (mantido)
df_final = enrich_with_tabela(df_final, mapping)
df_final = normalizar_cbhpm_para_numerico(df_final)  # CBHPM numérico com 2 casas
df_final = ensure_col_ab_cirurgiao(df_final)
df_final = ensure_cols_ac_to_ak(df_final)
df_final = preencher_cirurgiao_por_formula(df_final)
df_final = preencher_auxiliares_por_regra(df_final)
df_final = preencher_deflator_e_rep_regra(df_final)

min_global, max_global = compute_minmax_for_sheet(df_final)
nome_aba_futura = sheet_name_from_minmax(min_global, max_global)
chunks = chunk_df_for_excel_by_max_rows(df_final, EXCEL_MAX_LINHAS)

# Botões: Exportar + Recarregar
with c4:
    b1, b2 = st.columns([1, 1])

    with b1:
        export_disabled = usar_tabela and (excel_tabela_path is None or tabela_sheet_df is None or tabela_sheet_df.empty)
        if export_disabled:
            st.button("📤 Exportar Excel", disabled=True)
        else:
            excel_bytes = build_excel_with_formulas(chunks, tabela_sheet_df if usar_tabela else None)
            nome_base = f"HM_{nome_aba_futura}"
            st.download_button(
                "📤 Exportar Excel",
                data=excel_bytes,
                file_name=f"{nome_base}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    with b2:
        if st.button("🔄 Recarregar", use_container_width=True):
            st.rerun()

if usar_tabela and (excel_tabela_path is None or tabela_sheet_df is None or tabela_sheet_df.empty):
    st.warning("Para exportar com fórmulas (PROCV/VLOOKUP), é necessário o arquivo TABELA.xlsx na pasta e a aba 'TABELA' válida.")

st.markdown("### 2) Resultado acumulado (todos os TXT da pasta)")

if pd.isna(min_global) or pd.isna(max_global):
    st.warning("Não foi possível calcular o intervalo global de ENTRADA (datas inválidas/ausentes).")
else:
    st.caption(
        f"Intervalo global (ENTRADA): **{min_global.strftime('%d/%m/%Y')}** a **{max_global.strftime('%d/%m/%Y')}** "
        f"• Nome de aba sugerido (exportação): **{nome_aba_futura}**"
    )

if usar_tabela and "CÓD. TUSS" in df_final.columns:
    total_linhas = len(df_final)
    qtd_porte = (df_final["PORTE"].astype(str).str.strip() != "").sum() if "PORTE" in df_final.columns else 0
    qtd_cbhpm = (pd.to_numeric(df_final["CBHPM"], errors="coerce").fillna(0) != 0).sum() if "CBHPM" in df_final.columns else 0
    st.caption(
        f"Matches via TABELA — PORTE preenchido em **{qtd_porte:,}** linhas, "
        f"CBHPM preenchido em **{qtd_cbhpm:,}** linhas (de {total_linhas:,}).".replace(",", ".")
    )

df_view = formatar_exibicao(df_final)
st.dataframe(df_view, use_container_width=True, height=560)

st.markdown("### 3) Prévia da divisão por abas (para exportação futura)")
st.write(f"Total de linhas acumuladas: **{len(df_final):,}**".replace(",", "."))
st.write(f"Quantidade de abas necessárias (se exportasse agora): **{len(chunks)}**")

with st.expander("📌 Detalhes das abas planejadas"):
    for i, (sname, chunk) in enumerate(chunks, start=1):
        mn, mx = compute_minmax_for_sheet(chunk)
        mn_txt = mn.strftime("%d/%m/%Y") if pd.notna(mn) else "—"
        mx_txt = mx.strftime("%d/%m/%Y") if pd.notna(mx) else "—"
        st.write(
            f"**Aba {i}: {sname}** • Linhas: {len(chunk):,}".replace(",", ".")
            + f" • ENTRADA: {mn_txt} a {mx_txt}"
        )

with st.expander("🔎 Diagnóstico rápido"):
    st.write("Colunas:", list(df_final.columns))
    st.write("Amostra (processado):")
    st.dataframe(df_final.head(10), use_container_width=True)
    if usar_tabela and excel_tabela_path is not None:
        st.write(f"Excel TABELA usado: {excel_tabela_path.name}")
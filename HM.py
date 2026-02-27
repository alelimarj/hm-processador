# HM.py — HM – Processador Hospitalar TXT → Excel com Regras e Fórmulas
# (VERSÃO HOMOLOGADA + upload completo) — BASE HOMOLOGADA + AJUSTES:
# 1) Unificar "Exportar Excel" + "Baixar Excel" em um único botão (download_button)
# 2) Botão "Limpar base" (apaga apenas TXT)
# 3) QUANTIDADE exportada como valor numérico
# 4) CORREÇÃO DEFINITIVA CLOUD: base_dir em ./data/base_de_dados + debug de salvamento + tratamento de conflito arquivo/pasta

import os
import io
import glob
from datetime import datetime
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.worksheet import Worksheet


# =========================
# 0) STREAMLIT CONFIG (primeiro st.*)
# =========================
st.set_page_config(page_title="HM — Processador Hospitalar",
                   page_icon="🏥", layout="wide")


# =========================
# 1) CONFIG / PASTA BASE (CORREÇÃO DEFINITIVA CLOUD)
# =========================
def _ensure_dir(path: str) -> str:
    """
    Garante que 'path' seja diretório.
    Se existir e não for diretório (ex.: arquivo com mesmo nome), usa fallback.
    """
    path = os.path.normpath(path)

    # conflito: existe mas não é diretório
    if os.path.exists(path) and not os.path.isdir(path):
        # fallback seguro
        path = os.path.normpath("./data/base_de_dados")

    # cria se não existir
    if not os.path.isdir(path):
        os.makedirs(path, exist_ok=True)

    return path


def detect_base_dir() -> str:
    """
    Windows local: C:\\base_de_dados
    Streamlit Cloud/Linux: ./data/base_de_dados  (definitivo e mais estável)
    """
    win_path = r"C:\base_de_dados"

    if os.name == "nt":
        return _ensure_dir(win_path)

    # Cloud/Linux: caminho definitivo
    return _ensure_dir("./data/base_de_dados")


BASE_DIR = detect_base_dir()
TABELA_XLSX_PATH = os.path.join(BASE_DIR, "TABELA.xlsx")
CHUNK_SAFE_ROWS = 1_048_000  # margem de segurança (limite Excel 1.048.576)


# =========================
# 2) CABEÇALHOS FIXOS A..Y
# =========================
FIXED_HEADERS_A_TO_Y = [
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


# =========================
# 3) LEITURA TXT (ENCODINGS)
# =========================
def read_txt_as_df(path: str) -> pd.DataFrame:
    encodings = ["utf-8", "utf-8-sig", "latin-1"]
    last_err = None

    for enc in encodings:
        try:
            df = pd.read_csv(
                path,
                sep="\t",
                header=None,
                dtype=str,
                encoding=enc,
                engine="python",
                keep_default_na=False,
            )

            # garante 25 colunas (A..Y)
            if df.shape[1] < 25:
                for _ in range(25 - df.shape[1]):
                    df[df.shape[1]] = ""

            df = df.iloc[:, :25]
            df.columns = FIXED_HEADERS_A_TO_Y

            # tudo string
            for c in df.columns:
                df[c] = df[c].astype(str)

            return df

        except Exception as e:
            last_err = e
            continue

    raise RuntimeError(
        f"Falha ao ler TXT {os.path.basename(path)}. Último erro: {last_err}")


def parse_date_safe(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
    return dt


def format_ddmmaa(dt: datetime) -> str:
    return dt.strftime("%d-%m-%y")


# =========================
# 4) RENOMEIO AUTOMÁTICO TXT
# =========================
def rename_txt_by_entrada(path: str) -> str:
    df = read_txt_as_df(path)
    entrada_dt = parse_date_safe(df["ENTRADA"])

    if entrada_dt.notna().sum() == 0:
        return path  # sem data válida: não renomeia

    mn = entrada_dt.min()
    mx = entrada_dt.max()

    new_name = f"{format_ddmmaa(mn)}_a_{format_ddmmaa(mx)}_HM.txt"
    new_path = os.path.join(os.path.dirname(path), new_name)

    # evita conflito: se já é o mesmo
    if os.path.abspath(new_path) == os.path.abspath(path):
        return path

    # evita sobrescrever
    if os.path.exists(new_path):
        base, ext = os.path.splitext(new_name)
        i = 2
        while True:
            candidate = os.path.join(os.path.dirname(path), f"{base}_{i}{ext}")
            if not os.path.exists(candidate):
                new_path = candidate
                break
            i += 1

    os.rename(path, new_path)
    return new_path


# =========================
# 5) TABELA.xlsx (integração)
# =========================
def load_tabela_df(path: str) -> pd.DataFrame:
    if not os.path.exists(path):
        return pd.DataFrame()

    try:
        df = pd.read_excel(path, sheet_name="TABELA", dtype=object)
        return df
    except ValueError:
        raise RuntimeError(
            "O arquivo TABELA.xlsx foi encontrado, mas NÃO possui a aba 'TABELA'.")


def build_tabela_map(tabela_df: pd.DataFrame) -> pd.DataFrame:
    if tabela_df.empty:
        return pd.DataFrame(columns=["KEY_TUSS", "PORTE", "CBHPM", "QTD_AUX_TABELA"])

    cols_lower = {c: str(c).strip().lower() for c in tabela_df.columns}
    key_col = None
    for c, cl in cols_lower.items():
        if "id" in cl and "proced" in cl:
            key_col = c
            break

    if key_col is None:
        if len(tabela_df.columns) > 4:
            key_col = tabela_df.columns[4]  # fallback coluna E
        else:
            return pd.DataFrame(columns=["KEY_TUSS", "PORTE", "CBHPM", "QTD_AUX_TABELA"])

    def safe_iloc(idx: int):
        return tabela_df.columns[idx] if len(tabela_df.columns) > idx else None

    porte_col = safe_iloc(8)
    qtd_aux_col = safe_iloc(10)
    cbhpm_col = safe_iloc(15)

    out = pd.DataFrame()
    out["KEY_TUSS"] = tabela_df[key_col].astype(str).str.strip()
    out["PORTE"] = tabela_df[porte_col] if porte_col is not None else None
    out["CBHPM"] = tabela_df[cbhpm_col] if cbhpm_col is not None else None
    out["QTD_AUX_TABELA"] = tabela_df[qtd_aux_col] if qtd_aux_col is not None else None
    return out


def integrate_tabela(main_df: pd.DataFrame, tabela_map: pd.DataFrame) -> pd.DataFrame:
    if main_df.empty:
        return main_df

    if tabela_map.empty:
        for col in ["PORTE", "CBHPM", "QTD_AUX_TABELA"]:
            if col not in main_df.columns:
                main_df[col] = ""
        return main_df

    tmp = tabela_map.copy()
    tmp["KEY_TUSS"] = tmp["KEY_TUSS"].astype(str).str.strip()

    df = main_df.copy()
    df["CÓD. TUSS"] = df["CÓD. TUSS"].astype(str).str.strip()

    df = df.merge(
        tmp[["KEY_TUSS", "PORTE", "CBHPM", "QTD_AUX_TABELA"]],
        how="left",
        left_on="CÓD. TUSS",
        right_on="KEY_TUSS",
    ).drop(columns=["KEY_TUSS"])

    return df


# =========================
# 6) COLUNAS ADICIONAIS (após CBHPM)
# =========================
ADDITIONAL_COLS_ORDER = [
    "CIRURGIÃO",
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


def ensure_additional_columns(df: pd.DataFrame) -> pd.DataFrame:
    for col in ["PORTE", "CBHPM", "QTD_AUX_TABELA"]:
        if col not in df.columns:
            df[col] = ""

    out = df.copy()

    for col in ADDITIONAL_COLS_ORDER:
        if col not in out.columns:
            out[col] = ""

    cols = list(out.columns)
    if "CBHPM" in cols:
        idx = cols.index("CBHPM")
        left = cols[: idx + 1]
        right = [c for c in cols[idx + 1:] if c not in ADDITIONAL_COLS_ORDER]
        out = out[left + ADDITIONAL_COLS_ORDER + right]

    return out


# =========================
# 7) REGRAS DE CÁLCULO (preview em python)
# =========================
def to_float_safe(x, default=0.0) -> float:
    if x is None:
        return default
    s = str(x).strip()
    if s == "" or s.lower() == "nan":
        return default

    s2 = s.replace(".", "").replace(",", ".") if ("," in s) else s
    try:
        return float(s2)
    except:
        return default


def compute_preview_values(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    cbhpm = out["CBHPM"].apply(to_float_safe)
    via = out["VIA DE ACESSO"].apply(lambda v: to_float_safe(v, default=1.0))
    qtd = out["QUANTIDADE"].apply(lambda v: to_float_safe(v, default=0.0))
    qtd_aux = out["QTD_AUX_TABELA"].apply(
        lambda v: int(to_float_safe(v, default=0.0)))

    cir = cbhpm * via * qtd
    aux1 = cir * 0.3 * (qtd_aux >= 1)
    aux2 = cir * 0.2 * (qtd_aux >= 2)
    aux3 = aux1 * 0.1 * (qtd_aux >= 3)
    deflator = (aux1 + aux2 + aux3) * 0.2
    regra = (cir + aux1 + aux2 + aux3) - deflator

    out["CIRURGIÃO"] = cir
    out["1º AUXILIAR"] = aux1
    out["2º AUXILIAR"] = aux2
    out["3º AUXILIAR"] = aux3
    out["DEFLATOR"] = deflator
    out["VALOR DE REP. REGRA"] = regra

    return out


# =========================
# 8) PROCESSAMENTO ACUMULATIVO + CHUNK EXCEL
# =========================
def list_txt_files() -> list[str]:
    return sorted(glob.glob(os.path.join(BASE_DIR, "*.txt")))


def process_all_txts() -> pd.DataFrame:
    paths = list_txt_files()
    if not paths:
        return pd.DataFrame(columns=FIXED_HEADERS_A_TO_Y)

    # renomeia antes de concatenar (homologado)
    for p in paths:
        try:
            rename_txt_by_entrada(p)
        except:
            pass

    paths = list_txt_files()

    dfs = []
    for p in paths:
        dfs.append(read_txt_as_df(p))

    main = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame(
        columns=FIXED_HEADERS_A_TO_Y)
    return main


def chunk_dataframe_for_excel(df: pd.DataFrame) -> list[pd.DataFrame]:
    if df.empty:
        return [df]

    chunks = []
    n = len(df)
    start = 0
    while start < n:
        end = min(start + CHUNK_SAFE_ROWS, n)
        chunks.append(df.iloc[start:end].copy())
        start = end
    return chunks


def chunk_sheet_name(df_chunk: pd.DataFrame) -> str:
    if df_chunk.empty:
        return "SEM_DADOS"
    entrada_dt = parse_date_safe(df_chunk["ENTRADA"])
    if entrada_dt.notna().sum() == 0:
        return "SEM_DATA"
    mn = entrada_dt.min()
    mx = entrada_dt.max()
    return f"{format_ddmmaa(mn)}_a_{format_ddmmaa(mx)}"


# =========================
# 9) EXPORTAÇÃO EXCEL (com fórmulas reais + formatos)
# =========================
def set_col_format(ws: Worksheet, col_idx: int, number_format: str, start_row: int = 2):
    for r in range(start_row, ws.max_row + 1):
        ws.cell(row=r, column=col_idx).number_format = number_format


def try_write_numeric(ws: Worksheet, row: int, col: int, value_str):
    s = str(value_str).strip()
    if s == "" or s.lower() == "nan":
        ws.cell(row=row, column=col).value = None
        return

    if s.isdigit():
        ws.cell(row=row, column=col).value = int(s)
        return

    v = to_float_safe(s, default=None)
    if v is None:
        ws.cell(row=row, column=col).value = s
    else:
        ws.cell(row=row, column=col).value = v


def normalize_numeric_columns_for_excel(ws: Worksheet, cols: list[str]):
    """
    Ajuste seguro: garante que campos usados em fórmula sejam numéricos no Excel,
    e atende QUANTIDADE como VALOR (numérico).
    """
    numeric_cols_general = ["CBHPM", "VIA DE ACESSO",
                            "QUANTIDADE", "QTD_AUX_TABELA", "VALOR DO PROC."]

    for cname in numeric_cols_general:
        if cname in cols:
            idx = cols.index(cname) + 1
            for r in range(2, ws.max_row + 1):
                raw = ws.cell(row=r, column=idx).value
                try_write_numeric(ws, r, idx, raw)

            if cname == "QUANTIDADE":
                for r in range(2, ws.max_row + 1):
                    ws.cell(row=r, column=idx).number_format = "0.########"

    for code_col in ["CÓD. TUSS", "CÓD. PRODUTO"]:
        if code_col in cols:
            idx = cols.index(code_col) + 1
            for r in range(2, ws.max_row + 1):
                raw = ws.cell(row=r, column=idx).value
                try_write_numeric(ws, r, idx, raw)
                ws.cell(row=r, column=idx).number_format = "0"


def build_excel_bytes(final_df: pd.DataFrame, tabela_df: pd.DataFrame) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    # Aba TABELA (homologado)
    ws_tab = wb.create_sheet("TABELA")
    if tabela_df is not None and not tabela_df.empty:
        ws_tab.append(list(tabela_df.columns))
        for _, row in tabela_df.iterrows():
            ws_tab.append([row.get(c) for c in tabela_df.columns])
    else:
        ws_tab.append(["(Sem TABELA.xlsx carregado)"])

    chunks = chunk_dataframe_for_excel(final_df)

    for chunk in chunks:
        name = chunk_sheet_name(chunk)
        sheet_name = name[:31]

        base_name = sheet_name
        i = 2
        while sheet_name in wb.sheetnames:
            sheet_name = (base_name[:28] + f"_{i}")[:31]
            i += 1

        ws = wb.create_sheet(sheet_name)

        # header
        ws.append(list(chunk.columns))
        header_font = Font(bold=True)
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=c)
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True)

        cols = list(chunk.columns)

        def col_letter(col_name: str) -> str:
            idx = cols.index(col_name) + 1
            return get_column_letter(idx)

        # dados + fórmulas
        for r_idx, (_, row) in enumerate(chunk.iterrows(), start=2):
            for c_idx, col_name in enumerate(cols, start=1):
                ws.cell(row=r_idx, column=c_idx).value = row.get(col_name, "")

            if all(x in cols for x in ["CBHPM", "VIA DE ACESSO", "QUANTIDADE", "QTD_AUX_TABELA"]):
                cb = f"{col_letter('CBHPM')}{r_idx}"
                via = f"{col_letter('VIA DE ACESSO')}{r_idx}"
                qt = f"{col_letter('QUANTIDADE')}{r_idx}"
                qa = f"{col_letter('QTD_AUX_TABELA')}{r_idx}"

                cir_col = col_letter(
                    "CIRURGIÃO") if "CIRURGIÃO" in cols else None
                a1_col = col_letter(
                    "1º AUXILIAR") if "1º AUXILIAR" in cols else None
                a2_col = col_letter(
                    "2º AUXILIAR") if "2º AUXILIAR" in cols else None
                a3_col = col_letter(
                    "3º AUXILIAR") if "3º AUXILIAR" in cols else None
                def_col = col_letter(
                    "DEFLATOR") if "DEFLATOR" in cols else None
                reg_col = col_letter(
                    "VALOR DE REP. REGRA") if "VALOR DE REP. REGRA" in cols else None

                if cir_col:
                    ws[f"{cir_col}{r_idx}"].value = f"={cb}*{via}*{qt}"
                if a1_col and cir_col:
                    ws[f"{a1_col}{r_idx}"].value = f"=IF({qa}>=1,{cir_col}{r_idx}*0.3,0)"
                if a2_col and cir_col:
                    ws[f"{a2_col}{r_idx}"].value = f"=IF({qa}>=2,{cir_col}{r_idx}*0.2,0)"
                if a3_col and a1_col:
                    ws[f"{a3_col}{r_idx}"].value = f"=IF({qa}>=3,{a1_col}{r_idx}*0.1,0)"
                if def_col and a1_col and a2_col and a3_col:
                    ws[f"{def_col}{r_idx}"].value = f"=({a1_col}{r_idx}+{a2_col}{r_idx}+{a3_col}{r_idx})*0.2"
                if reg_col and cir_col and a1_col and a2_col and a3_col and def_col:
                    ws[f"{reg_col}{r_idx}"].value = f"=({cir_col}{r_idx}+{a1_col}{r_idx}+{a2_col}{r_idx}+{a3_col}{r_idx})-{def_col}{r_idx}"

        # garante números (inclui QUANTIDADE como valor)
        normalize_numeric_columns_for_excel(ws, cols)

        # Formatações homologadas:
        # CBHPM: #.##0,00
        if "CBHPM" in cols:
            cb_idx = cols.index("CBHPM") + 1
            set_col_format(ws, cb_idx, "#.##0,00", start_row=2)

        # largura leve
        for c in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(c)].width = 18

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# =========================
# 10) UPLOAD: salvar sem sobrescrever
# =========================
def safe_save_uploaded_file(uploaded_file, target_dir: str, forced_name: str | None = None) -> str:
    """
    Salva evitando sobrescrever. Retorna o caminho final.
    """
    target_dir = _ensure_dir(target_dir)

    original_name = forced_name if forced_name else uploaded_file.name
    base, ext = os.path.splitext(original_name)
    candidate = os.path.join(target_dir, original_name)

    if not os.path.exists(candidate):
        with open(candidate, "wb") as out:
            out.write(uploaded_file.getbuffer())
        return candidate

    i = 2
    while True:
        new_name = f"{base}_{i}{ext}"
        candidate = os.path.join(target_dir, new_name)
        if not os.path.exists(candidate):
            with open(candidate, "wb") as out:
                out.write(uploaded_file.getbuffer())
            return candidate
        i += 1


# =========================
# 11) LIMPAR BASE (TXT)
# =========================
def limpar_base_txt() -> int:
    """
    Apaga apenas arquivos .txt da pasta base. Retorna quantidade apagada.
    """
    removed = 0
    for p in glob.glob(os.path.join(BASE_DIR, "*.txt")):
        try:
            os.remove(p)
            removed += 1
        except:
            pass
    return removed


# =========================
# 12) STREAMLIT UI
# =========================
st.title("🏥 HM — Processador Hospitalar TXT → Excel")

with st.expander("📌 Pasta base detectada", expanded=True):
    st.write(f"**Pasta base em uso:** `{BASE_DIR}`")
    st.caption(
        "Windows usa `C:\\base_de_dados`. No Streamlit Cloud usa `./data/base_de_dados` (mais estável).")

col_up1, col_up2 = st.columns([2, 1], gap="large")

with col_up1:
    st.subheader("📤 Upload de arquivos")
    txt_files = st.file_uploader(
        "Envie os arquivos .txt (múltiplos)",
        type=["txt"],
        accept_multiple_files=True
    )
    tabela_file = st.file_uploader(
        "Envie o arquivo TABELA.xlsx (aba TABELA)",
        type=["xlsx"],
        accept_multiple_files=False
    )

    save_clicked = st.button(
        "💾 Salvar uploads na pasta base", use_container_width=True)

    # ✅ CORREÇÃO DEFINITIVA: feedback claro + erro visível se falhar
    if save_clicked:
        saved_paths = []
        try:
            if txt_files:
                for f in txt_files:
                    p = safe_save_uploaded_file(f, BASE_DIR)
                    saved_paths.append(p)

            if tabela_file is not None:
                p = safe_save_uploaded_file(
                    tabela_file, BASE_DIR, forced_name="TABELA.xlsx")
                saved_paths.append(p)

            if saved_paths:
                st.success("Arquivos salvos na pasta base:")
                for p in saved_paths:
                    try:
                        st.write(f"- `{p}` ({os.path.getsize(p)} bytes)")
                    except:
                        st.write(f"- `{p}`")
                st.rerun()
            else:
                st.warning("Nenhum arquivo foi enviado para salvar.")

        except Exception as e:
            st.error(
                f"Falha ao salvar uploads na pasta base ({BASE_DIR}): {e}")

with col_up2:
    st.subheader("📁 Conteúdo atual da base")
    txt_list = list_txt_files()
    st.write(f"**TXT na pasta:** {len(txt_list)}")
    if txt_list:
        st.caption("Arquivos encontrados:")
        st.write("\n".join([f"- {os.path.basename(p)}" for p in txt_list]))
    else:
        st.info("Nenhum TXT encontrado na pasta base.")

    st.write("---")
    has_tabela = os.path.exists(TABELA_XLSX_PATH)
    st.write(
        f"**TABELA.xlsx:** {'✅ encontrado' if has_tabela else '❌ não encontrado'}")

st.write("---")

# Botões principais (Recarregar / Exportar único)
btn_col1, btn_col2 = st.columns([1, 1], gap="large")

recarregar = btn_col1.button(
    "🔄 Recarregar (processar TXT da base)", use_container_width=True)

# Processa sempre que clicar ou se ainda não tiver dados
if recarregar or "final_df" not in st.session_state:
    try:
        main_df = process_all_txts()
        tabela_df = load_tabela_df(TABELA_XLSX_PATH) if os.path.exists(
            TABELA_XLSX_PATH) else pd.DataFrame()
        tabela_map = build_tabela_map(
            tabela_df) if not tabela_df.empty else pd.DataFrame()

        final_df = integrate_tabela(main_df, tabela_map)
        final_df = ensure_additional_columns(final_df)

        preview_df = compute_preview_values(final_df)

        st.session_state["main_df"] = main_df
        st.session_state["tabela_df"] = tabela_df
        st.session_state["final_df"] = final_df
        st.session_state["preview_df"] = preview_df

        # invalida cache de excel para forçar novo quando dados mudarem
        st.session_state.pop("excel_bytes", None)
        st.session_state.pop("excel_filename", None)

    except Exception as e:
        st.error(f"Erro no processamento: {e}")

final_df = st.session_state.get("final_df", pd.DataFrame())
preview_df = st.session_state.get("preview_df", pd.DataFrame())
tabela_df = st.session_state.get("tabela_df", pd.DataFrame())

st.subheader("✅ Preview (com regras em Python)")
if preview_df is None or preview_df.empty:
    st.info("Sem dados para exibir. Envie TXT, clique em **Salvar uploads**, depois clique em **Recarregar**.")
else:
    st.dataframe(preview_df.head(200), use_container_width=True)

# ===== Exportação em UM ÚNICO BOTÃO (download direto) =====
# gera bytes uma vez por sessão (ou quando recarregar)
if final_df is not None and not final_df.empty:
    if "excel_bytes" not in st.session_state:
        try:
            st.session_state["excel_bytes"] = build_excel_bytes(
                final_df,
                tabela_df if isinstance(
                    tabela_df, pd.DataFrame) else pd.DataFrame()
            )
            st.session_state["excel_filename"] = f"HM_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        except Exception as e:
            st.session_state["excel_bytes"] = None
            st.session_state["excel_filename"] = None
            st.error(f"Erro ao preparar Excel: {e}")

    excel_bytes = st.session_state.get("excel_bytes")
    excel_filename = st.session_state.get("excel_filename")

    if excel_bytes:
        exported = btn_col2.download_button(
            label="📥 Exportar Excel (com fórmulas)",
            data=excel_bytes,
            file_name=excel_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        if exported:
            st.success(
                "Excel exportado com sucesso (inclui aba TABELA e fórmulas reais).")
    else:
        btn_col2.button("📥 Exportar Excel (com fórmulas)",
                        use_container_width=True, disabled=True)
else:
    btn_col2.button("📥 Exportar Excel (com fórmulas)",
                    use_container_width=True, disabled=True)

st.write("---")

# ===== Botão Limpar base =====
st.subheader("🧹 Manutenção da base")

col_l1, col_l2 = st.columns([2, 1], gap="large")
with col_l1:
    st.caption(
        "Esta ação apaga **somente** os arquivos **.txt** da pasta base. O `TABELA.xlsx` é mantido.")
    confirm_clear = st.checkbox(
        "Confirmo que desejo apagar TODOS os .txt da pasta base.", value=False)

with col_l2:
    if st.button("🧹 Limpar base (apagar TXT)", use_container_width=True, disabled=not confirm_clear):
        removed = limpar_base_txt()
        # limpa sessão
        for k in ["main_df", "tabela_df", "final_df", "preview_df", "excel_bytes", "excel_filename"]:
            st.session_state.pop(k, None)
        st.success(f"Base limpa: {removed} arquivo(s) .txt removido(s).")
        st.rerun()

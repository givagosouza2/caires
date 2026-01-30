import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Juntar 25 condições", layout="wide")
st.title("Extrair colunas por nome e juntar tudo em um arquivo")

COLS_WANTED = [
    "K",
    "Início global (s)",
    "Fim global (s)",
    "Duração global (s)",
    "Comp1 início (s)",
    "Comp1 fim (s)",
    "Comp1 duração (s)",
    "Comp2 início (s)",
    "Comp2 fim (s)",
    "Comp2 duração (s)",
]

def read_any_table(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv") or name.endswith(".txt"):
        # utf-8-sig lida com BOM
        return pd.read_csv(uploaded_file, encoding="utf-8-sig")
    elif name.endswith(".xlsx") or name.endswith(".xlsm") or name.endswith(".xls"):
        return pd.read_excel(uploaded_file, sheet_name=0, engine="openpyxl")
    else:
        raise ValueError(f"Formato não suportado: {uploaded_file.name}")

def normalize_col(s: str) -> str:
    """Normaliza para comparar nomes de colunas (remove espaços extras)."""
    return " ".join(str(s).strip().split())

def build_col_map(df: pd.DataFrame):
    """Mapa: nome_normalizado -> nome_original"""
    return {normalize_col(c): c for c in df.columns}

st.info("Faça upload dos 25 arquivos (ou quantos quiser). O app vai extrair as colunas por nome e juntar tudo.")

files = st.file_uploader(
    "Arquivos (.csv, .txt, .xlsx, .xlsm, .xls)",
    type=["csv", "txt", "xlsx", "xlsm", "xls"],
    accept_multiple_files=True
)

if not files:
    st.stop()

all_parts = []
errors = []

for f in files:
    try:
        df = read_any_table(f)

        col_map = build_col_map(df)
        wanted_norm = [normalize_col(c) for c in COLS_WANTED]

        missing = [COLS_WANTED[i] for i, wn in enumerate(wanted_norm) if wn not in col_map]
        if missing:
            errors.append((f.name, f"Faltando colunas: {missing}"))
            continue

        # Seleciona mantendo nomes originais do arquivo
        selected = df[[col_map[wn] for wn in wanted_norm]].copy()

        # Renomeia para os nomes padronizados (iguais para todos)
        selected.columns = COLS_WANTED

        # Rastreabilidade
        selected.insert(0, "Arquivo", f.name)

        all_parts.append(selected)

    except Exception as e:
        errors.append((f.name, str(e)))

if errors:
    st.error("Alguns arquivos não puderam ser processados:")
    for fname, msg in errors:
        st.write(f"- **{fname}**: {msg}")

if not all_parts:
    st.warning("Nenhum arquivo foi consolidado (todos falharam ou estavam sem as colunas).")
    st.stop()

final_df = pd.concat(all_parts, ignore_index=True)

st.success(f"Consolidado! Linhas totais: {len(final_df)} | Arquivos OK: {len(all_parts)}")
st.dataframe(final_df, use_container_width=True)

# Download CSV
csv_bytes = final_df.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    "⬇️ Baixar CSV consolidado",
    data=csv_bytes,
    file_name="consolidado_25_condicoes.csv",
    mime="text/csv"
)

# Download Excel
output = io.BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    final_df.to_excel(writer, index=False, sheet_name="consolidado")
st.download_button(
    "⬇️ Baixar Excel consolidado",
    data=output.getvalue(),
    file_name="consolidado_25_condicoes.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

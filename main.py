import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Extrator de Blocos (F2:H6 e L2:S6)", layout="wide")
st.title("Extrator de dados por condição (25 arquivos)")

# -----------------------------
# Helpers
# -----------------------------
def col_letter_to_index(letter: str) -> int:
    """Excel column letter (A, B, ..., Z, AA, AB, ...) -> 0-based index"""
    letter = letter.strip().upper()
    idx = 0
    for ch in letter:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1

def extract_excel_like_range(df: pd.DataFrame, col_start: str, col_end: str, row_start: int, row_end: int, header_in_row1: bool = True):
    """
    Extrai um range no estilo Excel: colunas por letras (ex: F-H), linhas por números (ex: 2-6).

    Se header_in_row1=True:
      - Linha 1 é cabeçalho
      - Linha 2 (Excel) corresponde à primeira linha de dados (iloc[0])
    """
    c0 = col_letter_to_index(col_start)
    c1 = col_letter_to_index(col_end)

    # Converter linhas Excel -> iloc
    # Excel row 2 -> iloc 0 quando há cabeçalho na linha 1
    offset = 2 if header_in_row1 else 1  # se há header, subtrai 2; se não, subtrai 1
    r0 = row_start - offset
    r1 = row_end - offset

    # iloc end é exclusivo, então +1
    return df.iloc[r0:r1 + 1, c0:c1 + 1]

def read_any_table(uploaded_file) -> pd.DataFrame:
    """Lê CSV ou Excel para DataFrame."""
    name = uploaded_file.name.lower()

    if name.endswith(".csv") or name.endswith(".txt"):
        # utf-8-sig lida com BOM (comum em CSVs gerados por alguns softwares)
        return pd.read_csv(uploaded_file, encoding="utf-8-sig")
    elif name.endswith(".xlsx") or name.endswith(".xlsm") or name.endswith(".xls"):
        # para Excel: tenta primeira aba
        return pd.read_excel(uploaded_file, sheet_name=0, engine="openpyxl")
    else:
        raise ValueError(f"Formato não suportado: {uploaded_file.name}")

def sanitize_condition_label(s: str) -> str:
    s = s.strip()
    s = re.sub(r"\s+", " ", s)
    return s

# -----------------------------
# UI - condições
# -----------------------------
st.sidebar.header("Configuração")

default_conditions = "\n".join([f"Cond_{i:02d}" for i in range(1, 26)])
conditions_text = st.sidebar.text_area(
    "Lista de condições (1 por linha)",
    value=default_conditions,
    height=260
)

header_in_row1 = st.sidebar.checkbox("Arquivo tem cabeçalho na linha 1 (recomendado)", value=True)

conditions = [sanitize_condition_label(x) for x in conditions_text.splitlines() if x.strip()]
if len(conditions) == 0:
    st.warning("Informe pelo menos 1 condição na barra lateral.")
    st.stop()

st.write(f"**Condições ativas:** {len(conditions)}")

st.info("Para cada condição, faça upload do arquivo correspondente. Em seguida, clique em **Processar**.")

# Uploaders
uploaded_by_condition = {}
cols = st.columns(3)
for i, cond in enumerate(conditions):
    with cols[i % 3]:
        uploaded_by_condition[cond] = st.file_uploader(
            f"📄 {cond}",
            type=["csv", "txt", "xlsx", "xlsm", "xls"],
            key=f"uploader_{i}"
        )

# -----------------------------
# Processamento
# -----------------------------
if st.button("🚀 Processar", type="primary"):
    results_long = []   # formato longo: condition, block, row, col, value
    results_wide = []   # formato largo: condition + colunas do bloco

    errors = []

    for cond, up in uploaded_by_condition.items():
        if up is None:
            continue

        try:
            df = read_any_table(up)

            # Extrair ranges
            block1 = extract_excel_like_range(df, "F", "H", 2, 6, header_in_row1=header_in_row1)   # F2:H6
            block2 = extract_excel_like_range(df, "L", "S", 2, 6, header_in_row1=header_in_row1)   # L2:S6

            # Guardar "wide" (mantendo colunas originais)
            b1 = block1.copy()
            b1.insert(0, "Condition", cond)
            b1.insert(1, "Block", "F2:H6")
            results_wide.append(b1)

            b2 = block2.copy()
            b2.insert(0, "Condition", cond)
            b2.insert(1, "Block", "L2:S6")
            results_wide.append(b2)

            # Guardar "long" (facilita estatística / filtros depois)
            for block_name, block_df in [("F2:H6", block1), ("L2:S6", block2)]:
                tmp = block_df.copy()
                tmp["__row__"] = range(2, 2 + len(tmp))  # rotula como linhas Excel 2..6
                tmp_long = tmp.melt(id_vars="__row__", var_name="Column", value_name="Value")
                tmp_long.insert(0, "Condition", cond)
                tmp_long.insert(1, "Block", block_name)
                tmp_long = tmp_long.rename(columns={"__row__": "ExcelRow"})
                results_long.append(tmp_long)

        except Exception as e:
            errors.append((cond, up.name, str(e)))

    if errors:
        st.error("Alguns arquivos falharam ao processar:")
        for cond, fname, msg in errors:
            st.write(f"- **{cond}** ({fname}): {msg}")

    if not results_wide:
        st.warning("Nenhum arquivo foi processado. Envie pelo menos 1 arquivo e tente novamente.")
        st.stop()

    df_wide = pd.concat(results_wide, ignore_index=True)
    df_long = pd.concat(results_long, ignore_index=True)

    st.success("Processamento concluído!")

    tab1, tab2 = st.tabs(["📌 Preview (wide)", "🔎 Preview (long)"])

    with tab1:
        st.dataframe(df_wide, use_container_width=True)

    with tab2:
        st.dataframe(df_long, use_container_width=True)

    # Exportar para Excel (2 abas)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_wide.to_excel(writer, index=False, sheet_name="wide_blocks")
        df_long.to_excel(writer, index=False, sheet_name="long_blocks")

    st.download_button(
        "⬇️ Baixar resultado (Excel)",
        data=output.getvalue(),
        file_name="extracao_blocos_25condicoes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Exportar também CSV do long (se quiser)
    st.download_button(
        "⬇️ Baixar resultado (CSV long)",
        data=df_long.to_csv(index=False).encode("utf-8-sig"),
        file_name="extracao_blocos_long.csv",
        mime="text/csv"
    )

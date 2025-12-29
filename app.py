# -*- coding: utf-8 -*-
import re
import io
import streamlit as st
from pypdf import PdfReader
import pandas as pd

# =========================
# Configuração da página
# =========================
st.set_page_config(
    page_title="PDF Lince → Excel (Perdas)",
    layout="wide"
)

st.title("📄 PDF Lince → Excel (Perdas)")
st.caption("Converte PDFs de Perdas por Departamento em Excel padronizado.")

# =========================
# Inputs fixos
# =========================
col1, col2, col3 = st.columns(3)

with col1:
    mes = st.text_input("Mês", placeholder="Ex: Dezembro")

with col2:
    semana = st.text_input("Semana", placeholder="Ex: Semana 1")

with col3:
    setor = st.text_input("Setor", placeholder="Ex: Padaria")

uploaded_files = st.file_uploader(
    "Envie os PDFs",
    type=["pdf"],
    accept_multiple_files=True
)

# =========================
# Função de parsing
# =========================
def parse_pdf(file):
    reader = PdfReader(file)
    registros = []

    # Regex robusta para o padrão do Lince
    pattern = re.compile(
        r"^\d{5}\s+(.*?)\s+(KG|UN)\s+(?:KG|UN)\s+\d+,\d+\s+-\s+(\d+,\d+)\s+(\d+,\d+)$"
    )

    for page in reader.pages:
        text = page.extract_text()
        if not text:
            continue

        for line in text.split("\n"):
            line = line.strip()

            match = pattern.match(line)
            if match:
                produto = match.group(1).strip()
                quantidade = float(match.group(3).replace(",", "."))
                valor = float(match.group(4).replace(",", "."))

                registros.append({
                    "Produto": produto,
                    "Setor": setor,
                    "Mês": mes,
                    "Semana": semana,
                    "Quantidade": quantidade,
                    "Valor": valor
                })

    return registros

# =========================
# Processamento
# =========================
if uploaded_files and mes and semana and setor:
    dados = []

    for pdf in uploaded_files:
        dados.extend(parse_pdf(pdf))

    if dados:
        df = pd.DataFrame(dados)

        st.success(f"{len(df)} registros extraídos com sucesso.")
        st.dataframe(df, use_container_width=True)

        # Exportar Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Perdas")

        st.download_button(
            label="⬇️ Baixar Excel",
            data=output.getvalue(),
            file_name="perdas_lince.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Nenhum dado válido foi encontrado nos PDFs.")
else:
    st.info("Preencha Mês, Semana, Setor e envie ao menos um PDF.")


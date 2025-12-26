# -*- coding: utf-8 -*-

import re
import io
import pandas as pd
import streamlit as st
from datetime import datetime
from pypdf import PdfReader

# =========================
# CONFIGURAÇÃO DA PÁGINA
# =========================
st.set_page_config(
    page_title="Lince → Excel | Perdas",
    page_icon="📄",
    layout="wide"
)

st.title("📄 Lince → Excel (Perdas por Departamento)")
st.caption(
    "Faça upload de PDFs do Lince e gere um Excel padronizado "
    "com Produto, Setor, Mês, Semana, Quantidade e Valor."
)

# =========================
# PARÂMETROS MANUAIS
# =========================
st.subheader("📌 Parâmetros da Planilha")

col1, col2, col3 = st.columns(3)

with col1:
    mes_manual = st.text_input(
        "Mês (MM/AAAA)",
        placeholder="Ex: 12/2025"
    )

with col2:
    semana_manual = st.number_input(
        "Semana",
        min_value=1,
        max_value=53,
        step=1
    )

with col3:
    setor_manual = st.text_input(
        "Setor",
        placeholder="Ex: Padaria / Confeitaria / Restaurante"
    )

st.divider()

# =========================
# FUNÇÕES AUXILIARES
# =========================
def limpar_produto(nome):
    nome = re.sub(r'^\d+\s*-\s*', '', nome)
    nome = re.sub(r'\s+(UN|KG|G|PCT)\s*$', '', nome)
    return nome.strip()


def parse_pdf(file, mes_manual, semana_manual, setor_manual):
    reader = PdfReader(file)
    registros = []

    setor_atual = None
    mes = mes_manual if mes_manual else None
    semana = semana_manual if semana_manual else None
    setor_fixo = setor_manual if setor_manual else None

    for page in reader.pages:
        texto = page.extract_text()
        if not texto:
            continue

        linhas = texto.splitlines()

        for linha in linhas:
            linha = linha.strip()

            # Captura automática do setor (caso não seja manual)
            if re.match(r'\d{4}\s+.+\s-\s*$', linha):
                if not setor_fixo:
                    setor_atual = linha.split(" ", 1)[1].replace("-", "").strip()
                continue

            # Ignorar totais
            if linha.startswith("Total"):
                continue

            # Linha de produto começa com código
            if not re.match(r'\d{5}', linha):
                continue

            # Extrair nome do produto
            nome_match = re.match(r'\d{5}\s+(.+?)\s+(UN|KG)', linha)
            if not nome_match:
                continue

            produto = limpar_produto(nome_match.group(1))

            # Extrair todos os números da linha
            numeros = re.findall(r'\d+,\d+|\d+', linha)

            if len(numeros) < 2:
                continue

            try:
                quantidade = float(numeros[-2].replace(",", "."))
                valor = float(numeros[-1].replace(",", "."))
            except ValueError:
                continue

            registros.append({
                "Produto": produto,
                "Setor": setor_fixo if setor_fixo else setor_atual,
                "Mês": mes,
                "Semana": semana,
                "Quantidade": quantidade,
                "Valor": valor
            })

    return registros


# =========================
# UPLOAD DOS PDFS
# =========================
files = st.file_uploader(
    "📤 Envie os PDFs do Lince",
    type=["pdf"],
    accept_multiple_files=True
)

if files:
    if not mes_manual or not semana_manual or not setor_manual:
        st.error("⚠️ Preencha Mês, Semana e Setor antes de processar os PDFs.")
    else:
        dados = []

        for f in files:
            dados.extend(
                parse_pdf(
                    f,
                    mes_manual,
                    semana_manual,
                    setor_manual
                )
            )

        if not dados:
            st.warning("Nenhum dado válido foi encontrado nos PDFs.")
        else:
            df = pd.DataFrame(dados)

            st.success(f"✅ {len(df)} registros extraídos com sucesso")
            st.dataframe(df, use_container_width=True)

            # =========================
            # EXPORTAR EXCEL
            # =========================
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False, sheet_name="Perdas")

            st.download_button(
                "⬇️ Baixar Excel",
                data=buffer.getvalue(),
                file_name="perdas_lince.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

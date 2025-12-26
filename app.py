# -*- coding: utf-8 -*-

import re
import io
import pandas as pd
import streamlit as st
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
    "Upload de PDFs do Lince para gerar Excel padronizado."
)

# =========================
# PARÂMETROS MANUAIS
# =========================
st.subheader("📌 Parâmetros da Planilha")

c1, c2, c3 = st.columns(3)

with c1:
    mes_manual = st.text_input("Mês (MM/AAAA)", placeholder="12/2025")

with c2:
    semana_manual = st.number_input("Semana", 1, 53, 49)

with c3:
    setor_manual = st.text_input(
        "Setor",
        placeholder="Padaria / Confeitaria / Restaurante"
    )

st.divider()

# =========================
# FUNÇÕES
# =========================
def limpar_produto(texto):
    texto = re.sub(r'^\d{5}\s+', '', texto)
    texto = re.sub(r'\s+(UN|KG|G|PCT)(\s+(UN|KG|G|PCT))*$', '', texto)
    return texto.strip()


def parse_pdf(file, mes, semana, setor):
    reader = PdfReader(file)
    registros = []

    for page in reader.pages:
        texto = page.extract_text()
        if not texto:
            continue

        for linha in texto.splitlines():
            linha = linha.strip()

            # Linha precisa começar com código
            if not re.match(r'^\d{5}\s+', linha):
                continue

            # Extrair todos os números
            numeros = re.findall(r'\d+,\d+|\d+', linha)
            if len(numeros) < 2:
                continue

            try:
                quantidade = float(numeros[-2].replace(",", "."))
                valor = float(numeros[-1].replace(",", "."))
            except ValueError:
                continue

            # Extrair nome do produto
            partes = linha.split()
            produto_tokens = []

            for token in partes[1:]:
                if token in ("UN", "KG", "G", "PCT"):
                    break
                produto_tokens.append(token)

            produto = limpar_produto(" ".join(produto_tokens))

            if not produto:
                continue

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
# UPLOAD
# =========================
files = st.file_uploader(
    "📤 Envie os PDFs do Lince",
    type="pdf",
    accept_multiple_files=True
)

if files:
    if not mes_manual or not setor_manual:
        st.error("Preencha Mês e Setor.")
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
            st.error("❌ Nenhum dado válido foi encontrado nos PDFs.")
        else:
            df = pd.DataFrame(dados)

            st.success(f"✅ {len(df)} registros extraídos")
            st.dataframe(df, use_container_width=True)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False, sheet_name="Perdas")

            st.download_button(
                "⬇️ Baixar Excel",
                data=buffer.getvalue(),
                file_name="perdas_lince.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

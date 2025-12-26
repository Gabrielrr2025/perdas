# -*- coding: utf-8 -*-
import re
import io
import pandas as pd
import streamlit as st
from datetime import datetime
from pypdf import PdfReader

# =========================
# Configuração da página
# =========================
st.set_page_config(
    page_title="Lince → Excel | Perdas",
    page_icon="📄",
    layout="wide"
)

st.title("📄 Lince → Excel (Perdas por Departamento)")
st.caption(
    "Envie PDFs do Lince (Padaria, Confeitaria ou Restaurante) "
    "e gere um Excel padronizado."
)

# =========================
# Funções auxiliares
# =========================
def extrair_periodo(texto):
    """
    Extrai mês e semana a partir do período do PDF
    Ex: 26/11/2025 a 02/12/2025
    """
    m = re.search(r'(\d{2}/\d{2}/\d{4}).+?(\d{2}/\d{2}/\d{4})', texto)
    if not m:
        return None, None

    data_fim = datetime.strptime(m.group(2), "%d/%m/%Y")
    mes = data_fim.strftime("%m/%Y")
    semana = data_fim.isocalendar().week
    return mes, semana


def limpar_produto(nome):
    """
    Remove códigos e unidades duplicadas
    """
    nome = re.sub(r'^\d+\s*-\s*', '', nome)   # remove código
    nome = re.sub(r'\s+(UN|KG|G|PCT)\s*$', '', nome)
    return nome.strip()


def parse_pdf(file):
    reader = PdfReader(file)
    registros = []

    setor_atual = None
    mes = None
    semana = None

    for page in reader.pages:
        texto = page.extract_text()
        if not texto:
            continue

        if mes is None:
            mes, semana = extrair_periodo(texto)

        linhas = texto.splitlines()

        for linha in linhas:
            linha = linha.strip()

            # Setor
            if linha.startswith("000") and "-" in linha and " - " in linha:
                setor_atual = linha.split(" - ", 1)[1].strip()
                continue

            # Ignorar totais
            if linha.startswith("Total"):
                continue

            # Linhas de produto
            m = re.match(
                r'(\d+)\s+(.+?)\s+(UN|KG)\s+([\d,]+)\s+([\d,]+)',
                linha
            )

            if m:
                produto = limpar_produto(m.group(2))
                quantidade = float(m.group(4).replace(",", "."))
                valor = float(m.group(5).replace(",", "."))

                registros.append({
                    "Produto": produto,
                    "Setor": setor_atual,
                    "Mês": mes,
                    "Semana": semana,
                    "Quantidade": quantidade,
                    "Valor": valor
                })

    return registros


# =========================
# Upload de arquivos
# =========================
files = st.file_uploader(
    "📤 Envie os PDFs do Lince",
    type=["pdf"],
    accept_multiple_files=True
)

if files:
    dados = []

    for f in files:
        dados.extend(parse_pdf(f))

    if not dados:
        st.warning("Nenhum dado válido encontrado nos PDFs.")
    else:
        df = pd.DataFrame(dados)

        st.success(f"✅ {len(df)} registros extraídos com sucesso")

        st.dataframe(df, use_container_width=True)

        # Exportar Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Perdas")

        st.download_button(
            "⬇️ Baixar Excel",
            data=buffer.getvalue(),
            file_name="perdas_lince.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )



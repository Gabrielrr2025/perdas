# -*- coding: utf-8 -*-

import re
import io
import unicodedata
from math import ceil
from datetime import datetime

import streamlit as st
import pandas as pd
from pypdf import PdfReader
import xlsxwriter

# =========================
# CONFIGURAÇÃO DA PÁGINA
# =========================
st.set_page_config(
    page_title="Lince → Excel | Perdas",
    page_icon="📄",
    layout="wide"
)

st.title("📄 Lince → Excel (Perdas por Departamento)")
st.caption("Parser robusto do Lince — tolerante a quebras, unidades duplicadas e layouts inconsistentes.")

# =========================
# PARÂMETROS MANUAIS
# =========================
st.subheader("📌 Parâmetros da Planilha")

c1, c2, c3 = st.columns(3)

with c1:
    mes_manual = st.text_input("Mês (MM/AAAA)", value=datetime.today().strftime("%m/%Y"))

with c2:
    semana_manual = st.text_input("Semana", placeholder="Ex: 49")

with c3:
    setor_manual = st.text_input("Setor", placeholder="Padaria / Confeitaria / Restaurante")

st.divider()

# =========================
# FUNÇÕES UTILITÁRIAS
# =========================
def br_to_float(txt: str):
    if not txt:
        return None
    t = txt.strip()
    try:
        return float(t.replace(".", "").replace(",", "."))
    except Exception:
        return None

def is_num_token(tok: str) -> bool:
    return bool(re.fullmatch(r"[0-9][0-9\.\,]*", tok or ""))

def dec_places(tok: str) -> int:
    if "," in tok:
        return len(tok.split(",")[-1])
    return 0

def extract_text(file) -> str:
    reader = PdfReader(file)
    out = []
    for page in reader.pages:
        try:
            out.append(page.extract_text() or "")
        except Exception:
            out.append("")
    return "\n".join(out)

def glue_wrapped_lines(lines):
    glued = []
    i = 0
    while i < len(lines):
        cur = lines[i]
        nxt = lines[i+1] if i + 1 < len(lines) else ""
        cur_nums = sum(is_num_token(t) for t in cur.split())
        nxt_nums = sum(is_num_token(t) for t in nxt.split())

        if cur_nums < 2 and nxt_nums >= 2:
            glued.append((cur + " " + nxt).strip())
            i += 2
        else:
            glued.append(cur)
            i += 1
    return glued

def clean_tokens(tokens):
    out = []
    removed_code = False
    for i, t in enumerate(tokens):
        if not removed_code and i == 0 and re.fullmatch(r"\d{3,6}", t):
            removed_code = True
            continue
        if t in ("UN", "KG", "G", "PCT"):
            continue
        out.append(t)
    return out

# =========================
# PARSER PRINCIPAL (ROBUSTO)
# =========================
def parse_lince(texto):
    linhas = [re.sub(r"\s{2,}", " ", ln).strip() for ln in texto.splitlines()]
    lixo = (
        "Período", "Sub Departamento", "Setor:", "Total do Departamento",
        "Total Geral", "www.grupotecnoweb.com.br", "Curva ABC"
    )
    linhas = [ln for ln in linhas if ln and not any(x in ln for x in lixo)]

    linhas = glue_wrapped_lines(linhas)

    itens = []

    for ln in linhas:
        toks = ln.split()
        if not toks:
            continue

        toks = clean_tokens(toks)
        if len(toks) < 4:
            continue

        idx = len(toks)
        while idx > 0 and is_num_token(toks[idx-1]):
            idx -= 1

        head = toks[:idx]
        tail = toks[idx:]

        if len(tail) < 2:
            continue

        i_qtd = None
        for j in range(len(tail)-1, -1, -1):
            if dec_places(tail[j]) == 3:
                i_qtd = j
                break

        if i_qtd is not None and i_qtd + 1 < len(tail):
            qtd = br_to_float(tail[i_qtd])
            val = br_to_float(tail[i_qtd + 1])
        else:
            qtd = br_to_float(tail[-2])
            val = br_to_float(tail[-1])

        if qtd is None or val is None or qtd < 0 or val < 0:
            continue

        nome = " ".join([t for t in head if not is_num_token(t)]).strip()
        if not nome:
            continue

        itens.append({
            "Produto": nome,
            "Quantidade": round(qtd, 3),
            "Valor": round(val, 2)
        })

    # AGREGAÇÃO
    agg = {}
    for it in itens:
        k = it["Produto"]
        if k not in agg:
            agg[k] = {"Produto": k, "Quantidade": 0.0, "Valor": 0.0}
        agg[k]["Quantidade"] += it["Quantidade"]
        agg[k]["Valor"] += it["Valor"]

    return list(agg.values())

# =========================
# UPLOAD
# =========================
files = st.file_uploader(
    "📤 Envie os PDFs do Lince (Perdas por Departamento)",
    type="pdf",
    accept_multiple_files=True
)

if files:
    if not mes_manual or not setor_manual:
        st.error("Preencha Mês e Setor.")
        st.stop()

    registros = []

    for f in files:
        texto = extract_text(f)
        itens = parse_lince(texto)
        for it in itens:
            registros.append({
                "Produto": it["Produto"],
                "Setor": setor_manual,
                "Mês": mes_manual,
                "Semana": semana_manual,
                "Quantidade": it["Quantidade"],
                "Valor": it["Valor"]
            })

    if not registros:
        st.error("❌ Nenhum dado válido foi identificado.")
        st.stop()

    df = pd.DataFrame(registros)
    st.success(f"✅ {len(df)} produtos consolidados")
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
else:
    st.info("Envie ao menos um PDF para iniciar.")


# -*- coding: utf-8 -*-
import io
import re
from datetime import datetime

import streamlit as st
from pypdf import PdfReader
import xlsxwriter


# =========================
# Config
# =========================
st.set_page_config(page_title="Lince → Excel (Perdas)", page_icon="📄", layout="wide")
st.title("📄 Lince → Excel (Perdas por Departamento)")
st.caption(
    "Envie PDFs do Lince (Perdas por Departamento) e gere Excel padronizado: "
    "Produto | Setor | Mês | Semana | Quantidade | Valor."
)


# =========================
# Constantes
# =========================
SETORES_FIXOS = [
    "Padaria",
    "Lanchonete",
    "Confeitaria Fina",
    "Confeitaria Trad",
    "Restaurante",
    "Frios",
    "Salgados",
]

MESES_PT = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
}


# =========================
# Utilitários
# =========================
def br_to_float(txt: str):
    if txt is None:
        return None
    t = str(txt).strip()
    if not t:
        return None
    try:
        return float(t.replace(".", "").replace(",", "."))
    except Exception:
        return None


def is_num_token(tok: str) -> bool:
    """Verifica se é um token numérico (preço)"""
    return bool(re.fullmatch(r"[0-9]+[\.\,][0-9]{2}", (tok or "").strip()))


def extract_text_with_pypdf(file) -> str:
    reader = PdfReader(file)
    texts = []
    for page in reader.pages:
        try:
            texts.append(page.extract_text() or "")
        except Exception:
            texts.append("")
    return "\n".join(texts)


def parse_periodo(text: str):
    t = " ".join((text or "").split())
    m = re.search(
        r"Per[ií]odo:\s*(\d{2}/\d{2}/\d{4}).*?(\d{2}/\d{2}/\d{4})",
        t, flags=re.IGNORECASE
    )
    if not m:
        return (None, None)
    try:
        return (
            datetime.strptime(m.group(1), "%d/%m/%Y").date(),
            datetime.strptime(m.group(2), "%d/%m/%Y").date(),
        )
    except Exception:
        return (None, None)


def sugestao_mes_semana(dt_ini):
    if not dt_ini:
        hoje = datetime.today().date()
        return MESES_PT.get(hoje.month, ""), (hoje.day - 1) // 7 + 1
    return MESES_PT.get(dt_ini.month, ""), (dt_ini.day - 1) // 7 + 1


def clean_produto_name(nome: str) -> str:
    nome = (nome or "").strip()
    nome = re.sub(r"\s{2,}", " ", nome)
    return nome


# =========================
# PARSER CORRIGIDO
# =========================
def parse_perdas_lince(text: str):
    """
    Parser para 'Perdas por Departamento' (Lince)
    
    Formato da linha:
    CODIGO NOME_PRODUTO UNIDADE UNIDADE PRECO QUANTIDADE VALOR-
    
    Exemplo:
    001681 SALG COQUETEL ASSADO KG KG 69,90 7,64 534,18-
    
    Regras:
    - Quantidade e Valor são os dois últimos números antes do '-'
    - Nome é tudo entre CODIGO e os 3 últimos números (preço, qtd, valor)
    """
    lines = [
        re.sub(r"\s{2,}", " ", (ln or "")).strip()
        for ln in (text or "").splitlines()
    ]

    lixo = (
        "SHOPPING DO PAO", "Perdas por Departamento", "Pag.",
        "Período:", "Periodo:", "UN Preço Qtde Venda", "PreçoQtde Venda",
        "Sub Departamento:", "Setor:",
        "Total do Departamento", "Total Geral",
        "www.grupotecnoweb.com.br", "Lince "
    )

    lines = [ln for ln in lines if ln and not any(k in ln for k in lixo)]
    itens = []

    for ln in lines:
        toks = ln.split()
        if not toks:
            continue

        # Deve ter código do produto
        if not re.fullmatch(r"\d{3,10}", toks[0]):
            continue

        # Deve ter o hífen no final
        if "-" not in toks:
            continue

        idx_hifen = toks.index("-")
        
        # Precisa ter pelo menos: CODIGO NOME PRECO QTD VALOR -
        if idx_hifen < 4:
            continue

        # Antes do hífen temos: [CODIGO, NOME..., PRECO, QTD, VALOR]
        antes = toks[:idx_hifen]
        
        # Os 2 últimos números antes do hífen são QTD e VALOR
        try:
            valor = br_to_float(antes[-1])
            qtd = br_to_float(antes[-2])
        except (IndexError, ValueError):
            continue

        if qtd is None or valor is None:
            continue

        # Remove CODIGO, QTD, VALOR e PREÇO (último número com formato X,XX)
        nome_tokens = antes[1:-2]  # Remove código e os 2 últimos (qtd, valor)
        
        # Remove o preço (último token com formato XX,XX)
        while nome_tokens and is_num_token(nome_tokens[-1]):
            nome_tokens.pop()

        produto = clean_produto_name(" ".join(nome_tokens))
        if not produto:
            continue

        itens.append({
            "produto": produto,
            "quantidade": float(qtd),
            "valor": float(valor)
        })

    # Consolidação por produto
    agg = {}
    for it in itens:
        k = it["produto"]
        if k not in agg:
            agg[k] = {"produto": k, "quantidade": 0.0, "valor": 0.0}
        agg[k]["quantidade"] += it["quantidade"]
        agg[k]["valor"] += it["valor"]

    return sorted(agg.values(), key=lambda x: x["valor"], reverse=True)


def build_excel(rows, setor, mes, semana):
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {"in_memory": True})
    ws = wb.add_worksheet("Dados")

    headers = ["Produto", "Setor", "Mês", "Semana", "Quantidade", "Valor"]
    header_fmt = wb.add_format({"bold": True, "border": 1})
    num3 = wb.add_format({"num_format": "0.000"})
    money = wb.add_format({"num_format": "#,##0.00"})

    for c, h in enumerate(headers):
        ws.write(0, c, h, header_fmt)

    for i, r in enumerate(rows, start=1):
        ws.write(i, 0, r["produto"])
        ws.write(i, 1, setor)
        ws.write(i, 2, mes)
        ws.write_number(i, 3, int(semana))
        ws.write_number(i, 4, round(r["quantidade"], 3), num3)
        ws.write_number(i, 5, round(r["valor"], 2), money)

    ws.set_column(0, 0, 45)
    ws.set_column(1, 1, 20)
    ws.set_column(2, 2, 12)
    ws.set_column(3, 3, 8)
    ws.set_column(4, 4, 12)
    ws.set_column(5, 5, 14)

    wb.close()
    output.seek(0)
    return output.getvalue()


# =========================
# UI
# =========================
uploads = st.file_uploader(
    "Envie 1 ou vários PDFs do Lince (Perdas por Departamento)",
    type=["pdf"],
    accept_multiple_files=True
)

col1, col2, col3 = st.columns(3)
with col1:
    setor = st.selectbox("Setor", SETORES_FIXOS)

sug_mes = MESES_PT.get(datetime.today().month, "")
sug_sem = (datetime.today().day - 1) // 7 + 1

if uploads:
    try:
        base_text = extract_text_with_pypdf(uploads[0])
        dt_ini, _ = parse_periodo(base_text)
        sug_mes, sug_sem = sugestao_mes_semana(dt_ini)
    except Exception:
        pass

with col2:
    mes = st.text_input("Mês", value=sug_mes)

with col3:
    semana = st.text_input("Semana", value=str(sug_sem))

st.markdown("---")

if uploads:
    # Validações
    if not mes.strip():
        st.error("⚠️ Por favor, preencha o mês!")
        st.stop()
    
    if not semana.strip().isdigit():
        st.error("⚠️ A semana deve ser um número!")
        st.stop()

    all_rows = []
    progress = st.progress(0)
    erros = []

    for i, f in enumerate(uploads):
        try:
            text = extract_text_with_pypdf(f)
            rows = parse_perdas_lince(text)
            if not rows:
                erros.append(f"⚠️ Nenhum dado encontrado em: {f.name}")
            all_rows.extend(rows)
        except Exception as e:
            erros.append(f"❌ Erro ao processar {f.name}: {str(e)}")
        
        progress.progress((i + 1) / len(uploads))

    # Mostra erros se houver
    if erros:
        for erro in erros:
            st.warning(erro)

    if not all_rows:
        st.error("❌ Nenhum dado foi extraído dos PDFs. Verifique se são arquivos do Lince (Perdas por Departamento).")
        st.stop()

    # Consolidação entre PDFs
    agg = {}
    for r in all_rows:
        k = r["produto"]
        if k not in agg:
            agg[k] = {"produto": k, "quantidade": 0.0, "valor": 0.0}
        agg[k]["quantidade"] += r["quantidade"]
        agg[k]["valor"] += r["valor"]

    final_rows = sorted(agg.values(), key=lambda x: x["valor"], reverse=True)

    st.success(f"✅ {len(final_rows)} produtos processados com sucesso!")
    
    st.subheader("Prévia dos dados")
    st.dataframe(
        [{
            "Produto": r["produto"],
            "Quantidade": round(r["quantidade"], 3),
            "Valor": f"R$ {round(r['valor'], 2):.2f}"
        } for r in final_rows],
        use_container_width=True,
        height=420
    )

    if st.button("📥 Gerar Excel", type="primary"):
        try:
            excel = build_excel(final_rows, setor, mes.strip(), semana.strip())
            nome = f"perdas_{setor}_{mes}_sem{semana}.xlsx".replace(" ", "_")
            st.download_button(
                "⬇️ Baixar Excel", 
                data=excel, 
                file_name=nome,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"❌ Erro ao gerar Excel: {str(e)}")

else:
    st.info("📤 Envie pelo menos um PDF para começar.")

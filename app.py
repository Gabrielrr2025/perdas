import io
import re
import unicodedata
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
    return bool(re.fullmatch(r"[0-9][0-9\.\,]*", (tok or "").strip()))


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
    nome = re.sub(r"^\-\s*", "", nome)
    return nome


# =========================
# PARSER CORRIGIDO
# =========================
def parse_perdas_lince(text: str):
    """
    Parser robusto para 'Perdas por Departamento' (Lince)

    Regras:
    - Quantidade = número logo após o '-'
    - Valor = último número da linha
    - Nome = tudo entre o código e o preço
    - Mantém 'KG' quando fizer parte do nome
    """
    lines = [
        re.sub(r"\s{2,}", " ", (ln or "")).strip()
        for ln in (text or "").splitlines()
    ]

    lixo = (
        "SHOPPING DO PAO", "Perdas por Departamento", "Pag.",
        "Período:", "Periodo:", "UN Preço Qtde Venda",
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

        # código do produto
        if not re.fullmatch(r"\d{3,10}", toks[0]):
            continue

        if "-" not in toks:
            continue

        idx_hifen = toks.index("-")
        if idx_hifen + 2 >= len(toks):
            continue

        qtd = br_to_float(toks[idx_hifen + 1])
        valor = br_to_float(toks[idx_hifen + 2])

        if qtd is None or valor is None:
            continue

        # antes do hífen: COD + NOME + UNIDADE + PREÇO
        antes = toks[1:idx_hifen]

        # remove o preço (último número antes do hífen)
        while antes and is_num_token(antes[-1]):
            antes.pop()

        produto = clean_produto_name(" ".join(antes))
        if not produto:
            continue

        itens.append({
            "produto": produto,
            "quantidade": float(qtd),
            "valor": float(valor)
        })

    # consolidação por produto
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
    base_text = extract_text_with_pypdf(uploads[0])
    dt_ini, _ = parse_periodo(base_text)
    sug_mes, sug_sem = sugestao_mes_semana(dt_ini)

with col2:
    mes = st.text_input("Mês", value=sug_mes)

with col3:
    semana = st.text_input("Semana", value=str(sug_sem))

st.markdown("---")

if uploads:
    all_rows = []
    for f in uploads:
        text = extract_text_with_pypdf(f)
        all_rows.extend(parse_perdas_lince(text))

    # consolidação entre PDFs
    agg = {}
    for r in all_rows:
        k = r["produto"]
        if k not in agg:
            agg[k] = {"produto": k, "quantidade": 0.0, "valor": 0.0}
        agg[k]["quantidade"] += r["quantidade"]
        agg[k]["valor"] += r["valor"]

    final_rows = sorted(agg.values(), key=lambda x: x["valor"], reverse=True)

    st.subheader("Prévia dos dados")
    st.dataframe(
        [{
            "Produto": r["produto"],
            "Quantidade": round(r["quantidade"], 3),
            "Valor": round(r["valor"], 2)
        } for r in final_rows],
        use_container_width=True,
        height=420
    )

    if st.button("Gerar Excel"):
        excel = build_excel(final_rows, setor, mes.strip(), semana.strip())
        nome = f"perdas_{setor}_{mes}_sem{semana}.xlsx".replace(" ", "_")
        st.download_button("⬇️ Baixar Excel", data=excel, file_name=nome)

else:
    st.info("Envie pelo menos um PDF para começar.")

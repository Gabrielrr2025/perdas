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
st.caption("Envie PDFs do Lince (Perdas por Departamento) e gere Excel padronizado: Produto | Setor | Mês | Semana | Quantidade | Valor.")


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
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
    7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
}


# =========================
# Utilitários
# =========================
def _norm(s: str) -> str:
    s = unicodedata.normalize("NFD", s or "").encode("ascii", "ignore").decode("ascii")
    return (s or "").upper()

def br_to_float(txt: str):
    if txt is None:
        return None
    t = str(txt).strip()
    if not t:
        return None
    # formato BR: 3.896,54
    try:
        return float(t.replace(".", "").replace(",", "."))
    except Exception:
        return None

def is_num_token(tok: str) -> bool:
    # aceita 12,07 / 3.896,54 / 7,00
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
    """
    Tenta extrair datas do campo 'Período:' do Lince.
    Retorna (dt_ini, dt_fim) ou (None, None).
    """
    t = " ".join((text or "").split())
    m = re.search(r"Per[ií]odo:\s*(\d{2}/\d{2}/\d{4}).*?(\d{2}/\d{2}/\d{4})", t, flags=re.IGNORECASE)
    if not m:
        return (None, None)
    try:
        dt_ini = datetime.strptime(m.group(1), "%d/%m/%Y").date()
        dt_fim = datetime.strptime(m.group(2), "%d/%m/%Y").date()
        return (dt_ini, dt_fim)
    except Exception:
        return (None, None)

def sugestao_mes_semana(dt_ini):
    """
    Sugere Mês (nome PT) e Semana (1-5) com base na data inicial.
    """
    if not dt_ini:
        # fallback: mês atual e semana 1
        hoje = datetime.today().date()
        mes = MESES_PT.get(hoje.month, f"{hoje.month:02d}/{hoje.year}")
        semana = (hoje.day - 1) // 7 + 1
        return mes, semana

    mes = MESES_PT.get(dt_ini.month, f"{dt_ini.month:02d}/{dt_ini.year}")
    semana = (dt_ini.day - 1) // 7 + 1
    return mes, semana

def clean_produto_name(nome: str) -> str:
    nome = (nome or "").strip()
    nome = re.sub(r"\s{2,}", " ", nome)
    nome = re.sub(r"^\-\s*", "", nome).strip()
    return nome

def parse_perdas_lince(text: str):
    """
    Parser específico para 'Perdas por Departamento' (Lince).
    Espera linhas do tipo:
      001685 SANDUICHE A METRO KG KG 65,90 - 0,66 43,49
      004090 PAO DE QUEIJO UN UN 7,99 - 1,00 7,99

    Retorna lista agregada:
      [{"produto": str, "quantidade": float, "valor": float}, ...]
    """
    lines = [re.sub(r"\s{2,}", " ", (ln or "")).strip() for ln in (text or "").splitlines()]
    # remove lixo e cabeçalhos
    lixo = (
        "SHOPPING DO PAO", "Perdas por Departamento", "Pag.", "Período:", "Periodo:",
        "UN Preço Qtde Venda", "Sub Departamento:", "Setor:",
        "Total do Departamento", "Total Geral", "www.grupotecnoweb.com.br", "Lince "
    )
    lines = [ln for ln in lines if ln and not any(k in ln for k in lixo)]

    itens = []

    for ln in lines:
        toks = ln.split()
        if not toks:
            continue

        # primeira coluna costuma ser o código numérico do produto
        if not re.fullmatch(r"\d{3,10}", toks[0]):
            continue

        # coletar tokens numéricos do final (ignorando "-")
        # Ex.: [..., "65,90", "-", "0,66", "43,49"]
        num_tokens = []
        for tok in reversed(toks):
            if tok == "-":
                continue
            if is_num_token(tok):
                num_tokens.append(tok)
                if len(num_tokens) >= 3:
                    break
            else:
                # assim que bater num texto, para de voltar
                # (mas só depois de já ter encontrado algum número)
                if num_tokens:
                    break

        if len(num_tokens) < 2:
            continue

        # num_tokens está reverso (ex.: ["43,49","0,66","65,90"])
        valor_tok = num_tokens[0]
        qtd_tok = num_tokens[1]
        valor = br_to_float(valor_tok)
        qtd = br_to_float(qtd_tok)
        if valor is None or qtd is None:
            continue
        if valor < 0 or qtd < 0:
            continue

        # tentar extrair nome do produto:
        # padrão: [COD, ...NOME..., UN|KG, UN|KG, PRECO, -, QTD, VALOR]
        # nome = tokens entre COD e primeiro UN/KG
        nome_tokens = []
        unidades = {"UN", "KG"}

        # procurar o primeiro token UN/KG depois do código
        idx_un = None
        for i in range(1, len(toks)):
            if toks[i] in unidades:
                idx_un = i
                break

        if idx_un is not None and idx_un > 1:
            nome_tokens = toks[1:idx_un]
        else:
            # fallback: tira o código e remove a cauda numérica
            # acha posição do último número (valor) e pega o que vier antes
            # (ainda pode conter UN/KG, mas funciona como fallback)
            last_num_idx = None
            for i in range(len(toks) - 1, 0, -1):
                if is_num_token(toks[i]):
                    last_num_idx = i
                    break
            if last_num_idx and last_num_idx > 1:
                nome_tokens = toks[1:last_num_idx - 1]  # -1 p/ tirar qtd também
            else:
                nome_tokens = toks[1:]

        produto = clean_produto_name(" ".join(nome_tokens))
        if not produto or not re.search(r"[A-Za-zÀ-ÖØ-öø-ÿ]", produto):
            continue

        itens.append({"produto": produto, "quantidade": float(qtd), "valor": float(valor)})

    # agrega por produto (somando qtd e valor)
    agg = {}
    for it in itens:
        k = it["produto"]
        if k not in agg:
            agg[k] = {"produto": k, "quantidade": 0.0, "valor": 0.0}
        agg[k]["quantidade"] += it["quantidade"]
        agg[k]["valor"] += it["valor"]

    # ordena por valor desc
    return sorted(agg.values(), key=lambda x: x["valor"], reverse=True)


def build_excel(rows, setor, mes, semana):
    """
    Gera Excel em memória no padrão:
      Produto | Setor | Mês | Semana | Quantidade | Valor
    """
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {"in_memory": True})
    ws = wb.add_worksheet("Dados")

    headers = ["Produto", "Setor", "Mês", "Semana", "Quantidade", "Valor"]
    header_fmt = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1})
    num3 = wb.add_format({"num_format": "0.000"})
    money = wb.add_format({"num_format": "#,##0.00"})

    for c, h in enumerate(headers):
        ws.write(0, c, h, header_fmt)

    for i, r in enumerate(rows, start=1):
        ws.write(i, 0, r["produto"])
        ws.write(i, 1, setor)
        ws.write(i, 2, mes)
        ws.write_number(i, 3, int(semana) if str(semana).strip().isdigit() else 0)
        ws.write_number(i, 4, round(float(r["quantidade"]), 3), num3)
        ws.write_number(i, 5, round(float(r["valor"]), 2), money)

    ws.set_column(0, 0, 45)
    ws.set_column(1, 1, 18)
    ws.set_column(2, 2, 12)
    ws.set_column(3, 3, 9)
    ws.set_column(4, 4, 12)
    ws.set_column(5, 5, 12)

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

colA, colB, colC = st.columns([1.2, 1.0, 1.0])
with colA:
    setor = st.selectbox("Setor (fixo)", SETORES_FIXOS, index=0)

# Sugestões: usa o primeiro PDF como base (se existir)
sug_mes = MESES_PT.get(datetime.today().month, "Mês")
sug_sem = (datetime.today().day - 1) // 7 + 1

if uploads:
    base_text = extract_text_with_pypdf(uploads[0])
    dt_ini, dt_fim = parse_periodo(base_text)
    sug_mes, sug_sem = sugestao_mes_semana(dt_ini)

with colB:
    mes = st.text_input("Mês (digitável)", value=str(sug_mes))

with colC:
    semana = st.text_input("Semana (digitável)", value=str(sug_sem))

st.markdown("---")

if uploads:
    all_rows = []
    detalhes = []

    for f in uploads:
        text = extract_text_with_pypdf(f)
        rows = parse_perdas_lince(text)
        all_rows.extend(rows)

        dt_ini, dt_fim = parse_periodo(text)
        detalhes.append({
            "arquivo": f.name,
            "periodo": f"{dt_ini.strftime('%d/%m/%Y') if dt_ini else '?'} a {dt_fim.strftime('%d/%m/%Y') if dt_fim else '?'}",
            "itens": len(rows),
            "valor_total": sum(r["valor"] for r in rows) if rows else 0.0
        })

    # consolida entre PDFs também
    agg = {}
    for r in all_rows:
        k = r["produto"]
        if k not in agg:
            agg[k] = {"produto": k, "quantidade": 0.0, "valor": 0.0}
        agg[k]["quantidade"] += r["quantidade"]
        agg[k]["valor"] += r["valor"]
    final_rows = sorted(agg.values(), key=lambda x: x["valor"], reverse=True)

    st.subheader("Prévia (consolidado)")
    st.write(f"Arquivos: **{len(uploads)}** | Produtos (após consolidação): **{len(final_rows)}**")

    # tabela simples (top 50)
    preview = final_rows[:50]
    st.dataframe(
        [{
            "Produto": r["produto"],
            "Quantidade": round(r["quantidade"], 3),
            "Valor": round(r["valor"], 2),
        } for r in preview],
        use_container_width=True,
        height=420
    )

    with st.expander("Detalhes por arquivo"):
        st.dataframe(
            [{
                "Arquivo": d["arquivo"],
                "Período (detectado)": d["periodo"],
                "Itens": d["itens"],
                "Valor total (R$)": round(d["valor_total"], 2),
            } for d in detalhes],
            use_container_width=True
        )

    st.markdown("---")
    if st.button("Gerar Excel (.xlsx)"):
        excel_bytes = build_excel(final_rows, setor=setor, mes=mes.strip(), semana=semana.strip())
        fname = f"perdas_{setor}_{mes.strip()}_sem{semana.strip()}.xlsx".replace(" ", "_")
        st.download_button("⬇️ Baixar Excel", data=excel_bytes, file_name=fname)

else:
    st.info("Envie pelo menos 1 PDF para começar.")

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
    """Converte string BR (1.234,56) para float"""
    if txt is None:
        return None
    t = str(txt).strip()
    if not t:
        return None
    try:
        return float(t.replace(".", "").replace(",", "."))
    except Exception:
        return None


def extract_text_with_pypdf(file) -> str:
    """Extrai texto do PDF"""
    reader = PdfReader(file)
    texts = []
    for page in reader.pages:
        try:
            texts.append(page.extract_text() or "")
        except Exception:
            texts.append("")
    return "\n".join(texts)


def parse_periodo(text: str):
    """Extrai período do relatório"""
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
    """Sugere mês e semana baseado na data"""
    if not dt_ini:
        hoje = datetime.today().date()
        return MESES_PT.get(hoje.month, ""), (hoje.day - 1) // 7 + 1
    return MESES_PT.get(dt_ini.month, ""), (dt_ini.day - 1) // 7 + 1


# =========================
# PARSER MELHORADO
# =========================
def parse_perdas_lince(text: str):
    """
    Parser robusto para PDFs do Lince (Perdas por Departamento)
    
    Formato esperado:
    CODIGO NOME_PRODUTO [UNIDADE] PRECO QUANTIDADE VALOR-
    
    Exemplos:
    001681 SALG COQUETEL ASSADO KG KG 69,90 7,64 534,18-
    000415 PAO FERM PROL PROV/PAT ALHO/PRES/QUEI PRATO KGKG 60,90 6,26 380,99-
    001500 BOLO ABACAXI KG KG 36,90 2,34 86,42-
    """
    
    lines = text.splitlines()
    
    # Palavras que indicam linhas inúteis
    lixo = [
        "SHOPPING DO PAO", "Perdas por Departamento", "Pag.",
        "Período:", "Periodo:", "UN Preço", "Qtde Venda", "PreçoQtde",
        "Sub Departamento:", "Setor:",
        "Total do Departamento", "Total Geral",
        "www.grupotecnoweb.com.br", "Lince", "MATRIZ"
    ]
    
    itens = []
    
    for linha in lines:
        linha = linha.strip()
        if not linha:
            continue
            
        # Pula linhas de cabeçalho/rodapé
        if any(palavra in linha for palavra in lixo):
            continue
        
        # Regex para extrair dados da linha
        # Formato: CODIGO NOME PRECO QTD VALOR-
        # Captura: código (6 dígitos), depois tudo até encontrar 3 números no formato BR
        match = re.match(
            r'^(\d{6})\s+'  # Código do produto (6 dígitos)
            r'(.+?)'  # Nome do produto (não-greedy)
            r'\s+(\d+[,\.]\d{2})'  # Preço (formato: 12,34 ou 12.34)
            r'\s+(\d+[,\.]\d{2,3})'  # Quantidade (formato: 1,23 ou 12,345)
            r'\s+(\d+[,\.]\d{2})-',  # Valor (formato: 123,45)
            linha
        )
        
        if match:
            codigo = match.group(1)
            nome = match.group(2).strip()
            preco = match.group(3)
            qtd_str = match.group(4)
            valor_str = match.group(5)
            
            # Limpa o nome (remove unidades duplicadas e espaços extras)
            nome = re.sub(r'\s+', ' ', nome)
            
            # Converte quantidade e valor
            qtd = br_to_float(qtd_str)
            valor = br_to_float(valor_str)
            
            if qtd is not None and valor is not None and qtd > 0 and valor > 0:
                itens.append({
                    "codigo": codigo,
                    "produto": nome,
                    "quantidade": float(qtd),
                    "valor": float(valor)
                })
    
    # Consolidação por produto
    agg = {}
    for item in itens:
        chave = item["produto"]
        if chave not in agg:
            agg[chave] = {
                "produto": chave,
                "quantidade": 0.0,
                "valor": 0.0
            }
        agg[chave]["quantidade"] += item["quantidade"]
        agg[chave]["valor"] += item["valor"]
    
    # Ordena por valor (maior primeiro)
    resultado = sorted(agg.values(), key=lambda x: x["valor"], reverse=True)
    
    return resultado


def build_excel(rows, setor, mes, semana):
    """Gera arquivo Excel com os dados"""
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {"in_memory": True})
    ws = wb.add_worksheet("Dados")

    headers = ["Produto", "Setor", "Mês", "Semana", "Quantidade", "Valor"]
    header_fmt = wb.add_format({"bold": True, "border": 1, "bg_color": "#D3D3D3"})
    num3 = wb.add_format({"num_format": "0.000", "border": 1})
    money = wb.add_format({"num_format": "#,##0.00", "border": 1})
    text_fmt = wb.add_format({"border": 1})
    center_fmt = wb.add_format({"border": 1, "align": "center"})

    # Cabeçalhos
    for c, h in enumerate(headers):
        ws.write(0, c, h, header_fmt)

    # Dados
    for i, r in enumerate(rows, start=1):
        ws.write(i, 0, r["produto"], text_fmt)
        ws.write(i, 1, setor, text_fmt)
        ws.write(i, 2, mes, text_fmt)
        ws.write_number(i, 3, int(semana), center_fmt)
        ws.write_number(i, 4, round(r["quantidade"], 3), num3)
        ws.write_number(i, 5, round(r["valor"], 2), money)

    # Largura das colunas
    ws.set_column(0, 0, 50)  # Produto
    ws.set_column(1, 1, 20)  # Setor
    ws.set_column(2, 2, 12)  # Mês
    ws.set_column(3, 3, 8)   # Semana
    ws.set_column(4, 4, 12)  # Quantidade
    ws.set_column(5, 5, 14)  # Valor

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
    progress_bar = st.progress(0)
    status_text = st.empty()
    erros = []

    for i, f in enumerate(uploads):
        status_text.text(f"Processando: {f.name}...")
        try:
            text = extract_text_with_pypdf(f)
            rows = parse_perdas_lince(text)
            
            if not rows:
                erros.append(f"⚠️ Nenhum dado encontrado em: {f.name}")
            else:
                all_rows.extend(rows)
                
        except Exception as e:
            erros.append(f"❌ Erro ao processar {f.name}: {str(e)}")
        
        progress_bar.progress((i + 1) / len(uploads))
    
    status_text.empty()
    progress_bar.empty()

    # Mostra erros se houver
    if erros:
        for erro in erros:
            st.warning(erro)

    if not all_rows:
        st.error("❌ Nenhum dado foi extraído dos PDFs. Verifique se são arquivos do Lince (Perdas por Departamento).")
        st.info("💡 O arquivo deve conter linhas no formato: CODIGO PRODUTO UNIDADE PRECO QUANTIDADE VALOR-")
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

    # Calcula totais
    total_qtd = sum(r["quantidade"] for r in final_rows)
    total_valor = sum(r["valor"] for r in final_rows)

    st.success(f"✅ {len(final_rows)} produtos processados | Total: R$ {total_valor:,.2f}")
    
    st.subheader("Prévia dos dados")
    st.dataframe(
        [{
            "Produto": r["produto"],
            "Quantidade": f"{r['quantidade']:.3f}",
            "Valor": f"R$ {r['valor']:.2f}"
        } for r in final_rows],
        use_container_width=True,
        height=420
    )

    if st.button("📥 Gerar Excel", type="primary", use_container_width=True):
        try:
            excel = build_excel(final_rows, setor, mes.strip(), semana.strip())
            nome = f"perdas_{setor}_{mes}_sem{semana}.xlsx".replace(" ", "_")
            st.download_button(
                "⬇️ Baixar Excel", 
                data=excel, 
                file_name=nome,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.balloons()
        except Exception as e:
            st.error(f"❌ Erro ao gerar Excel: {str(e)}")

else:
    st.info("📤 Envie pelo menos um PDF para começar.")
    
    # Instruções
    with st.expander("ℹ️ Como usar"):
        st.markdown("""
        1. **Faça upload** de um ou mais PDFs do Lince (Perdas por Departamento)
        2. **Selecione o setor** no dropdown
        3. **Confira** o mês e semana (preenchidos automaticamente)
        4. **Visualize** a prévia dos dados
        5. **Clique em "Gerar Excel"** para baixar o arquivo
        
        **Formato esperado do PDF:**
        ```
        CODIGO NOME_PRODUTO UNIDADE PRECO QUANTIDADE VALOR-
        ```
        
        **Exemplo:**
        ```
        001681 SALG COQUETEL ASSADO KG KG 69,90 7,64 534,18-
        ```
        """)

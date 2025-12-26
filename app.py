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
# FUNÇÕES MELHORADAS
# =========================

def limpar_produto(texto):
    """Remove código inicial e unidades do final do nome do produto"""
    # Remove código inicial (5 dígitos)
    texto = re.sub(r'^\d{5}\s+', '', texto)
    # Remove unidades do final
    texto = re.sub(r'\s+(UN|KG|G|PCT|L|ML|CX)(\s+(UN|KG|G|PCT|L|ML|CX))*$', '', texto, flags=re.IGNORECASE)
    return texto.strip()

def extrair_numeros_finais(linha):
    """Extrai os dois últimos números da linha (quantidade e valor)"""
    # Busca todos os números (inteiros e decimais)
    numeros = re.findall(r'\d+[,\.]\d+|\d+', linha)
    
    if len(numeros) < 2:
        return None, None
    
    try:
        # Pega os dois últimos números
        quantidade_str = numeros[-2].replace(",", ".")
        valor_str = numeros[-1].replace(",", ".")
        
        quantidade = float(quantidade_str)
        valor = float(valor_str)
        
        return quantidade, valor
    except (ValueError, IndexError):
        return None, None

def extrair_produto(linha):
    """Extrai o nome do produto da linha"""
    # Remove código inicial
    linha_sem_codigo = re.sub(r'^\d{5}\s+', '', linha)
    
    # Separa tokens
    tokens = linha_sem_codigo.split()
    produto_tokens = []
    
    # Pega tokens até encontrar unidade ou número
    for token in tokens:
        # Se encontrar unidade, para
        if token.upper() in ("UN", "KG", "G", "PCT", "L", "ML", "CX"):
            break
        # Se for um número (possível quantidade/valor), para
        if re.match(r'^\d+[,\.]?\d*$', token):
            break
        produto_tokens.append(token)
    
    produto = " ".join(produto_tokens).strip()
    return produto

def parse_pdf(file, mes, semana, setor):
    """Extrai dados do PDF do Lince"""
    reader = PdfReader(file)
    registros = []
    linhas_debug = []
    
    for page_num, page in enumerate(reader.pages, 1):
        texto = page.extract_text()
        if not texto:
            continue
        
        for linha in texto.splitlines():
            linha = linha.strip()
            
            # Debug: armazena linha original
            if linha:
                linhas_debug.append(linha)
            
            # Pula linhas vazias
            if not linha:
                continue
            
            # Verifica se linha começa com código de 5 dígitos
            match_codigo = re.match(r'^(\d{5})\s+(.+)$', linha)
            if not match_codigo:
                continue
            
            codigo = match_codigo.group(1)
            resto_linha = match_codigo.group(2)
            
            # Extrai quantidade e valor
            quantidade, valor = extrair_numeros_finais(linha)
            if quantidade is None or valor is None:
                continue
            
            # Extrai produto
            produto = extrair_produto(linha)
            if not produto or len(produto) < 3:
                continue
            
            # Adiciona registro
            registros.append({
                "Código": codigo,
                "Produto": produto,
                "Setor": setor,
                "Mês": mes,
                "Semana": semana,
                "Quantidade": quantidade,
                "Valor": valor
            })
    
    return registros, linhas_debug

# =========================
# DEBUG MODE
# =========================
debug_mode = st.sidebar.checkbox("🔍 Modo Debug", value=False)

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
        st.error("⚠️ Preencha os campos Mês e Setor antes de processar.")
    else:
        dados = []
        todas_linhas_debug = []
        
        # Processa cada arquivo
        for idx, f in enumerate(files, 1):
            st.info(f"📄 Processando arquivo {idx}/{len(files)}: {f.name}")
            registros, linhas_debug = parse_pdf(f, mes_manual, semana_manual, setor_manual)
            dados.extend(registros)
            todas_linhas_debug.extend(linhas_debug)
        
        # Modo Debug
        if debug_mode and todas_linhas_debug:
            st.subheader("🔍 Debug: Linhas extraídas do PDF")
            st.text_area(
                "Primeiras 50 linhas do PDF",
                "\n".join(todas_linhas_debug[:50]),
                height=300
            )
        
        # Verifica se encontrou dados
        if not dados:
            st.error("❌ Nenhum dado válido foi encontrado nos PDFs.")
            st.warning("""
            **Possíveis causas:**
            - O formato do PDF não corresponde ao esperado
            - As linhas não começam com código de 5 dígitos
            - Não há números de quantidade e valor nas linhas
            
            **Ative o Modo Debug** na barra lateral para ver as linhas extraídas do PDF.
            """)
        else:
            # Cria DataFrame
            df = pd.DataFrame(dados)
            
            # Estatísticas
            st.success(f"✅ {len(df)} registros extraídos com sucesso!")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total de Produtos", len(df))
            with col2:
                st.metric("Quantidade Total", f"{df['Quantidade'].sum():.2f}")
            with col3:
                st.metric("Valor Total", f"R$ {df['Valor'].sum():.2f}")
            
            st.divider()
            
            # Visualização dos dados
            st.subheader("📊 Dados Extraídos")
            st.dataframe(df, use_container_width=True)
            
            # Resumo por produto
            if st.checkbox("📈 Ver resumo por produto"):
                resumo = df.groupby("Produto").agg({
                    "Quantidade": "sum",
                    "Valor": "sum"
                }).sort_values("Valor", ascending=False)
                st.dataframe(resumo, use_container_width=True)
            
            # Exportar Excel
            st.divider()
            buffer = io.BytesIO()
            
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                # Sheet principal
                df.to_excel(writer, index=False, sheet_name="Perdas")
                
                # Sheet de resumo
                resumo = df.groupby("Produto").agg({
                    "Quantidade": "sum",
                    "Valor": "sum"
                }).reset_index()
                resumo.to_excel(writer, index=False, sheet_name="Resumo")
                
                # Formatação
                workbook = writer.book
                money_fmt = workbook.add_format({'num_format': 'R$ #,##0.00'})
                num_fmt = workbook.add_format({'num_format': '#,##0.00'})
                
                for sheet_name in ["Perdas", "Resumo"]:
                    worksheet = writer.sheets[sheet_name]
                    worksheet.set_column('F:F', 12, num_fmt)  # Quantidade
                    worksheet.set_column('G:G', 15, money_fmt)  # Valor
            
            st.download_button(
                "⬇️ Baixar Excel",
                data=buffer.getvalue(),
                file_name=f"perdas_lince_{setor_manual.lower().replace(' ', '_')}_{mes_manual.replace('/', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

st.divider()
st.caption("💡 Dica: Use o Modo Debug para verificar como as linhas estão sendo extraídas do PDF.")

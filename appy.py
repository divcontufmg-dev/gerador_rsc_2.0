import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import zipfile
from datetime import datetime

# Configuração da página
st.set_page_config(page_title="Gerador Automático RSC", layout="wide")

st.title("Gerador de Declarações RSC em Lote")
st.write("Faça o upload de uma planilha Excel contendo os dados dos servidores para gerar as declarações automaticamente.")

# Mapeamento de meses para a formatação da data
meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", 
         "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]

def gerar_documento(nome, cargo, siape, unidade, sistemas_selecionados):
    doc = Document()
    
    # Título
    titulo = doc.add_paragraph()
    run_titulo = titulo.add_run("DECLARAÇÃO")
    run_titulo.bold = True
    run_titulo.font.size = Pt(14)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph("")
    
    # Primeiro parágrafo
    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    p1.add_run("Declaramos, para os devidos fins de comprovação junto ao processo de Reconhecimento de Saberes e Competências (RSC), em atendimento ao disposto no Anexo IV do Decreto nº 13.048, de 3 de julho de 2026, que o(a) servidor(a) ")
    
    p1.add_run(nome).bold = True
    p1.add_run(", ocupante do cargo de ")
    p1.add_run(cargo).bold = True
    p1.add_run(", matrícula SIAPE nº ")
    p1.add_run(siape).bold = True
    p1.add_run(", lotado(a) no(a) ")
    p1.add_run(unidade).bold = True
    p1.add_run(", possui ou possuiu o perfil de operador/executor desempenhando atividades que demandam acesso e utilização dos seguintes sistemas estruturantes do governo federal:")
    
    # Lista numérica de sistemas
    for i, sis in enumerate(sistemas_selecionados, 1):
        p_sis = doc.add_paragraph(f"{i})\t{sis};")
        p_sis.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # Segundo parágrafo
    p2 = doc.add_paragraph("Esta declaração é a expressão da verdade e destina-se exclusivamente à instrução de processo administrativo de concessão da Gratificação por Reconhecimento de Saberes e Competências (GRSC).")
    p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    doc.add_paragraph("")
    
    # Data
    hoje = datetime.now()
    data_atual = f"Belo Horizonte, {hoje.day} de {meses[hoje.month - 1]} de {hoje.year}"
    
    p_data = doc.add_paragraph(data_atual)
    p_data.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    doc.add_paragraph("\n")
    
    # Assinatura
    assinatura = doc.add_paragraph()
    assinatura.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_ass = assinatura.add_run("ELÍZIO MARCOS DOS REIS\n")
    run_ass.bold = True
    assinatura.add_run("Diretor do Departamento de Contabilidade e Finanças\n")
    assinatura.add_run("Pró-Reitoria de Planejamento e Desenvolvimento - UFMG")
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def gerar_zip_declaracoes(df, sistemas_selecionados):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for index, row in df.iterrows():
            nome = str(row.get('nome', f'Servidor_{index}'))
            
            # Novo mapeamento de colunas baseado na planilha enviada
            cargo = str(row.get('unidade', ''))
            siape = str(row.get('cpf', '')) 
            unidade_dept = str(row.get('uadnome', ''))
            
            docx_buffer = gerar_documento(nome, cargo, siape, unidade_dept, sistemas_selecionados)
            nome_arquivo = f"Declaracao_RSC_{nome.replace(' ', '_')}.docx"
            zip_file.writestr(nome_arquivo, docx_buffer.getvalue())
            
    zip_buffer.seek(0)
    return zip_buffer

st.subheader("1. Configurações da Declaração")
st.write("Quais sistemas devem constar nas declarações deste lote?")

col_sys1, col_sys2, col_sys3 = st.columns(3)
with col_sys1:
    siafi = st.checkbox("SIAFI", value=True)
with col_sys2:
    scdp = st.checkbox("SCDP", value=True)
with col_sys3:
    tesouro = st.checkbox("Tesouro Gerencial")

st.divider()

st.subheader("2. Carregar Dados")
st.info("A planilha está mapeada para utilizar as colunas: **NOME**, **UNIDADE** (como Cargo), **CPF** (como SIAPE) e **UADNOME** (como Departamento).")

uploaded_file = st.file_uploader("Envie a planilha Excel (.xlsx)", type=["xlsx", "xls"])

if uploaded_file:
    # Tratando as células vazias para não renderizarem o valor zero durante a extração
    df = pd.read_excel(uploaded_file, dtype=str).fillna("")
    
    df.columns = df.columns.str.strip().str.lower()
    
    # Atualizando o filtro para buscar as novas colunas
    colunas_obrigatorias = ['nome', 'unidade', 'cpf', 'uadnome']
    colunas_presentes = [col for col in colunas_obrigatorias if col in df.columns]
    
    if len(colunas_presentes) < 4:
        st.error(f"Erro: A planilha enviada não contém todas as colunas esperadas. Colunas encontradas: {', '.join(colunas_presentes)}")
    else:
        st.success(f"Planilha carregada com sucesso! {len(df)} servidores encontrados.")
        st.dataframe(df[['nome', 'unidade', 'cpf', 'uadnome']].head())
        
        sistemas_selecionados = []
        if siafi: sistemas_selecionados.append("Sistema de Administração Financeira do Governo Federal (SIAFI)")
        if scdp: sistemas_selecionados.append("Sistema Concessão de Diárias e Passagens (SCDP)")
        if tesouro: sistemas_selecionados.append("Tesouro Gerencial")
        
        if not sistemas_selecionados:
            st.warning("⚠️ Selecione pelo menos um sistema estruturante na etapa 1.")
        else:
            if st.button("Gerar Declarações (Arquivo ZIP)", type="primary"):
                with st.spinner("Gerando documentos..."):
                    zip_file = gerar_zip_declaracoes(df, sistemas_selecionados)
                    
                st.download_button(
                    label="📦 Baixar Lote Completo (ZIP)",
                    data=zip_file,
                    file_name="Declaracoes_RSC_Lote_Atualizado.zip",
                    mime="application/zip",
                    type="primary",
                    use_container_width=True
                )

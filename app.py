import streamlit as st
import pandas as pd
from datetime import datetime
import io
import zipfile
import re
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Configuração da página
st.set_page_config(
    page_title="Gerador de Declarações - UFF",
    page_icon="📄",
    layout="wide"
)

# Funções auxiliares (mantidas iguais às originais)
# ... suas funções formatar_nome, obter_grau_curso, etc ...

# Interface principal
st.title("Gerador de Declarações e Cronograma de Defesas")
st.markdown("### Universidade Federal Fluminense - Instituto de Química")
st.markdown("---")

# Sidebar com instruções
with st.sidebar:
    st.header("Instruções")
    st.markdown("""
    1. Faça upload do arquivo Excel com os dados das defesas
    2. Configure o período letivo
    3. Clique em "Processar"
    4. Baixe os arquivos gerados
    
    **Formato esperado:**
    - Arquivo Excel (.xlsx ou .xls)
    - Colunas necessárias:
        - Nome do aluno
        - Matrícula
        - Curso
        - Título da Defesa
        - Orientador
        - Membro titular 1
        - Membro Suplente
        - Escolha a data para a defesa
        - Coluna com "horario" no nome
    """)

# Upload do arquivo
uploaded_file = st.file_uploader(
    "Selecione o arquivo Excel com os dados das defesas",
    type=['xlsx', 'xls'],
    help="Arquivo Excel exportado do Google Forms"
)

# Configurações
col1, col2 = st.columns(2)
with col1:
    periodo_letivo = st.text_input("Período Letivo", value="2025.2")
with col2:
    gerar_cronograma = st.checkbox("Gerar Cronograma", value=True)
    gerar_declaracoes = st.checkbox("Gerar Declarações", value=True)

if uploaded_file is not None:
    # Detectar período no nome do arquivo
    padroes_periodo = [r'(\d{4}\.\d)', r'(\d{4}-\d)', r'(\d{4}_\d)']
    for padrao in padroes_periodo:
        match = re.search(padrao, uploaded_file.name)
        if match:
            periodo_detectado = match.group(1).replace('-', '.').replace('_', '.')
            st.info(f"Período detectado no nome do arquivo: {periodo_detectado}")
            break

    if st.button("Processar Arquivo", type="primary"):
        try:
            # Carregar todas as abas
            xls = pd.ExcelFile(uploaded_file)
            aba_escolhida = st.selectbox("Selecione a aba da planilha:", xls.sheet_names)
            df = pd.read_excel(uploaded_file, sheet_name=aba_escolhida)
            st.success(f"Planilha carregada com {len(df)} registros da aba '{aba_escolhida}'")
            
            with st.expander("Preview dos dados"):
                st.dataframe(df.head())
            
            # Geração das declarações
            if gerar_declaracoes:
                zip_declaracoes = io.BytesIO()
                with zipfile.ZipFile(zip_declaracoes, 'w', zipfile.ZIP_DEFLATED) as zf:
                    contador = 0
                    start_index = 0
                    if 'Carimbo de data/hora' in str(df.iloc[0, 0]):
                        start_index = 1
                    
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    total_registros = len(df) - start_index
                    
                    for index in range(start_index, len(df)):
                        linha = df.iloc[index]
                        if 'Nome do aluno' in linha and pd.notna(linha['Nome do aluno']) and str(linha['Nome do aluno']).strip() not in ['', 'Nome do aluno']:
                            status_text.text(f"Gerando declaração: {linha['Nome do aluno']}")
                            doc = gerar_documento_word(linha)
                            nome_aluno_arquivo = str(linha['Nome do aluno']).replace(' ', '_').replace('/', '_').replace('\\', '_')
                            nome_arquivo_word = f"Declaracao_{nome_aluno_arquivo}.docx"
                            doc_buffer = io.BytesIO()
                            doc.save(doc_buffer)
                            doc_buffer.seek(0)
                            zf.writestr(nome_arquivo_word, doc_buffer.getvalue())
                            contador += 1
                        progress_bar.progress((index - start_index + 1) / total_registros)
                    status_text.text(f"{contador} declarações geradas!")
                
                zip_declaracoes.seek(0)
                st.download_button(
                    label=f"Baixar Declarações ({contador} arquivos)",
                    data=zip_declaracoes,
                    file_name=f"Declaracoes_Defesa_{periodo_letivo}.zip",
                    mime="application/zip"
                )
            
            # Geração do cronograma
            if gerar_cronograma:
                doc_cronograma = gerar_cronograma_defesas(df, periodo_letivo)
                cronograma_buffer = io.BytesIO()
                doc_cronograma.save(cronograma_buffer)
                cronograma_buffer.seek(0)
                st.download_button(
                    label="Baixar Cronograma",
                    data=cronograma_buffer,
                    file_name=f"Cronograma_Defesas_{periodo_letivo}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            
            st.balloons()
        except Exception as e:
            st.error(f"Erro ao processar arquivo: {str(e)}")
            st.exception(e)

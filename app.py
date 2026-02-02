# ==============================================
# GERADOR DE DECLARACOES E CRONOGRAMA DE DEFESAS
# Versao Streamlit Cloud - UFF Instituto de Quimica
# ==============================================

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

# ===== CONFIGURACAO DA PAGINA =====
st.set_page_config(
    page_title="Gerador de Declaracoes - UFF",
    page_icon="📄",
    layout="wide"
)

# ===== ESTILOS CSS =====
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1e3a5f;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-box {
        padding: 1rem;
        background-color: #d4edda;
        border-radius: 0.5rem;
        border-left: 4px solid #28a745;
    }
</style>
""", unsafe_allow_html=True)

# ===== FUNCOES AUXILIARES =====
def formatar_nome(nome):
    if pd.isna(nome) or nome == "":
        return ""
    return str(nome).title()

def obter_grau_curso(curso):
    if pd.isna(curso) or curso == "":
        return "Bacharelado em Quimica"
    curso = str(curso).lower().strip()
    if 'industrial' in curso:
        return "Bacharelado em Quimica Industrial"
    elif 'licenciatura' in curso:
        return "Licenciatura em Quimica"
    else:
        return "Bacharelado em Quimica"

def formatar_horario(horario):
    if pd.isna(horario) or horario == "":
        return "horario a definir"
    horario_str = str(horario)
    if '(' in horario_str and ')' in horario_str:
        horario_str = horario_str.split('(')[0].strip()
    separadores = ['/', 'as', '-', 'as', 'a']
    for sep in separadores:
        if sep in horario_str:
            horario_str = horario_str.split(sep)[0].strip()
            break
    horario_str = horario_str.replace(' ', '')
    if 'h' not in horario_str.lower():
        horario_str += 'h'
    return horario_str

def formatar_data_sem_dia_semana(data_input):
    if pd.isna(data_input) or data_input == "":
        return "data a definir"
    data_str = str(data_input)
    if '(' in data_str and ')' in data_str:
        data_str = data_str.split('(')[0].strip()
    try:
        data_obj = pd.to_datetime(data_str, dayfirst=True, errors='coerce')
        if not pd.isna(data_obj):
            return data_obj.strftime('%d/%m/%Y')
        else:
            return data_str
    except:
        return data_str

def extrair_horario_local(horario_str):
    if pd.isna(horario_str) or horario_str == "":
        return "horario a definir", "local a definir"
    horario_str = str(horario_str)
    local = "Sala a definir"
    if '(' in horario_str and ')' in horario_str:
        match = re.search(r'((.*?))', horario_str)
        if match:
            local = match.group(1)
        horario_str = horario_str.split('(')[0].strip()
    horario = horario_str.strip()
    return horario, local

def configurar_paragrafo(document, texto, negrito=False, italico=False, tamanho=12, alinhamento='left'):
    p = document.add_paragraph()
    if alinhamento == 'center':
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif alinhamento == 'right':
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(texto)
    run.font.size = Pt(tamanho)
    if negrito:
        run.bold = True
    if italico:
        run.italic = True
    return p

def set_cell_background(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)

def set_cell_border(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        tcBorders.append(border)
    tcPr.append(tcBorders)

def criar_tabela_cronograma_unica(doc, dados_defesas, periodo_letivo):
    total_linhas = 1 + 1 + len(dados_defesas)
    tabela = doc.add_table(rows=total_linhas, cols=6)
    tabela.style = 'Table Grid'
    tabela.autofit = False
    widths = [Inches(1.4), Inches(1.0), Inches(0.8), Inches(0.8), Inches(1.3), Inches(2.2)]
    for row in tabela.rows:
        for idx, width in enumerate(widths):
            if idx < len(row.cells):
                row.cells[idx].width = width
    linha_info = tabela.rows[0]
    cell_info = linha_info.cells[0].merge(linha_info.cells[5])
    cell_info.text = ""
    textos_info = [
        f"Informamos abaixo o cronograma das defesas de monografia referente ao periodo letivo {periodo_letivo}.",
        "",
        "Lembramos que os estudantes de nossos cursos que assistirem as defesas de monografia tem",
        "direito ao aproveitamento da respectiva carga horaria como atividade complementar, bastando",
        "assinar a respectiva lista de presenca.",
        "",
        "Obs.: As defesas serao realizadas no Anfiteatro ou em salas do Instituto de Quimica. O local",
        'segue na coluna "Horario/ local"'
    ]
    for texto in textos_info:
        p = cell_info.add_paragraph()
        if texto != "":
            run = p.add_run(texto)
            run.font.size = Pt(11)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = 1.0
    set_cell_border(cell_info)
    linha_cabecalho = tabela.rows[1]
    cabecalho_campos = ["Nome do aluno", "Data", "Horario", "Local", "Orientador", "Titulo do trabalho"]
    for i, campo in enumerate(cabecalho_campos):
        if i < len(linha_cabecalho.cells):
            cell = linha_cabecalho.cells[i]
            cell.text = campo
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.bold = True
                    run.font.size = Pt(11)
            set_cell_background(cell, "D9D9D9")
            set_cell_border(cell)
    for idx, dados in enumerate(dados_defesas):
        linha_idx = idx + 2
        if linha_idx < len(tabela.rows):
            linha_dados = tabela.rows[linha_idx]
        else:
            linha_dados = tabela.add_row()
        campos = [
            dados['nome_aluno'],
            dados['data'],
            dados['horario'],
            dados['local'],
            dados['orientador'],
            dados['titulo']
        ]
        for i, valor in enumerate(campos):
            if i < len(linha_dados.cells):
                cell = linha_dados.cells[i]
                cell.text = valor
                set_cell_border(cell)
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    for run in paragraph.runs:
                        run.font.size = Pt(10)
    return tabela

def gerar_cronograma_defesas(df, periodo_letivo="2025.2"):
    doc = Document()
    sections = doc.sections
    for section in sections:
        section.top_margin = Pt(40)
        section.bottom_margin = Pt(40)
        section.left_margin = Pt(40)
        section.right_margin = Pt(40)
    titulo = f"Cronograma de defesas de monografia {periodo_letivo}"
    configurar_paragrafo(doc, titulo, negrito=True, tamanho=14, alinhamento='center')
    doc.add_paragraph()
    dados_defesas = []
    start_index = 0
    if 'Carimbo de data/hora' in str(df.iloc[0, 0]):
        start_index = 1
    for index in range(start_index, len(df)):
        linha = df.iloc[index]
        if pd.notna(linha.get('Nome do aluno')) and str(linha['Nome do aluno']).strip() not in ['', 'Nome do aluno']:
            nome_aluno = formatar_nome(linha['Nome do aluno'])
            try:
                data_input = str(linha['Escolha a data para a defesa'])
                if '(' in data_input and ')' in data_input:
                    data_sem_parenteses = data_input.split('(')[0].strip()
                else:
                    data_sem_parenteses = data_input
                data_ordenacao = pd.to_datetime(data_sem_parenteses, dayfirst=True, errors='coerce')
                data_display = formatar_data_sem_dia_semana(data_input)
            except Exception:
                data_ordenacao = datetime.max
                data_display = formatar_data_sem_dia_semana(str(linha['Escolha a data para a defesa']))
            horarios = []
            for col in linha.index:
                if 'horario' in str(col).lower() and pd.notna(linha[col]) and linha[col] != "":
                    horarios.append(str(linha[col]))
            if horarios:
                horario, local = extrair_horario_local(horarios[0])
            else:
                horario, local = "horario a definir", "local a definir"
            try:
                horario_ordenacao = re.search(r'(d{1,2})h', horario)
                if horario_ordenacao:
                    hora = int(horario_ordenacao.group(1))
                else:
                    hora = 0
            except:
                hora = 0
            orientador = formatar_nome(linha['Orientador'])
            titulo_trabalho = linha['Titulo da Defesa']
            dados_defesas.append({
                'nome_aluno': nome_aluno,
                'data': data_display,
                'horario': horario,
                'local': local,
                'orientador': orientador,
                'titulo': titulo_trabalho,
                'data_ordenacao': data_ordenacao if not pd.isna(data_ordenacao) else datetime.max,
                'hora_ordenacao': hora
            })
    dados_defesas.sort(key=lambda x: (x['data_ordenacao'], x['hora_ordenacao']))
    criar_tabela_cronograma_unica(doc, dados_defesas, periodo_letivo)
    for _ in range(6):
        doc.add_paragraph()
    return doc

def gerar_documento_word(dados):
    nome_aluno = formatar_nome(dados['Nome do aluno'])
    matricula = str(dados['Matricula'])
    titulo = dados["Titulo da Defesa"]
    curso_original = dados['Curso']
    curso_formatado = obter_grau_curso(curso_original)
    try:
        data_input = str(dados['Escolha a data para a defesa'])
        if '(' in data_input and ')' in data_input:
            data_input = data_input.split('(')[0].strip()
        data_defesa = pd.to_datetime(data_input, dayfirst=True).strftime('%d/%m/%Y')
    except Exception:
        data_defesa = str(dados['Escolha a data para a defesa'])
        if '(' in data_defesa and ')' in data_defesa:
            data_defesa = data_defesa.split('(')[0].strip()
    horarios = []
    for col in dados.index:
        if 'horario' in str(col).lower() and pd.notna(dados[col]) and dados[col] != "":
            horarios.append(str(dados[col]))
    horario = formatar_horario(horarios[0] if horarios else "")
    orientador = formatar_nome(dados['Orientador'])
    membro_titular1 = formatar_nome(dados['Membro titular 1'])
    membro_titular2 = formatar_nome(dados.get('Membro Titular 2', ''))
    membro_suplente = formatar_nome(dados['Membro Suplente'])
    coorientador = formatar_nome(dados.get('Coorientador', ''))
    doc = Document()
    sections = doc.sections
    for section in sections:
        section.top_margin = Pt(50)
        section.bottom_margin = Pt(50)
        section.left_margin = Pt(60)
        section.right_margin = Pt(60)
    configurar_paragrafo(doc, "SERVICO PUBLICO FEDERAL", tamanho=12, alinhamento='center')
    configurar_paragrafo(doc, "MINISTERIO DA EDUCACAO", tamanho=12, alinhamento='center')
    configurar_paragrafo(doc, "UNIVERSIDADE FEDERAL FLUMINENSE", tamanho=12, alinhamento='center')
    configurar_paragrafo(doc, "PRO-REITORIA DE GRADUACAO", tamanho=12, alinhamento='center')
    doc.add_paragraph()
    configurar_paragrafo(doc, "DECLARACAO", negrito=True, tamanho=14, alinhamento='center')
    doc.add_paragraph()
    texto_declaracao = f'Declaro, para os devidos fins, que a Banca Examinadora da Monografia de Final de Curso do(a) estudante {nome_aluno.upper()}, matriculado(a) sob o numero {matricula}, cujo titulo e "{titulo}" defendida e aprovada no dia {data_defesa}, as {horario}, para obtencao do grau {curso_formatado}, foi composta pelos seguintes membros: TITULARES: {orientador} (Orientador(a))'
    if membro_titular1 and membro_titular1 != "":
        texto_declaracao += f", {membro_titular1}"
    if membro_titular2 and membro_titular2 != "":
        texto_declaracao += f", {membro_titular2}"
    texto_declaracao += f"; SUPLENTES: {membro_suplente}"
    if coorientador and coorientador != "":
        texto_declaracao += f", {coorientador} (Coorientador(a))"
    texto_declaracao += "."
    p_declaracao = configurar_paragrafo(doc, texto_declaracao, tamanho=12)
    p_declaracao.paragraph_format.first_line_indent = Pt(0)
    for _ in range(12):
        doc.add_paragraph()
    data_emissao = datetime.now().strftime('%d/%m/%Y as %H:%M:%S')
    configurar_paragrafo(doc, "Universidade Federal Fluminense", tamanho=12, alinhamento='center')
    configurar_paragrafo(doc, f"Niteroi, {data_emissao}", tamanho=12, alinhamento='center')
    doc.add_paragraph()
    configurar_paragrafo(doc, "CPF [CPF_DO_COORDENADOR]", tamanho=10, alinhamento='center')
    doc.add_paragraph()
    texto_rodape1 = "Este documento foi gerado pelo Sistema Academico da Universidade Federal Fluminense - IdUFF."
    texto_rodape2 = 'Para verificar a autenticidade deste documento, acesse https://inscricao.id.uff.br no link "Validar declaracao"'
    configurar_paragrafo(doc, texto_rodape1, tamanho=9, alinhamento='center')
    configurar_paragrafo(doc, texto_rodape2, tamanho=9, alinhamento='center')
    doc.add_paragraph()
    configurar_paragrafo(doc, "[CODIGO_DE_VALIDACAO]", tamanho=9, alinhamento='center')
    return doc

# ===== INTERFACE PRINCIPAL =====
st.markdown('<p class="main-header">Gerador de Declaracoes e Cronograma</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Universidade Federal Fluminense - Instituto de Quimica</p>', unsafe_allow_html=True)

# Sidebar com instrucoes
with st.sidebar:
    st.header("Instrucoes")
    st.markdown("""
    1. Faca upload do arquivo Excel com os dados das defesas
    2. Configure o periodo letivo
    3. Clique em "Processar"
    4. Baixe os arquivos gerados
    
    **Formato esperado:**
    - Arquivo Excel (.xlsx)
    - Aba: "Respostas ao formulario 1"
    - Colunas necessarias:
        - Nome do aluno
        - Matricula
        - Curso
        - Titulo da Defesa
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

# Configuracoes
col1, col2 = st.columns(2)
with col1:
    periodo_letivo = st.text_input("Periodo Letivo", value="2025.2")
with col2:
    gerar_cronograma = st.checkbox("Gerar Cronograma", value=True)
    gerar_declaracoes = st.checkbox("Gerar Declaracoes", value=True)

if uploaded_file is not None:
    padroes_periodo = [r'(d{4}.d)', r'(d{4}-d)', r'(d{4}_d)']
    for padrao in padroes_periodo:
        match = re.search(padrao, uploaded_file.name)
        if match:
            periodo_detectado = match.group(1).replace('-', '.').replace('_', '.')
            st.info(f"Periodo detectado no nome do arquivo: {periodo_detectado}")
            break

    if st.button("Processar Arquivo", type="primary", use_container_width=True):
        try:
            df = pd.read_excel(uploaded_file, sheet_name='Respostas ao formulario 1')
            st.success(f"Planilha carregada com {len(df)} registros")
            
            with st.expander("Preview dos dados"):
                st.dataframe(df.head())
            
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
                        if pd.notna(linha.get('Nome do aluno')) and str(linha['Nome do aluno']).strip() not in ['', 'Nome do aluno']:
                            status_text.text(f"Gerando declaracao: {linha['Nome do aluno']}")
                            doc = gerar_documento_word(linha)
                            nome_aluno_arquivo = str(linha['Nome do aluno']).replace(' ', '_').replace('/', '_').replace('\', '_')
                            nome_arquivo_word = f"Declaracao_{nome_aluno_arquivo}.docx"
                            doc_buffer = io.BytesIO()
                            doc.save(doc_buffer)
                            doc_buffer.seek(0)
                            zf.writestr(nome_arquivo_word, doc_buffer.getvalue())
                            contador += 1
                        progress_bar.progress((index - start_index + 1) / total_registros)
                    status_text.text(f"{contador} declaracoes geradas!")
                
                zip_declaracoes.seek(0)
                st.download_button(
                    label=f"Download Declaracoes ({contador} arquivos)",
                    data=zip_declaracoes,
                    file_name=f"Declaracoes_Defesa_{periodo_letivo}.zip",
                    mime="application/zip",
                    use_container_width=True
                )
            
            if gerar_cronograma:
                doc_cronograma = gerar_cronograma_defesas(df, periodo_letivo)
                cronograma_buffer = io.BytesIO()
                doc_cronograma.save(cronograma_buffer)
                cronograma_buffer.seek(0)
                st.download_button(
                    label="Download Cronograma",
                    data=cronograma_buffer,
                    file_name=f"Cronograma_Defesas_{periodo_letivo}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            
            st.balloons()
        except Exception as e:
            st.error(f"Erro ao processar arquivo: {str(e)}")
            st.exception(e)

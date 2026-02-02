# ==============================================
# GERADOR DE DECLARACOES E CRONOGRAMA DE DEFESAS
# Versao Streamlit Cloud - UFF Instituto de Quimica
# CORRIGIDO PARA STREAMLIT CLOUD
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
import sys
import traceback

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
    .stProgress > div > div > div > div {
        background-color: #1e3a5f;
    }
</style>
""", unsafe_allow_html=True)

# ===== INICIALIZAR STATE =====
if 'processando' not in st.session_state:
    st.session_state.processando = False
if 'arquivo_processado' not in st.session_state:
    st.session_state.arquivo_processado = False

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
        match = re.search(r'\((.*?)\)', horario_str)
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
    try:
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), color)
        tcPr.append(shd)
    except Exception as e:
        st.warning(f"Erro ao definir fundo da célula: {e}")

def set_cell_border(cell):
    try:
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
    except Exception as e:
        st.warning(f"Erro ao definir borda da célula: {e}")

def criar_tabela_cronograma_unica(doc, dados_defesas, periodo_letivo):
    try:
        total_linhas = 2 + len(dados_defesas)  # 1 linha para info + 1 para cabeçalho + linhas de dados
        tabela = doc.add_table(rows=total_linhas, cols=6)
        tabela.style = 'Table Grid'
        
        # Definir larguras das colunas
        widths = [Inches(1.4), Inches(1.0), Inches(0.8), Inches(0.8), Inches(1.3), Inches(2.2)]
        for row in tabela.rows:
            for idx, width in enumerate(widths[:6]):  # Garantir apenas 6 colunas
                row.cells[idx].width = width
        
        # Linha de informações
        linha_info = tabela.rows[0]
        cell_info = linha_info.cells[0]
        for i in range(1, 6):
            cell_info = cell_info.merge(linha_info.cells[i])
        
        textos_info = [
            f"Informamos abaixo o cronograma das defesas de monografia referente ao período letivo {periodo_letivo}.",
            "",
            "Lembramos que os estudantes de nossos cursos que assistirem às defesas de monografia têm",
            "direito ao aproveitamento da respectiva carga horária como atividade complementar, bastando",
            "assinar a respectiva lista de presença.",
            "",
            "Obs.: As defesas serão realizadas no Anfiteatro ou em salas do Instituto de Química. O local",
            'segue na coluna "Horário/ local"'
        ]
        
        for texto in textos_info:
            p = cell_info.add_paragraph()
            if texto != "":
                run = p.add_run(texto)
                run.font.size = Pt(11)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.line_spacing = 1.0
        
        # Cabeçalho da tabela
        linha_cabecalho = tabela.rows[1]
        cabecalho_campos = ["Nome do aluno", "Data", "Horário", "Local", "Orientador", "Título do trabalho"]
        
        for i, campo in enumerate(cabecalho_campos):
            cell = linha_cabecalho.cells[i]
            cell.text = campo
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.bold = True
                    run.font.size = Pt(11)
        
        # Preencher dados
        for idx, dados in enumerate(dados_defesas):
            linha_idx = idx + 2
            if linha_idx < len(tabela.rows):
                linha_dados = tabela.rows[linha_idx]
            else:
                # Adicionar linha se necessário
                linha_dados = tabela.add_row()
            
            campos = [
                dados.get('nome_aluno', ''),
                dados.get('data', ''),
                dados.get('horario', ''),
                dados.get('local', ''),
                dados.get('orientador', ''),
                dados.get('titulo', '')
            ]
            
            for i, valor in enumerate(campos):
                if i < len(linha_dados.cells):
                    cell = linha_dados.cells[i]
                    cell.text = str(valor) if valor else ''
                    for paragraph in cell.paragraphs:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        for run in paragraph.runs:
                            run.font.size = Pt(10)
        
        return tabela
    except Exception as e:
        st.error(f"Erro ao criar tabela: {e}")
        return None

def gerar_cronograma_defesas(df, periodo_letivo="2025.2"):
    try:
        doc = Document()
        
        # Configurar margens
        for section in doc.sections:
            section.top_margin = Pt(40)
            section.bottom_margin = Pt(40)
            section.left_margin = Pt(40)
            section.right_margin = Pt(40)
        
        # Título
        titulo = f"Cronograma de defesas de monografia {periodo_letivo}"
        configurar_paragrafo(doc, titulo, negrito=True, tamanho=14, alinhamento='center')
        doc.add_paragraph()
        
        # Processar dados
        dados_defesas = []
        start_index = 0
        
        # Verificar se a primeira linha é cabeçalho ou dados
        primeira_celula = str(df.iloc[0, 0]) if len(df) > 0 else ""
        if 'Carimbo de data/hora' in primeira_celula:
            start_index = 1
        
        # Coletar nomes de colunas disponíveis
        colunas_disponiveis = df.columns.tolist()
        st.info(f"Colunas disponíveis: {', '.join(colunas_disponiveis[:5])}...")
        
        for index in range(start_index, len(df)):
            try:
                linha = df.iloc[index]
                
                # Verificar coluna de nome do aluno
                coluna_nome = None
                for col in df.columns:
                    if 'nome' in str(col).lower() and 'aluno' in str(col).lower():
                        coluna_nome = col
                        break
                
                if coluna_nome is None:
                    # Tentar encontrar qualquer coluna com 'nome'
                    for col in df.columns:
                        if 'nome' in str(col).lower():
                            coluna_nome = col
                            break
                
                if coluna_nome and coluna_nome in linha:
                    nome_aluno = formatar_nome(linha[coluna_nome])
                    if nome_aluno and nome_aluno.strip() not in ['', 'Nome do aluno']:
                        # Extrair data
                        coluna_data = None
                        for col in df.columns:
                            if 'data' in str(col).lower() and 'defesa' in str(col).lower():
                                coluna_data = col
                                break
                        
                        data_display = "data a definir"
                        data_ordenacao = datetime.max
                        if coluna_data and coluna_data in linha:
                            try:
                                data_input = str(linha[coluna_data])
                                data_display = formatar_data_sem_dia_semana(data_input)
                                data_ordenacao = pd.to_datetime(data_display, dayfirst=True, errors='coerce')
                                if pd.isna(data_ordenacao):
                                    data_ordenacao = datetime.max
                            except:
                                pass
                        
                        # Extrair horário
                        horario_col = None
                        for col in df.columns:
                            if 'horário' in str(col).lower() or 'horario' in str(col).lower():
                                horario_col = col
                                break
                        
                        horario = "horário a definir"
                        local = "local a definir"
                        if horario_col and horario_col in linha:
                            horario_str = str(linha[horario_col])
                            horario, local = extrair_horario_local(horario_str)
                        
                        # Extrair hora para ordenação
                        hora_ordenacao = 0
                        try:
                            match = re.search(r'(\d{1,2})h', horario.lower())
                            if match:
                                hora_ordenacao = int(match.group(1))
                        except:
                            pass
                        
                        # Orientador
                        orientador_col = None
                        for col in df.columns:
                            if 'orientador' in str(col).lower():
                                orientador_col = col
                                break
                        
                        orientador = ""
                        if orientador_col and orientador_col in linha:
                            orientador = formatar_nome(linha[orientador_col])
                        
                        # Título
                        titulo_col = None
                        for col in df.columns:
                            if 'título' in str(col).lower() or 'titulo' in str(col).lower():
                                titulo_col = col
                                break
                        
                        titulo_trabalho = ""
                        if titulo_col and titulo_col in linha:
                            titulo_trabalho = str(linha[titulo_col])
                        
                        dados_defesas.append({
                            'nome_aluno': nome_aluno,
                            'data': data_display,
                            'horario': horario,
                            'local': local,
                            'orientador': orientador,
                            'titulo': titulo_trabalho,
                            'data_ordenacao': data_ordenacao,
                            'hora_ordenacao': hora_ordenacao
                        })
            except Exception as e:
                st.warning(f"Erro ao processar linha {index}: {e}")
                continue
        
        # Ordenar dados
        dados_defesas.sort(key=lambda x: (x['data_ordenacao'], x['hora_ordenacao']))
        
        # Criar tabela
        if dados_defesas:
            tabela = criar_tabela_cronograma_unica(doc, dados_defesas, periodo_letivo)
            if tabela is None:
                st.error("Falha ao criar tabela do cronograma")
        else:
            st.warning("Nenhum dado válido encontrado para gerar o cronograma")
            configurar_paragrafo(doc, "Nenhuma defesa agendada para este período.", tamanho=12, alinhamento='center')
        
        # Adicionar espaços no final
        for _ in range(3):
            doc.add_paragraph()
            
        return doc
        
    except Exception as e:
        st.error(f"Erro crítico ao gerar cronograma: {e}")
        # Criar documento mínimo em caso de erro
        doc = Document()
        configurar_paragrafo(doc, f"ERRO: Não foi possível gerar o cronograma: {str(e)}", 
                           tamanho=12, alinhamento='center')
        return doc

def gerar_documento_word(dados):
    try:
        # Mapear colunas
        colunas = {}
        for col in dados.index:
            col_lower = str(col).lower()
            if 'nome' in col_lower and 'aluno' in col_lower:
                colunas['nome'] = col
            elif 'matricula' in col_lower or 'matrícula' in col_lower:
                colunas['matricula'] = col
            elif 'curso' in col_lower:
                colunas['curso'] = col
            elif 'título' in col_lower or 'titulo' in col_lower:
                colunas['titulo'] = col
            elif 'data' in col_lower and 'defesa' in col_lower:
                colunas['data'] = col
            elif 'horário' in col_lower or 'horario' in col_lower:
                colunas['horario'] = col
            elif 'orientador' in col_lower:
                colunas['orientador'] = col
        
        # Extrair valores com fallback
        nome_aluno = formatar_nome(dados.get(colunas.get('nome', ''), ''))
        matricula = str(dados.get(colunas.get('matricula', ''), ''))
        curso_original = dados.get(colunas.get('curso', ''), '')
        curso_formatado = obter_grau_curso(curso_original)
        titulo = dados.get(colunas.get('titulo', ''), 'Título não informado')
        
        # Data
        data_input = str(dados.get(colunas.get('data', ''), ''))
        try:
            if '(' in data_input and ')' in data_input:
                data_input = data_input.split('(')[0].strip()
            data_defesa = pd.to_datetime(data_input, dayfirst=True, errors='coerce')
            if pd.isna(data_defesa):
                data_defesa_str = data_input
            else:
                data_defesa_str = data_defesa.strftime('%d/%m/%Y')
        except:
            data_defesa_str = data_input
        
        # Horário
        horario_input = str(dados.get(colunas.get('horario', ''), ''))
        horario = formatar_horario(horario_input)
        
        # Membros da banca (simplificado)
        orientador = formatar_nome(dados.get(colunas.get('orientador', ''), ''))
        
        # Criar documento
        doc = Document()
        
        # Configurar margens
        for section in doc.sections:
            section.top_margin = Pt(50)
            section.bottom_margin = Pt(50)
            section.left_margin = Pt(60)
            section.right_margin = Pt(60)
        
        # Cabeçalho
        configurar_paragrafo(doc, "SERVIÇO PÚBLICO FEDERAL", tamanho=12, alinhamento='center')
        configurar_paragrafo(doc, "MINISTÉRIO DA EDUCAÇÃO", tamanho=12, alinhamento='center')
        configurar_paragrafo(doc, "UNIVERSIDADE FEDERAL FLUMINENSE", tamanho=12, alinhamento='center')
        configurar_paragrafo(doc, "PRÓ-REITORIA DE GRADUAÇÃO", tamanho=12, alinhamento='center')
        doc.add_paragraph()
        
        # Título
        configurar_paragrafo(doc, "DECLARAÇÃO", negrito=True, tamanho=14, alinhamento='center')
        doc.add_paragraph()
        
        # Texto da declaração (versão simplificada)
        texto_declaracao = f'Declaro, para os devidos fins, que o(a) estudante {nome_aluno.upper()}, '
        texto_declaracao += f'matriculado(a) sob o número {matricula}, '
        texto_declaracao += f'defendeu e foi aprovado(a) no dia {data_defesa_str}, '
        texto_declaracao += f'às {horario}, para obtenção do grau de {curso_formatado}, '
        texto_declaracao += f'sob a orientação de {orientador}.'
        
        if titulo and titulo != "Título não informado":
            texto_declaracao += f'\n\nTítulo do trabalho: "{titulo}"'
        
        p_declaracao = configurar_paragrafo(doc, texto_declaracao, tamanho=12)
        p_declaracao.paragraph_format.space_after = Pt(12)
        
        # Espaço para assinatura
        for _ in range(10):
            doc.add_paragraph()
        
        # Rodapé
        data_emissao = datetime.now().strftime('%d/%m/%Y')
        configurar_paragrafo(doc, "Universidade Federal Fluminense", tamanho=12, alinhamento='center')
        configurar_paragrafo(doc, f"Niterói, {data_emissao}", tamanho=12, alinhamento='center')
        
        doc.add_paragraph()
        configurar_paragrafo(doc, "[Assinatura do Coordenador]", tamanho=10, alinhamento='center', italico=True)
        
        return doc
        
    except Exception as e:
        st.error(f"Erro ao gerar documento: {e}")
        # Retornar documento de erro
        doc = Document()
        configurar_paragrafo(doc, f"ERRO: {str(e)}", tamanho=12, alinhamento='center')
        return doc

# ===== INTERFACE PRINCIPAL =====
st.markdown('<p class="main-header">Gerador de Declarações e Cronograma</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Universidade Federal Fluminense - Instituto de Química</p>', unsafe_allow_html=True)

# Sidebar com instruções
with st.sidebar:
    st.header("Instruções")
    st.markdown("""
    1. Faça upload do arquivo Excel com os dados das defesas
    2. Configure o período letivo
    3. Clique em "Processar"
    4. Baixe os arquivos gerados
    
    **Formato esperado:**
    - Arquivo Excel (.xlsx)
    - Colunas necessárias:
        - Nome do aluno
        - Matrícula
        - Curso
        - Título da Defesa
        - Orientador
        - Data da defesa
        - Horário
    """)
    
    st.markdown("---")
    st.info("**Dica:** Se o processamento travar, tente:")
    st.markdown("""
    1. Reduzir o tamanho do arquivo
    2. Verificar o formato das colunas
    3. Recarregar a página
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

# Container para resultados
result_container = st.container()

if uploaded_file is not None:
    # Detectar período no nome do arquivo
    periodo_detectado = None
    padroes_periodo = [r'(\d{4}\.\d)', r'(\d{4}-\d)', r'(\d{4}_\d)']
    for padrao in padroes_periodo:
        match = re.search(padrao, uploaded_file.name)
        if match:
            periodo_detectado = match.group(1).replace('-', '.').replace('_', '.')
            st.info(f"Período detectado no nome do arquivo: {periodo_detectado}")
            break
    
    if periodo_detectado:
        periodo_letivo = periodo_detectado
    
    # Botão de processamento
    if st.button("🔄 Processar Arquivo", type="primary", use_container_width=True):
        st.session_state.processando = True
        
        try:
            # Ler arquivo
            df = pd.read_excel(uploaded_file)
            st.success(f"✅ Planilha carregada com {len(df)} registros")
            
            # Mostrar preview
            with st.expander("📊 Visualizar dados (primeiras 5 linhas)"):
                st.dataframe(df.head())
                st.write(f"**Colunas disponíveis:** {', '.join(df.columns.tolist()[:10])}...")
            
            # Processar em container separado
            with result_container:
                st.markdown("---")
                st.subheader("📁 Resultados")
                
                # Gerar declarações
                if gerar_declaracoes:
                    st.info("🔄 Gerando declarações individuais...")
                    
                    try:
                        zip_buffer = io.BytesIO()
                        contador = 0
                        
                        # Determinar início dos dados
                        start_index = 0
                        if len(df) > 0:
                            primeira_celula = str(df.iloc[0, 0])
                            if 'Carimbo' in primeira_celula or 'carimbo' in primeira_celula:
                                start_index = 1
                        
                        # Progresso
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        total_registros = max(1, len(df) - start_index)
                        
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                            for idx in range(start_index, len(df)):
                                try:
                                    linha = df.iloc[idx]
                                    
                                    # Verificar se tem nome
                                    tem_nome = False
                                    for col in df.columns:
                                        if 'nome' in str(col).lower() and pd.notna(linha.get(col, '')):
                                            if str(linha[col]).strip() not in ['', 'Nome do aluno']:
                                                tem_nome = True
                                                break
                                    
                                    if tem_nome:
                                        # Gerar documento
                                        doc = gerar_documento_word(linha)
                                        
                                        # Nome do arquivo
                                        nome_aluno = "Aluno"
                                        for col in df.columns:
                                            if 'nome' in str(col).lower():
                                                nome_aluno = str(linha[col]).split()[0] if pd.notna(linha[col]) else "Aluno"
                                                break
                                        
                                        nome_arquivo = f"Declaracao_{nome_aluno}_{idx}.docx"
                                        
                                        # Salvar no ZIP
                                        doc_buffer = io.BytesIO()
                                        doc.save(doc_buffer)
                                        doc_buffer.seek(0)
                                        zip_file.writestr(nome_arquivo, doc_buffer.getvalue())
                                        contador += 1
                                        
                                        # Atualizar progresso
                                        if idx % max(1, total_registros // 10) == 0:
                                            status_text.text(f"Processando... {idx - start_index + 1}/{total_registros}")
                                        
                                    progress_bar.progress((idx - start_index + 1) / total_registros)
                                    
                                except Exception as e:
                                    st.warning(f"Erro na linha {idx}: {e}")
                                    continue
                        
                        progress_bar.empty()
                        status_text.empty()
                        
                        if contador > 0:
                            zip_buffer.seek(0)
                            st.success(f"✅ {contador} declarações geradas com sucesso!")
                            
                            st.download_button(
                                label=f"📥 Baixar Declarações ({contador} arquivos)",
                                data=zip_buffer,
                                file_name=f"Declaracoes_{periodo_letivo}.zip",
                                mime="application/zip",
                                use_container_width=True
                            )
                        else:
                            st.warning("⚠️ Nenhuma declaração foi gerada. Verifique os dados do arquivo.")
                            
                    except Exception as e:
                        st.error(f"❌ Erro ao gerar declarações: {e}")
                
                # Gerar cronograma
                if gerar_cronograma:
                    st.info("🔄 Gerando cronograma...")
                    
                    try:
                        doc_cronograma = gerar_cronograma_defesas(df, periodo_letivo)
                        
                        if doc_cronograma:
                            cronograma_buffer = io.BytesIO()
                            doc_cronograma.save(cronograma_buffer)
                            cronograma_buffer.seek(0)
                            
                            st.success("✅ Cronograma gerado com sucesso!")
                            
                            st.download_button(
                                label="📥 Baixar Cronograma",
                                data=cronograma_buffer,
                                file_name=f"Cronograma_Defesas_{periodo_letivo}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                    except Exception as e:
                        st.error(f"❌ Erro ao gerar cronograma: {e}")
                        st.exception(e)
                
                st.balloons()
                st.session_state.arquivo_processado = True
                
        except Exception as e:
            st.error(f"❌ Erro crítico ao processar arquivo: {e}")
            st.code(traceback.format_exc())
            
        finally:
            st.session_state.processando = False

# Mensagem de ajuda
if not uploaded_file:
    st.info("👈 **Para começar:** Faça upload de um arquivo Excel na barra lateral.")
    
# Adicionar informações de debug no final (opcional)
with st.expander("ℹ️ Informações de debug", expanded=False):
    st.write(f"Python version: {sys.version}")
    st.write(f"Streamlit version: {st.__version__}")
    st.write(f"Pandas version: {pd.__version__}")

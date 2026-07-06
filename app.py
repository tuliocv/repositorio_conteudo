import streamlit as st
import pandas as pd
import requests
import io
import zipfile
import re
import os

st.set_page_config(page_title="Gerenciador de Catálogos e PDFs", page_icon="📚", layout="wide")

st.title("📚 Gerenciador de Download de Conteúdos (PDFs)")
st.markdown("---")

# Função para sanitizar nomes de arquivos e pastas
def limpar_nome_arquivo(nome):
    if pd.isna(nome):
        return "sem_nome"
    # Remove caracteres inválidos para nomes de arquivos
    nome_limpo = re.sub(r'[\\/*?:"<>|]', "", str(nome))
    return nome_limpo.strip()

def criar_slug_curso(nome_curso):
    if pd.isna(nome_curso):
        return "curso_sem_nome"
    # Transforma em minúsculo e substitui espaços por underline
    slug = str(nome_curso).lower().strip()
    slug = re.sub(r'\s+', '_', slug)
    # Remove caracteres especiais comuns de pastas
    slug = re.sub(r'[\\/*?:"<>|]', "", slug)
    return slug

# 1. Upload do Arquivo Base
st.sidebar.header("📁 Upload da Base de Dados")
arquivo_carregado = st.sidebar.file_uploader("Carregue o arquivo Excel unificado (.xlsx)", type=["xlsx", "csv"])

if arquivo_carregado is not None:
    try:
        # Carrega a base de dados
        if arquivo_carregado.name.endswith('.csv'):
            df = pd.read_csv(arquivo_carregado)
        else:
            df = pd.read_excel(arquivo_carregado)
        
        # Validação básica de colunas necessárias
        colunas_obrigatorias = ['ANO', 'PERÍODO', 'CURSO', 'UC', 'CATÁLOGO', 'ORDEM AULA', 'ATIVIDADE', 'PDF']
        colunas_faltando = [col for col in colunas_obrigatorias if col not in df.columns]
        
        if colunas_faltando:
            st.error(f"❌ O arquivo carregado está faltando as seguintes colunas obrigatórias: {', '.join(colunas_faltando)}")
        else:
            st.success("✅ Base de dados carregada com sucesso!")
            
            # 2. Filtros Cascata
            st.subheader("🔍 Filtros de Seleção")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                anos_disponiveis = sorted(df['ANO'].dropna().unique())
                ano_selecionado = st.selectbox("1. Selecione o ANO:", anos_disponiveis)
            
            df_ano = df[df['ANO'] == ano_selecionado]
            
            with col2:
                periodos_disponiveis = sorted(df_ano['PERÍODO'].dropna().unique())
                periodo_selecionado = st.selectbox("2. Selecione o PERÍODO:", periodos_disponiveis)
                
            df_periodo = df_ano[df_ano['PERÍODO'] == periodo_selecionado]
            
            with col3:
                cursos_disponiveis = sorted(df_periodo['CURSO'].dropna().unique())
                curso_selecionado = st.selectbox("3. Selecione o CURSO:", cursos_disponiveis)
            
            df_filtrado = df_periodo[df_periodo['CURSO'] == curso_selecionado]
            
            st.markdown("---")
            
            # 3. Apresentação das Métricas
            st.subheader("📊 Resumo dos Dados Disponíveis")
            
            qtd_uc = df_filtrado['UC'].nunique()
            qtd_catalogo = df_filtrado['CATÁLOGO'].nunique()
            total_linhas = len(df_filtrado)
            
            m1, m2, m3 = st.columns(3)
            m1.metric("Quantidade de UCs Únicas", qtd_uc)
            m2.metric("Quantidade de Catálogos Únicos", qtd_catalogo)
            m3.metric("Total de Arquivos/Aulas", total_linhas)
            
            with st.expander("👀 Visualizar linhas que serão processadas"):
                st.dataframe(df_filtrado[['UC', 'CATÁLOGO', 'ORDEM AULA', 'AULA', 'ATIVIDADE', 'PDF']])
            
            st.markdown("---")
            
            # 4. Processamento e Download
            st.subheader("⚙️ Processamento e Download dos PDFs")
            st.info("Ao clicar no botão abaixo, o sistema fará o download de cada PDF da lista filtrada e estruturará as pastas conforme solicitado.")
            
            if st.button("🚀 Iniciar Processamento dos PDFs", type="primary"):
                df_com_pdf = df_filtrado[df_filtrado['PDF'].notna() & df_filtrado['PDF'].astype(str).str.startswith('http')]
                
                if len(df_com_pdf) == 0:
                    st.warning("⚠️ Nenhuma linha com link de PDF válido encontrado para os filtros selecionados.")
                else:
                    zip_buffer = io.BytesIO()
                    
                    progresso_bar = st.progress(0)
                    status_text = st.empty()
                    
                    nome_pasta_curso = criar_slug_curso(curso_selecionado)
                    
                    sucessos = 0
                    erros = 0
                    
                    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                        # Iteração usando iteração de pandas
                        total = len(df_com_pdf)
                        for idx, row in df_com_pdf.reset_index().iterrows():
                            # Atualiza barra de progresso
                            porcentagem = (idx + 1) / total
                            progresso_bar.progress(porcentagem)
                            status_text.text(f"Baixando arquivo {idx+1} de {total}: {row['ATIVIDADE']}")
                            
                            try:
                                response = requests.get(row['PDF'], timeout=15)
                                
                                if response.status_code == 200:
                                    ordem_aula = limpar_nome_arquivo(row['ORDEM AULA'])
                                    uc = limpar_nome_arquivo(row['UC'])
                                    atividade = limpar_nome_arquivo(row['ATIVIDADE'])
                                    
                                    nome_pdf = f"{ordem_aula} - {uc} - {atividade}.pdf"
                                    
                                    caminho_no_zip = os.path.join(
                                        nome_pasta_curso, 
                                        "04_conteudos_especificos", 
                                        "conteudo", 
                                        nome_pdf
                                    )
                                    
                                    zip_file.writestr(caminho_no_zip, response.content)
                                    sucessos += 1
                                else:
                                    erros += 1
                            except Exception as e:
                                erros += 1
                                continue
                    
                    status_text.text("✨ Processamento concluído!")
                    st.balloons()
                    
                    if sucessos > 0:
                        st.success(f"📊 Download concluído! {sucessos} PDFs foram baixados e estruturados com sucesso.")
                    if erros > 0:
                        st.warning(f"⚠️ Houve erro ou timeout no download de {erros} arquivos PDF. Verifique os links na planilha.")
                    
                    zip_buffer.seek(0)
                    st.download_button(
                        label=f"📥 Baixar Pasta do Curso ({nome_pasta_curso}.zip)",
                        data=zip_buffer,
                        file_name=f"{nome_pasta_curso}.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")

else:
    st.info("💡 Por favor, carregue o arquivo Excel na barra lateral para iniciar o aplicativo.")
    st.markdown('''
    ### Estrutura esperada da Planilha:
    Certifique-se de que seu arquivo contenha exatamente estas colunas no cabeçalho:
    `CATÁLOGO`, `UC`, `CURSO`, `PERÍODO`, `ANO`, `ID_CATÁLOGO`, `CATÁLOGO2`, `ORDEM AULA`, `AULA`, `ORDEM ATIVIDADE`, `ATIVIDADE`, `HTML`, `PDF`
    ''')

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
    nome_limpo = re.sub(r'[\\/*?:"<>|]', "", str(nome))
    return nome_limpo.strip()

def criar_slug_curso(nome_curso):
    if pd.isna(nome_curso):
        return "curso_sem_nome"
    slug = str(nome_curso).lower().strip()
    slug = re.sub(r'\s+', '_', slug)
    slug = re.sub(r'[\\/*?:"<>|]', "", slug)
    return slug

st.sidebar.header("📁 Upload da Base de Dados")
arquivo_carregado = st.sidebar.file_uploader("Carregue o arquivo Excel unificado (.xlsx)", type=["xlsx", "xls", "csv"])

if arquivo_carregado is not None:
    try:
        if arquivo_carregado.name.endswith('.csv'):
            df = pd.read_csv(arquivo_carregado)
        else:
            # Força o uso do openpyxl e lida com possíveis erros de formatação
            df = pd.read_excel(arquivo_carregado, engine='openpyxl')
        
        # Opcional: Limpar nomes das colunas (remover espaços extras)
        df.columns = [col.strip() for col in df.columns]
        
        colunas_obrigatorias = ['ANO', 'PERÍODO', 'CURSO', 'UC', 'CATÁLOGO', 'ORDEM AULA', 'ORDEM ATIVIDADE', 'ATIVIDADE', 'PDF']
        colunas_faltando = [col for col in colunas_obrigatorias if col not in df.columns]
        
        if colunas_faltando:
            st.error(f"❌ Faltam as seguintes colunas no arquivo: {', '.join(colunas_faltando)}")
            st.info(f"Colunas encontradas: {', '.join(df.columns.tolist())}")
        else:
            st.success("✅ Base de dados carregada com sucesso!")
            
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
            
            st.subheader("📊 Resumo dos Dados Disponíveis")
            qtd_uc = df_filtrado['UC'].nunique()
            qtd_catalogo = df_filtrado['CATÁLOGO'].nunique()
            total_linhas = len(df_filtrado)
            
            m1, m2, m3 = st.columns(3)
            m1.metric("Quantidade de UCs Únicas", qtd_uc)
            m2.metric("Quantidade de Catálogos Únicos", qtd_catalogo)
            m3.metric("Total de Arquivos/Aulas", total_linhas)
            
            with st.expander("👀 Visualizar linhas que serão processadas"):
                st.dataframe(df_filtrado[['UC', 'CATÁLOGO', 'ORDEM AULA', 'ORDEM ATIVIDADE', 'ATIVIDADE', 'PDF']])
            
            st.markdown("---")
            
            st.subheader("⚙️ Processamento e Download dos PDFs")
            
            if st.button("🚀 Iniciar Processamento (Sequencial)", type="primary"):
                df_com_pdf = df_filtrado[df_filtrado['PDF'].notna() & df_filtrado['PDF'].astype(str).str.startswith('http')]
                
                if len(df_com_pdf) == 0:
                    st.warning("⚠️ Nenhum PDF válido encontrado.")
                else:
                    zip_buffer = io.BytesIO()
                    progresso_bar = st.progress(0)
                    status_text = st.empty()
                    
                    nome_pasta_curso = criar_slug_curso(curso_selecionado)
                    sucessos = 0
                    erros = 0
                    total = len(df_com_pdf)
                    
                    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                        for idx, row in df_com_pdf.reset_index().iterrows():
                            porcentagem = (idx + 1) / total
                            progresso_bar.progress(porcentagem)
                            status_text.text(f"Baixando arquivo {idx+1} de {total}: {row['ATIVIDADE']}")
                            
                            try:
                                response = requests.get(row['PDF'], timeout=15)
                                if response.status_code == 200:
                                    ordem_aula = limpar_nome_arquivo(row['ORDEM AULA'])
                                    ordem_atividade = limpar_nome_arquivo(row['ORDEM ATIVIDADE'])
                                    uc = limpar_nome_arquivo(row['UC'])
                                    atividade = limpar_nome_arquivo(row['ATIVIDADE'])
                                    
                                    nome_pdf = f"{ordem_aula}.{ordem_atividade} - {uc} - {atividade}.pdf"
                                    
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
                            except Exception:
                                erros += 1
                                continue
                    
                    status_text.text("✨ Processamento concluído!")
                    st.balloons()
                    
                    if sucessos > 0:
                        st.success(f"📊 Download concluído! {sucessos} PDFs foram baixados.")
                    if erros > 0:
                        st.warning(f"⚠️ Houve erro no download de {erros} arquivos.")
                    
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
        st.info("Dica: Se for um erro do Excel (ex: openpyxl ou xlrd), certifique-se de que o arquivo está em formato .xlsx ou .csv limpo.")
else:
    st.info("💡 Por favor, carregue o arquivo Excel na barra lateral.")

import streamlit as st
import pandas as pd
import requests
import io
import zipfile
import re
import os
import shutil

st.set_page_config(page_title="Gerenciador de Catálogos e PDFs", page_icon="📚", layout="wide")

st.title("📚 Gerenciador de Download de Conteúdos (PDFs)")
st.markdown("---")

def limpar_nome_arquivo(nome):
    if pd.isna(nome) or str(nome).strip() == "":
        return "sem_nome"
    nome_limpo = re.sub(r'[\\\\/*?:"<>|]', "", str(nome))
    return nome_limpo.strip()

def criar_slug_curso(nome_curso):
    if pd.isna(nome_curso) or str(nome_curso).strip() == "":
        return "curso_sem_nome"
    slug = str(nome_curso).lower().strip()
    slug = re.sub(r'\s+', '_', slug)
    slug = re.sub(r'[\\\\/*?:"<>|]', "", slug)
    return slug

PASTA_TMP = "/tmp/download_workspace"
ARQUIVO_ZIP_FINAL = "/tmp/resultado_curso.zip"

st.sidebar.header("📁 Upload da Base de Dados")
# Aceita explicitamente formatos mais leves como CSV se necessário
arquivo_carregado = st.sidebar.file_uploader("Carregue o arquivo Excel unificado (.xlsx)", type=["xlsx", "xls", "csv"])

if arquivo_carregado is not None:
    try:
        colunas_obrigatorias = ['ANO', 'PERÍODO', 'CURSO', 'UC', 'CATÁLOGO', 'ORDEM AULA', 'ORDEM ATIVIDADE', 'ATIVIDADE', 'PDF']
        
        # Mapeamento de tipos para reduzir drasticamente o uso de memória RAM (Memória OOM)
        tipos_colunas = {col: str for col in colunas_obrigatorias}
        
        # Otimização: Carrega APENAS as colunas estritamente necessárias
        if arquivo_carregado.name.endswith('.csv'):
            df = pd.read_csv(arquivo_carregado, usecols=colunas_obrigatorias, dtype=tipos_colunas)
        else:
            df = pd.read_excel(arquivo_carregado, engine='openpyxl', usecols=colunas_obrigatorias, dtype=tipos_colunas)
        
        # Remove espaços ocultos nos nomes das colunas carregadas
        df.columns = [col.strip() for col in df.columns]
        
        st.success("✅ Base de dados otimizada carregada com sucesso!")
        
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
            st.dataframe(df_filtrado)
        
        st.markdown("---")
        st.subheader("⚙️ Processamento e Download dos PDFs")
        
        with st.form(key="form_processamento"):
            st.write("Clique abaixo para processar salvando em disco temporário.")
            botao_disparar = st.form_submit_button("🚀 Iniciar Processamento de Baixo Consumo", type="primary")
        
        if botao_disparar:
            if os.path.exists(PASTA_TMP):
                shutil.rmtree(PASTA_TMP)
            if os.path.exists(ARQUIVO_ZIP_FINAL):
                os.remove(ARQUIVO_ZIP_FINAL)
                
            df_com_pdf = df_filtrado[df_filtrado['PDF'].notna()].copy()
            df_com_pdf['PDF_STR'] = df_com_pdf['PDF'].astype(str).str.strip()
            df_com_pdf = df_com_pdf[df_com_pdf['PDF_STR'].str.startswith('http')]
            
            if len(df_com_pdf) == 0:
                st.warning("⚠️ Nenhum PDF válido encontrado.")
            else:
                progresso_bar = st.progress(0)
                status_text = st.empty()
                
                nome_pasta_curso = criar_slug_curso(curso_selecionado)
                sucessos = 0
                erros = 0
                total = len(df_com_pdf)
                
                relatorio_erros = []
                
                with zipfile.ZipFile(ARQUIVO_ZIP_FINAL, "w", compression=zipfile.ZIP_DEFLATED) as zip_file:
                    lista_dados = df_com_pdf.to_dict(orient="records")
                    
                    for idx, row in enumerate(lista_dados):
                        porcentagem = (idx + 1) / total
                        progresso_bar.progress(porcentagem)
                        
                        atividade_nome = str(row.get('ATIVIDADE', 'Sem Nome'))
                        uc_nome = str(row.get('UC', 'Sem_UC'))
                        status_text.text(f"Baixando {idx+1}/{total}: {atividade_nome}")
                        
                        link_pdf = row.get('PDF_STR', '')
                        
                        try:
                            response = requests.get(link_pdf, timeout=20)
                            if response.status_code == 200:
                                ordem_aula = limpar_nome_arquivo(row.get('ORDEM AULA', '0'))
                                ordem_atv = limpar_nome_arquivo(row.get('ORDEM ATIVIDADE', '0'))
                                atv = limpar_nome_arquivo(atividade_nome)
                                uc = limpar_nome_arquivo(uc_nome)
                                
                                nome_pdf = f"{ordem_aula}.{ordem_atv} - {uc} - {atv}.pdf"
                                caminho_no_zip = f"{nome_pasta_curso}/04_conteudos_especificos/conteudo/{nome_pdf}"
                                
                                zip_file.writestr(caminho_no_zip, response.content)
                                sucessos += 1
                            else:
                                erros += 1
                                relatorio_erros.append({
                                    "UC": uc_nome,
                                    "Atividade": atividade_nome,
                                    "Motivo": f"Erro HTTP {response.status_code}",
                                    "Link": link_pdf
                                })
                        except Exception:
                            erros += 1
                            relatorio_erros.append({
                                "UC": uc_nome,
                                        "Atividade": atividade_nome,
                                "Motivo": "Timeout/Falha de Conexão",
                                "Link": link_pdf
                            })
                            continue
                            
                status_text.text("✨ Compactação em disco concluída!")
                st.balloons()
                
                if sucessos > 0:
                    st.success(f"📊 {sucessos} PDFs foram estruturados com sucesso!")
                    st.session_state["zip_disponivel_no_disco"] = True
                    st.session_state["nome_do_curso_zip"] = nome_pasta_curso
                    
                if erros > 0:
                    st.warning(f"⚠️ {erros} links apresentaram falhas.")
                    with st.expander("🚨 Ver detalhes das falhas"):
                        df_erros = pd.DataFrame(relatorio_erros)
                        st.dataframe(df_erros, use_container_width=True)
        
        if st.session_state.get("zip_disponivel_no_disco") and os.path.exists(ARQUIVO_ZIP_FINAL):
            st.markdown("### 📥 Seu arquivo está pronto:")
            with open(ARQUIVO_ZIP_FINAL, "rb") as f:
                st.download_button(
                    label=f"Baixar {st.session_state.get('nome_do_curso_zip')}.zip",
                    data=f,
                    file_name=f"{st.session_state.get('nome_do_curso_zip')}.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                
    except Exception as e:
        st.error(f"Erro crítico ao ler o arquivo: {e}")
        st.info("Garanta que o arquivo carregado possui as colunas necessárias.")
else:
    st.session_state["zip_disponivel_no_disco"] = False
    st.info("💡 Por favor, carregue o arquivo Excel na barra lateral.")

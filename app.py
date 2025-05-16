import streamlit as st
from docx import Document
import pandas as pd
import plotly.express as px
# import spacy # Descomente se for usar NLP avançado
# nlp = spacy.load("pt_core_news_sm") # Exemplo para português

def extrair_dados_docx(uploaded_file):
    """Extrai textos e tabelas de um arquivo DOCX."""
    try:
        document = Document(uploaded_file)
        textos = [p.text for p in document.paragraphs if p.text.strip()]
        tabelas_dfs = []
        for i, table in enumerate(document.tables):
            data = []
            keys = None
            for j, row in enumerate(table.rows):
                text_cells = [cell.text.strip() for cell in row.cells]
                if j == 0:  # Assume primeira linha como cabeçalho
                    keys = text_cells
                    continue
                if keys:
                    # Garante que haja o mesmo número de chaves e valores
                    if len(keys) == len(text_cells):
                        data.append(dict(zip(keys, text_cells)))
                    else:
                        # Se não, tenta preencher com None ou loga um aviso
                        # st.warning(f"Tabela {i+1}, linha {j+1} tem contagem de células diferente do cabeçalho.")
                        # Opção: preencher com None para colunas faltantes
                        filled_row_data = {}
                        for k_idx, key in enumerate(keys):
                            filled_row_data[key] = text_cells[k_idx] if k_idx < len(text_cells) else None
                        data.append(filled_row_data)


            if data:
                try:
                    df = pd.DataFrame(data)
                    # Tentativa de conversão de tipos (mais robusta)
                    for col in df.columns:
                        try:
                            # Tenta converter para numérico (lida com vírgula como decimal)
                            df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', '.', regex=False))
                        except ValueError:
                            # Se falhar, tenta converter para datetime
                            try:
                                df[col] = pd.to_datetime(df[col])
                            except ValueError:
                                # Se falhar, mantém como objeto (string)
                                pass 
                    tabelas_dfs.append({"id": f"tabela_{i+1}", "dataframe": df, "nome": f"Tabela {i+1}"})
                except Exception as e:
                    st.warning(f"Não foi possível processar completamente a tabela {i+1}: {e}")
        return textos, tabelas_dfs
    except Exception as e:
        st.error(f"Erro ao ler o arquivo DOCX: {e}")
        return [], []

def sugerir_visualizacoes(textos, tabelas_dfs):
    """Gera sugestões de visualizações baseadas nos dados extraídos."""
    sugestoes = []
    sugestao_id_counter = 0

    for tabela_info in tabelas_dfs:
        df = tabela_info["dataframe"]
        nome_tabela = tabela_info["nome"]
        
        colunas_numericas = df.select_dtypes(include=['number']).columns.tolist()
        colunas_categoricas = df.select_dtypes(include=['object', 'category']).columns.tolist()
        colunas_datas = df.select_dtypes(include=['datetime']).columns.tolist()

        # Gráfico de Barras: 1 categórica + 1 numérica
        if colunas_categoricas and colunas_numericas:
            for cat_col in colunas_categoricas:
                for num_col in colunas_numericas:
                    if df[cat_col].nunique() > 1 and df[cat_col].nunique() < 50: # Evita excesso de barras
                        sugestao_id_counter += 1
                        sugestoes.append({
                            "id": f"sug_{sugestao_id_counter}", "tipo": "bar",
                            "fonte_dados_id": tabela_info["id"],
                            "x_col": cat_col, "y_col": num_col,
                            "titulo": f"Barras: {num_col} por {cat_col} ({nome_tabela})"
                        })
                        break # Pega a primeira combinação numérica para esta categórica
                # break # Se quiser só uma sugestão de barra por tabela

        # Gráfico de Pizza: 1 categórica (poucas categorias) + 1 numérica
        if colunas_categoricas and colunas_numericas:
            for cat_col in colunas_categoricas:
                if 1 < df[cat_col].nunique() < 10: # Bom para pizza
                    for num_col in colunas_numericas:
                        sugestao_id_counter += 1
                        sugestoes.append({
                            "id": f"sug_{sugestao_id_counter}", "tipo": "pie",
                            "fonte_dados_id": tabela_info["id"],
                            "names_col": cat_col, "values_col": num_col,
                            "titulo": f"Pizza: {num_col} por {cat_col} ({nome_tabela})"
                        })
                        break # Pega a primeira combinação numérica para esta categórica
                # break

        # Gráfico de Linha: 1 data + 1 numérica
        if colunas_datas and colunas_numericas:
            for date_col in colunas_datas:
                for num_col in colunas_numericas:
                    sugestao_id_counter += 1
                    sugestoes.append({
                        "id": f"sug_{sugestao_id_counter}", "tipo": "line",
                        "fonte_dados_id": tabela_info["id"],
                        "x_col": date_col, "y_col": num_col,
                        "titulo": f"Linha: {num_col} ao longo de {date_col} ({nome_tabela})"
                    })
                    break
                # break
        
        # Gráfico de Dispersão: 2 numéricas
        if len(colunas_numericas) >= 2:
            for i in range(len(colunas_numericas)):
                for j in range(i + 1, len(colunas_numericas)):
                    sugestao_id_counter += 1
                    sugestoes.append({
                        "id": f"sug_{sugestao_id_counter}", "tipo": "scatter",
                        "fonte_dados_id": tabela_info["id"],
                        "x_col": colunas_numericas[i], "y_col": colunas_numericas[j],
                        "titulo": f"Dispersão: {colunas_numericas[i]} vs {colunas_numericas[j]} ({nome_tabela})"
                    })
                    break # Só uma dispersão para o primeiro par encontrado
                break # Só um par de colunas numéricas para dispersão por tabela

    # TODO: Sugestões baseadas em texto (NLP ou Gemini API)
    return sugestoes

# --- Interface Streamlit ---
st.set_page_config(layout="wide")
st.title("Gerador de Dashboard a partir de DOCX 📄➡️📊")
st.markdown("Faça upload de um arquivo DOCX contendo tabelas e textos para análise e geração de gráficos.")

# Inicialização do estado da sessão
if 'sugestoes_geradas' not in st.session_state:
    st.session_state.sugestoes_geradas = []
if 'dados_extraidos' not in st.session_state:
    st.session_state.dados_extraidos = {"textos": [], "tabelas_dfs": []}
if 'sugestoes_validadas' not in st.session_state:
    st.session_state.sugestoes_validadas = {}
if 'arquivo_processado' not in st.session_state:
    st.session_state.arquivo_processado = None


uploaded_file = st.file_uploader("Escolha um arquivo DOCX", type=["docx"])

if uploaded_file is not None:
    # Se um novo arquivo for carregado, resetar o estado anterior
    if st.session_state.arquivo_processado != uploaded_file.name:
        st.session_state.sugestoes_geradas = []
        st.session_state.dados_extraidos = {"textos": [], "tabelas_dfs": []}
        st.session_state.sugestoes_validadas = {}
        st.session_state.arquivo_processado = uploaded_file.name # Marca o arquivo como processado

    if not st.session_state.sugestoes_geradas: # Processar apenas se não houver sugestões para o arquivo atual
        with st.spinner("Lendo e analisando o documento... Por favor, aguarde."):
            textos, tabelas_dfs = extrair_dados_docx(uploaded_file)
            st.session_state.dados_extraidos = {"textos": textos, "tabelas_dfs": tabelas_dfs}
            
            if not tabelas_dfs and not textos:
                st.warning("Nenhum dado extraível (texto ou tabela) encontrado no documento.")
            else:
                st.success(f"Documento '{uploaded_file.name}' lido com sucesso!")
                if tabelas_dfs:
                    st.write(f"Encontradas {len(tabelas_dfs)} tabelas.")
                    # Preview das tabelas (opcional, pode poluir)
                    # for t_info in tabelas_dfs:
                    #     with st.expander(f"Preview da {t_info['nome']}"):
                    #         st.dataframe(t_info['dataframe'].head())

                st.session_state.sugestoes_geradas = sugerir_visualizacoes(textos, tabelas_dfs)
                if not st.session_state.sugestoes_geradas:
                    st.info("Não foram encontradas sugestões automáticas de gráficos com base nos dados tabulares.")
                else:
                    st.success(f"{len(st.session_state.sugestoes_geradas)} sugestões de gráficos encontradas!")
                    # Inicializa o estado de validação para novas sugestões
                    for sugestao in st.session_state.sugestoes_geradas:
                        if sugestao['id'] not in st.session_state.sugestoes_validadas:
                            st.session_state.sugestoes_validadas[sugestao['id']] = {
                                "aceito": True, "tipo_grafico": sugestao['tipo'],
                                "x_col": sugestao.get('x_col'), "y_col": sugestao.get('y_col'),
                                "names_col": sugestao.get('names_col'), "values_col": sugestao.get('values_col'),
                                "titulo": sugestao['titulo']
                            }

# Exibir sugestões e permitir validação
if st.session_state.sugestoes_geradas:
    st.sidebar.header("⚙️ Valide as Sugestões")
    
    for i, sugestao_original in enumerate(st.session_state.sugestoes_geradas):
        s_id = sugestao_original['id']
        config_atual = st.session_state.sugestoes_validadas[s_id]

        with st.sidebar.expander(f"Sugestão: {config_atual['titulo']}", expanded=False):
            config_atual['aceito'] = st.checkbox("Incluir gráfico?", value=config_atual['aceito'], key=f"aceito_{s_id}")
            
            df_correspondente = next((t['dataframe'] for t in st.session_state.dados_extraidos['tabelas_dfs'] if t['id'] == sugestao_original['fonte_dados_id']), None)
            
            if df_correspondente is not None:
                opcoes_colunas = df_correspondente.columns.tolist()
                tipos_graficos_disponiveis = ['bar', 'line', 'pie', 'scatter', 'funnel', 'sunburst'] # Adicionar mais
                
                idx_tipo_atual = tipos_graficos_disponiveis.index(config_atual['tipo_grafico']) if config_atual['tipo_grafico'] in tipos_graficos_disponiveis else 0
                config_atual['tipo_grafico'] = st.selectbox("Tipo", options=tipos_graficos_disponiveis, index=idx_tipo_atual, key=f"tipo_{s_id}")

                if config_atual['tipo_grafico'] in ['bar', 'line', 'scatter']:
                    config_atual['x_col'] = st.selectbox("Eixo X", options=opcoes_colunas, index=opcoes_colunas.index(config_atual['x_col']) if config_atual['x_col'] in opcoes_colunas else 0, key=f"x_col_{s_id}")
                    config_atual['y_col'] = st.selectbox("Eixo Y", options=opcoes_colunas, index=opcoes_colunas.index(config_atual['y_col']) if config_atual['y_col'] in opcoes_colunas else (1 if len(opcoes_colunas)>1 else 0), key=f"y_col_{s_id}")
                elif config_atual['tipo_grafico'] in ['pie', 'funnel']:
                    config_atual['names_col'] = st.selectbox("Nomes/Categorias", options=opcoes_colunas, index=opcoes_colunas.index(config_atual['names_col']) if config_atual['names_col'] in opcoes_colunas else 0, key=f"names_col_{s_id}")
                    config_atual['values_col'] = st.selectbox("Valores", options=opcoes_colunas, index=opcoes_colunas.index(config_atual['values_col']) if config_atual['values_col'] in opcoes_colunas else (1 if len(opcoes_colunas)>1 else 0), key=f"values_col_{s_id}")
                # TODO: Adicionar opções para sunburst (path, values)
                
                config_atual['titulo'] = st.text_input("Título do Gráfico", value=config_atual['titulo'], key=f"titulo_{s_id}")
            st.session_state.sugestoes_validadas[s_id] = config_atual # Salva alterações

    if st.sidebar.button("Gerar Dashboard com Gráficos Selecionados", type="primary", use_container_width=True):
        st.header("🚀 Dashboard Gerado")
        graficos_para_exibir = []
        for sugestao_original in st.session_state.sugestoes_geradas:
            s_id = sugestao_original['id']
            config_atual = st.session_state.sugestoes_validadas[s_id]
            
            if config_atual['aceito']:
                df_grafico = next((t['dataframe'] for t in st.session_state.dados_extraidos['tabelas_dfs'] if t['id'] == sugestao_original['fonte_dados_id']), None)
                
                if df_grafico is not None:
                    try:
                        fig = None
                        tipo_grafico = config_atual['tipo_grafico']
                        titulo = config_atual['titulo']
                        
                        if tipo_grafico == 'bar' and config_atual.get('x_col') and config_atual.get('y_col'):
                            fig = px.bar(df_grafico, x=config_atual['x_col'], y=config_atual['y_col'], title=titulo)
                        elif tipo_grafico == 'line' and config_atual.get('x_col') and config_atual.get('y_col'):
                            fig = px.line(df_grafico, x=config_atual['x_col'], y=config_atual['y_col'], title=titulo, markers=True)
                        elif tipo_grafico == 'scatter' and config_atual.get('x_col') and config_atual.get('y_col'):
                            fig = px.scatter(df_grafico, x=config_atual['x_col'], y=config_atual['y_col'], title=titulo)
                        elif tipo_grafico == 'pie' and config_atual.get('names_col') and config_atual.get('values_col'):
                            fig = px.pie(df_grafico, names=config_atual['names_col'], values=config_atual['values_col'], title=titulo)
                        elif tipo_grafico == 'funnel' and config_atual.get('names_col') and config_atual.get('values_col'):
                            # Plotly Express não tem funnel para 'names' diretamente, é mais para x e y
                            # Usando uma adaptação ou outra lib, ou simplificando para x/y
                            # Por agora, vou usar x como categoria e y como valor para o funnel
                             fig = px.funnel(df_grafico, y=config_atual['names_col'], x=config_atual['values_col'], title=titulo)


                        if fig:
                            graficos_para_exibir.append(fig)
                        elif config_atual['aceito']: # Se estava aceito mas não gerou fig
                             st.warning(f"Não foi possível gerar o gráfico '{titulo}' com as configurações atuais. Verifique as colunas selecionadas.")

                    except Exception as e:
                        st.error(f"Erro ao gerar gráfico '{config_atual['titulo']}': {e}")
                elif config_atual['aceito']:
                    st.warning(f"Não foi possível encontrar os dados para o gráfico: {config_atual['titulo']}")
        
        if graficos_para_exibir:
            num_cols_dashboard = 2 # min(2, len(graficos_para_exibir)) # até 2 colunas
            cols_dashboard = st.columns(num_cols_dashboard) 
            for i, fig_dash in enumerate(graficos_para_exibir):
                with cols_dashboard[i % num_cols_dashboard]:
                    st.plotly_chart(fig_dash, use_container_width=True)
        elif any(s['aceito'] for s_id, s in st.session_state.sugestoes_validadas.items() if s_id in [sug_orig['id'] for sug_orig in st.session_state.sugestoes_geradas]):
             st.info("Nenhum gráfico pôde ser gerado com as seleções atuais. Verifique as configurações dos gráficos na barra lateral.")
        else:
            st.info("Nenhum gráfico foi selecionado para o dashboard.")

elif uploaded_file is None and st.session_state.arquivo_processado is not None: 
    # Limpar estado se o arquivo for removido após ter sido processado
    st.session_state.sugestoes_geradas = []
    st.session_state.dados_extraidos = {"textos": [], "tabelas_dfs": []}
    st.session_state.sugestoes_validadas = {}
    st.session_state.arquivo_processado = None
    st.experimental_rerun() 

import streamlit as st
from docx import Document
import pandas as pd
import plotly.express as px
import google.generativeai as genai
import json
import os

# --- Configuração da Chave da API do Gemini ---
def get_gemini_api_key():
    try:
        return st.secrets["GOOGLE_API_KEY"]
    except (FileNotFoundError, KeyError):
        api_key = os.environ.get("GOOGLE_API_KEY")
        if api_key:
            return api_key
        # Não mostra st.error aqui para não poluir se a chave não for usada imediatamente.
        # A função que usa a chave deve lidar com a ausência dela.
        return None

# --- Funções de Extração e Sugestão ---
def extrair_dados_docx(uploaded_file):
    """Extrai textos e tabelas de um arquivo DOCX com tratamento de tipos aprimorado."""
    try:
        document = Document(uploaded_file)
        textos = [p.text for p in document.paragraphs if p.text.strip()]
        tabelas_dfs = []
        for i, table in enumerate(document.tables):
            data = []
            keys = None
            nome_tabela_doc = f"Tabela_{i+1}" 

            for j, row in enumerate(table.rows):
                text_cells = [cell.text.strip() for cell in row.cells]
                if j == 0:
                    keys = text_cells
                    continue
                if keys:
                    if len(keys) == len(text_cells):
                        data.append(dict(zip(keys, text_cells)))
                    else:
                        filled_row_data = {}
                        for k_idx, key in enumerate(keys):
                            filled_row_data[key] = text_cells[k_idx] if k_idx < len(text_cells) else None
                        data.append(filled_row_data)
            
            if data:
                try:
                    df = pd.DataFrame(data)
                    for col in df.columns:
                        original_series = df[col].copy()

                        # 1. Tentar converter para numérico
                        try:
                            cleaned_series = df[col].astype(str).str.strip().str.replace(',', '.', regex=False)
                            cleaned_series = cleaned_series.str.replace(r'[R$\s%]', '', regex=True)
                            df[col] = pd.to_numeric(cleaned_series, errors='raise')
                            continue 
                        except (ValueError, TypeError):
                            df[col] = original_series.copy()

                        # 2. Tentar converter para datetime
                        try:
                            possible_formats = [
                                '%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', 
                                '%d-%m-%Y', '%m-%d-%Y', '%Y', '%Y%m%d', '%d.%m.%Y'
                            ]
                            converted_with_format = False
                            # Tenta converter para string primeiro para lidar com tipos mistos antes do datetime
                            temp_col_str = df[col].astype(str)

                            for fmt in possible_formats:
                                try:
                                    temp_series_fmt = pd.to_datetime(temp_col_str, format=fmt, errors='coerce')
                                    if temp_series_fmt.notna().sum() > len(df[col]) * 0.5 : # Se mais da metade for convertida
                                        df[col] = temp_series_fmt
                                        converted_with_format = True
                                        break 
                                except (ValueError, TypeError):
                                    continue
                            
                            if not converted_with_format:
                                inferred_series = pd.to_datetime(temp_col_str, errors='coerce', infer_datetime_format=True)
                                if inferred_series.notna().sum() > len(df[col]) * 0.5:
                                    df[col] = inferred_series
                                else:
                                    df[col] = original_series.copy()
                        except Exception:
                            df[col] = original_series.copy()
                    
                    # Garante que todas as colunas restantes sejam string se forem 'object' para evitar problemas com Arrow
                    for col in df.columns:
                        if df[col].dtype == 'object':
                            df[col] = df[col].astype(str).fillna('') # Preenche NaNs em colunas de string com string vazia

                    tabelas_dfs.append({"id": f"tabela_{i+1}", "dataframe": df, "nome": nome_tabela_doc})
                except Exception as e:
                    st.warning(f"Não foi possível processar completamente a tabela {i+1} ({nome_tabela_doc}): {e}")
        return textos, tabelas_dfs
    except Exception as e:
        st.error(f"Erro crítico ao ler o arquivo DOCX: {e}")
        return [], []

def obter_sugestoes_da_llm(texto_doc_completo, tabelas_dfs_list):
    api_key = get_gemini_api_key()
    if not api_key:
        st.warning("Nenhuma sugestão da LLM pôde ser gerada: chave da API do Gemini não configurada.")
        return []

    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel(
            model_name="gemini-1.5-flash-latest",
             safety_settings=[ 
                {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
                {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
                {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
                {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"},
            ]
        )

        tabelas_prompt_str = ""
        for tabela_info in tabelas_dfs_list:
            df = tabela_info["dataframe"]
            nome_tabela = tabela_info["nome"]
            id_tabela = tabela_info["id"]
            
            df_sample = df.head(10) 
            if len(df.columns) > 15: 
                df_sample = df_sample.iloc[:, :15]
            
            markdown_tabela = df_sample.to_markdown(index=False)
            tabelas_prompt_str += f"\n--- {nome_tabela} (ID para referência: {id_tabela}) ---\n"
            tabelas_prompt_str += f"Colunas disponíveis (e seus tipos de dados inferidos): "
            col_types_str = ", ".join([f"{col} ({str(dtype)})" for col, dtype in df.dtypes.items()])
            tabelas_prompt_str += f"{col_types_str}\n"
            tabelas_prompt_str += f"{markdown_tabela}\n"

        max_texto_len = 80000 
        texto_doc_para_prompt = texto_doc_completo[:max_texto_len]
        if len(texto_doc_completo) > max_texto_len:
            texto_doc_para_prompt += "\n[TEXTO DO DOCUMENTO TRUNCADO ...]"

        prompt_completo = f"""
        Você é um assistente especialista em análise de dados e visualização.
        Analise o seguinte conteúdo de um DOCX: texto e tabelas (com seus IDs, nomes de colunas e tipos de dados inferidos).

        [TEXTO DO DOCUMENTO]
        {texto_doc_para_prompt}
        [FIM DO TEXTO]

        [TABELAS DO DOCUMENTO]
        {tabelas_prompt_str}
        [FIM DAS TABELAS]

        Baseado no conteúdo, gere uma lista JSON de sugestões de visualizações. Cada objeto na lista deve ter:
        - "sugestao_id": string, ID único (ex: "llm_sug_1").
        - "titulo_sugerido": string, título descritivo.
        - "tipo_grafico_sugerido": string (ex: "pizza", "bar", "line", "scatter", "diagrama_swot_lista", "tabela_informativa").
        - "fonte_dados_id": string, ID da tabela (ex: "tabela_1") ou descrição da seção do texto (ex: "texto_swot_ifood").
        - "parametros_grafico": objeto com parâmetros. USE OS NOMES EXATOS DAS COLUNAS E SEUS TIPOS DE DADOS FORNECIDOS.
            - Para gráficos (pizza, bar, line, scatter): {{ "eixo_x": "NomeColunaX", "eixo_y": "NomeColunaY", "valores": "NomeColunaValores", "categorias": "NomeColunaCategorias" }}
            - Para "diagrama_swot_lista": {{ "forcas": ["Ponto 1"], "fraquezas": ["Ponto A"], "oportunidades": ["Ponto X"], "ameacas": ["Ponto Z"] }}
            - Para "tabela_informativa": {{ "id_tabela_original": "ID_da_Tabela_Referenciada" }}
            - Se um gráfico requer dados extraídos do texto (não de uma tabela), adicione um campo "dados_diretos": [{{ "col1_nome": val1, "col2_nome": val2 }}, ...]
        - "justificativa": string, breve explicação.

        Exemplo SWOT: {{"sugestao_id": "llm_swot1", "titulo_sugerido": "SWOT XYZ", "tipo_grafico_sugerido": "diagrama_swot_lista", "fonte_dados_id": "texto_swot_xyz", "parametros_grafico": {{"forcas": ["F1"], "fraquezas": ["Fr1"], "oportunidades": ["O1"], "ameacas": ["A1"]}}, "justificativa": "SWOT XYZ."}}
        Exemplo PIZZA para tabela_2 com colunas 'Player' (object) e 'Market Share (%)' (float64): {{"sugestao_id": "llm_pizza1", "titulo_sugerido": "Market Share", "tipo_grafico_sugerido": "pizza", "fonte_dados_id": "tabela_2", "parametros_grafico": {{"categorias": "Player", "valores": "Market Share (%)"}}, "justificativa": "Market share."}}
        Exemplo BAR com DADOS DIRETOS (extraídos do texto): {{"sugestao_id": "llm_bar_direto", "titulo_sugerido": "Downloads por Player", "tipo_grafico_sugerido": "bar", "fonte_dados_id": "texto_metricas_downloads", "parametros_grafico": {{"dados_diretos": [{{"Player": "AppA", "Downloads": 1000}}, {{"Player": "AppB", "Downloads": 1500}}], "eixo_x": "Player", "eixo_y": "Downloads"}}, "justificativa": "Downloads."}}
        Retorne APENAS a lista JSON válida.
        """
        
        with st.spinner("🤖 Consultando o Gemini para obter sugestões avançadas..."):
            # st.text_area("Debug: Prompt enviado ao Gemini", prompt_completo, height=200)
            response = model.generate_content(prompt_completo)
            
        cleaned_response_text = response.text.strip()
        if cleaned_response_text.startswith("```json"):
            cleaned_response_text = cleaned_response_text[7:].strip()
        if cleaned_response_text.endswith("```"):
            cleaned_response_text = cleaned_response_text[:-3].strip()
        
        # st.text_area("Debug: Resposta do Gemini (após limpeza)", cleaned_response_text, height=200)

        sugestoes_llm = json.loads(cleaned_response_text)
        
        if isinstance(sugestoes_llm, list) and all(isinstance(item, dict) for item in sugestoes_llm):
             st.success(f"{len(sugestoes_llm)} sugestões avançadas recebidas do Gemini!")
             return sugestoes_llm
        else:
            st.error("A resposta do Gemini não está no formato de lista JSON esperado.")
            return []

    except json.JSONDecodeError as json_e:
        st.error(f"Erro ao decodificar JSON da resposta do Gemini: {json_e}")
        if 'response' in locals() and hasattr(response, 'text'):
            st.code(response.text, language="text")
        return []
    except Exception as e:
        st.error(f"Erro ao chamar a API do Gemini: {e}")
        import traceback
        st.text(traceback.format_exc())
        return []

# --- Interface Streamlit ---
st.set_page_config(layout="wide")
st.title("Gerador de Dashboard a partir de DOCX 📄➡️📊 (com Gemini AI)")
st.markdown("Faça upload de um arquivo DOCX. A IA analisará o conteúdo e sugerirá visualizações.")

if 'sugestoes_geradas' not in st.session_state: st.session_state.sugestoes_geradas = []
if 'dados_extraidos' not in st.session_state: st.session_state.dados_extraidos = {"textos_list": [], "tabelas_dfs": []}
if 'sugestoes_validadas' not in st.session_state: st.session_state.sugestoes_validadas = {}
if 'arquivo_processado_nome' not in st.session_state: st.session_state.arquivo_processado_nome = None
if 'debug_dataframes_shown' not in st.session_state: st.session_state.debug_dataframes_shown = False

uploaded_file = st.file_uploader("Escolha um arquivo DOCX", type=["docx"], key="file_uploader")
show_debug_info = st.sidebar.checkbox("Mostrar Informações de Depuração", value=False)

if uploaded_file is not None:
    if st.session_state.arquivo_processado_nome != uploaded_file.name:
        st.session_state.sugestoes_geradas = []
        st.session_state.dados_extraidos = {"textos_list": [], "tabelas_dfs": []}
        st.session_state.sugestoes_validadas = {}
        st.session_state.arquivo_processado_nome = uploaded_file.name
        st.session_state.debug_dataframes_shown = False

    if not st.session_state.sugestoes_geradas:
        with st.spinner("Lendo e pré-processando o documento..."):
            textos_list, tabelas_dfs_list = extrair_dados_docx(uploaded_file)
            st.session_state.dados_extraidos = {"textos_list": textos_list, "tabelas_dfs": tabelas_dfs_list}
        
        if not tabelas_dfs_list and not textos_list:
            st.warning("Nenhum dado extraível (texto ou tabela) encontrado no documento.")
        else:
            st.success(f"Documento '{uploaded_file.name}' lido!")
            if show_debug_info and not st.session_state.debug_dataframes_shown:
                with st.expander("DEBUG: DataFrames Extraídos e Tipos de Dados", expanded=False):
                    for t_info in st.session_state.dados_extraidos['tabelas_dfs']:
                        st.write(f"ID: {t_info['id']}, Nome: {t_info['nome']}")
                        st.dataframe(t_info['dataframe'].head().astype(str)) # .astype(str) para evitar erro do Arrow
                        st.write("Tipos de dados das colunas:", t_info['dataframe'].dtypes)
                        st.divider()
                st.session_state.debug_dataframes_shown = True

            texto_completo_doc = "\n\n".join(textos_list)
            sugestoes_da_llm = obter_sugestoes_da_llm(texto_completo_doc, tabelas_dfs_list)
            st.session_state.sugestoes_geradas = sugestoes_da_llm

            if not st.session_state.sugestoes_geradas:
                st.info("Nenhuma sugestão foi gerada pela IA para este documento.")
            else:
                for sugestao in st.session_state.sugestoes_geradas:
                    s_id = sugestao.get('sugestao_id', f"sug_{hash(sugestao.get('titulo_sugerido', ''))}")
                    if s_id not in st.session_state.sugestoes_validadas:
                        st.session_state.sugestoes_validadas[s_id] = {
                            "aceito": True,
                            "tipo_grafico": sugestao.get('tipo_grafico_sugerido', 'desconhecido'),
                            "titulo": sugestao.get('titulo_sugerido', 'Título não fornecido'),
                            "fonte_dados_id": sugestao.get('fonte_dados_id', 'desconhecido'),
                            "parametros_grafico_completos": sugestao.get('parametros_grafico', {}),
                            "justificativa": sugestao.get('justificativa', 'N/A')
                        }

if st.session_state.sugestoes_geradas:
    st.sidebar.header("⚙️ Valide as Sugestões da IA")
    for sugestao_original in st.session_state.sugestoes_geradas:
        s_id = sugestao_original.get('sugestao_id', f"sug_{hash(sugestao_original.get('titulo_sugerido', ''))}")
        if s_id not in st.session_state.sugestoes_validadas: continue
        config_atual = st.session_state.sugestoes_validadas[s_id]
        with st.sidebar.expander(f"Sug.: {config_atual['titulo']}", expanded=False):
            st.caption(f"Tipo: {config_atual['tipo_grafico']} | Fonte: {config_atual['fonte_dados_id']}")
            st.markdown(f"**Justificativa IA:** *{config_atual.get('justificativa', 'N/A')}*")
            config_atual['aceito'] = st.checkbox("Incluir?", value=config_atual['aceito'], key=f"aceito_{s_id}")
            config_atual['titulo'] = st.text_input("Título", value=config_atual['titulo'], key=f"titulo_{s_id}")
            params_grafico = config_atual['parametros_grafico_completos']
            if config_atual['tipo_grafico'] in ['bar', 'line', 'pie', 'scatter'] and \
               not str(config_atual['fonte_dados_id']).startswith("texto_") and \
               not params_grafico.get("dados_diretos"): # Não mostra edição de colunas se dados_diretos existem
                df_correspondente = next((t['dataframe'] for t in st.session_state.dados_extraidos['tabelas_dfs'] if t['id'] == config_atual['fonte_dados_id']), None)
                if df_correspondente is not None:
                    opcoes_colunas = [""] + df_correspondente.columns.tolist()
                    if config_atual['tipo_grafico'] in ['bar', 'line', 'scatter']:
                        params_grafico['eixo_x'] = st.selectbox("Eixo X", options=opcoes_colunas, index=opcoes_colunas.index(params_grafico.get('eixo_x', "")) if params_grafico.get('eixo_x', "") in opcoes_colunas else 0, key=f"x_col_{s_id}")
                        params_grafico['eixo_y'] = st.selectbox("Eixo Y", options=opcoes_colunas, index=opcoes_colunas.index(params_grafico.get('eixo_y', "")) if params_grafico.get('eixo_y', "") in opcoes_colunas else 0, key=f"y_col_{s_id}")
                    elif config_atual['tipo_grafico'] == 'pie':
                        params_grafico['categorias'] = st.selectbox("Categorias", options=opcoes_colunas, index=opcoes_colunas.index(params_grafico.get('categorias', "")) if params_grafico.get('categorias', "") in opcoes_colunas else 0, key=f"names_col_{s_id}")
                        params_grafico['valores'] = st.selectbox("Valores", options=opcoes_colunas, index=opcoes_colunas.index(params_grafico.get('valores', "")) if params_grafico.get('valores', "") in opcoes_colunas else 0, key=f"values_col_{s_id}")
            st.session_state.sugestoes_validadas[s_id] = config_atual

    if st.sidebar.button("Gerar Dashboard com Gráficos Selecionados", type="primary", use_container_width=True):
        st.header("🚀 Dashboard Gerado")
        if show_debug_info:
            with st.expander("DEBUG: Parâmetros Validados para Plotagem", expanded=False):
                for s_id_debug, config_debug in st.session_state.sugestoes_validadas.items():
                    if config_debug['aceito']:
                        st.write(f"ID: {s_id_debug}, Título: {config_debug['titulo']}, Tipo: {config_debug['tipo_grafico']}")
                        st.write(f"Fonte: {config_debug['fonte_dados_id']}")
                        st.json(config_debug['parametros_grafico_completos'])
                        if not str(config_debug['fonte_dados_id']).startswith("texto_") and not config_debug['parametros_grafico_completos'].get("dados_diretos"):
                            df_debug_plot = next((t['dataframe'] for t in st.session_state.dados_extraidos['tabelas_dfs'] if t['id'] == config_debug['fonte_dados_id']), None)
                            if df_debug_plot is not None:
                                st.dataframe(df_debug_plot.head().astype(str))
                                st.write(df_debug_plot.dtypes)
                        st.divider()
        
        elementos_renderizados_count = 0
        num_cols_dashboard = 2 
        cols_dashboard = st.columns(num_cols_dashboard)
        col_idx = 0

        for s_id, config_atual in st.session_state.sugestoes_validadas.items():
            if config_atual['aceito']:
                tipo_grafico = config_atual['tipo_grafico']
                titulo = config_atual['titulo']
                params_completos = config_atual['parametros_grafico_completos']
                fonte_id = config_atual['fonte_dados_id']
                df_para_plotar = None
                elemento_renderizado_nesta_iteracao = False

                with cols_dashboard[col_idx % num_cols_dashboard]:
                    try:
                        if params_completos.get("dados_diretos"):
                            try:
                                df_para_plotar = pd.DataFrame(params_completos["dados_diretos"])
                            except Exception as e_df_direto:
                                st.warning(f"'{titulo}': Não foi possível criar DataFrame de 'dados_diretos'. Erro: {e_df_direto}")
                                continue
                        elif not str(fonte_id).startswith("texto_"):
                             df_para_plotar = next((t['dataframe'] for t in st.session_state.dados_extraidos['tabelas_dfs'] if t['id'] == fonte_id), None)
                             if df_para_plotar is None:
                                 st.warning(f"'{titulo}': Dados da tabela '{fonte_id}' não encontrados.")
                                 continue
                        
                        if df_para_plotar is not None: # Verifica se temos um DataFrame para plotar
                            if tipo_grafico == 'bar' and params_completos.get('eixo_x') and params_completos.get('eixo_y'):
                                st.plotly_chart(px.bar(df_para_plotar, x=params_completos['eixo_x'], y=params_completos['eixo_y'], title=titulo), use_container_width=True)
                                elemento_renderizado_nesta_iteracao = True
                            elif tipo_grafico == 'line' and params_completos.get('eixo_x') and params_completos.get('eixo_y'):
                                st.plotly_chart(px.line(df_para_plotar, x=params_completos['eixo_x'], y=params_completos['eixo_y'], title=titulo, markers=True), use_container_width=True)
                                elemento_renderizado_nesta_iteracao = True
                            elif tipo_grafico == 'scatter' and params_completos.get('eixo_x') and params_completos.get('eixo_y'):
                                st.plotly_chart(px.scatter(df_para_plotar, x=params_completos['eixo_x'], y=params_completos['eixo_y'], title=titulo), use_container_width=True)
                                elemento_renderizado_nesta_iteracao = True
                            elif tipo_grafico == 'pie' and params_completos.get('categorias') and params_completos.get('valores'):
                                st.plotly_chart(px.pie(df_para_plotar, names=params_completos['categorias'], values=params_completos['valores'], title=titulo), use_container_width=True)
                                elemento_renderizado_nesta_iteracao = True
                            elif tipo_grafico == 'tabela_informativa':
                                st.subheader(titulo)
                                st.dataframe(df_para_plotar.astype(str), use_container_width=True)
                                elemento_renderizado_nesta_iteracao = True
                        
                        # Elementos que não dependem de df_para_plotar (como SWOT)
                        if tipo_grafico == 'diagrama_swot_lista':
                            st.subheader(titulo)
                            c1, c2 = st.columns(2)
                            swot_items = {"forcas": "Forças 💪", "fraquezas": "Fraquezas 📉", "oportunidades": "Oportunidades 🚀", "ameacas": "Ameaças ⚠️"}
                            current_col_map = {0: c1, 1:c1, 2:c2, 3:c2}
                            for i_key_idx, (key, header) in enumerate(swot_items.items()):
                                with current_col_map[i_key_idx]:
                                    st.markdown(f"##### {header}")
                                    items = params_completos.get(key, ["N/A"])
                                    if not items: items = ["N/A"] 
                                    for item in items: st.markdown(f"- {item}")
                            elemento_renderizado_nesta_iteracao = True
                        
                        if not elemento_renderizado_nesta_iteracao and config_atual['aceito']:
                             st.warning(f"Não foi possível gerar visualização para '{titulo}' (tipo: {tipo_grafico}).")
                    except Exception as e:
                        st.error(f"Erro ao gerar '{titulo}': {e}")
                
                if elemento_renderizado_nesta_iteracao:
                    col_idx +=1
                    elementos_renderizados_count +=1
        
        if elementos_renderizados_count == 0 and any(c['aceito'] for c in st.session_state.sugestoes_validadas.values()):
            st.info("Nenhum elemento pôde ser gerado para o dashboard com as seleções atuais.")
        elif elementos_renderizados_count == 0:
            st.info("Nenhum elemento foi selecionado para o dashboard.")

elif uploaded_file is None and st.session_state.arquivo_processado_nome is not None: 
    st.session_state.sugestoes_geradas = []
    st.session_state.dados_extraidos = {"textos_list": [], "tabelas_dfs": []}
    st.session_state.sugestoes_validadas = {}
    st.session_state.arquivo_processado_nome = None
    st.session_state.debug_dataframes_shown = False
    if "file_uploader" in st.session_state: 
        st.session_state.file_uploader = None
    st.experimental_rerun()

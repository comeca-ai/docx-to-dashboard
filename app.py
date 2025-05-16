import streamlit as st
from docx import Document
import pandas as pd
import plotly.express as px
import google.generativeai as genai
import json
import os
import traceback # Para logs de erro mais detalhados

# --- Configuração da Chave da API do Gemini ---
def get_gemini_api_key():
    try:
        return st.secrets["GOOGLE_API_KEY"]
    except (FileNotFoundError, KeyError):
        api_key = os.environ.get("GOOGLE_API_KEY")
        if api_key:
            return api_key
        return None # A função que usa a chave deve lidar com a ausência dela.

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
            # Tenta pegar um nome/título para a tabela do parágrafo anterior (heurística)
            nome_tabela_doc = f"Tabela_{i+1}_Conteudo" 
            try:
                prev_element = table._element.getprevious()
                if prev_element is not None and prev_element.tag.endswith('p'):
                    # Extrai texto do parágrafo
                    p_text = "".join(node.text for node in prev_element.xpath('.//w:t'))
                    if p_text.strip():
                        nome_tabela_doc = p_text.strip().replace(":", "").replace("\n", " ").strip()
                        nome_tabela_doc = nome_tabela_doc[:50] + (nome_tabela_doc[50:] and '...') # Limita o tamanho
            except Exception:
                pass # Mantém o nome padrão se falhar

            for j, row in enumerate(table.rows):
                text_cells = [cell.text.strip() for cell in row.cells]
                if j == 0: # Assume primeira linha como cabeçalho
                    keys = [key.replace("\n", " ").strip() if key else f"Coluna_{k_idx+1}" for k_idx, key in enumerate(text_cells)]
                    continue
                if keys:
                    row_data = {}
                    for k_idx, key_name in enumerate(keys):
                        row_data[key_name] = text_cells[k_idx] if k_idx < len(text_cells) else None
                    data.append(row_data)
            
            if data:
                try:
                    df = pd.DataFrame(data)
                    for col_name in df.columns:
                        col_series = df[col_name].copy()
                        try:
                            # 1. Tentar converter para numérico
                            cleaned_series = col_series.astype(str).str.strip().str.replace(r'\.(?=\d{3})', '', regex=True).str.replace(',', '.', regex=False)
                            cleaned_series = cleaned_series.str.replace(r'[R$\s%()]', '', regex=True) # Remove R$, espaço, %, parênteses (para negativos)
                            # Trata negativos que podem estar com '-' no final ou entre parênteses
                            is_negative = cleaned_series.str.endswith('-') | (cleaned_series.str.startswith('(') & cleaned_series.str.endswith(')'))
                            cleaned_series_num = cleaned_series.str.replace(r'[-()]', '', regex=True)
                            
                            numeric_col = pd.to_numeric(cleaned_series_num, errors='raise')
                            numeric_col[is_negative.fillna(False)] *= -1 # Aplica sinal negativo
                            df[col_name] = numeric_col
                            continue 
                        except (ValueError, TypeError, AttributeError):
                            df[col_name] = col_series.copy() # Reverte se falhar

                        # 2. Tentar converter para datetime
                        try:
                            temp_col_str = df[col_name].astype(str)
                            possible_formats = [
                                '%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', 
                                '%d-%m-%Y', '%m-%d-%Y', '%Y', '%Y%m%d', '%d.%m.%Y'
                            ]
                            converted_with_format = False
                            for fmt in possible_formats:
                                try:
                                    temp_series_fmt = pd.to_datetime(temp_col_str, format=fmt, errors='coerce')
                                    if temp_series_fmt.notna().sum() > len(df[col_name]) * 0.5:
                                        df[col_name] = temp_series_fmt
                                        converted_with_format = True
                                        break 
                                except (ValueError, TypeError):
                                    continue
                            if not converted_with_format:
                                inferred_series = pd.to_datetime(temp_col_str, errors='coerce', infer_datetime_format=True)
                                if inferred_series.notna().sum() > len(df[col_name]) * 0.5:
                                    df[col_name] = inferred_series
                                else:
                                    df[col_name] = col_series.copy()
                        except Exception:
                            df[col_name] = col_series.copy()
                    
                    for col_name in df.columns: # Fallback final para string
                        if df[col_name].dtype == 'object':
                            df[col_name] = df[col_name].astype(str).fillna('')
                    
                    tabelas_dfs.append({"id": f"tabela_{i+1}", "dataframe": df, "nome": nome_tabela_doc})
                except Exception as e_df:
                    st.warning(f"Não foi possível processar DataFrame para tabela {i+1} ({nome_tabela_doc}): {e_df}")
        return textos, tabelas_dfs
    except Exception as e_doc:
        st.error(f"Erro crítico ao ler o arquivo DOCX: {e_doc}")
        return [], []

def obter_sugestoes_da_llm(texto_doc_completo, tabelas_dfs_list):
    api_key = get_gemini_api_key()
    if not api_key:
        st.warning("Nenhuma sugestão da LLM pôde ser gerada: chave da API do Gemini não configurada.")
        return []
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel(model_name="gemini-1.5-flash-latest",
                                      safety_settings=[{"category": c, "threshold": "BLOCK_NONE"} for c in ["HARM_CATEGORY_HARASSMENT", "HARM_CATEGORY_HATE_SPEECH", "HARM_CATEGORY_SEXUALLY_EXPLICIT", "HARM_CATEGORY_DANGEROUS_CONTENT"]])
        tabelas_prompt_str = ""
        for tabela_info in tabelas_dfs_list:
            df = tabela_info["dataframe"]
            nome_tabela = tabela_info["nome"]
            id_tabela = tabela_info["id"]
            df_sample = df.head(7) 
            if len(df.columns) > 10: df_sample = df_sample.iloc[:, :10]
            markdown_tabela = df_sample.to_markdown(index=False)
            tabelas_prompt_str += f"\n--- Tabela '{nome_tabela}' (ID para referência: {id_tabela}) ---\n"
            col_types_str = ", ".join([f"'{col}' ({str(dtype)})" for col, dtype in df.dtypes.items()])
            tabelas_prompt_str += f"Colunas disponíveis: {col_types_str}\n{markdown_tabela}\n"

        max_texto_len = 70000 
        texto_doc_para_prompt = texto_doc_completo[:max_texto_len] + ("\n[TEXTO TRUNCADO...]" if len(texto_doc_completo) > max_texto_len else "")

        prompt_completo = f"""
        Você é um assistente especialista em análise de dados e visualização.
        Analise o seguinte conteúdo de um DOCX: texto e tabelas (com seus IDs, nomes de colunas e tipos de dados inferidos).

        [TEXTO DO DOCUMENTO]
        {texto_doc_para_prompt}
        [FIM DO TEXTO]

        [TABELAS DO DOCUMENTO]
        {tabelas_prompt_str}
        [FIM DAS TABELAS]

        Baseado no conteúdo, gere uma lista JSON de sugestões de visualizações. Cada objeto deve ter:
        - "sugestao_id": string (ex: "llm_sug_1").
        - "titulo_sugerido": string.
        - "tipo_grafico_sugerido": string (ex: "pizza", "bar", "line", "scatter", "diagrama_swot_lista", "tabela_informativa", "metrica").
        - "fonte_dados_id": string (ID da tabela ex: "tabela_1", ou descrição da seção do texto ex: "texto_swot_ifood").
        - "parametros_grafico": objeto. USE NOMES EXATOS DAS COLUNAS E CONSIDERE SEUS TIPOS.
            - Para "bar", "line", "scatter": {{"eixo_x": "NomeColunaX", "eixo_y": "NomeColunaY"}} (Y deve ser numérico, X pode ser categórico/data/numérico).
            - Para "pizza": {{"categorias": "NomeColunaCategorias", "valores": "NomeColunaValores"}} (Valores deve ser numérico).
            - Para "diagrama_swot_lista": {{"forcas": ["Ponto 1"], "fraquezas": ["Fr1"], "oportunidades": ["O1"], "ameacas": ["A1"]}}.
            - Para "tabela_informativa": {{"id_tabela_original": "ID_da_Tabela_Referenciada"}}.
            - Para "metrica": {{"valor": "ValorPrincipalKPI", "delta": "ValorDeMudancaOpcional", "descricao_curta_kpi": "Breve descrição"}}.
            - Se um gráfico (bar, line, pie, scatter) requer DADOS EXTRAÍDOS DIRETAMENTE DO TEXTO (não de uma tabela listada), adicione: {{"dados_diretos": [{{"NomeEixoX": valX1, "NomeEixoY": valY1}}, ...], "eixo_x": "NomeEixoX", "eixo_y": "NomeEixoY"}}.
        - "justificativa": string.

        Exemplo SWOT: {{"sugestao_id": "llm_swot1", "titulo_sugerido": "SWOT XYZ", "tipo_grafico_sugerido": "diagrama_swot_lista", "fonte_dados_id": "texto_swot_xyz", "parametros_grafico": {{"forcas": ["F1"], "fraquezas": ["Fr1"], "oportunidades": ["O1"], "ameacas": ["A1"]}}, "justificativa": "SWOT XYZ."}}
        Exemplo PIZZA para tabela_2 com colunas 'Player' (object) e 'Market Share (%)' (float64): {{"sugestao_id": "llm_pizza1", "titulo_sugerido": "Market Share", "tipo_grafico_sugerido": "pizza", "fonte_dados_id": "tabela_2", "parametros_grafico": {{"categorias": "Player", "valores": "Market Share (%)"}}, "justificativa": "Market share."}}
        Exemplo BAR com DADOS DIRETOS: {{"sugestao_id": "llm_bar_direto", "titulo_sugerido": "Downloads", "tipo_grafico_sugerido": "bar", "fonte_dados_id": "texto_metricas", "parametros_grafico": {{"dados_diretos": [{{"Aplicativo": "AppA", "Downloads": 1000}}, {{"Aplicativo": "AppB", "Downloads": 1500}}], "eixo_x": "Aplicativo", "eixo_y": "Downloads"}}, "justificativa": "Downloads."}}
        Exemplo METRICA: {{"sugestao_id": "llm_kpi1", "titulo_sugerido": "Receita Total", "tipo_grafico_sugerido": "metrica", "fonte_dados_id": "texto_sumario", "parametros_grafico": {{"valor": "US$ 1.2 Bi", "delta": "20% vs Ano Anterior"}}, "justificativa": "Receita total."}}
        Retorne APENAS a lista JSON válida.
        """
        with st.spinner("🤖 Consultando o Gemini para obter sugestões avançadas..."):
            # st.text_area("Debug: Prompt enviado ao Gemini", prompt_completo, height=300)
            response = model.generate_content(prompt_completo)
        cleaned_response_text = response.text.strip()
        if cleaned_response_text.startswith("```json"): cleaned_response_text = cleaned_response_text[7:].strip()
        if cleaned_response_text.endswith("```"): cleaned_response_text = cleaned_response_text[:-3].strip()
        # st.text_area("Debug: Resposta do Gemini (após limpeza)", cleaned_response_text, height=300)
        sugestoes_llm = json.loads(cleaned_response_text)
        if isinstance(sugestoes_llm, list) and all(isinstance(item, dict) for item in sugestoes_llm):
             st.success(f"{len(sugestoes_llm)} sugestões avançadas recebidas do Gemini!")
             return sugestoes_llm
        st.error("Resposta do Gemini não é lista JSON esperada."); return []
    except json.JSONDecodeError as e:
        st.error(f"Erro ao decodificar JSON do Gemini: {e}")
        if 'response' in locals(): st.code(response.text, language="text")
        return []
    except Exception as e:
        st.error(f"Erro ao chamar API Gemini: {e}"); st.text(traceback.format_exc()); return []

# --- Interface Streamlit e Lógica Principal ---
st.set_page_config(layout="wide"); st.title("Gerador de Dashboard a partir de DOCX 📄➡️📊 (com Gemini AI)")
st.markdown("Faça upload de um arquivo DOCX. A IA analisará o conteúdo e sugerirá visualizações.")

if 'sugestoes_geradas' not in st.session_state: st.session_state.sugestoes_geradas = []
if 'dados_extraidos' not in st.session_state: st.session_state.dados_extraidos = {"textos_list": [], "tabelas_dfs": []}
if 'sugestoes_validadas' not in st.session_state: st.session_state.sugestoes_validadas = {}
if 'arquivo_processado_nome' not in st.session_state: st.session_state.arquivo_processado_nome = None
if 'debug_info_shown' not in st.session_state: st.session_state.debug_info_shown = False

uploaded_file = st.file_uploader("Escolha um arquivo DOCX", type=["docx"], key="file_uploader")
show_debug_info = st.sidebar.checkbox("Mostrar Informações de Depuração", value=st.session_state.get('show_debug_checkbox', False), key="show_debug_checkbox")
st.session_state.show_debug_checkbox = show_debug_info


if uploaded_file is not None:
    if st.session_state.arquivo_processado_nome != uploaded_file.name:
        st.session_state.sugestoes_geradas = []
        st.session_state.dados_extraidos = {"textos_list": [], "tabelas_dfs": []}
        st.session_state.sugestoes_validadas = {}
        st.session_state.arquivo_processado_nome = uploaded_file.name
        st.session_state.debug_info_shown = False

    if not st.session_state.sugestoes_geradas:
        with st.spinner("Lendo e pré-processando o documento..."):
            textos_list, tabelas_dfs_list = extrair_dados_docx(uploaded_file)
            st.session_state.dados_extraidos = {"textos_list": textos_list, "tabelas_dfs": tabelas_dfs_list}
        
        if not tabelas_dfs_list and not textos_list:
            st.warning("Nenhum dado extraível (texto ou tabela) encontrado no documento.")
        else:
            st.success(f"Documento '{uploaded_file.name}' lido!")
            if show_debug_info and not st.session_state.debug_info_shown: # Mostrar debug apenas uma vez por upload
                with st.expander("DEBUG: DataFrames Extraídos e Tipos de Dados (após extração)", expanded=False):
                    for t_info in st.session_state.dados_extraidos['tabelas_dfs']:
                        st.write(f"ID: {t_info['id']}, Nome da Tabela no DOCX: {t_info['nome']}")
                        st.dataframe(t_info['dataframe'].head().astype(str))
                        st.write("Tipos de dados das colunas:", t_info['dataframe'].dtypes)
                        st.divider()
                st.session_state.debug_info_shown = True
            
            texto_completo_doc = "\n\n".join(textos_list)
            sugestoes_da_llm = obter_sugestoes_da_llm(texto_completo_doc, tabelas_dfs_list)
            st.session_state.sugestoes_geradas = sugestoes_da_llm

            if not st.session_state.sugestoes_geradas:
                st.info("Nenhuma sugestão foi gerada pela IA para este documento.")
            else:
                for i, sugestao in enumerate(st.session_state.sugestoes_geradas):
                    s_id = sugestao.get('sugestao_id', f"sug_llm_{i}_{hash(sugestao.get('titulo_sugerido', ''))}")
                    sugestao['sugestao_id'] = s_id # Garante que o ID está na sugestão original
                    if s_id not in st.session_state.sugestoes_validadas:
                        st.session_state.sugestoes_validadas[s_id] = {
                            "aceito": True, "tipo_grafico": sugestao.get('tipo_grafico_sugerido', 'desconhecido'),
                            "titulo": sugestao.get('titulo_sugerido', 'Título não fornecido'),
                            "fonte_dados_id": sugestao.get('fonte_dados_id', 'desconhecido'),
                            "parametros_grafico_completos": sugestao.get('parametros_grafico', {}),
                            "justificativa": sugestao.get('justificativa', 'N/A')
                        }

if st.session_state.sugestoes_geradas:
    st.sidebar.header("⚙️ Valide as Sugestões da IA")
    for sugestao_original in st.session_state.sugestoes_geradas:
        s_id = sugestao_original['sugestao_id']
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
               not params_grafico.get("dados_diretos"):
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

    if st.sidebar.button("Gerar Dashboard", type="primary", use_container_width=True):
        st.header("🚀 Dashboard Gerado")
        
        # Renderizar KPIs primeiro
        kpi_sugestoes = {s_id: s_conf for s_id, s_conf in st.session_state.sugestoes_validadas.items() if s_conf['aceito'] and s_conf['tipo_grafico'] == 'metrica'}
        if kpi_sugestoes:
            num_kpis = len(kpi_sugestoes)
            kpi_cols = st.columns(min(num_kpis, 4)) # Máximo de 4 KPIs por linha
            for i, (s_id_kpi, kpi_conf) in enumerate(kpi_sugestoes.items()):
                with kpi_cols[i % min(num_kpis, 4)]:
                    params_kpi = kpi_conf['parametros_grafico_completos']
                    st.metric(label=kpi_conf['titulo'], value=str(params_kpi.get('valor', 'N/A')), delta=str(params_kpi.get('delta', '')), help=params_kpi.get('descricao_curta_kpi'))
            st.divider()

        if show_debug_info:
            with st.expander("DEBUG: Parâmetros Validados para Plotagem (após validação da sidebar)", expanded=False):
                for s_id_debug, config_debug in st.session_state.sugestoes_validadas.items():
                    if config_debug['aceito'] and config_debug['tipo_grafico'] != 'metrica': # Não repetir KPIs
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
            if config_atual['aceito'] and config_atual['tipo_grafico'] != 'metrica': # KPIs já renderizados
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
                        
                        if df_para_plotar is not None:
                            if tipo_grafico == 'bar' and params_completos.get('eixo_x') and params_completos.get('eixo_y'):
                                st.plotly_chart(px.bar(df_para_plotar, x=params_completos['eixo_x'], y=params_completos['eixo_y'], title=titulo), use_container_width=True)
                                elemento_renderizado_nesta_iteracao = True
                            # Adicione aqui elif para line, scatter, pie, garantindo que df_para_plotar existe e as colunas estão nos params
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
                                st.dataframe(df_para_plotar.astype(str), use_container_width=True) # .astype(str) para exibição
                                elemento_renderizado_nesta_iteracao = True
                        
                        if tipo_grafico == 'diagrama_swot_lista': # Não depende de df_para_plotar
                            st.subheader(titulo)
                            c1, c2 = st.columns(2)
                            swot_items = {"forcas": "Forças 💪", "fraquezas": "Fraquezas 📉", "oportunidades": "Oportunidades 🚀", "ameacas": "Ameaças ⚠️"}
                            current_col_map = {0: c1, 1:c1, 2:c2, 3:c2}
                            for i_key_idx, (key, header) in enumerate(swot_items.items()):
                                with current_col_map[i_key_idx % 4]: # Garante que sempre pega uma coluna válida
                                    st.markdown(f"##### {header}")
                                    items = params_completos.get(key, ["N/A"])
                                    if not items or not isinstance(items, list): items = ["N/A (dados ausentes ou formato incorreto)"] 
                                    for item in items: st.markdown(f"- {item}")
                            elemento_renderizado_nesta_iteracao = True
                        elif tipo_grafico == 'mapa':
                             st.info(f"Visualização de mapa para '{titulo}' ainda não implementada ou dados geoespaciais não fornecidos pela IA.")
                             elemento_renderizado_nesta_iteracao = True # Considera renderizado como info
                        
                        if not elemento_renderizado_nesta_iteracao and config_atual['aceito']:
                             st.warning(f"Não foi possível gerar visualização para '{titulo}' (tipo: {tipo_grafico}).")
                    except Exception as e:
                        st.error(f"Erro ao gerar '{titulo}': {e}")
                        # st.text(traceback.format_exc()) # Descomentar para debug detalhado do erro de plotagem
                
                if elemento_renderizado_nesta_iteracao:
                    col_idx +=1
                    elementos_renderizados_count +=1
        
        if elementos_renderizados_count == 0 and any(c['aceito'] and c['tipo_grafico'] != 'metrica' for c in st.session_state.sugestoes_validadas.values()):
            st.info("Nenhum gráfico/tabela (além de KPIs) pôde ser gerado com as seleções atuais.")
        elif elementos_renderizados_count == 0 and not kpi_sugestoes: # Se não há nem KPIs
            st.info("Nenhum elemento foi selecionado ou pôde ser gerado para o dashboard.")

elif uploaded_file is None and st.session_state.arquivo_processado_nome is not None: 
    st.session_state.sugestoes_geradas = []
    st.session_state.dados_extraidos = {"textos_list": [], "tabelas_dfs": []}
    st.session_state.sugestoes_validadas = {}
    st.session_state.arquivo_processado_nome = None
    st.session_state.debug_info_shown = False
    if "file_uploader" in st.session_state: st.session_state.file_uploader = None
    st.experimental_rerun()

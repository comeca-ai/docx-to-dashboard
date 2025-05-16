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
        # Para deploy no Streamlit Community Cloud
        return st.secrets["GOOGLE_API_KEY"]
    except (FileNotFoundError, KeyError):
        # Para desenvolvimento local via variável de ambiente
        api_key = os.environ.get("GOOGLE_API_KEY")
        if api_key:
            return api_key
        st.error("Chave da API do Gemini (GOOGLE_API_KEY) não configurada. "
                 "Por favor, configure-a nos segredos do Streamlit ou como variável de ambiente.")
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
            nome_tabela_doc = f"Tabela_{i+1}" # Nome padrão

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
                        original_series = df[col].copy() # Salva original para possível reversão

                        # 1. Tentar converter para numérico
                        try:
                            # Remove espaços extras, substitui vírgula por ponto para decimais
                            cleaned_series = df[col].astype(str).str.strip().str.replace(',', '.', regex=False)
                            # Remove símbolos monetários comuns e % antes de tentar converter para numérico
                            cleaned_series = cleaned_series.str.replace(r'[R$\s%]', '', regex=True)
                            
                            df[col] = pd.to_numeric(cleaned_series, errors='raise')
                            # st.write(f"Coluna '{col}' (Tabela {nome_tabela_doc}) convertida para NUMÉRICO.") # DEBUG
                            continue 
                        except (ValueError, TypeError):
                            df[col] = original_series.copy() # Reverte se falhar

                        # 2. Tentar converter para datetime (se não for numérico)
                        # Ser mais seletivo: tentar converter para data apenas se o nome da coluna ou conteúdo sugerir
                        # Ou, se preferir, pode tentar em todas que não são numéricas.
                        # Aqui, vamos tentar de forma um pouco mais genérica, mas com formatos explícitos primeiro.
                        try:
                            possible_formats = [
                                '%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', 
                                '%d-%m-%Y', '%m-%d-%Y', '%Y', '%Y%m%d'
                            ] # Adicione mais formatos comuns se necessário
                            
                            temp_series = pd.Series(index=df.index, dtype='object')
                            all_na = True
                            
                            # Tenta converter com formatos explícitos
                            converted_with_format = False
                            for fmt in possible_formats:
                                try:
                                    temp_series_fmt = pd.to_datetime(df[col].astype(str), format=fmt, errors='coerce')
                                    # Se a maioria não for NaT com este formato, considera um bom candidato
                                    if temp_series_fmt.notna().sum() > temp_series_fmt.isna().sum() or temp_series_fmt.notna().sum() > len(df[col])*0.5 :
                                        df[col] = temp_series_fmt
                                        # st.write(f"Coluna '{col}' (Tabela {nome_tabela_doc}) convertida para DATETIME com formato {fmt}.") # DEBUG
                                        converted_with_format = True
                                        break 
                                except (ValueError, TypeError):
                                    continue
                            
                            if not converted_with_format:
                                # Se nenhum formato explícito funcionou bem, tenta a inferência do Pandas
                                # Isso pode gerar o UserWarning se o formato não for óbvio
                                inferred_series = pd.to_datetime(df[col].astype(str), errors='coerce', infer_datetime_format=True)
                                if inferred_series.notna().sum() > inferred_series.isna().sum() or inferred_series.notna().sum() > len(df[col])*0.5:
                                    df[col] = inferred_series
                                    # st.write(f"Coluna '{col}' (Tabela {nome_tabela_doc}) convertida para DATETIME por INFERÊNCIA.") # DEBUG
                                else:
                                    # Se a inferência resultou em muitos NaT, reverte para o original (string)
                                    df[col] = original_series.copy()
                                    # st.write(f"Inferência de data para '{col}' (Tabela {nome_tabela_doc}) resultou em muitos NaT. Mantida como string.") # DEBUG
                        
                        except Exception: # Qualquer outro erro na conversão de data
                            df[col] = original_series.copy() # Reverte para o original (string)
                            # st.write(f"Falha geral ao converter '{col}' (Tabela {nome_tabela_doc}) para datetime. Mantida como string.") # DEBUG
                    
                    # st.write(f"DataFrame final para {nome_tabela_doc} (ID: tabela_{i+1}):") # DEBUG
                    # st.dataframe(df.head()) # DEBUG
                    # st.write(df.dtypes) # DEBUG
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
        st.warning("Nenhuma sugestão da LLM pôde ser gerada pois a chave da API não está configurada.")
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

        tabelas_markdown_dict = {}
        tabelas_prompt_str = ""
        for tabela_info in tabelas_dfs_list:
            df = tabela_info["dataframe"]
            nome_tabela = tabela_info["nome"]
            id_tabela = tabela_info["id"]
            
            df_sample = df.head(10) # Amostra das primeiras 10 linhas
            if len(df.columns) > 15: # Limita o número de colunas na amostra
                df_sample = df_sample.iloc[:, :15]
            
            markdown_tabela = df_sample.to_markdown(index=False)
            tabelas_prompt_str += f"\n--- {nome_tabela} (ID para referência: {id_tabela}) ---\n"
            tabelas_prompt_str += f"Colunas disponíveis: {', '.join(df.columns.tolist())}\n" # Lista explícita de colunas
            tabelas_prompt_str += f"{markdown_tabela}\n"


        max_texto_len = 80000 
        texto_doc_para_prompt = texto_doc_completo[:max_texto_len]
        if len(texto_doc_completo) > max_texto_len:
            texto_doc_para_prompt += "\n[TEXTO DO DOCUMENTO TRUNCADO DEVIDO AO TAMANHO LIMITE PARA ESTE PROMPT]"

        prompt_completo = f"""
        Você é um assistente especialista em análise de dados e visualização.
        Analise o seguinte conteúdo extraído de um documento DOCX. O conteúdo inclui parágrafos de texto e representações de tabelas em formato Markdown. Para cada tabela, listei explicitamente as colunas disponíveis.

        [INÍCIO DO TEXTO DO DOCUMENTO]
        {texto_doc_para_prompt}
        [FIM DO TEXTO DO DOCUMENTO]

        [INÍCIO DAS TABELAS DO DOCUMENTO]
        {tabelas_prompt_str}
        [FIM DAS TABELAS DO DOCUMENTO]

        Com base no conteúdo fornecido:
        1. Identifique todas as possíveis análises e visualizações de dados que poderiam ser geradas.
        2. Considere dados explícitos em tabelas e informações implícitas ou descritas no texto (como análises SWOT com seus pontos, comparações de market share, tendências, etc.).
        3. Para cada análise/visualização sugerida, forneça as seguintes informações em formato JSON, como uma lista de objetos. Cada objeto deve ter as chaves:
            - "sugestao_id": Um identificador único para a sugestão (ex: "llm_sug_1").
            - "titulo_sugerido": Um título descritivo para o gráfico/análise (ex: "Análise SWOT do iFood").
            - "tipo_grafico_sugerido": O tipo de gráfico recomendado (ex: "pizza", "bar", "line", "scatter", "diagrama_swot_lista", "tabela_informativa"). Use "tabela_informativa" para tabelas que devem ser exibidas como estão. Use nomes de colunas EXATOS como listados para cada tabela.
            - "fonte_dados_id": O ID da tabela de origem (ex: "tabela_1", "tabela_2", conforme o ID fornecido) ou uma descrição da seção do texto (ex: "texto_swot_ifood").
            - "parametros_grafico": Um objeto com os parâmetros específicos. Use os nomes exatos das colunas da tabela de origem. Exemplos:
                - Para gráficos (pizza, bar, line, scatter): {{ "eixo_x": "NomeColunaX", "eixo_y": "NomeColunaY", "valores": "NomeColunaValores", "categorias": "NomeColunaCategorias" }}
                - Para "diagrama_swot_lista": {{ "forcas": ["Ponto 1"], "fraquezas": ["Ponto A"], "oportunidades": ["Ponto X"], "ameacas": ["Ponto Z"] }}
                - Para "tabela_informativa": {{ "id_tabela_original": "ID_da_Tabela_no_DOCX" }}
            - "justificativa": Uma breve explicação do que a visualização mostraria.

        Exemplo de SWOT:
        {{
          "sugestao_id": "llm_swot1", "titulo_sugerido": "Análise SWOT XYZ", "tipo_grafico_sugerido": "diagrama_swot_lista",
          "fonte_dados_id": "texto_secao_analise_swot_xyz",
          "parametros_grafico": {{ "forcas": ["F1", "F2"], "fraquezas": ["Fr1"], "oportunidades": ["O1"], "ameacas": ["A1"] }},
          "justificativa": "Entendimento estratégico da empresa XYZ."
        }}
        Exemplo de gráfico de pizza para uma tabela com ID 'tabela_2' e colunas 'Player' e 'Market Share Estimado (%)':
        {{
          "sugestao_id": "llm_pizza1", "titulo_sugerido": "Market Share", "tipo_grafico_sugerido": "pizza",
          "fonte_dados_id": "tabela_2",
          "parametros_grafico": {{ "categorias": "Player", "valores": "Market Share Estimado (%)" }},
          "justificativa": "Distribuição do market share."
        }}
        Retorne APENAS a lista JSON válida, nada mais. Certifique-se de que os nomes de colunas nos 'parametros_grafico' correspondam exatamente aos nomes de colunas fornecidos para cada tabela.
        """
        
        with st.spinner("🤖 Consultando o Gemini para obter sugestões avançadas... Isso pode levar um momento."):
            # st.text_area("Debug: Prompt enviado ao Gemini", prompt_completo, height=200) # Descomentar para depuração
            response = model.generate_content(prompt_completo)
            
        cleaned_response_text = response.text.strip()
        if cleaned_response_text.startswith("```json"):
            cleaned_response_text = cleaned_response_text[7:].strip()
        if cleaned_response_text.endswith("```"):
            cleaned_response_text = cleaned_response_text[:-3].strip()
        
        # st.text_area("Debug: Resposta do Gemini (após limpeza)", cleaned_response_text, height=200) # Descomentar para depuração

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

# Inicialização do estado da sessão
if 'sugestoes_geradas' not in st.session_state:
    st.session_state.sugestoes_geradas = []
if 'dados_extraidos' not in st.session_state:
    st.session_state.dados_extraidos = {"textos_list": [], "tabelas_dfs": []}
if 'sugestoes_validadas' not in st.session_state:
    st.session_state.sugestoes_validadas = {}
if 'arquivo_processado_nome' not in st.session_state:
    st.session_state.arquivo_processado_nome = None
if 'debug_dataframes_shown' not in st.session_state: # Novo estado para controlar exibição de debug
    st.session_state.debug_dataframes_shown = False


uploaded_file = st.file_uploader("Escolha um arquivo DOCX", type=["docx"], key="file_uploader")

# Checkbox para mostrar/ocultar informações de depuração
show_debug_info = st.sidebar.checkbox("Mostrar Informações de Depuração", value=False)


if uploaded_file is not None:
    if st.session_state.arquivo_processado_nome != uploaded_file.name:
        st.session_state.sugestoes_geradas = []
        st.session_state.dados_extraidos = {"textos_list": [], "tabelas_dfs": []}
        st.session_state.sugestoes_validadas = {}
        st.session_state.arquivo_processado_nome = uploaded_file.name
        st.session_state.debug_dataframes_shown = False # Reseta o debug para novo arquivo

    if not st.session_state.sugestoes_geradas: # Processar apenas se não houver sugestões para o arquivo atual
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
                        st.dataframe(t_info['dataframe'].head())
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
                    s_id = sugestao.get('sugestao_id', f"sug_{hash(sugestao.get('titulo_sugerido', ''))}") # ID fallback
                    if s_id not in st.session_state.sugestoes_validadas:
                        st.session_state.sugestoes_validadas[s_id] = {
                            "aceito": True,
                            "tipo_grafico": sugestao.get('tipo_grafico_sugerido', 'desconhecido'),
                            "titulo": sugestao.get('titulo_sugerido', 'Título não fornecido'),
                            "fonte_dados_id": sugestao.get('fonte_dados_id', 'desconhecido'),
                            "parametros_grafico_completos": sugestao.get('parametros_grafico', {}),
                            "justificativa": sugestao.get('justificativa', 'N/A')
                        }

# Exibir sugestões e permitir validação
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
            if config_atual['tipo_grafico'] in ['bar', 'line', 'pie', 'scatter'] and not str(config_atual['fonte_dados_id']).startswith("texto_"):
                df_correspondente = next((t['dataframe'] for t in st.session_state.dados_extraidos['tabelas_dfs'] if t['id'] == config_atual['fonte_dados_id']), None)
                if df_correspondente is not None:
                    opcoes_colunas = [""] + df_correspondente.columns.tolist()
                    
                    if config_atual['tipo_grafico'] in ['bar', 'line', 'scatter']:
                        x_atual = params_grafico.get('eixo_x', "")
                        y_atual = params_grafico.get('eixo_y', "")
                        params_grafico['eixo_x'] = st.selectbox("Eixo X", options=opcoes_colunas, index=opcoes_colunas.index(x_atual) if x_atual in opcoes_colunas else 0, key=f"x_col_{s_id}")
                        params_grafico['eixo_y'] = st.selectbox("Eixo Y", options=opcoes_colunas, index=opcoes_colunas.index(y_atual) if y_atual in opcoes_colunas else 0, key=f"y_col_{s_id}")
                    elif config_atual['tipo_grafico'] == 'pie':
                        cat_atual = params_grafico.get('categorias', "")
                        val_atual = params_grafico.get('valores', "")
                        params_grafico['categorias'] = st.selectbox("Categorias", options=opcoes_colunas, index=opcoes_colunas.index(cat_atual) if cat_atual in opcoes_colunas else 0, key=f"names_col_{s_id}")
                        params_grafico['valores'] = st.selectbox("Valores", options=opcoes_colunas, index=opcoes_colunas.index(val_atual) if val_atual in opcoes_colunas else 0, key=f"values_col_{s_id}")
            st.session_state.sugestoes_validadas[s_id] = config_atual

    if st.sidebar.button("Gerar Dashboard com Gráficos Selecionados", type="primary", use_container_width=True):
        st.header("🚀 Dashboard Gerado")
        
        if show_debug_info:
            with st.expander("DEBUG: Parâmetros Validados das Sugestões para Plotagem", expanded=False):
                for s_id_debug, config_debug in st.session_state.sugestoes_validadas.items():
                    if config_debug['aceito']:
                        st.write(f"ID Sugestão: {s_id_debug}, Título: {config_debug['titulo']}, Tipo: {config_debug['tipo_grafico']}")
                        st.write(f"Fonte Dados ID: {config_debug['fonte_dados_id']}")
                        st.write(f"Parâmetros Completos: {config_debug['parametros_grafico_completos']}")
                        # Tenta mostrar o DF se for de tabela
                        if not str(config_debug['fonte_dados_id']).startswith("texto_"):
                            df_debug_plot = next((t['dataframe'] for t in st.session_state.dados_extraidos['tabelas_dfs'] if t['id'] == config_debug['fonte_dados_id']), None)
                            if df_debug_plot is not None:
                                st.write(f"DataFrame para '{config_debug['fonte_dados_id']}':")
                                st.dataframe(df_debug_plot.head())
                                st.write(df_debug_plot.dtypes)
                        st.divider()

        elementos_renderizados_count = 0
        
        # Tenta organizar em colunas
        num_cols_dashboard = 2 
        cols_dashboard = st.columns(num_cols_dashboard)
        col_idx = 0

        for s_id, config_atual in st.session_state.sugestoes_validadas.items():
            if config_atual['aceito']:
                tipo_grafico = config_atual['tipo_grafico']
                titulo = config_atual['titulo']
                params_completos = config_atual['parametros_grafico_completos']
                fonte_id = config_atual['fonte_dados_id']
                df_grafico = None
                elemento_renderizado_nesta_iteracao = False

                with cols_dashboard[col_idx % num_cols_dashboard]:
                    try:
                        if not str(fonte_id).startswith("texto_"):
                             df_grafico = next((t['dataframe'] for t in st.session_state.dados_extraidos['tabelas_dfs'] if t['id'] == fonte_id), None)
                             if df_grafico is None:
                                 st.warning(f"Gráfico '{titulo}': Dados da tabela '{fonte_id}' não encontrados.")
                                 continue
                        
                        if tipo_grafico == 'bar' and params_completos.get('eixo_x') and params_completos.get('eixo_y') and df_grafico is not None:
                            st.plotly_chart(px.bar(df_grafico, x=params_completos['eixo_x'], y=params_completos['eixo_y'], title=titulo), use_container_width=True)
                            elemento_renderizado_nesta_iteracao = True
                        elif tipo_grafico == 'line' and params_completos.get('eixo_x') and params_completos.get('eixo_y') and df_grafico is not None:
                            st.plotly_chart(px.line(df_grafico, x=params_completos['eixo_x'], y=params_completos['eixo_y'], title=titulo, markers=True), use_container_width=True)
                            elemento_renderizado_nesta_iteracao = True
                        elif tipo_grafico == 'scatter' and params_completos.get('eixo_x') and params_completos.get('eixo_y') and df_grafico is not None:
                            st.plotly_chart(px.scatter(df_grafico, x=params_completos['eixo_x'], y=params_completos['eixo_y'], title=titulo), use_container_width=True)
                            elemento_renderizado_nesta_iteracao = True
                        elif tipo_grafico == 'pie' and params_completos.get('categorias') and params_completos.get('valores') and df_grafico is not None:
                            st.plotly_chart(px.pie(df_grafico, names=params_completos['categorias'], values=params_completos['valores'], title=titulo), use_container_width=True)
                            elemento_renderizado_nesta_iteracao = True
                        elif tipo_grafico == 'diagrama_swot_lista':
                            st.subheader(titulo)
                            c1, c2 = st.columns(2)
                            swot_items = {"forcas": "Forças 💪", "fraquezas": "Fraquezas 📉", "oportunidades": "Oportunidades 🚀", "ameacas": "Ameaças ⚠️"}
                            current_col = c1
                            for i_key, (key, header) in enumerate(swot_items.items()):
                                if i_key >= 2 : current_col = c2 # Muda para segunda coluna do SWOT
                                with current_col:
                                    st.markdown(f"##### {header}")
                                    items = params_completos.get(key, ["N/A"])
                                    if not items: items = ["N/A"] # Garante que não é lista vazia
                                    for item in items: st.markdown(f"- {item}")
                            elemento_renderizado_nesta_iteracao = True
                        elif tipo_grafico == 'tabela_informativa':
                            id_tabela_original = params_completos.get('id_tabela_original', fonte_id)
                            df_tabela_info = next((t['dataframe'] for t in st.session_state.dados_extraidos['tabelas_dfs'] if t['id'] == id_tabela_original), df_grafico) # Usa df_grafico como fallback
                            if df_tabela_info is not None:
                                st.subheader(titulo)
                                st.dataframe(df_tabela_info, use_container_width=True)
                                elemento_renderizado_nesta_iteracao = True
                            else:
                                st.warning(f"Tabela para '{titulo}' (ID: {id_tabela_original}) não encontrada.")
                        
                        if not elemento_renderizado_nesta_iteracao and config_atual['aceito']:
                             st.warning(f"Não foi possível gerar visualização para '{titulo}' (tipo: {tipo_grafico}). Verifique os parâmetros e a lógica de plotagem.")

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

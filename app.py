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
    """Extrai textos e tabelas de um arquivo DOCX."""
    try:
        document = Document(uploaded_file)
        textos = [p.text for p in document.paragraphs if p.text.strip()]
        tabelas_dfs = []
        for i, table in enumerate(document.tables):
            data = []
            keys = None
            nome_tabela_doc = f"Tabela_{i+1}" # Nome padrão

            # Tenta pegar um título para a tabela se houver um parágrafo anterior
            # Isso é uma heurística e pode não funcionar sempre.
            # try:
            #     if table._element.getprevious() is not None and table._element.getprevious().tag.endswith('p'):
            #         paragrafo_anterior = table._element.getprevious()
            #         texto_paragrafo_anterior = "".join([run.text for run in paragrafo_anterior.xpath('.//w:t')])
            #         if texto_paragrafo_anterior.lower().startswith("tabela"):
            #             nome_tabela_doc = texto_paragrafo_anterior.strip()
            # except Exception:
            #     pass


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
                        try:
                            df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', '.', regex=False))
                        except ValueError:
                            try:
                                df[col] = pd.to_datetime(df[col], errors='coerce')
                            except ValueError:
                                pass
                    tabelas_dfs.append({"id": f"tabela_{i+1}", "dataframe": df, "nome": nome_tabela_doc})
                except Exception as e:
                    st.warning(f"Não foi possível processar completamente a tabela {i+1} ({nome_tabela_doc}): {e}")
        return textos, tabelas_dfs
    except Exception as e:
        st.error(f"Erro ao ler o arquivo DOCX: {e}")
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
             safety_settings=[ # Permite mais conteúdo, ajuste se necessário
                {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
                {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
                {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
                {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"},
            ]
        )

        tabelas_markdown_dict = {}
        for tabela_info in tabelas_dfs_list:
            df = tabela_info["dataframe"]
            nome_tabela = tabela_info["nome"]
            # Limitar o número de linhas e colunas para não exceder o prompt
            df_sample = df.head(10)
            if len(df.columns) > 10:
                df_sample = df_sample.iloc[:, :10]
            tabelas_markdown_dict[nome_tabela] = df_sample.to_markdown(index=False)

        tabelas_prompt_str = ""
        for nome_tabela, markdown_tabela in tabelas_markdown_dict.items():
            tabelas_prompt_str += f"\n--- {nome_tabela} (ID para referência: {tabela_info['id']}) ---\n{markdown_tabela}\n"

        max_texto_len = 80000 # Ajuste conforme necessário e limites do modelo (gemini-1.5-flash tem contexto grande)
        texto_doc_para_prompt = texto_doc_completo[:max_texto_len]
        if len(texto_doc_completo) > max_texto_len:
            texto_doc_para_prompt += "\n[TEXTO DO DOCUMENTO TRUNCADO DEVIDO AO TAMANHO LIMITE PARA ESTE PROMPT]"

        prompt_completo = f"""
        Você é um assistente especialista em análise de dados e visualização.
        Analise o seguinte conteúdo extraído de um documento DOCX. O conteúdo inclui parágrafos de texto e representações de tabelas em formato Markdown.

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
            - "tipo_grafico_sugerido": O tipo de gráfico recomendado (ex: "pizza", "bar", "line", "scatter", "diagrama_swot_lista", "tabela_informativa"). Use "tabela_informativa" para tabelas que devem ser exibidas como estão.
            - "fonte_dados_id": O ID da tabela de origem (ex: "tabela_1", "tabela_2") ou uma descrição da seção do texto (ex: "texto_swot_ifood").
            - "parametros_grafico": Um objeto com os parâmetros específicos. Exemplos:
                - Para gráficos (pizza, bar, line, scatter): {{ "eixo_x": "NomeColunaX", "eixo_y": "NomeColunaY", "valores": "NomeColunaValores", "categorias": "NomeColunaCategorias" }}
                - Para "diagrama_swot_lista": {{ "forcas": ["Ponto 1"], "fraquezas": ["Ponto A"], "oportunidades": ["Ponto X"], "ameacas": ["Ponto Z"] }}
                - Para "tabela_informativa": {{ "nome_tabela_original_no_doc": "Nome da Tabela como aparece no texto ou seu ID" }}
            - "justificativa": Uma breve explicação do que a visualização mostraria.

        Exemplo de SWOT:
        {{
          "sugestao_id": "llm_swot1", "titulo_sugerido": "Análise SWOT XYZ", "tipo_grafico_sugerido": "diagrama_swot_lista",
          "fonte_dados_id": "texto_secao_analise_swot_xyz",
          "parametros_grafico": {{ "forcas": ["F1", "F2"], "fraquezas": ["Fr1"], "oportunidades": ["O1"], "ameacas": ["A1"] }},
          "justificativa": "Entendimento estratégico da empresa XYZ."
        }}
        Exemplo de gráfico de pizza para uma tabela com ID 'tabela_2':
        {{
          "sugestao_id": "llm_pizza1", "titulo_sugerido": "Market Share", "tipo_grafico_sugerido": "pizza",
          "fonte_dados_id": "tabela_2",
          "parametros_grafico": {{ "categorias": "Player", "valores": "Market Share Estimado (%)" }},
          "justificativa": "Distribuição do market share."
        }}
        Retorne APENAS a lista JSON válida, nada mais.
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
    st.session_state.dados_extraidos = {"textos_list": [], "tabelas_dfs": []} # Renomeado para clareza
if 'sugestoes_validadas' not in st.session_state:
    st.session_state.sugestoes_validadas = {}
if 'arquivo_processado_nome' not in st.session_state:
    st.session_state.arquivo_processado_nome = None

uploaded_file = st.file_uploader("Escolha um arquivo DOCX", type=["docx"], key="file_uploader")

if uploaded_file is not None:
    if st.session_state.arquivo_processado_nome != uploaded_file.name:
        st.session_state.sugestoes_geradas = []
        st.session_state.dados_extraidos = {"textos_list": [], "tabelas_dfs": []}
        st.session_state.sugestoes_validadas = {}
        st.session_state.arquivo_processado_nome = uploaded_file.name

    if not st.session_state.sugestoes_geradas:
        with st.spinner("Lendo e pré-processando o documento..."):
            textos_list, tabelas_dfs_list = extrair_dados_docx(uploaded_file)
            st.session_state.dados_extraidos = {"textos_list": textos_list, "tabelas_dfs": tabelas_dfs_list}
        
        if not tabelas_dfs_list and not textos_list:
            st.warning("Nenhum dado extraível (texto ou tabela) encontrado no documento.")
        else:
            st.success(f"Documento '{uploaded_file.name}' lido!")
            texto_completo_doc = "\n\n".join(textos_list) # Juntar parágrafos
            
            sugestoes_da_llm = obter_sugestoes_da_llm(texto_completo_doc, tabelas_dfs_list)
            st.session_state.sugestoes_geradas = sugestoes_da_llm

            if not st.session_state.sugestoes_geradas:
                st.info("Nenhuma sugestão foi gerada pela IA para este documento.")
            else:
                for sugestao in st.session_state.sugestoes_geradas:
                    s_id = sugestao['sugestao_id']
                    if s_id not in st.session_state.sugestoes_validadas:
                        st.session_state.sugestoes_validadas[s_id] = {
                            "aceito": True,
                            "tipo_grafico": sugestao['tipo_grafico_sugerido'],
                            "titulo": sugestao['titulo_sugerido'],
                            "fonte_dados_id": sugestao['fonte_dados_id'],
                            "parametros_grafico_completos": sugestao['parametros_grafico'],
                            "justificativa": sugestao.get('justificativa', '')
                        }

# Exibir sugestões e permitir validação
if st.session_state.sugestoes_geradas:
    st.sidebar.header("⚙️ Valide as Sugestões da IA")
    
    for sugestao_original in st.session_state.sugestoes_geradas:
        s_id = sugestao_original['sugestao_id']
        if s_id not in st.session_state.sugestoes_validadas: continue # Segurança
            
        config_atual = st.session_state.sugestoes_validadas[s_id]

        with st.sidebar.expander(f"Sugestão: {config_atual['titulo']}", expanded=False):
            st.caption(f"Tipo: {config_atual['tipo_grafico']} | Fonte: {config_atual['fonte_dados_id']}")
            st.markdown(f"**Justificativa da IA:** *{config_atual.get('justificativa', 'N/A')}*")

            config_atual['aceito'] = st.checkbox("Incluir no dashboard?", value=config_atual['aceito'], key=f"aceito_{s_id}")
            config_atual['titulo'] = st.text_input("Título", value=config_atual['titulo'], key=f"titulo_{s_id}")
            
            # Permitir edição de parâmetros para gráficos comuns se vierem de tabelas
            if config_atual['tipo_grafico'] in ['bar', 'line', 'pie', 'scatter'] and not config_atual['fonte_dados_id'].startswith("texto_"):
                df_correspondente = next((t['dataframe'] for t in st.session_state.dados_extraidos['tabelas_dfs'] if t['id'] == config_atual['fonte_dados_id']), None)
                if df_correspondente is not None:
                    opcoes_colunas = [""] + df_correspondente.columns.tolist() # Adicionar opção vazia
                    params_grafico = config_atual['parametros_grafico_completos']

                    if config_atual['tipo_grafico'] in ['bar', 'line', 'scatter']:
                        x_atual = params_grafico.get('eixo_x', "")
                        y_atual = params_grafico.get('eixo_y', "")
                        params_grafico['eixo_x'] = st.selectbox("Eixo X", options=opcoes_colunas, index=opcoes_colunas.index(x_atual) if x_atual in opcoes_colunas else 0, key=f"x_col_{s_id}")
                        params_grafico['eixo_y'] = st.selectbox("Eixo Y", options=opcoes_colunas, index=opcoes_colunas.index(y_atual) if y_atual in opcoes_colunas else 0, key=f"y_col_{s_id}")
                    elif config_atual['tipo_grafico'] == 'pie':
                        cat_atual = params_grafico.get('categorias', "")
                        val_atual = params_grafico.get('valores', "")
                        params_grafico['categorias'] = st.selectbox("Categorias (Nomes)", options=opcoes_colunas, index=opcoes_colunas.index(cat_atual) if cat_atual in opcoes_colunas else 0, key=f"names_col_{s_id}")
                        params_grafico['valores'] = st.selectbox("Valores", options=opcoes_colunas, index=opcoes_colunas.index(val_atual) if val_atual in opcoes_colunas else 0, key=f"values_col_{s_id}")
            st.session_state.sugestoes_validadas[s_id] = config_atual


    if st.sidebar.button("Gerar Dashboard com Gráficos Selecionados", type="primary", use_container_width=True):
        st.header("🚀 Dashboard Gerado")
        elementos_dashboard = [] # Pode conter figuras Plotly ou dicts para renderização customizada

        for s_id, config_atual in st.session_state.sugestoes_validadas.items():
            if config_atual['aceito']:
                tipo_grafico = config_atual['tipo_grafico']
                titulo = config_atual['titulo']
                params_completos = config_atual['parametros_grafico_completos']
                fonte_id = config_atual['fonte_dados_id']
                df_grafico = None

                if not fonte_id.startswith("texto_"): # Tenta carregar DataFrame se a fonte for uma tabela
                     df_grafico = next((t['dataframe'] for t in st.session_state.dados_extraidos['tabelas_dfs'] if t['id'] == fonte_id), None)
                     if df_grafico is None:
                         st.warning(f"Não foi possível encontrar dados da tabela '{fonte_id}' para o gráfico '{titulo}'.")
                         continue
                
                try:
                    if tipo_grafico == 'bar' and params_completos.get('eixo_x') and params_completos.get('eixo_y') and df_grafico is not None:
                        elementos_dashboard.append(px.bar(df_grafico, x=params_completos['eixo_x'], y=params_completos['eixo_y'], title=titulo))
                    elif tipo_grafico == 'line' and params_completos.get('eixo_x') and params_completos.get('eixo_y') and df_grafico is not None:
                        elementos_dashboard.append(px.line(df_grafico, x=params_completos['eixo_x'], y=params_completos['eixo_y'], title=titulo, markers=True))
                    elif tipo_grafico == 'scatter' and params_completos.get('eixo_x') and params_completos.get('eixo_y') and df_grafico is not None:
                        elementos_dashboard.append(px.scatter(df_grafico, x=params_completos['eixo_x'], y=params_completos['eixo_y'], title=titulo))
                    elif tipo_grafico == 'pie' and params_completos.get('categorias') and params_completos.get('valores') and df_grafico is not None:
                        elementos_dashboard.append(px.pie(df_grafico, names=params_completos['categorias'], values=params_completos['valores'], title=titulo))
                    
                    elif tipo_grafico == 'diagrama_swot_lista':
                        elementos_dashboard.append({"tipo": "swot", "titulo": titulo, "params": params_completos})
                    
                    elif tipo_grafico == 'tabela_informativa' and df_grafico is not None:
                         elementos_dashboard.append({"tipo": "tabela", "titulo": titulo, "dataframe": df_grafico})
                    
                    # Adicionar mais tipos de gráficos aqui se a LLM sugerir
                    
                    elif config_atual['aceito']: # Se estava aceito mas não gerou fig
                         st.warning(f"Não foi possível gerar o gráfico '{titulo}' (tipo: {tipo_grafico}). Verifique os parâmetros e a lógica de plotagem.")

                except Exception as e:
                    st.error(f"Erro ao gerar elemento '{titulo}': {e}")
        
        if elementos_dashboard:
            num_cols_dashboard = 2 
            cols_dashboard = st.columns(num_cols_dashboard) 
            col_idx = 0
            for i, elemento in enumerate(elementos_dashboard):
                with cols_dashboard[col_idx % num_cols_dashboard]:
                    if hasattr(elemento, 'update_layout'): # Se for uma figura Plotly
                        st.plotly_chart(elemento, use_container_width=True)
                        col_idx += 1
                    elif isinstance(elemento, dict):
                        if elemento["tipo"] == "swot":
                            st.subheader(elemento["titulo"])
                            params = elemento["params"]
                            c1, c2 = st.columns(2)
                            with c1:
                                st.markdown("##### Forças 💪")
                                for item in params.get('forcas', ["N/A"]): st.markdown(f"- {item}")
                                st.markdown("##### Fraquezas 📉")
                                for item in params.get('fraquezas', ["N/A"]): st.markdown(f"- {item}")
                            with c2:
                                st.markdown("##### Oportunidades 🚀")
                                for item in params.get('oportunidades', ["N/A"]): st.markdown(f"- {item}")
                                st.markdown("##### Ameaças ⚠️")
                                for item in params.get('ameacas', ["N/A"]): st.markdown(f"- {item}")
                            col_idx += 1 # Considera como um elemento renderizado
                        elif elemento["tipo"] == "tabela":
                            st.subheader(elemento["titulo"])
                            st.dataframe(elemento["dataframe"], use_container_width=True)
                            col_idx += 1
        else:
            st.info("Nenhum elemento foi selecionado ou pôde ser gerado para o dashboard.")


elif uploaded_file is None and st.session_state.arquivo_processado_nome is not None: 
    st.session_state.sugestoes_geradas = []
    st.session_state.dados_extraidos = {"textos_list": [], "tabelas_dfs": []}
    st.session_state.sugestoes_validadas = {}
    st.session_state.arquivo_processado_nome = None
    if "file_uploader" in st.session_state: # Limpa o uploader se ele ainda tiver um arquivo carregado na sessão
        st.session_state.file_uploader = None
    st.experimental_rerun()

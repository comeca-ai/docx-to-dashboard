import streamlit as st
from docx import Document
import pandas as pd
import plotly.express as px
import google.generativeai as genai
import json
import os
import traceback
import re 

# --- 1. Configuração da Chave da API do Gemini ---
def get_gemini_api_key():
    try:
        return st.secrets["GOOGLE_API_KEY"]
    except (FileNotFoundError, KeyError):
        api_key = os.environ.get("GOOGLE_API_KEY")
        if api_key:
            return api_key
        return None

# --- 2. Funções de Processamento do Documento e Interação com Gemini ---
def parse_value_range_or_text(val_str):
    """Tenta extrair um número de uma string, lidando com intervalos e texto."""
    if pd.isna(val_str) or val_str == '':
        return None
    
    # Tenta extrair o primeiro número flutuante ou inteiro da string
    # Remove separadores de milhar (ponto) antes de trocar vírgula decimal
    cleaned_val_str = str(val_str).replace('.', '', val_str.count('.') -1 if val_str.count('.') > 1 and ',' in val_str else val_str.count('.'))
    cleaned_val_str = cleaned_val_str.replace(',', '.')
    
    numbers = re.findall(r"[-+]?\d*\.\d+|\d+", cleaned_val_str) # Encontra números decimais ou inteiros
    
    if numbers:
        try:
            # Se for um intervalo como "70 - 86", pega o primeiro
            return float(numbers[0])
        except ValueError:
            return None # Não conseguiu converter para float
    return None # Se nenhum número for encontrado, ou se for texto não conversível


def clean_and_convert_to_numeric(series_data):
    """Tenta limpar e converter uma série Pandas para numérico de forma mais robusta."""
    if not isinstance(series_data, pd.Series):
        s = pd.Series(series_data)
    else:
        s = series_data.copy()

    # Aplica a função de parsing em cada elemento
    # Converte para string primeiro para garantir que .apply funcione em todos os elementos
    parsed_series = s.astype(str).apply(parse_value_range_or_text)
    
    # Tenta converter toda a série resultante para numérico
    # Se a maioria dos valores puder ser convertida, usa a série convertida
    numeric_col = pd.to_numeric(parsed_series, errors='coerce')
    if numeric_col.notna().sum() > s.notna().sum() * 0.3: # Se pelo menos 30% dos não-nulos originais viraram números
        return numeric_col
    
    # Fallback: tenta converter a série original de forma mais direta se o parsing falhou muito
    return pd.to_numeric(s, errors='coerce')


def extrair_conteudo_docx(uploaded_file):
    """Extrai texto e tabelas de um arquivo DOCX, com tratamento de tipos aprimorado."""
    try:
        document = Document(uploaded_file)
        textos = [p.text for p in document.paragraphs if p.text.strip()]
        tabelas_data = [] 

        for i, table_obj in enumerate(document.tables):
            data_rows = []
            keys = None
            nome_tabela = f"Tabela Documento {i+1}"
            try:
                prev_el = table_obj._element.getprevious()
                if prev_el is not None and prev_el.tag.endswith('p'):
                    p_text = "".join(node.text for node in prev_el.xpath('.//w:t')).strip()
                    if p_text and len(p_text) < 100 : nome_tabela = p_text.replace(":", "").strip()[:80] # Limita nome
            except: pass

            for row_idx, row in enumerate(table_obj.rows):
                text_cells = [cell.text.strip() for cell in row.cells]
                if row_idx == 0:
                    keys = [key.replace("\n", " ").strip() if key else f"Coluna_{k_idx+1}" for k_idx, key in enumerate(text_cells)]
                    continue
                if keys:
                    row_data = {}
                    for k_idx, key_name in enumerate(keys):
                         row_data[key_name] = text_cells[k_idx] if k_idx < len(text_cells) else None
                    data_rows.append(row_data)
            
            if data_rows:
                df = pd.DataFrame(data_rows)
                for col in df.columns:
                    original_col_data = df[col].copy()
                    
                    # 1. Tentar converter para numérico
                    converted_numeric = clean_and_convert_to_numeric(df[col])
                    if converted_numeric.notna().sum() >= len(df[col]) * 0.3:  # Critério mais flexível
                        df[col] = converted_numeric
                        continue
                    else: 
                         df[col] = original_col_data.copy() # Reverte se a conversão numérica não foi boa

                    # 2. Tentar converter para datetime (se não numérico)
                    try:
                        temp_col_str = df[col].astype(str) # Trabalha com strings para conversão de data
                        possible_formats = ['%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', 
                                            '%d-%m-%Y', '%m-%d-%Y', '%Y', '%Y%m%d', '%d.%m.%Y', '%Y-%m', '%m-%Y']
                        converted_with_format = False
                        for fmt in possible_formats:
                            try:
                                dt_series = pd.to_datetime(temp_col_str, format=fmt, errors='coerce')
                                if dt_series.notna().sum() > len(df[col]) * 0.5: # Se mais da metade for convertida
                                    df[col] = dt_series
                                    converted_with_format = True
                                    break
                            except (ValueError, TypeError): continue
                        
                        if not converted_with_format:
                            inferred_dt_series = pd.to_datetime(temp_col_str, errors='coerce') 
                            if inferred_dt_series.notna().sum() > len(df[col]) * 0.5:
                                 df[col] = inferred_dt_series
                            # else: # Se inferência falhou muito, mantém como estava (string do original)
                            #    df[col] = original_col_data.astype(str).fillna('') # Já é string do original
                    except Exception:
                        df[col] = original_col_data.astype(str).fillna('') # Garante que é string no erro
                
                # Fallback final para garantir que colunas 'object' sejam string para evitar erro do Arrow
                for col in df.columns:
                    if df[col].dtype == 'object':
                        df[col] = df[col].astype(str).fillna('')

                tabelas_data.append({"id": f"doc_tabela_{i+1}", "nome": nome_tabela, "dataframe": df})
        return "\n\n".join(textos), tabelas_data
    except Exception as e:
        st.error(f"Erro crítico ao ler DOCX: {e}"); return "", []

def analisar_documento_com_gemini(texto_doc, tabelas_info_list):
    api_key = get_gemini_api_key()
    if not api_key:
        st.warning("Chave da API do Gemini não configurada. Não é possível gerar sugestões da IA."); return []
    try:
        genai.configure(api_key=api_key)
        safety_settings_config = [
            {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"},
        ]
        model = genai.GenerativeModel(model_name="gemini-1.5-flash-latest", safety_settings=safety_settings_config)
        
        tabelas_prompt_str = ""
        for t_info in tabelas_info_list:
            df = t_info["dataframe"]
            df_sample = df.head(5) # Amostra menor para o prompt
            if len(df.columns) > 8: df_sample = df_sample.iloc[:, :8] # Limita colunas na amostra
            markdown_tabela = df_sample.to_markdown(index=False)
            tabelas_prompt_str += f"\n--- Tabela '{t_info['nome']}' (ID para referência: {t_info['id']}) ---\n"
            col_types_str = ", ".join([f"'{col}' (tipo inferido: {str(dtype)})" for col, dtype in df.dtypes.items()])
            tabelas_prompt_str += f"Colunas e tipos: {col_types_str}\nAmostra de dados:\n{markdown_tabela}\n"

        max_texto_len = 50000 # Reduzido para testar se é problema de tamanho de prompt
        texto_doc_para_prompt = texto_doc[:max_texto_len] + ("\n[TEXTO TRUNCADO...]" if len(texto_doc) > max_texto_len else "")

        prompt = f"""
        Você é um assistente de análise de dados e visualização. Analise o texto e as tabelas de um documento.
        [TEXTO DO DOCUMENTO]
        {texto_doc_para_prompt}
        [FIM DO TEXTO]
        [TABELAS DO DOCUMENTO (com ID, nome, colunas, tipos e amostra de dados)]
        {tabelas_prompt_str}
        [FIM DAS TABELAS]

        Gere uma lista JSON de sugestões de visualizações. CADA objeto na lista DEVE ter:
        - "id": string (ex: "gemini_sug_1").
        - "titulo": string (título para a visualização).
        - "tipo_sugerido": string ("kpi", "tabela_dados", "lista_swot", "grafico_barras", "grafico_pizza", "grafico_linha", "grafico_dispersao").
        - "fonte_id": string (ID da tabela ex: "doc_tabela_1", ou descrição textual da fonte ex: "texto_sumario_executivo").
        - "parametros": objeto JSON com dados e configurações. É CRUCIAL usar NOMES EXATOS de colunas das tabelas como fornecidos.
            - Para "kpi": {{"valor": "ValorDoKPI", "delta": "Mudança (opcional)", "descricao": "Contexto do KPI"}}
            - Para "tabela_dados": {{"id_tabela_original": "ID_da_Tabela_Referenciada"}}
            - Para "lista_swot": {{"forcas": ["Ponto Força 1"], "fraquezas": ["Ponto Fraqueza 1"], "oportunidades": ["Ponto Oportunidade 1"], "ameacas": ["Ponto Ameaça 1"]}}
            - Para "grafico_barras", "grafico_linha", "grafico_dispersao":
                Se baseado em TABELA (use o "fonte_id" da tabela): {{"eixo_x": "NOME_EXATO_COLUNA_X", "eixo_y": "NOME_EXATO_COLUNA_Y"}} (eixo_y geralmente numérico).
                Se DADOS EXTRAÍDOS DO TEXTO: {{"dados": [{{"NomeQualquerParaEixoX": "CategoriaA", "NomeQualquerParaEixoY": ValorNumericoA}}, ...], "eixo_x": "NomeQualquerParaEixoX", "eixo_y": "NomeQualquerParaEixoY"}} (Certifique-se que ValorNumericoA é de fato um número).
            - Para "grafico_pizza":
                Se baseado em TABELA: {{"categorias": "NOME_EXATO_COLUNA_CATEGORIAS", "valores": "NOME_EXATO_COLUNA_VALORES_NUMERICOS"}}
                Se DADOS EXTRAÍDOS DO TEXTO: {{"dados": [{{"NomeQualquerParaCategoria": "CategoriaA", "NomeQualquerParaValor": ValorNumericoA}}, ...], "categorias": "NomeQualquerParaCategoria", "valores": "NomeQualquerParaValor"}} (Certifique-se que ValorNumericoA é de fato um número).
        - "justificativa": string (breve explicação da utilidade).

        INSTRUÇÕES IMPORTANTES:
        1.  Para gráficos (barras, pizza, linha, dispersão) baseados em TABELAS, use os nomes EXATOS das colunas fornecidos. Verifique se a coluna de VALOR (eixo_y, valores) é de fato numérica (float64, int64) conforme os "tipos inferidos" da tabela. Se não for numérica, NÃO sugira um gráfico que exija valor numérico para essa coluna, A MENOS que você possa extrair um valor numérico dela (ex: de '70%' extrair 70.0).
        2.  Para Market Share que pode ter valores como '70% - 86%', extraia o primeiro número como float (ex: 70.0) para o campo de valores do gráfico de pizza.
        3.  Para Cobertura Geográfica, se tiver apenas Player e Cidades, sugira "tabela_dados". Não sugira "mapa" a menos que haja dados de coordenadas explícitos.
        4.  Para SWOTs comparativos de uma tabela, gere sugestões "lista_swot" INDIVIDUAIS para CADA player da tabela, não um SWOT agregado.
        5.  Se extrair dados do texto para gráficos, garanta que os valores numéricos SEJAM NÚMEROS no JSON (não strings de números).
        Retorne APENAS a lista JSON válida.
        """
        with st.spinner("🤖 Gemini está analisando o documento..."):
            # st.text_area("Debug: Prompt Gemini", prompt, height=200)
            response = model.generate_content(prompt)
        cleaned_text = response.text.strip().lstrip("```json").rstrip("```").strip()
        # st.text_area("Debug: Resposta Gemini", cleaned_text, height=200)
        sugestoes = json.loads(cleaned_text)
        if isinstance(sugestoes, list) and all(isinstance(item, dict) for item in sugestoes):
             st.success(f"{len(sugestoes)} sugestões recebidas do Gemini!")
             return sugestoes
        st.error("Resposta do Gemini não é lista JSON esperada."); return []
    except json.JSONDecodeError as e:
        st.error(f"Erro ao decodificar JSON do Gemini: {e}")
        if 'response' in locals() and hasattr(response, 'text'): st.code(response.text, language="text")
        return []
    except Exception as e: st.error(f"Erro na comunicação com Gemini: {e}"); st.text(traceback.format_exc()); return []

# --- 3. Interface Streamlit e Lógica de Apresentação ---
st.set_page_config(layout="wide")
st.title("✨ Apps com Gemini: DOCX para Insights Visuais")
st.markdown("Faça upload de um DOCX e deixe o Gemini sugerir como visualizar suas informações.")

if "sugestoes_gemini" not in st.session_state: st.session_state.sugestoes_gemini = []
if "conteudo_docx" not in st.session_state: st.session_state.conteudo_docx = {"texto": "", "tabelas": []}
if "config_sugestoes" not in st.session_state: st.session_state.config_sugestoes = {}
if "nome_arquivo_atual" not in st.session_state: st.session_state.nome_arquivo_atual = None
if 'debug_checkbox_key_main' not in st.session_state: st.session_state.debug_checkbox_key_main = False # Chave específica

uploaded_file = st.file_uploader("Selecione seu arquivo DOCX", type="docx", key="file_uploader_widget_main")
show_debug_info = st.sidebar.checkbox("Mostrar Informações de Depuração", 
                                    value=st.session_state.debug_checkbox_key_main, 
                                    key="debug_checkbox_widget_main")
st.session_state.debug_checkbox_key_main = show_debug_info


if uploaded_file:
    if st.session_state.nome_arquivo_atual != uploaded_file.name: 
        st.session_state.sugestoes_gemini = []
        st.session_state.conteudo_docx = {"texto": "", "tabelas": []}
        st.session_state.config_sugestoes = {}
        st.session_state.nome_arquivo_atual = uploaded_file.name

    if not st.session_state.sugestoes_gemini: 
        texto_doc, tabelas_doc = extrair_conteudo_docx(uploaded_file)
        st.session_state.conteudo_docx = {"texto": texto_doc, "tabelas": tabelas_doc}
        
        if texto_doc or tabelas_doc:
            st.success(f"Documento '{uploaded_file.name}' lido com sucesso.")
            if show_debug_info:
                with st.expander("Debug: Conteúdo Extraído do DOCX (após processamento de tipos)", expanded=False):
                    st.text_area("Texto Extraído (amostra)", texto_doc[:1500], height=100) # Amostra menor
                    for t_info_debug in tabelas_doc:
                        st.write(f"ID: {t_info_debug['id']}, Nome da Tabela: {t_info_debug['nome']}")
                        try:
                            st.dataframe(t_info_debug['dataframe'].head().astype(str)) 
                        except Exception as e_df_display_debug:
                            st.warning(f"Não foi possível exibir head do DF {t_info_debug['id']} com st.dataframe: {e_df_display_debug}")
                            st.text(f"Head como string:\n{t_info_debug['dataframe'].head().to_string()}")
                        st.write("Tipos de dados das colunas (após conversão):", t_info_debug['dataframe'].dtypes)
                        st.divider()
            
            sugestoes = analisar_documento_com_gemini(texto_doc, tabelas_doc)
            st.session_state.sugestoes_gemini = sugestoes
            for sug_idx_init, sug_init in enumerate(sugestoes):
                s_id_init = sug_init.get("id", f"sug_{sug_idx_init}_{hash(sug_init.get('titulo',''))}")
                sug_init["id"] = s_id_init 
                if s_id_init not in st.session_state.config_sugestoes:
                    st.session_state.config_sugestoes[s_id_init] = {
                        "aceito": True, "titulo_editado": sug_init.get("titulo", "Sem Título"),
                        "dados_originais": sug_init }
        else:
            st.warning("Nenhum conteúdo (texto ou tabelas) extraído do documento.")

if st.session_state.sugestoes_gemini:
    st.sidebar.header("⚙️ Configurar Visualizações Sugeridas")
    for sug_original_sidebar in st.session_state.sugestoes_gemini:
        s_id_sidebar = sug_original_sidebar['id'] 
        if s_id_sidebar not in st.session_state.config_sugestoes: 
             st.session_state.config_sugestoes[s_id_sidebar] = { # Inicializa se faltar
                "aceito": True, "titulo_editado": sug_original_sidebar.get("titulo", "Sem Título"),
                "dados_originais": sug_original_sidebar }
        config_sidebar = st.session_state.config_sugestoes[s_id_sidebar]

        with st.sidebar.expander(f"{config_sidebar['titulo_editado']}", expanded=False):
            st.caption(f"Tipo: {sug_original_sidebar.get('tipo_sugerido')} | Fonte: {sug_original_sidebar.get('fonte_id')}")
            st.markdown(f"**Justificativa IA:** *{sug_original_sidebar.get('justificativa', 'N/A')}*")
            config_sidebar["aceito"] = st.checkbox("Incluir no Dashboard?", value=config_sidebar["aceito"], key=f"aceito_{s_id_sidebar}")
            config_sidebar["titulo_editado"] = st.text_input("Título para Dashboard", value=config_sidebar["titulo_editado"], key=f"titulo_{s_id_sidebar}")
            tipo_sug_sidebar = sug_original_sidebar.get("tipo_sugerido")
            params_sug_sidebar = config_sidebar["dados_originais"].get("parametros",{}) # Pega dos dados originais para edição
            if tipo_sug_sidebar in ["grafico_barras", "grafico_pizza", "grafico_linha", "grafico_dispersao"] and \
               not params_sug_sidebar.get("dados") and \
               str(sug_original_sidebar.get("fonte_id")).startswith("doc_tabela_"):
                df_correspondente_sidebar = next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"] == sug_original_sidebar.get("fonte_id")), None)
                if df_correspondente_sidebar is not None:
                    opcoes_colunas_sidebar = [""] + df_correspondente_sidebar.columns.tolist()
                    if tipo_sug_sidebar in ["grafico_barras", "grafico_linha", "grafico_dispersao"]:
                        params_sug_sidebar["eixo_x"] = st.selectbox("Eixo X", options=opcoes_colunas_sidebar, index=opcoes_colunas_sidebar.index(params_sug_sidebar.get("eixo_x", "")) if params_sug_sidebar.get("eixo_x", "") in opcoes_colunas_sidebar else 0, key=f"param_x_{s_id_sidebar}")
                        params_sug_sidebar["eixo_y"] = st.selectbox("Eixo Y", options=opcoes_colunas_sidebar, index=opcoes_colunas_sidebar.index(params_sug_sidebar.get("eixo_y", "")) if params_s_id_sidebar.get("eixo_y", "") in opcoes_colunas_sidebar else 0, key=f"param_y_{s_id_sidebar}")
                    elif tipo_sug_sidebar == "grafico_pizza":
                        params_sug_sidebar["categorias"] = st.selectbox("Categorias (Nomes)", options=opcoes_colunas_sidebar, index=opcoes_colunas_sidebar.index(params_sug_sidebar.get("categorias", "")) if params_sug_sidebar.get("categorias", "") in opcoes_colunas_sidebar else 0, key=f"param_cat_{s_id_sidebar}")
                        params_sug_sidebar["valores"] = st.selectbox("Valores", options=opcoes_colunas_sidebar, index=opcoes_colunas_sidebar.index(params_sug_sidebar.get("valores", "")) if params_sug_sidebar.get("valores", "") in opcoes_colunas_sidebar else 0, key=f"param_val_{s_id_sidebar}")

if st.session_state.sugestoes_gemini:
    if st.sidebar.button("🚀 Gerar Dashboard", type="primary", use_container_width=True):
        st.header("📊 Dashboard de Insights")
        kpis_para_renderizar = []
        outros_elementos_para_dashboard = []
        for s_id_render, config_render in st.session_state.config_sugestoes.items():
            if config_render["aceito"]:
                sug_original_render = config_render["dados_originais"]
                item_render = {"titulo": config_render["titulo_editado"], "tipo": sug_original_render.get("tipo_sugerido"),
                               "parametros": sug_original_render.get("parametros", {}),"fonte_id": sug_original_render.get("fonte_id")}
                if item_render["tipo"] == "kpi": kpis_para_renderizar.append(item_render)
                else: outros_elementos_para_dashboard.append(item_render)
        
        if kpis_para_renderizar:
            kpi_cols_render = st.columns(min(len(kpis_para_renderizar), 4))
            for i_kpi, kpi_item_render in enumerate(kpis_para_renderizar):
                with kpi_cols_render[i_kpi % min(len(kpis_para_renderizar), 4)]:
                    st.metric(label=kpi_item_render["titulo"], value=str(kpi_item_render["parametros"].get("valor", "N/A")),
                              delta=str(kpi_item_render["parametros"].get("delta", "")), help=kpi_item_render["parametros"].get("descricao"))
            if outros_elementos_para_dashboard: st.divider()

        if show_debug_info and (kpis_para_renderizar or outros_elementos_para_dashboard):
             with st.expander("Debug: Configurações Finais dos Elementos para Dashboard (após validação)", expanded=False):
                # ... (código de debug dos parâmetros como na versão anterior) ...
                pass # Simplificado para brevidade, mas mantenha seu código de debug aqui se útil

        if outros_elementos_para_dashboard:
            item_cols_render = st.columns(2)
            col_idx_render = 0
            for item_render_main in outros_elementos_para_dashboard:
                with item_cols_render[col_idx_render % 2]:
                    st.subheader(item_render_main["titulo"])
                    try:
                        df_plot_main = None; elemento_renderizado_nesta_iteracao = False
                        if item_render_main["parametros"].get("dados"):
                            try: df_plot_main = pd.DataFrame(item_render_main["parametros"]["dados"])
                            except Exception as e_df_direto: st.warning(f"'{item_render_main['titulo']}': Erro DF de 'dados': {e_df_direto}"); continue
                        elif str(item_render_main["fonte_id"]).startswith("doc_tabela_"):
                            df_plot_main = next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"] == item_render_main["fonte_id"]), None)
                        
                        tipo_render = item_render_main["tipo"]; params_render = item_render_main["parametros"]

                        if tipo_render == "tabela_dados":
                            id_tabela_render = params_render.get("id_tabela_original", item_render_main["fonte_id"])
                            df_tabela_render = next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"] == id_tabela_render), None)
                            if df_tabela_render is not None: st.dataframe(df_tabela_render.astype(str).fillna("-")); elemento_renderizado_nesta_iteracao = True
                            else: st.warning(f"Tabela '{id_tabela_render}' não encontrada para '{item_render_main['titulo']}'.")
                        
                        elif tipo_render == "lista_swot":
                            # ... (código de renderização do SWOT como antes) ...
                            swot_data_render = params_render
                            c1_swot, c2_swot = st.columns(2)
                            swot_map_render = {"forcas": ("Forças 💪", c1_swot), "fraquezas": ("Fraquezas 📉", c1_swot), 
                                               "oportunidades": ("Oportunidades 🚀", c2_swot), "ameacas": ("Ameaças ⚠️", c2_swot)}
                            for key_swot, (header_swot, col_target_swot) in swot_map_render.items():
                                with col_target_swot:
                                    st.markdown(f"##### {header_swot}")
                                    points_swot = swot_data_render.get(key_swot, ["N/A (informação não fornecida)"])
                                    if not points_swot or not isinstance(points_swot, list): points_swot = ["N/A (dados ausentes ou formato incorreto)"] 
                                    for point_swot in points_swot: st.markdown(f"- {point_swot}")
                            elemento_renderizado_nesta_iteracao = True
                        
                        elif df_plot_main is not None:
                            x_col, y_col = params_render.get("eixo_x"), params_render.get("eixo_y")
                            cat_col, val_col = params_render.get("categorias"), params_render.get("valores")
                            
                            plot_func = None
                            plot_args = {}

                            if tipo_render == "grafico_barras" and x_col and y_col: plot_func, plot_args = px.bar, {"x": x_col, "y": y_col}
                            elif tipo_render == "grafico_linha" and x_col and y_col: plot_func, plot_args = px.line, {"x": x_col, "y": y_col, "markers": True}
                            elif tipo_render == "grafico_dispersao" and x_col and y_col: plot_func, plot_args = px.scatter, {"x": x_col, "y": y_col}
                            elif tipo_render == "grafico_pizza" and cat_col and val_col: plot_func, plot_args = px.pie, {"names": cat_col, "values": val_col}

                            if plot_func and all(k in df_plot_main.columns for k in plot_args.values() if isinstance(k, str)):
                                st.plotly_chart(plot_func(df_plot_main, title=item_render_main["titulo"], **plot_args), use_container_width=True)
                                elemento_renderizado_nesta_iteracao = True
                            elif plot_func: # Se a função foi definida mas as colunas não bateram
                                st.warning(f"Colunas ausentes/incorretas para '{item_render_main['titulo']}' (tipo: {tipo_render}). Parâmetros: {plot_args}. Colunas no DF: {df_plot_main.columns.tolist()}")
                        
                        if tipo_render == 'mapa': # Placeholder para mapa
                             st.info(f"Visualização de mapa para '{item_render_main['titulo']}' ainda não implementada."); elemento_renderizado_nesta_iteracao = True
                        
                        if not elemento_renderizado_nesta_iteracao and item_render_main["tipo"] not in ["kpi", "tabela_dados", "lista_swot", "mapa"] and df_plot_main is None:
                            st.info(f"'{item_render_main['titulo']}' (tipo: {tipo_render}) não pôde ser gerado. Dados insuficientes (ex: fonte textual sem 'dados' nos parâmetros, ou tabela não encontrada).")
                        elif not elemento_renderizado_nesta_iteracao and item_render_main["tipo"] not in ["kpi", "tabela_dados", "lista_swot", "mapa"]:
                             st.warning(f"Não foi possível gerar visualização para '{item_render_main['titulo']}' (tipo: {tipo_render}). Verifique os parâmetros e o DF.")

                    except Exception as e_render_main: st.error(f"Erro ao renderizar '{item_render_main['titulo']}': {e_render_main}")
                
                if elemento_renderizado_nesta_iteracao: col_idx_render += 1; elementos_renderizados_count +=1
        
        if elementos_renderizados_count == 0 and any(c['aceito'] and c['dados_originais']['tipo_sugerido'] != 'kpi' for c in st.session_state.config_sugestoes.values()):
            st.info("Nenhum gráfico/tabela (além de KPIs) pôde ser gerado com as seleções atuais.")
        elif elementos_renderizados_count == 0 and not kpis_para_renderizar:
            st.info("Nenhum elemento foi selecionado ou pôde ser gerado para o dashboard.")

elif uploaded_file is None and st.session_state.nome_arquivo_atual is not None:
    st.session_state.sugestoes_gemini = []; st.session_state.conteudo_docx = {"texto": "", "tabelas": []}
    st.session_state.config_sugestoes = {}; st.session_state.nome_arquivo_atual = None
    st.session_state.debug_checkbox_key_main = False
    if "file_uploader_widget_main" in st.session_state: st.session_state.file_uploader_widget_main = None 
    st.experimental_rerun()

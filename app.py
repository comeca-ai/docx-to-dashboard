import streamlit as st
from docx import Document
import pandas as pd
import plotly.express as px
import google.generativeai as genai
import json
import os
import traceback
import re # Para limpeza numérica mais avançada

# --- 1. Configuração da Chave da API do Gemini ---
def get_gemini_api_key():
    try:
        # Tenta pegar dos segredos do Streamlit (para deploy)
        return st.secrets["GOOGLE_API_KEY"]
    except (FileNotFoundError, KeyError): # Erros comuns se secrets.toml não existe ou chave não definida
        # Tenta pegar de variável de ambiente (para desenvolvimento local)
        api_key = os.environ.get("GOOGLE_API_KEY")
        if api_key:
            return api_key
        # Não mostra st.error aqui para não poluir se a chave não for usada imediatamente.
        # A função que usa a chave deve lidar com a ausência dela.
        return None

# --- 2. Funções de Processamento do Documento e Interação com Gemini ---

def clean_and_convert_to_numeric(series_data):
    """Tenta limpar e converter uma série Pandas para numérico de forma mais robusta."""
    # Garante que estamos trabalhando com uma Série Pandas
    if not isinstance(series_data, pd.Series):
        s = pd.Series(series_data)
    else:
        s = series_data.copy() # Trabalha com uma cópia para não modificar a original inesperadamente

    if s.dtype == 'object' or isinstance(s.dtype, pd.StringDtype):
        s_str = s.astype(str).str.strip()
        
        # Tenta extrair o primeiro número de strings que podem conter texto e números
        # Ex: "70% - 86%" -> 70.0, "US$ 1.2 Bilhão" -> 1.2
        # Esta regex tenta encontrar o primeiro número (inteiro ou decimal)
        extracted_numbers = s_str.str.extract(r'(\d[\d.,]*)').iloc[:, 0]
        
        if extracted_numbers.notna().any(): # Se algum número foi extraído
            # Limpeza adicional nos números extraídos
            # Remove pontos usados como separadores de milhar ANTES de trocar vírgula por ponto decimal
            cleaned_numbers = extracted_numbers.str.replace(r'\.(?=\d{3}(?:,|\.|$))', '', regex=True)
            cleaned_numbers = cleaned_numbers.str.replace(',', '.', regex=False)
            # Remove quaisquer outros caracteres não numéricos (exceto o ponto decimal e sinal negativo no início)
            cleaned_numbers = cleaned_numbers.str.replace(r'[^\d.-]', '', regex=True)
            
            numeric_col = pd.to_numeric(cleaned_numbers, errors='coerce')
            
            # Se a conversão resultou em muitos NaNs, talvez a extração não foi boa.
            # Comparamos com a série original antes da extração regex.
            if numeric_col.notna().sum() < s.notna().sum() * 0.3: # Se menos de 30% dos não-nulos originais viraram números
                return pd.to_numeric(s, errors='coerce') # Tenta converter a série original (menos agressivo)
            return numeric_col
        else: # Se a regex não extraiu nada, tenta uma conversão direta na string original
            return pd.to_numeric(s_str, errors='coerce')

    return pd.to_numeric(s, errors='coerce') # Tenta converter diretamente se não for objeto/string


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
                    if p_text and len(p_text) < 100 : nome_tabela = p_text.replace(":", "").strip()
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
                    if converted_numeric.notna().sum() > len(df[col]) * 0.3: 
                        df[col] = converted_numeric
                        continue
                    else: 
                         df[col] = original_col_data.copy()

                    # 2. Tentar converter para datetime
                    try:
                        temp_col_str = df[col].astype(str)
                        possible_formats = [
                            '%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', 
                            '%d-%m-%Y', '%m-%d-%Y', '%Y', '%Y%m%d', '%d.%m.%Y',
                            '%Y-%m', '%m-%Y' 
                        ]
                        converted_with_format = False
                        for fmt in possible_formats:
                            try:
                                dt_series = pd.to_datetime(temp_col_str, format=fmt, errors='coerce')
                                if dt_series.notna().sum() > len(df[col]) * 0.5:
                                    df[col] = dt_series
                                    converted_with_format = True
                                    break
                            except (ValueError, TypeError):
                                continue
                        if not converted_with_format:
                            # O argumento infer_datetime_format é depreciado e o comportamento padrão é True.
                            # Podemos remover o argumento para evitar o warning.
                            inferred_dt_series = pd.to_datetime(temp_col_str, errors='coerce') 
                            if inferred_dt_series.notna().sum() > len(df[col]) * 0.5:
                                 df[col] = inferred_dt_series
                            else: 
                                df[col] = original_col_data.astype(str).fillna('')
                    except Exception:
                        df[col] = original_col_data.astype(str).fillna('')
                
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
            df_sample = df.head(7) 
            if len(df.columns) > 10: df_sample = df_sample.iloc[:, :10]
            markdown_tabela = df_sample.to_markdown(index=False)
            tabelas_prompt_str += f"\n--- Tabela '{t_info['nome']}' (ID para referência: {t_info['id']}) ---\n"
            col_types_str = ", ".join([f"'{col}' (tipo inferido: {str(dtype)})" for col, dtype in df.dtypes.items()])
            tabelas_prompt_str += f"Colunas e tipos: {col_types_str}\nAmostra:\n{markdown_tabela}\n"

        max_texto_len = 60000
        texto_doc_para_prompt = texto_doc[:max_texto_len] + ("\n[TEXTO TRUNCADO...]" if len(texto_doc) > max_texto_len else "")

        prompt = f"""
        Você é um assistente de análise de dados e visualização. Analise o texto e as tabelas de um documento.
        [TEXTO DO DOCUMENTO]
        {texto_doc_para_prompt}
        [FIM DO TEXTO]
        [TABELAS DO DOCUMENTO (com ID, nome, colunas, tipos e amostra)]
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
                Se DADOS EXTRAÍDOS DO TEXTO: {{"dados": [{{"NomeQualquerParaEixoX": "CategoriaA", "NomeQualquerParaEixoY": ValorNumericoA}}, ...], "eixo_x": "NomeQualquerParaEixoX", "eixo_y": "NomeQualquerParaEixoY"}}
            - Para "grafico_pizza":
                Se baseado em TABELA: {{"categorias": "NOME_EXATO_COLUNA_CATEGORIAS", "valores": "NOME_EXATO_COLUNA_VALORES_NUMERICOS"}}
                Se DADOS EXTRAÍDOS DO TEXTO: {{"dados": [{{"NomeQualquerParaCategoria": "CategoriaA", "NomeQualquerParaValor": ValorNumericoA}}, ...], "categorias": "NomeQualquerParaCategoria", "valores": "NomeQualquerParaValor"}}
        - "justificativa": string (breve explicação da utilidade).

        Exemplos de `parametros` para GRÁFICOS DE TABELAS (assumindo que a tabela 'doc_tabela_1' tem colunas 'Ano' (int64) e 'Vendas' (float64)):
        - Gráfico de Linha: {{"eixo_x": "Ano", "eixo_y": "Vendas"}}
        - Gráfico de Barras: {{"eixo_x": "Ano", "eixo_y": "Vendas"}}

        Exemplo de `parametros` para GRÁFICO DE PIZZA COM DADOS EXTRAÍDOS DO TEXTO (assumindo que os valores extraídos são numéricos):
        {{"dados": [{{"Região": "Norte", "Faturamento": 50000.0}}, {{"Região": "Sul", "Faturamento": 75000.0}}], "categorias": "Região", "valores": "Faturamento"}}

        Certifique-se de que os NOMES DE COLUNAS nos 'parametros' correspondam EXATAMENTE aos nomes de colunas e tipos de dados fornecidos na descrição das tabelas. Se uma coluna de valor não for numérica (ex: 'object' contendo '70% - 80%'), extraia um valor numérico representativo (ex: média, ou o primeiro número como float) se possível, ou não sugira o gráfico se não for viável tratar como numérico. Para valores como '70% - 80%', se for para 'Market Share', use a média (ex: 78.0). Para '17,35 Bilhões', extraia 17.35 (e a unidade 'Bilhões' pode ser parte do título ou justificativa).
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

# Gerenciamento de estado
if "sugestoes_gemini" not in st.session_state: st.session_state.sugestoes_gemini = []
if "conteudo_docx" not in st.session_state: st.session_state.conteudo_docx = {"texto": "", "tabelas": []}
if "config_sugestoes" not in st.session_state: st.session_state.config_sugestoes = {}
if "nome_arquivo_atual" not in st.session_state: st.session_state.nome_arquivo_atual = None
# Chave para o widget checkbox de debug
if 'debug_checkbox_key' not in st.session_state: st.session_state.debug_checkbox_key = False


uploaded_file = st.file_uploader("Selecione seu arquivo DOCX", type="docx", key="file_uploader_widget")

# Usar o valor do session_state para o checkbox
show_debug_info = st.sidebar.checkbox(
    "Mostrar Informações de Depuração", 
    value=st.session_state.debug_checkbox_key, 
    key="debug_checkbox_widget" # A chave do widget que atualiza st.session_state.debug_checkbox_key
)
# Sincroniza o valor lido do widget de volta para a chave de estado que usamos para lógica
st.session_state.debug_checkbox_key = show_debug_info


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
                with st.expander("Debug: Conteúdo Extraído do DOCX (após processamento de tipos)"):
                    st.text_area("Texto Extraído (amostra)", texto_doc[:2000], height=100)
                    for t_info in tabelas_doc:
                        st.write(f"ID: {t_info['id']}, Nome da Tabela: {t_info['nome']}")
                        try:
                            st.dataframe(t_info['dataframe'].head().astype(str)) 
                        except Exception as e_df_display:
                            st.warning(f"Não foi possível exibir head do DF {t_info['id']} com st.dataframe: {e_df_display}")
                            st.text(f"Head como string:\n{t_info['dataframe'].head().to_string()}")
                        st.write("Tipos de dados das colunas (após conversão):", t_info['dataframe'].dtypes)
                        st.divider()
            
            sugestoes = analisar_documento_com_gemini(texto_doc, tabelas_doc)
            st.session_state.sugestoes_gemini = sugestoes
            for sug_idx_init, sug_init in enumerate(sugestoes):
                s_id_init = sug_init.get("id", f"sug_{sug_idx_init}_{hash(sug_init.get('titulo'))}")
                sug_init["id"] = s_id_init # Garante que a sugestão original tem o ID
                if s_id_init not in st.session_state.config_sugestoes:
                    st.session_state.config_sugestoes[s_id_init] = {
                        "aceito": True, 
                        "titulo_editado": sug_init.get("titulo", "Sem Título"),
                        "dados_originais": sug_init 
                    }
        else:
            st.warning("Nenhum conteúdo (texto ou tabelas) extraído do documento.")

if st.session_state.sugestoes_gemini:
    st.sidebar.header("⚙️ Configurar Visualizações Sugeridas")
    for sug_original_sidebar in st.session_state.sugestoes_gemini:
        s_id_sidebar = sug_original_sidebar['id'] # Usa o ID já garantido
        
        if s_id_sidebar not in st.session_state.config_sugestoes: # Segurança, deve ter sido inicializado
            st.session_state.config_sugestoes[s_id_sidebar] = {
                "aceito": True, 
                "titulo_editado": sug_original_sidebar.get("titulo", "Sem Título"),
                "dados_originais": sug_original_sidebar 
            }
        config_sidebar = st.session_state.config_sugestoes[s_id_sidebar]

        with st.sidebar.expander(f"{config_sidebar['titulo_editado']}", expanded=False):
            st.caption(f"Tipo: {sug_original_sidebar.get('tipo_sugerido')} | Fonte: {sug_original_sidebar.get('fonte_id')}")
            st.markdown(f"**Justificativa IA:** *{sug_original_sidebar.get('justificativa', 'N/A')}*")
            
            config_sidebar["aceito"] = st.checkbox("Incluir no Dashboard?", value=config_sidebar["aceito"], key=f"aceito_{s_id_sidebar}")
            config_sidebar["titulo_editado"] = st.text_input("Título para Dashboard", value=config_sidebar["titulo_editado"], key=f"titulo_{s_id_sidebar}")

            # Edição de parâmetros para gráficos comuns se não vierem de dados (diretos da LLM)
            tipo_sug_sidebar = sug_original_sidebar.get("tipo_sugerido")
            params_sug_sidebar = sug_original_sidebar.get("parametros",{})
            if tipo_sug_sidebar in ["grafico_barras", "grafico_pizza", "grafico_linha", "grafico_dispersao"] and \
               not params_sug_sidebar.get("dados") and \
               str(sug_original_sidebar.get("fonte_id")).startswith("doc_tabela_"):
                
                df_correspondente_sidebar = next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"] == sug_original_sidebar.get("fonte_id")), None)
                if df_correspondente_sidebar is not None:
                    opcoes_colunas_sidebar = [""] + df_correspondente_sidebar.columns.tolist()
                    # Atualiza os parâmetros na sugestão original dentro de config_sugestoes
                    editable_params = config_sidebar["dados_originais"]["parametros"] 

                    if tipo_sug_sidebar in ["grafico_barras", "grafico_linha", "grafico_dispersao"]:
                        editable_params["eixo_x"] = st.selectbox("Eixo X", options=opcoes_colunas_sidebar, index=opcoes_colunas_sidebar.index(editable_params.get("eixo_x", "")) if editable_params.get("eixo_x", "") in opcoes_colunas_sidebar else 0, key=f"param_x_{s_id_sidebar}")
                        editable_params["eixo_y"] = st.selectbox("Eixo Y", options=opcoes_colunas_sidebar, index=opcoes_colunas_sidebar.index(editable_params.get("eixo_y", "")) if editable_params.get("eixo_y", "") in opcoes_colunas_sidebar else 0, key=f"param_y_{s_id_sidebar}")
                    elif tipo_sug_sidebar == "grafico_pizza":
                        editable_params["categorias"] = st.selectbox("Categorias (Nomes)", options=opcoes_colunas_sidebar, index=opcoes_colunas_sidebar.index(editable_params.get("categorias", "")) if editable_params.get("categorias", "") in opcoes_colunas_sidebar else 0, key=f"param_cat_{s_id_sidebar}")
                        editable_params["valores"] = st.selectbox("Valores", options=opcoes_colunas_sidebar, index=opcoes_colunas_sidebar.index(editable_params.get("valores", "")) if editable_params.get("valores", "") in opcoes_colunas_sidebar else 0, key=f"param_val_{s_id_sidebar}")


if st.session_state.sugestoes_gemini:
    if st.sidebar.button("🚀 Gerar Dashboard", type="primary", use_container_width=True):
        st.header("📊 Dashboard de Insights")
        
        kpis_para_renderizar = []
        outros_elementos_para_dashboard = [] # Renomeado para evitar conflito

        for s_id_render, config_render in st.session_state.config_sugestoes.items():
            if config_render["aceito"]:
                sug_original_render = config_render["dados_originais"]
                item_render = {"titulo": config_render["titulo_editado"], 
                               "tipo": sug_original_render.get("tipo_sugerido"),
                               "parametros": sug_original_render.get("parametros", {}),
                               "fonte_id": sug_original_render.get("fonte_id")}
                if item_render["tipo"] == "kpi":
                    kpis_para_renderizar.append(item_render)
                else:
                    outros_elementos_para_dashboard.append(item_render) # Adiciona aqui
        
        if kpis_para_renderizar:
            kpi_cols_render = st.columns(min(len(kpis_para_renderizar), 4))
            for i_kpi, kpi_item_render in enumerate(kpis_para_renderizar):
                with kpi_cols_render[i_kpi % min(len(kpis_para_renderizar), 4)]:
                    st.metric(label=kpi_item_render["titulo"], 
                              value=str(kpi_item_render["parametros"].get("valor", "N/A")),
                              delta=str(kpi_item_render["parametros"].get("delta", "")),
                              help=kpi_item_render["parametros"].get("descricao"))
            if outros_elementos_para_dashboard: st.divider() # Adiciona divisor apenas se houver outros elementos

        if show_debug_info and (kpis_para_renderizar or outros_elementos_para_dashboard):
             with st.expander("Debug: Configurações Finais dos Elementos para Dashboard", expanded=False):
                if kpis_para_renderizar: st.json({"KPIs": kpis_para_renderizar})
                if outros_elementos_para_dashboard: st.json({"Outros Elementos": outros_elementos_para_dashboard})
                st.subheader("DataFrames Referenciados (se aplicável):")
                # Mostrar DFs que serão usados pelos gráficos/tabelas (se não forem dados diretos)
                ids_tabelas_usadas = set()
                for item_debug_df in kpis_para_renderizar + outros_elementos_para_dashboard:
                    if str(item_debug_df['fonte_id']).startswith("doc_tabela_") and not item_debug_df['parametros'].get("dados"):
                        ids_tabelas_usadas.add(item_debug_df['fonte_id'])
                    elif item_debug_df['tipo'] == "tabela_dados" and item_debug_df['parametros'].get("id_tabela_original"):
                        ids_tabelas_usadas.add(item_debug_df['parametros']['id_tabela_original'])
                
                for id_t_debug in ids_tabelas_usadas:
                    df_rel_debug = next((t['dataframe'] for t in st.session_state.conteudo_docx['tabelas'] if t['id'] == id_t_debug),None)
                    if df_rel_debug is not None: 
                        st.write(f"DataFrame ID: {id_t_debug} (head e dtypes):")
                        try: st.dataframe(df_rel_debug.head().astype(str))
                        except: st.text(df_rel_debug.head().to_string())
                        st.write(df_rel_debug.dtypes)
                        st.divider()


        if outros_elementos_para_dashboard:
            item_cols_render = st.columns(2)
            col_idx_render = 0
            for item_render_main in outros_elementos_para_dashboard:
                with item_cols_render[col_idx_render % 2]:
                    st.subheader(item_render_main["titulo"])
                    try:
                        df_plot_main = None
                        # Prioriza dados diretos, depois busca em tabelas extraídas
                        if item_render_main["parametros"].get("dados"):
                            try: df_plot_main = pd.DataFrame(item_render_main["parametros"]["dados"])
                            except Exception as e_df_direto_main: st.warning(f"'{item_render_main['titulo']}': Erro ao criar DF de 'dados': {e_df_direto_main}"); continue
                        elif str(item_render_main["fonte_id"]).startswith("doc_tabela_"):
                            df_plot_main = next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"] == item_render_main["fonte_id"]), None)
                        
                        tipo_render = item_render_main["tipo"]
                        params_render = item_render_main["parametros"]

                        if tipo_render == "tabela_dados":
                            id_tabela_render = params_render.get("id_tabela_original", item_render_main["fonte_id"])
                            df_tabela_render = next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"] == id_tabela_render), None)
                            if df_tabela_render is not None: st.dataframe(df_tabela_render.astype(str)) 
                            else: st.warning(f"Tabela '{id_tabela_render}' não encontrada para '{item_render_main['titulo']}'.")
                        
                        elif tipo_render == "lista_swot":
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
                        
                        elif df_plot_main is not None: # Gráficos que usam df_plot_main
                            x_col_render = params_render.get("eixo_x")
                            y_col_render = params_render.get("eixo_y")
                            cat_col_render = params_render.get("categorias")
                            val_col_render = params_render.get("valores")

                            if tipo_render == "grafico_barras" and x_col_render and y_col_render and x_col_render in df_plot_main.columns and y_col_render in df_plot_main.columns:
                                st.plotly_chart(px.bar(df_plot_main, x=x_col_render, y=y_col_render, title=item_render_main["titulo"]), use_container_width=True)
                            elif tipo_render == "grafico_linha" and x_col_render and y_col_render and x_col_render in df_plot_main.columns and y_col_render in df_plot_main.columns:
                                st.plotly_chart(px.line(df_plot_main, x=x_col_render, y=y_col_render, title=item_render_main["titulo"], markers=True), use_container_width=True)
                            elif tipo_render == "grafico_dispersao" and x_col_render and y_col_render and x_col_render in df_plot_main.columns and y_col_render in df_plot_main.columns:
                                st.plotly_chart(px.scatter(df_plot_main, x=x_col_render, y=y_col_render, title=item_render_main["titulo"]), use_container_width=True)
                            elif tipo_render == "grafico_pizza" and cat_col_render and val_col_render and cat_col_render in df_plot_main.columns and val_col_render in df_plot_main.columns:
                                st.plotly_chart(px.pie(df_plot_main, names=cat_col_render, values=val_col_render, title=item_render_main["titulo"]), use_container_width=True)
                            elif tipo_render not in ["tabela_dados", "lista_swot"]: # Se não é um tipo conhecido E tem df_plot_main
                                st.warning(f"Não foi possível gerar gráfico '{item_render_main['titulo']}' (tipo: {tipo_render}). Colunas X/Y ou Categorias/Valores podem estar ausentes/incorretas nos parâmetros ou no DataFrame. X: '{x_col_render}', Y: '{y_col_render}', Cat: '{cat_col_render}', Val: '{val_col_render}'. Colunas no DF: {df_plot_main.columns.tolist()}")
                        
                        elif tipo_render not in ["kpi", "tabela_dados", "lista_swot"]:
                            st.info(f"Visualização '{item_render_main['titulo']}' (tipo: {tipo_render}) não pôde ser gerada. Dados insuficientes (ex: fonte textual sem 'dados' nos parâmetros, ou tabela referenciada não encontrada).")
                    except Exception as e_render_main:
                        st.error(f"Erro ao renderizar '{item_render_main['titulo']}': {e_render_main}")
                        # st.text(traceback.format_exc()) # Descomentar para debug detalhado do erro de plotagem
                col_idx_render += 1
        
        if not kpis_para_renderizar and not outros_elementos_para_dashboard:
            st.info("Nenhum elemento selecionado ou passível de ser gerado para o dashboard.")

elif uploaded_file is None and st.session_state.nome_arquivo_atual is not None:
    st.session_state.sugestoes_gemini = []
    st.session_state.conteudo_docx = {"texto": "", "tabelas": []}
    st.session_state.config_sugestoes = {}
    st.session_state.nome_arquivo_atual = None
    st.session_state.debug_checkbox_key = False # Reseta o checkbox
    if "file_uploader_widget" in st.session_state: st.session_state.file_uploader_widget = None 
    st.experimental_rerun()

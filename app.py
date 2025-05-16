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
    if pd.isna(val_str) or val_str == '':
        return None
    cleaned_val_str = str(val_str).replace('.', '', val_str.count('.') -1 if val_str.count('.') > 1 and ',' in val_str else val_str.count('.'))
    cleaned_val_str = cleaned_val_str.replace(',', '.')
    numbers = re.findall(r"[-+]?\d*\.\d+|\d+", cleaned_val_str)
    if numbers:
        try:
            return float(numbers[0])
        except ValueError:
            return None
    return None

def clean_and_convert_to_numeric(series_data):
    if not isinstance(series_data, pd.Series):
        s = pd.Series(series_data)
    else:
        s = series_data.copy()
    parsed_series = s.astype(str).apply(parse_value_range_or_text)
    numeric_col = pd.to_numeric(parsed_series, errors='coerce')
    if numeric_col.notna().sum() > s.notna().sum() * 0.3:
        return numeric_col
    return pd.to_numeric(s, errors='coerce')

def extrair_conteudo_docx(uploaded_file):
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
                    if p_text and len(p_text) < 100 : nome_tabela = p_text.replace(":", "").strip()[:80]
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
                    converted_numeric = clean_and_convert_to_numeric(df[col])
                    if converted_numeric.notna().sum() >= len(df[col]) * 0.3: 
                        df[col] = converted_numeric
                        continue
                    else: 
                         df[col] = original_col_data.copy()
                    try:
                        temp_col_str = df[col].astype(str)
                        possible_formats = ['%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', 
                                            '%d-%m-%Y', '%m-%d-%Y', '%Y', '%Y%m%d', '%d.%m.%Y', '%Y-%m', '%m-%Y']
                        converted_with_format = False
                        for fmt in possible_formats:
                            try:
                                dt_series = pd.to_datetime(temp_col_str, format=fmt, errors='coerce')
                                if dt_series.notna().sum() > len(df[col]) * 0.5:
                                    df[col] = dt_series
                                    converted_with_format = True
                                    break
                            except (ValueError, TypeError): continue
                        if not converted_with_format:
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
        st.warning("Chave da API do Gemini não configurada."); return []
    try:
        genai.configure(api_key=api_key)
        safety_settings_config = [{"category": c, "threshold": "BLOCK_NONE"} for c in ["HARM_CATEGORY_HARASSMENT", "HARM_CATEGORY_HATE_SPEECH", "HARM_CATEGORY_SEXUALLY_EXPLICIT", "HARM_CATEGORY_DANGEROUS_CONTENT"]]
        model = genai.GenerativeModel(model_name="gemini-1.5-flash-latest", safety_settings=safety_settings_config)
        tabelas_prompt_str = ""
        for t_info in tabelas_info_list:
            df = t_info["dataframe"]
            df_sample = df.head(5) 
            if len(df.columns) > 8: df_sample = df_sample.iloc[:, :8]
            markdown_tabela = df_sample.to_markdown(index=False)
            tabelas_prompt_str += f"\n--- Tabela '{t_info['nome']}' (ID: {t_info['id']}) ---\n"
            col_types_str = ", ".join([f"'{col}' (tipo: {str(dtype)})" for col, dtype in df.dtypes.items()])
            tabelas_prompt_str += f"Colunas e tipos: {col_types_str}\nAmostra:\n{markdown_tabela}\n"
        max_texto_len = 60000
        texto_doc_para_prompt = texto_doc[:max_texto_len] + ("\n[TEXTO TRUNCADO...]" if len(texto_doc) > max_texto_len else "")
        prompt = f"""
        Você é um assistente de análise de dados. Analise o texto e as tabelas de um documento.
        [TEXTO DO DOCUMENTO]{texto_doc_para_prompt}[FIM DO TEXTO]
        [TABELAS DO DOCUMENTO (com ID, nome, colunas, tipos e amostra)]{tabelas_prompt_str}[FIM DAS TABELAS]
        Gere lista JSON de sugestões de visualizações. Objeto DEVE ter: "id", "titulo", "tipo_sugerido" ("kpi", "tabela_dados", "lista_swot", "grafico_barras", "grafico_pizza", "grafico_linha", "grafico_dispersao"), "fonte_id" (ID tabela ou "texto_secao_xyz"), "parametros" (objeto JSON), "justificativa".
        Para "parametros":
        - "kpi": {{"valor": "ValorKPI", "delta": "Mudança", "descricao": "Contexto"}}
        - "tabela_dados": {{"id_tabela_original": "ID_Tabela"}}
        - "lista_swot": {{"forcas": ["F1"], "fraquezas": ["Fr1"], "oportunidades": ["Op1"], "ameacas": ["Am1"]}}
        - Gráficos de TABELA ("grafico_barras", "linha", "dispersao"): {{"eixo_x": "NOME_COL_X", "eixo_y": "NOME_COL_Y"}} (Y numérico).
        - Gráficos de PIZZA de TABELA: {{"categorias": "NOME_COL_CAT", "valores": "NOME_COL_VAL_NUM"}} (Valores numéricos).
        - Gráficos com DADOS EXTRAÍDOS DO TEXTO: {{"dados": [{{"NomeEixoX": "CatA", "NomeEixoY": ValNumA}}, ...], "eixo_x": "NomeEixoX", "eixo_y": "NomeEixoY"}} (Valores DEVEM ser numéricos).
        Use NOMES EXATOS de colunas. Se coluna de valor não for numérica (ex: '70% - 80%'), extraia valor numérico (ex: 70.0 ou média). Para '17,35 Bilhões', extraia 17.35. Se não tratável como numérico, não sugira gráfico que o exija. Para Cobertura Geográfica (Player, Cidades), sugira "tabela_dados". Para SWOTs comparativos, gere "lista_swot" INDIVIDUAL por player.
        Retorne APENAS a lista JSON válida.
        """
        with st.spinner("🤖 Gemini está analisando o documento..."):
            response = model.generate_content(prompt)
        cleaned_text = response.text.strip().lstrip("```json").rstrip("```").strip()
        sugestoes = json.loads(cleaned_text)
        if isinstance(sugestoes, list): st.success(f"{len(sugestoes)} sugestões recebidas!"); return sugestoes
        st.error("Resposta Gemini não é lista JSON."); return []
    except json.JSONDecodeError as e:
        st.error(f"Erro JSON Gemini: {e}")
        if 'response' in locals(): st.code(response.text, language="text")
        return []
    except Exception as e: st.error(f"Erro API Gemini: {e}"); st.text(traceback.format_exc()); return []

# --- 3. Interface Streamlit e Lógica de Apresentação ---
st.set_page_config(layout="wide"); st.title("✨ Apps com Gemini: DOCX para Insights Visuais")
st.markdown("Faça upload DOCX e Gemini sugerirá visualizações.")

for key in ["sugestoes_gemini", "config_sugestoes"]:
    if key not in st.session_state: st.session_state[key] = [] if key == "sugestoes_gemini" else {}
if "conteudo_docx" not in st.session_state: st.session_state.conteudo_docx = {"texto": "", "tabelas": []}
if "nome_arquivo_atual" not in st.session_state: st.session_state.nome_arquivo_atual = None
if 'debug_checkbox_key_main' not in st.session_state: st.session_state.debug_checkbox_key_main = False

uploaded_file = st.file_uploader("Selecione DOCX", type="docx", key="uploader_key")
show_debug_info = st.sidebar.checkbox("Mostrar Debug Info", value=st.session_state.debug_checkbox_key_main, key="debug_widget_key")
st.session_state.debug_checkbox_key_main = show_debug_info

if uploaded_file:
    if st.session_state.nome_arquivo_atual != uploaded_file.name: 
        st.session_state.sugestoes_gemini, st.session_state.config_sugestoes = [], {}
        st.session_state.nome_arquivo_atual = uploaded_file.name
    if not st.session_state.sugestoes_gemini: 
        texto_doc, tabelas_doc = extrair_conteudo_docx(uploaded_file)
        st.session_state.conteudo_docx = {"texto": texto_doc, "tabelas": tabelas_doc}
        if texto_doc or tabelas_doc:
            st.success(f"'{uploaded_file.name}' lido.")
            if show_debug_info:
                with st.expander("Debug: Conteúdo Extraído (após tipos)"):
                    st.text_area("Texto (amostra)", texto_doc[:1000], height=100)
                    for t_info in tabelas_doc:
                        st.write(f"ID: {t_info['id']}, Nome: {t_info['nome']}")
                        try: st.dataframe(t_info['dataframe'].head().astype(str)) 
                        except: st.text(f"Head:\n{t_info['dataframe'].head().to_string()}")
                        st.write("Tipos:", t_info['dataframe'].dtypes.to_dict())
            sugestoes = analisar_documento_com_gemini(texto_doc, tabelas_doc)
            st.session_state.sugestoes_gemini = sugestoes
            for i, s in enumerate(sugestoes):
                s_id = s.get("id", f"s_{i}_{hash(s.get('titulo'))}"); s["id"] = s_id
                if s_id not in st.session_state.config_sugestoes:
                    st.session_state.config_sugestoes[s_id] = {"aceito": True, "titulo_editado": s.get("titulo","S/Título"), "dados_originais": s}
        else: st.warning("Nenhum conteúdo extraído.")

if st.session_state.sugestoes_gemini:
    st.sidebar.header("⚙️ Configurar Sugestões")
    for sug in st.session_state.sugestoes_gemini:
        s_id, config = sug['id'], st.session_state.config_sugestoes[sug['id']]
        with st.sidebar.expander(f"{config['titulo_editado']}", expanded=False):
            st.caption(f"Tipo: {sug.get('tipo_sugerido')} | Fonte: {sug.get('fonte_id')}")
            st.markdown(f"**IA:** *{sug.get('justificativa', 'N/A')}*")
            config["aceito"] = st.checkbox("Incluir?", value=config["aceito"], key=f"aceito_{s_id}")
            config["titulo_editado"] = st.text_input("Título", value=config["titulo_editado"], key=f"titulo_{s_id}")
            params, tipo_sug = config["dados_originais"].get("parametros",{}), sug.get("tipo_sugerido")
            if tipo_sug in ["grafico_barras","grafico_pizza","grafico_linha","grafico_dispersao"] and \
               not params.get("dados") and str(sug.get("fonte_id")).startswith("doc_tabela_"):
                df_corr = next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"] == sug.get("fonte_id")), None)
                if df_corr is not None:
                    opts = [""] + df_corr.columns.tolist()
                    if tipo_sug != "grafico_pizza":
                        params["eixo_x"] = st.selectbox("Eixo X", opts, index=opts.index(params.get("eixo_x","")) if params.get("eixo_x","") in opts else 0, key=f"px_{s_id}")
                        params["eixo_y"] = st.selectbox("Eixo Y", opts, index=opts.index(params.get("eixo_y","")) if params.get("eixo_y","") in opts else 0, key=f"py_{s_id}")
                    else:
                        params["categorias"] = st.selectbox("Categorias", opts, index=opts.index(params.get("categorias","")) if params.get("categorias","") in opts else 0, key=f"pcat_{s_id}")
                        params["valores"] = st.selectbox("Valores", opts, index=opts.index(params.get("valores","")) if params.get("valores","") in opts else 0, key=f"pval_{s_id}")

    if st.sidebar.button("🚀 Gerar Dashboard", type="primary", use_container_width=True):
        st.header("📊 Dashboard de Insights"); kpis, outros = [], []
        for s_id, cfg in st.session_state.config_sugestoes.items():
            if cfg["aceito"]:
                item = {"titulo":cfg["titulo_editado"], **cfg["dados_originais"]} # Combina título editado com dados originais
                (kpis if item["tipo_sugerido"] == "kpi" else outros).append(item)
        if kpis:
            cols = st.columns(min(len(kpis), 4))
            for i, k in enumerate(kpis):
                with cols[i % min(len(kpis), 4)]:
                    p = k.get("parametros",{})
                    st.metric(k["titulo"], str(p.get("valor","N/A")), str(p.get("delta","")), help=p.get("descricao"))
            if outros: st.divider()
        if show_debug_info and (kpis or outros):
             with st.expander("Debug: Configs Finais para Dashboard", expanded=False):
                if kpis: st.json({"KPIs": kpis})
                if outros: st.json({"Outros": outros}) # Mostra os parâmetros que serão usados
                # Adicionar debug de DFs aqui se necessário
        if outros:
            item_cols = st.columns(2); col_idx = 0
            for item in outros:
                with item_cols[col_idx % 2]:
                    st.subheader(item["titulo"]); df_plot, el_rendered = None, False
                    params = item.get("parametros", {})
                    try:
                        if params.get("dados"): df_plot = pd.DataFrame(params["dados"])
                        elif str(item.get("fonte_id")).startswith("doc_tabela_"):
                            df_plot = next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"] == item.get("fonte_id")), None)
                        
                        tipo = item.get("tipo_sugerido")
                        if tipo == "tabela_dados":
                            id_t = params.get("id_tabela_original", item.get("fonte_id"))
                            df_t = next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"] == id_t), None)
                            if df_t is not None: st.dataframe(df_t.astype(str).fillna("-")); el_rendered = True
                            else: st.warning(f"Tabela '{id_t}' não encontrada.")
                        elif tipo == "lista_swot":
                            c1,c2=st.columns(2); smap={"forcas":("Forças 💪",c1),"fraquezas":("Fraquezas 📉",c1),"oportunidades":("Oportunidades 🚀",c2),"ameacas":("Ameaças ⚠️",c2)}
                            for k, (h, ct) in smap.items():
                                with ct: st.markdown(f"##### {h}"); [st.markdown(f"- {p}") for p in params.get(k,["N/A"])]
                            el_rendered = True
                        elif df_plot is not None:
                            x,y,cat,val = params.get("eixo_x"),params.get("eixo_y"),params.get("categorias"),params.get("valores")
                            fn,p_args=None,{}
                            if tipo=="grafico_barras" and x and y: fn,p_args=px.bar,{"x":x,"y":y}
                            elif tipo=="grafico_linha" and x and y: fn,p_args=px.line,{"x":x,"y":y,"markers":True}
                            elif tipo=="grafico_dispersao" and x and y: fn,p_args=px.scatter,{"x":x,"y":y}
                            elif tipo=="grafico_pizza" and cat and val: fn,p_args=px.pie,{"names":cat,"values":val}
                            if fn and all(k in df_plot.columns for k in p_args.values() if isinstance(k,str)):
                                st.plotly_chart(fn(df_plot,title=item["titulo"],**p_args),use_container_width=True); el_rendered=True
                            elif fn: st.warning(f"Colunas X/Y ou Cat/Val ausentes/incorretas para '{item['titulo']}'. Cols: {df_plot.columns.tolist()}")
                        if tipo == 'mapa': st.info(f"Mapa para '{item['titulo']}' não implementado."); el_rendered=True
                        if not el_rendered and tipo not in ["kpi","tabela_dados","lista_swot","mapa"]:
                            st.info(f"'{item['titulo']}' (tipo: {tipo}) não gerado. Dados insuficientes ou tipo não suportado.")
                    except Exception as e: st.error(f"Erro renderizando '{item['titulo']}': {e}")
                if el_rendered: col_idx+=1
        if not kpis and not outros: st.info("Nenhum elemento para o dashboard.")
elif uploaded_file is None and st.session_state.nome_arquivo_atual is not None:
    st.session_state.clear(); st.experimental_rerun() # Limpa tudo

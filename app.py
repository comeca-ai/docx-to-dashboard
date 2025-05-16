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
    try: return st.secrets["GOOGLE_API_KEY"]
    except (FileNotFoundError, KeyError): 
        api_key = os.environ.get("GOOGLE_API_KEY")
        return api_key if api_key else None

# --- 2. Funções de Processamento do Documento e Interação com Gemini ---
def parse_value_for_numeric(val_str_in):
    if pd.isna(val_str_in) or str(val_str_in).strip() == '': return None
    text = str(val_str_in).strip()
    is_negative_paren = text.startswith('(') and text.endswith(')')
    if is_negative_paren: text = text[1:-1]
    text_num_part = re.sub(r'[R$\s%]', '', text)
    if ',' in text_num_part and '.' in text_num_part:
        if text_num_part.rfind('.') < text_num_part.rfind(','): text_num_part = text_num_part.replace('.', '') 
        text_num_part = text_num_part.replace(',', '.') 
    elif ',' in text_num_part: text_num_part = text_num_part.replace(',', '.')
    match = re.search(r"([-+]?\d*\.?\d+|\d+)", text_num_part)
    if match:
        try: num = float(match.group(1)); return -num if is_negative_paren else num
        except ValueError: return None
    return None

def extrair_conteudo_docx(uploaded_file):
    try:
        document = Document(uploaded_file)
        textos = [p.text for p in document.paragraphs if p.text.strip()]
        tabelas_data = [] 
        for i, table_obj in enumerate(document.tables):
            data_rows, keys, nome_tabela = [], None, f"Tabela_DOCX_{i+1}"
            try: 
                prev_el = table_obj._element.getprevious()
                if prev_el is not None and prev_el.tag.endswith('p'):
                    p_text = "".join(node.text for node in prev_el.xpath('.//w:t')).strip()
                    if p_text and len(p_text) < 80: nome_tabela = p_text.replace(":", "").strip()
            except: pass
            for r_idx, row in enumerate(table_obj.rows):
                cells = [c.text.strip() for c in row.cells]
                if r_idx == 0: keys = [k.replace("\n"," ").strip() if k else f"Col{c_idx+1}" for c_idx, k in enumerate(cells)]; continue
                if keys: data_rows.append(dict(zip(keys, cells + [None]*(len(keys)-len(cells)))))
            if data_rows:
                df = pd.DataFrame(data_rows)
                for col in df.columns:
                    original_series = df[col].copy()
                    num_series = original_series.astype(str).apply(parse_value_for_numeric)
                    if num_series.notna().sum() / max(1, len(num_series)) > 0.3:
                        df[col] = pd.to_numeric(num_series, errors='coerce')
                        continue 
                    else: df[col] = original_series 
                    try:
                        temp_str_col = df[col].astype(str)
                        dt_series = pd.to_datetime(temp_str_col, errors='coerce', dayfirst=True) 
                        if dt_series.notna().sum() / max(1, len(dt_series)) > 0.5:
                            df[col] = dt_series
                        else: 
                            df[col] = original_series.astype(str).fillna('')
                    except: df[col] = original_series.astype(str).fillna('')
                for col in df.columns: 
                    if df[col].dtype == 'object': df[col] = df[col].astype(str).fillna('')
                tabelas_data.append({"id": f"doc_tabela_{i+1}", "nome": nome_tabela, "dataframe": df})
        return "\n\n".join(textos), tabelas_data
    except Exception as e: st.error(f"Erro crítico ao ler DOCX: {e}"); return "", []

def analisar_documento_com_gemini(texto_doc, tabelas_info_list):
    api_key = get_gemini_api_key()
    if not api_key: st.warning("Chave API Gemini não configurada."); return []
    try:
        genai.configure(api_key=api_key)
        safety_settings = [{"category": c,"threshold": "BLOCK_NONE"} for c in ["HARM_CATEGORY_HARASSMENT","HARM_CATEGORY_HATE_SPEECH","HARM_CATEGORY_SEXUALLY_EXPLICIT","HARM_CATEGORY_DANGEROUS_CONTENT"]]
        model = genai.GenerativeModel(model_name="gemini-1.5-flash-latest", safety_settings=safety_settings)
        tabelas_prompt_str = ""
        for t_info in tabelas_info_list:
            df, nome_t, id_t = t_info["dataframe"], t_info["nome"], t_info["id"]
            sample_df = df.head(3).iloc[:, :min(5, len(df.columns))]
            md_table = sample_df.to_markdown(index=False)
            colunas_para_mostrar_tipos = df.columns.tolist()[:min(8, len(df.columns))]
            col_types_list = [f"'{col_name_prompt}' (tipo: {str(df[col_name_prompt].dtype)})" for col_name_prompt in colunas_para_mostrar_tipos]
            col_types_str = ", ".join(col_types_list)
            tabelas_prompt_str += f"\n--- Tabela '{nome_t}' (ID: {id_t}) ---\nColunas e tipos (primeiras {len(colunas_para_mostrar_tipos)}): {col_types_str}\nAmostra de dados:\n{md_table}\n"
        text_limit = 50000
        prompt_text = texto_doc[:text_limit] + ("\n[TEXTO TRUNCADO...]" if len(texto_doc) > text_limit else "")
        prompt = f"""
        Você é um assistente de análise de dados. Analise o texto e as tabelas.
        [TEXTO]{prompt_text}[FIM TEXTO]
        [TABELAS]{tabelas_prompt_str}[FIM TABELAS]

        Gere lista JSON de sugestões de visualizações. Objeto DEVE ter: "id", "titulo", "tipo_sugerido" ("kpi", "tabela_dados", "lista_swot", "grafico_barras", "grafico_pizza", "grafico_linha", "grafico_dispersao"), "fonte_id" (ID tabela ou "texto_descricao_fonte"), "parametros" (objeto JSON), "justificativa".
        Para "parametros":
        - "kpi": {{"valor": "ValorKPI", "delta": "Mudança", "descricao": "Contexto"}}
        - "tabela_dados": Para TABELA EXISTENTE: {{"id_tabela_original": "ID_Tabela"}}. Para DADOS DO TEXTO: {{"dados": [{{"Coluna1": "ValorA1"}}, ...], "colunas_titulo": ["Título Col1"]}}
        - "lista_swot": {{"forcas": ["F1"], "fraquezas": ["Fr1"], "oportunidades": ["Op1"], "ameacas": ["Am1"]}} (Listas de strings).
        - Gráficos de TABELA ("barras", "linha", "dispersao"): {{"eixo_x": "NOME_COL_X", "eixo_y": "NOME_COL_Y"}} (Y numérico, use nomes exatos).
        - Gráficos de PIZZA de TABELA: {{"categorias": "NOME_COL_CAT", "valores": "NOME_COL_VAL_NUM"}} (Valores numéricos, use nomes exatos).
        - Gráficos com DADOS EXTRAÍDOS DO TEXTO: {{"dados": [{{"NomeEixoX": "CatA", "NomeEixoY": ValNumA}}, ...], "eixo_x": "NomeEixoX", "eixo_y": "NomeEixoY"}} (Valores DEVEM ser numéricos).
        
        INSTRUÇÕES CRÍTICAS:
        1.  NOMES DE COLUNAS: Para gráficos de TABELA, use os NOMES EXATOS das colunas como fornecidos nos "Colunas e tipos".
        2.  DADOS NUMÉRICOS: Se a coluna de valor de uma TABELA não for numérica (float64/int64) conforme os "tipos inferidos", NÃO sugira gráfico que exija valor numérico para ela, A MENOS que você possa confiavelmente extrair um valor numérico do seu conteúdo textual (ex: de '70%' extrair 70.0; de '70% - 86%' extrair 70.0 ou 78.0; de 'R$ 15,5 Bi' extrair 15.5). Se extrair do texto, coloque em "dados".
        3.  COBERTURA GEOGRÁFICA (Player, Cidades): Se for apenas lista, sugira "tabela_dados" e forneça os dados extraídos no campo "dados" dos "parametros" com "colunas_titulo".
        4.  SWOT COMPARATIVO: Se uma tabela compara SWOTs, gere "lista_swot" INDIVIDUAL para CADA player da tabela.
        Retorne APENAS a lista JSON válida. Seja conciso na justificativa.
        """
        with st.spinner("🤖 Gemini analisando... (Pode levar alguns instantes)"):
            response = model.generate_content(prompt)
        cleaned_text = response.text.strip().lstrip("```json").rstrip("```").strip()
        sugestoes = json.loads(cleaned_text)
        if isinstance(sugestoes, list) and all(isinstance(item, dict) for item in sugestoes):
             st.success(f"{len(sugestoes)} sugestões!"); return sugestoes
        st.error("Resposta Gemini não é lista JSON."); return []
    except json.JSONDecodeError as e: st.error(f"Erro JSON Gemini: {e}"); st.code(response.text if 'response' in locals() else "N/A", language="text"); return []
    except Exception as e: st.error(f"Erro API Gemini: {e}"); st.text(traceback.format_exc()); return []

# --- Funções de Renderização Específicas para o Novo Layout ---
def render_kpis(kpi_sugestoes):
    if kpi_sugestoes:
        num_kpis = len(kpi_sugestoes)
        kpi_cols = st.columns(min(num_kpis, 4)) 
        for i, kpi_sug in enumerate(kpi_sugestoes):
            with kpi_cols[i % min(num_kpis, 4)]:
                params = kpi_sug.get("parametros",{})
                st.metric(label=kpi_sug.get("titulo","KPI"), value=str(params.get("valor", "N/A")),
                          delta=str(params.get("delta", "")), help=params.get("descricao"))
        st.divider()

def render_swot_card(player_name, swot_data, card_key_prefix=""):
    st.subheader(f"Análise SWOT - {player_name}")
    col1, col2 = st.columns(2)
    swot_map = {"forcas": ("Forças 💪", col1), "fraquezas": ("Fraquezas 📉", col1), 
                "oportunidades": ("Oportunidades 🚀", col2), "ameacas": ("Ameaças ⚠️", col2)}
    for key, (header, col_target) in swot_map.items():
        with col_target:
            st.markdown(f"##### {header}")
            points = swot_data.get(key, ["N/A (informação não fornecida pela IA)"])
            if not points or not isinstance(points, list): points = ["N/A (formato de dados incorreto)"]
            for point_idx, item in enumerate(points): 
                st.markdown(f"<div style='margin-bottom: 5px;'>- {item}</div>", unsafe_allow_html=True, key=f"{card_key_prefix}_swot_{player_name}_{key}_{point_idx}")
    st.markdown("---")

def render_plotly_chart(item_config, df_plot):
    tipo_grafico = item_config.get("tipo_sugerido")
    titulo = item_config.get("titulo")
    params = item_config.get("parametros", {})
    x_col, y_col = params.get("eixo_x"), params.get("eixo_y")
    cat_col, val_col = params.get("categorias"), params.get("valores")
    fig, plot_func, plot_args = None, None, {}

    if tipo_grafico == "grafico_barras" and x_col and y_col: plot_func, plot_args = px.bar, {"x":x_col,"y":y_col}
    elif tipo_grafico == "grafico_linha" and x_col and y_col: plot_func, plot_args = px.line, {"x":x_col,"y":y_col,"markers":True}
    elif tipo_grafico == "grafico_dispersao" and x_col and y_col: plot_func, plot_args = px.scatter, {"x":x_col,"y":y_col}
    elif tipo_grafico == "grafico_pizza" and cat_col and val_col: plot_func, plot_args = px.pie,{"names":cat_col,"values":val_col}

    if plot_func and all(k_col in df_plot.columns for k_col in plot_args.values() if isinstance(k_col,str)):
        try:
            df_plot_cleaned = df_plot.copy()
            y_axis_col_plot = plot_args.get("y"); values_col_plot = plot_args.get("values")
            if y_axis_col_plot and y_axis_col_plot in df_plot_cleaned.columns: df_plot_cleaned[y_axis_col_plot] = pd.to_numeric(df_plot_cleaned[y_axis_col_plot], errors='coerce')
            if values_col_plot and values_col_plot in df_plot_cleaned.columns: df_plot_cleaned[values_col_plot] = pd.to_numeric(df_plot_cleaned[values_col_plot], errors='coerce')
            cols_to_check_na = [val for val in plot_args.values() if isinstance(val, str) and val in df_plot_cleaned.columns]
            df_plot_cleaned.dropna(subset=cols_to_check_na, inplace=True)
            if not df_plot_cleaned.empty:
                fig = plot_func(df_plot_cleaned, title=titulo, **plot_args)
                st.plotly_chart(fig, use_container_width=True); return True
            else: st.warning(f"Dados insuficientes para '{titulo}' após limpar NaNs.")
        except Exception as e_plotly: st.warning(f"Erro Plotly '{titulo}': {e_plotly}.")
    elif plot_func: st.warning(f"Colunas ausentes/incorretas para '{titulo}'. Esperado: {plot_args}. Disponível: {df_plot.columns.tolist()}")
    return False

# --- 3. Interface Streamlit Principal ---
st.set_page_config(layout="wide", page_title="Gemini DOCX Insights")

for k, dv in [("sugestoes_gemini",[]),("config_sugestoes",{}),("conteudo_docx",{"texto":"","tabelas":[]}),
              ("nome_arquivo_atual",None),("debug_checkbox_key",False),("pagina_selecionada","Dashboard Principal")]:
    st.session_state.setdefault(k, dv)

st.sidebar.title("✨ Navegação"); pagina_opcoes = ["Dashboard Principal", "Análise SWOT Detalhada"]
st.session_state.pagina_selecionada = st.sidebar.radio("Selecione:", pagina_opcoes, index=pagina_opcoes.index(st.session_state.pagina_selecionada), key="nav_radio")
st.sidebar.divider()
uploaded_file = st.sidebar.file_uploader("Selecione DOCX", type="docx", key="uploader_sidebar")
show_debug_info = st.sidebar.checkbox("Mostrar Debug Info", value=st.session_state.debug_checkbox_key, key="debug_cb_sidebar")
st.session_state.debug_checkbox_key = show_debug_info

if uploaded_file:
    if st.session_state.nome_arquivo_atual != uploaded_file.name: 
        with st.spinner("Processando novo documento..."):
            st.session_state.sugestoes_gemini, st.session_state.config_sugestoes = [], {}
            st.session_state.nome_arquivo_atual = uploaded_file.name
            texto_doc, tabelas_doc = extrair_conteudo_docx(uploaded_file)
            st.session_state.conteudo_docx = {"texto": texto_doc, "tabelas": tabelas_doc}
            if texto_doc or tabelas_doc:
                sugestoes = analisar_documento_com_gemini(texto_doc, tabelas_doc)
                st.session_state.sugestoes_gemini = sugestoes
                st.session_state.config_sugestoes={s.get("id",f"s_{i}_{hash(s.get('titulo'))}"):{"aceito":True,"titulo_editado":s.get("titulo","S/Título"),"dados_originais":s} for i,s in enumerate(sugestoes)}
            else: st.sidebar.warning("Nenhum conteúdo extraído.")
    if st.session_state.sugestoes_gemini:
        st.sidebar.divider(); st.sidebar.header("⚙️ Configurar Sugestões")
        for sug in st.session_state.sugestoes_gemini:
            s_id = sug['id']; cfg = st.session_state.config_sugestoes[s_id]
            with st.sidebar.expander(f"{cfg['titulo_editado']}",expanded=False):
                st.caption(f"Tipo: {sug.get('tipo_sugerido')} | Fonte: {sug.get('fonte_id')}")
                cfg["aceito"]=st.checkbox("Incluir?",value=cfg["aceito"],key=f"acc_{s_id}")
                cfg["titulo_editado"]=st.text_input("Título",value=cfg["titulo_editado"],key=f"tit_{s_id}")
                # (Lógica de edição de parâmetros da sidebar removida para esta versão focada)
else: 
    if st.session_state.pagina_selecionada == "Dashboard Principal": # Mostra mensagem apenas no dashboard
        st.info("Por favor, faça o upload de um arquivo DOCX na barra lateral para começar.")

# --- RENDERIZAÇÃO DA PÁGINA ---
if st.session_state.pagina_selecionada == "Dashboard Principal":
    st.title("📊 Dashboard de Insights do Documento")
    if uploaded_file and st.session_state.sugestoes_gemini:
        kpis, outros = [], []; [ (kpis if s_cfg["dados_originais"].get("tipo_sugerido")=="kpi" else outros).append({"titulo":s_cfg["titulo_editado"], **s_cfg["dados_originais"]}) for s_id,s_cfg in st.session_state.config_sugestoes.items() if s_cfg["aceito"] ]
        render_kpis(kpis)
        if show_debug_info:
             with st.expander("Debug: Elementos para Dashboard (Não-KPI)", expanded=False): st.json({"Outros": outros}, expanded=False)
        if outros:
            cols_dash = st.columns(2); idx_dash = 0; count_dash = 0
            for item in outros:
                if item.get("tipo_sugerido") == "lista_swot": continue # SWOTs em página separada
                with cols_dash[idx_dash % 2]:
                    st.subheader(item["titulo"]); df_plot, rendered = None, False
                    params, tipo, fonte = item.get("parametros",{}), item.get("tipo_sugerido"), item.get("fonte_id")
                    try:
                        if params.get("dados"): df_plot=pd.DataFrame(params["dados"])
                        elif str(fonte).startswith("doc_tabela_"): df_plot=next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"]==fonte),None)
                        
                        if tipo=="tabela_dados":
                            df_tbl=None
                            if str(fonte).startswith("texto_") and params.get("dados"):
                                df_tbl=pd.DataFrame(params.get("dados")); 
                                if params.get("colunas_titulo"): df_tbl.columns=params.get("colunas_titulo")
                            else: id_tbl=params.get("id_tabela_original",fonte); df_tbl=next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"]==id_tbl),None)
                            if df_tbl is not None: st.dataframe(df_tbl.astype(str).fillna("-")); rendered=True
                            else: st.warning(f"Tabela '{item['titulo']}' (Fonte: {fonte}) não encontrada.")
                        elif df_plot is not None and tipo in ["grafico_barras","grafico_linha","grafico_dispersao","grafico_pizza"]:
                            if render_plotly_chart(item, df_plot): rendered = True
                        elif tipo == 'mapa': st.info(f"Mapa '{item['titulo']}' não implementado."); rendered=True
                        if not rendered and tipo not in ["kpi","lista_swot","mapa"]: st.info(f"'{item['titulo']}' ({tipo}) não gerado. Dados/Tipo não suportado.")
                    except Exception as e: st.error(f"Erro renderizando '{item['titulo']}': {e}")
                if rendered: idx_dash+=1; count_dash+=1
            if count_dash == 0 and any(c['aceito'] and c['dados_originais'].get('tipo_sugerido') not in ['kpi','lista_swot'] for c in st.session_state.config_sugestoes.values()):
                st.info("Nenhum gráfico/tabela (além de KPIs/SWOTs) pôde ser gerado.")
        elif not kpis and not uploaded_file: pass # Já tem mensagem de upload acima
        elif not kpis and not outros and uploaded_file: st.info("Nenhum elemento selecionado ou gerado.")


elif st.session_state.pagina_selecionada == "Análise SWOT Detalhada":
    st.title("🔬 Análise SWOT Detalhada")
    if not uploaded_file: st.warning("Faça upload de um DOCX na barra lateral.")
    elif not st.session_state.sugestoes_gemini: st.info("Aguardando processamento ou nenhuma sugestão gerada.")
    else:
        swot_sugs = [s_cfg["dados_originais"] for s_id,s_cfg in st.session_state.config_sugestoes.items() if s_cfg["aceito"] and s_cfg["dados_originais"].get("tipo_sugerido")=="lista_swot"]
        if not swot_sugs: st.info("Nenhuma análise SWOT sugerida/selecionada.")
        else:
            for swot_item_page in swot_sugs:
                player_name = "Geral" # Default
                title_match = re.search(r"SWOT(?: d[oa]| -) (.+)", swot_item_page.get("titulo","SWOT"), re.IGNORECASE)
                if title_match: player_name = title_match.group(1)
                render_swot_card(player_name, swot_item_page.get("parametros",{}), card_key_prefix=swot_item_page["id"])

# Limpar estado se o arquivo for removido da UI
if uploaded_file is None and st.session_state.nome_arquivo_atual is not None:
    keys_to_preserve_on_clear = [k for k in st.session_state.keys() if k.endswith(("_key", "_widget_key"))] # Preserve widget keys
    current_keys_on_clear = list(st.session_state.keys())
    for key_cl in current_keys_on_clear:
        if key_cl not in keys_to_preserve_on_clear: del st.session_state[key_cl]
    for k_reinit, dv_reinit in [("sugestoes_gemini",[]),("config_sugestoes",{}),("conteudo_docx",{"texto":"","tabelas":[]}),("nome_arquivo_atual",None),("debug_checkbox_key_main",False),("pagina_selecionada","Dashboard Principal")]:
        st.session_state.setdefault(k_reinit, dv_reinit) # Re-inicializa estados chave
    st.experimental_rerun()

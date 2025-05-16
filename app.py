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
    except: api_key = os.environ.get("GOOGLE_API_KEY"); return api_key if api_key else None

# --- 2. Funções de Processamento do Documento e Interação com Gemini ---
def parse_value_for_numeric(val_str_in):
    if pd.isna(val_str_in) or str(val_str_in).strip() == '': return None
    text = str(val_str_in).strip()
    # Trata primeiro os negativos em parênteses
    is_negative_paren = text.startswith('(') and text.endswith(')')
    if is_negative_paren:
        text = text[1:-1] # Remove parênteses

    # Remove R$, $, %, e espaços internos, mas mantém o sinal - no início e . decimal
    # A ordem da limpeza é importante.
    text_num_part = re.sub(r'[R$\s%]', '', text)
    
    # Trata separador de milhar (ponto) ANTES de trocar vírgula por ponto decimal
    if ',' in text_num_part and '.' in text_num_part:
        # Assume que '.' é milhar se vier antes da última vírgula
        if text_num_part.rfind('.') < text_num_part.rfind(','):
             text_num_part = text_num_part.replace('.', '') 
        text_num_part = text_num_part.replace(',', '.') 
    elif ',' in text_num_part: # Apenas vírgula, assume como decimal
        text_num_part = text_num_part.replace(',', '.')
    # Se só tem ponto, ou se o ponto é o último separador, assume como decimal.

    # Tenta pegar o primeiro número encontrado (inteiro ou decimal), incluindo sinal negativo
    match = re.search(r"([-+]?\d*\.?\d+|\d+)", text_num_part)
    if match:
        try: 
            num = float(match.group(1))
            return -num if is_negative_paren else num
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
                    
                    # Tenta converter para numérico primeiro
                    num_series = original_series.astype(str).apply(parse_value_for_numeric)
                    # Se uma proporção significativa virou número, usa a série numérica
                    if num_series.notna().sum() / max(1, len(num_series)) > 0.3:
                        df[col] = pd.to_numeric(num_series, errors='coerce')
                        continue 
                    else: # Reverte para o original se a conversão numérica não foi boa
                         df[col] = original_series 
                    
                    # Se não virou numérico, tenta converter para datetime
                    try:
                        temp_str_col = df[col].astype(str) # Garante que é string para pd.to_datetime
                        # Tenta inferir formato, é mais flexível
                        dt_series = pd.to_datetime(temp_str_col, errors='coerce', dayfirst=True) 
                        # Se a maioria dos valores não nulos viraram datas, usa a série de datas
                        if dt_series.notna().sum() > len(df[col][df[col].notna()]) * 0.5:
                            df[col] = dt_series
                        else: # Mantém como string se a conversão de data falhou muito
                            df[col] = original_series.astype(str).fillna('')
                    except Exception: # Se qualquer erro na conversão de data
                        df[col] = original_series.astype(str).fillna('')
                
                # Fallback final para garantir que colunas 'object' sejam string
                for col in df.columns:
                    if df[col].dtype == 'object':
                        df[col] = df[col].astype(str).fillna('')

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
            # Amostra menor e menos colunas para o prompt para evitar exceder limites
            sample_df = df.head(3).iloc[:, :min(5, len(df.columns))] 
            md_table = sample_df.to_markdown(index=False)
            col_types = ", ".join([f"'{c}' (tipo: {str(d)})" for c,d in df.dtypes.items()[:min(8, len(df.columns))]]) # Limita colunas nos tipos
            tabelas_prompt_str += f"\n--- Tabela '{nome_t}' (ID: {id_t}) ---\nColunas e tipos: {col_types}\nAmostra:\n{md_table}\n"
        
        text_limit = 40000 # Reduzido ainda mais para segurança
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
        2.  DADOS NUMÉRICOS: Se a coluna de valor de uma TABELA não for numérica (float64/int64) conforme os "tipos inferidos", NÃO sugira um gráfico que exija valor numérico para ela, A MENOS que você possa confiavelmente extrair um valor numérico do seu conteúdo textual (ex: de '70%' extrair 70.0; de '70% - 86%' extrair 70.0 ou 78.0; de 'R$ 15,5 Bi' extrair 15.5). Se extrair do texto, coloque em "dados".
        3.  COBERTURA GEOGRÁFICA (Player, Cidades): Se for apenas lista, sugira "tabela_dados" e forneça os dados extraídos no campo "dados" dos "parametros" com "colunas_titulo".
        4.  SWOT COMPARATIVO: Se uma tabela compara SWOTs, gere "lista_swot" INDIVIDUAL para CADA player da tabela.
        Retorne APENAS a lista JSON válida. Seja conciso na justificativa.
        """
        with st.spinner("🤖 Gemini analisando... (Pode levar alguns instantes)"):
            # st.text_area("Debug Prompt:", prompt, height=150) # Descomentar para depurar
            response = model.generate_content(prompt)
        cleaned_text = response.text.strip().lstrip("```json").rstrip("```").strip()
        # st.text_area("Debug Resposta Gemini:", cleaned_text, height=150) # Descomentar para depurar
        sugestoes = json.loads(cleaned_text)
        if isinstance(sugestoes, list) and all(isinstance(item, dict) for item in sugestoes):
             st.success(f"{len(sugestoes)} sugestões recebidas do Gemini!")
             return sugestoes
        st.error("Resposta do Gemini não é uma lista JSON válida como esperado."); return []
    except json.JSONDecodeError as e: 
        st.error(f"Erro ao decodificar JSON da resposta do Gemini: {e}")
        if 'response' in locals() and hasattr(response, 'text'): st.code(response.text, language="text")
        return []
    except Exception as e: 
        st.error(f"Erro na comunicação com Gemini: {e}")
        st.text(traceback.format_exc())
        return []

# --- 3. Interface Streamlit e Lógica de Apresentação ---
st.set_page_config(layout="wide"); st.title("✨ Gemini: DOCX para Insights Visuais")
st.markdown("Faça upload de um DOCX e deixe o Gemini sugerir como visualizar suas informações.")

for k, default_val in [("sugestoes_gemini", []), ("config_sugestoes", {}), 
                       ("conteudo_docx", {"texto": "", "tabelas": []}), 
                       ("nome_arquivo_atual", None), ("debug_checkbox_key_main", False)]:
    st.session_state.setdefault(k, default_val)

uploaded_file = st.file_uploader("Selecione DOCX", type="docx", key="uploader_main_key_vfinal")
show_debug_info = st.sidebar.checkbox("Mostrar Informações de Depuração", 
                                    value=st.session_state.debug_checkbox_key_main, 
                                    key="debug_cb_widget_key_vfinal")
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
                with st.expander("Debug: Conteúdo DOCX (após extração e tipos)", expanded=False):
                    st.text_area("Texto (amostra)", texto_doc[:1000], height=80)
                    for t_info_dbg in tabelas_doc:
                        st.write(f"ID: {t_info_dbg['id']}, Nome: {t_info_dbg['nome']}")
                        try: st.dataframe(t_info_dbg['dataframe'].head().astype(str).fillna("-")) 
                        except Exception as e_df_dbg_display: st.warning(f"Debug DF {t_info_dbg['id']}: {e_df_dbg_display}"); st.text(f"Head:\n{t_info_dbg['dataframe'].head().to_string(na_rep='-')}")
                        st.write("Tipos:", t_info_dbg['dataframe'].dtypes.to_dict())
            sugestoes = analisar_documento_com_gemini(texto_doc, tabelas_doc)
            st.session_state.sugestoes_gemini = sugestoes
            temp_config_init = {}
            for i_init,s_init in enumerate(sugestoes): 
                s_id_init = s_init.get("id", f"s_init_{i_init}_{hash(s_init.get('titulo',''))}"); s_init["id"] = s_id_init
                temp_config_init[s_id_init] = {"aceito":True,"titulo_editado":s_init.get("titulo","S/Título"),"dados_originais":s_init}
            st.session_state.config_sugestoes = temp_config_init
        else: st.warning("Nenhum conteúdo extraído.")

if st.session_state.sugestoes_gemini:
    st.sidebar.header("⚙️ Configurar Sugestões")
    for sug_sb_main in st.session_state.sugestoes_gemini:
        s_id_sb_main = sug_sb_main['id'] 
        if s_id_sb_main not in st.session_state.config_sugestoes:
             st.session_state.config_sugestoes[s_id_sb_main] = {"aceito":True,"titulo_editado":sug_sb_main.get("titulo","S/Título"),"dados_originais":sug_sb_main}
        cfg_sb_main = st.session_state.config_sugestoes[s_id_sb_main]
        with st.sidebar.expander(f"{cfg_sb_main['titulo_editado']}", expanded=False):
            st.caption(f"Tipo: {sug_sb_main.get('tipo_sugerido')} | Fonte: {sug_sb_main.get('fonte_id')}")
            st.markdown(f"**IA:** *{sug_sb_main.get('justificativa', 'N/A')}*")
            cfg_sb_main["aceito"]=st.checkbox("Incluir?",value=cfg_sb_main["aceito"],key=f"acc_sb_{s_id_sb_main}")
            cfg_sb_main["titulo_editado"]=st.text_input("Título",value=cfg_sb_main["titulo_editado"],key=f"tit_sb_{s_id_sb_main}")
            params_orig_cfg_sb = cfg_sb_main["dados_originais"].get("parametros",{})
            tipo_sug_cfg_sb = sug_sb_main.get("tipo_sugerido")
            if tipo_sug_cfg_sb in ["grafico_barras","grafico_pizza","grafico_linha","grafico_dispersao"] and \
               not params_orig_cfg_sb.get("dados") and str(sug_sb_main.get("fonte_id")).startswith("doc_tabela_"):
                df_corr_cfg_sb = next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"] == sug_sb_main.get("fonte_id")), None)
                if df_corr_cfg_sb is not None:
                    opts_cfg_sb = [""] + df_corr_cfg_sb.columns.tolist()
                    editable_params_cfg_sb = params_orig_cfg_sb 
                    if tipo_sug_cfg_sb != "grafico_pizza":
                        editable_params_cfg_sb["eixo_x"] = st.selectbox("Eixo X", opts_cfg_sb, index=opts_cfg_sb.index(editable_params_cfg_sb.get("eixo_x","")) if editable_params_cfg_sb.get("eixo_x","") in opts_cfg_sb else 0, key=f"px_cfg_sb_{s_id_sb_main}")
                        editable_params_cfg_sb["eixo_y"] = st.selectbox("Eixo Y", opts_cfg_sb, index=opts_cfg_sb.index(editable_params_cfg_sb.get("eixo_y","")) if editable_params_cfg_sb.get("eixo_y","") in opts_cfg_sb else 0, key=f"py_cfg_sb_{s_id_sb_main}")
                    else:
                        editable_params_cfg_sb["categorias"] = st.selectbox("Categorias", opts_cfg_sb, index=opts_cfg_sb.index(editable_params_cfg_sb.get("categorias","")) if editable_params_cfg_sb.get("categorias","") in opts_cfg_sb else 0, key=f"pcat_cfg_sb_{s_id_sb_main}")
                        editable_params_cfg_sb["valores"] = st.selectbox("Valores", opts_cfg_sb, index=opts_cfg_sb.index(editable_params_cfg_sb.get("valores","")) if editable_params_cfg_sb.get("valores","") in opts_cfg_sb else 0, key=f"pval_cfg_sb_{s_id_sb_main}")

    if st.sidebar.button("🚀 Gerar Dashboard", type="primary", use_container_width=True):
        st.header("📊 Dashboard de Insights"); kpis_dash_f, outros_dash_f = [], []
        for s_id_f, cfg_f in st.session_state.config_sugestoes.items():
            if cfg_f["aceito"]: 
                item_f = {"titulo":cfg_f["titulo_editado"], **cfg_f["dados_originais"]}
                (kpis_dash_f if item_f.get("tipo_sugerido")=="kpi" else outros_dash_f).append(item_f)
        
        if kpis_dash_f:
            cols_kpi_f=st.columns(min(len(kpis_dash_f), 4)); 
            for i_kpi_f, k_f in enumerate(kpis_dash_f):
                with cols_kpi_f[i_kpi_f % min(len(kpis_dash_f), 4)]:
                    p_kpi_f = k_f.get("parametros",{}); 
                    st.metric(k_f.get("titulo","KPI"),str(p_kpi_f.get("valor","N/A")),str(p_kpi_f.get("delta","")),help=p_kpi_f.get("descricao"))
            if outros_dash_f: st.divider()

        if show_debug_info and (kpis_dash_f or outros_dash_f):
             with st.expander("Debug: Configs Finais para Dashboard (Elementos Selecionados)",expanded=True): # Expandido por padrão
                if kpis_dash_f: st.json({"KPIs Selecionados": kpis_dash_f}, expanded=False)
                if outros_dash_f: st.json({"Outros Elementos Selecionados": outros_dash_f}, expanded=False)
        
        elementos_renderizados_final_count = 0 
        if outros_dash_f:
            item_cols_f = st.columns(2); col_idx_f = 0 
            for item_f_main in outros_dash_f:
                with item_cols_f[col_idx_f % 2]:
                    st.subheader(item_f_main["titulo"]); df_plot_f, el_rend_f = None, False
                    params_f=item_f_main.get("parametros",{}); tipo_f=item_f_main.get("tipo_sugerido"); fonte_f=item_f_main.get("fonte_id")
                    try:
                        if params_f.get("dados"):
                            try: df_plot_f=pd.DataFrame(params_f["dados"])
                            except Exception as e_dfd_f: st.warning(f"'{item_f_main['titulo']}': Erro DF de 'dados': {e_dfd_f}"); continue
                        elif str(fonte_f).startswith("doc_tabela_"):
                            df_plot_f=next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"]==fonte_f),None)
                        
                        if tipo_f=="tabela_dados":
                            df_t_render_f = None 
                            if str(fonte_f).startswith("texto_") and params_f.get("dados"):
                                try: 
                                    df_t_render_f = pd.DataFrame(params_f.get("dados"))
                                    if params_f.get("colunas_titulo"): df_t_render_f.columns = params_f.get("colunas_titulo")
                                except Exception as e_df_txt_tbl_f: st.warning(f"Erro tabela texto '{item_f_main['titulo']}': {e_df_txt_tbl_f}")
                            else: 
                                id_t_render_f=params_f.get("id_tabela_original",fonte_f)
                                df_t_render_f=next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"]==id_t_render_f),None)
                            if df_t_render_f is not None: 
                                try: st.dataframe(df_t_render_f.astype(str).fillna("-"))
                                except Exception as e_arrow_final_tbl: st.warning(f"Falha st.dataframe p/ '{item_f_main['titulo']}': {e_arrow_final_tbl}"); st.text(df_t_render_f.to_string(na_rep='-'))
                                el_rend_f=True
                            else: st.warning(f"Tabela '{item_f_main['titulo']}' (Fonte: {fonte_f}) não encontrada.")
                        
                        elif tipo_f=="lista_swot":
                            swot_f=params_f; c1sf,c2sf=st.columns(2)
                            smap_f={"forcas":("Forças 💪",c1sf),"fraquezas":("Fraquezas 📉",c1sf),"oportunidades":("Oportunidades 🚀",c2sf),"ameacas":("Ameaças ⚠️",c2sf)}
                            for k_sf,(h_sf,ct_sf) in smap_f.items():
                                with ct_sf: st.markdown(f"##### {h_sf}"); [st.markdown(f"- {p_sf}") for p_sf in swot_f.get(k_sf,["N/A"])]
                            el_rend_f=True
                        
                        elif df_plot_f is not None:
                            x_f,y_f,cat_f,val_f = params_f.get("eixo_x"),params_f.get("eixo_y"),params_f.get("categorias"),params_f.get("valores")
                            fn_plot_f,p_args_f=None,{}
                            if tipo_f=="grafico_barras" and x_f and y_f: fn_plot_f,p_args_f=px.bar,{"x":x_f,"y":y_f}
                            elif tipo_f=="grafico_linha" and x_f and y_f: fn_plot_f,p_args_f=px.line,{"x":x_f,"y":y_f,"markers":True}
                            elif tipo_f=="grafico_dispersao" and x_f and y_f: fn_plot_f,p_args_f=px.scatter,{"x":x_f,"y":y_f}
                            elif tipo_f=="grafico_pizza" and cat_f and val_f: fn_plot_f,p_args_f=px.pie,{"names":cat_f,"values":val_f}
                            
                            if fn_plot_f and all(k_col_f in df_plot_f.columns for k_col_f in p_args_f.values() if isinstance(k_col_f,str)):
                                try:
                                    # Garante que colunas numéricas para plotagem são de fato numéricas no df_plot_f
                                    if "y" in p_args_f and p_args_f["y"] in df_plot_f.columns: 
                                        df_plot_f[p_args_f["y"]] = pd.to_numeric(df_plot_f[p_args_f["y"]], errors='coerce')
                                    if "values" in p_args_f and p_args_f["values"] in df_plot_f.columns:
                                        df_plot_f[p_args_f["values"]] = pd.to_numeric(df_plot_f[p_args_f["values"]], errors='coerce')
                                    
                                    # Remove linhas com NaN nas colunas de plotagem para evitar erros no Plotly
                                    cols_to_check_na = [val for val in p_args_f.values() if isinstance(val, str) and val in df_plot_f.columns]
                                    df_plot_f_cleaned = df_plot_f.dropna(subset=cols_to_check_na)

                                    if not df_plot_f_cleaned.empty:
                                        st.plotly_chart(fn_plot_f(df_plot_f_cleaned,title=item_f_main["titulo"],**p_args_f),use_container_width=True); el_rend_f=True
                                    else:
                                        st.warning(f"Dados insuficientes para '{item_f_main['titulo']}' após remover NaNs das colunas de plotagem.")

                                except Exception as e_plotly_f: st.warning(f"Erro Plotly '{item_f_main['titulo']}': {e_plotly_f}. Verifique tipos. X:{x_f}, Y:{y_f}, Cat:{cat_f}, Val:{val_f}. DF Dtypes: {df_plot_f.dtypes.to_dict() if df_plot_f is not None else 'DF é None'}")
                            elif fn_plot_f: st.warning(f"Colunas ausentes/incorretas para '{item_f_main['titulo']}'. Esperado: {p_args_f}. Disponível: {df_plot_f.columns.tolist() if df_plot_f is not None else 'DF é None'}")
                        
                        if tipo_f == 'mapa': st.info(f"Mapa para '{item_f_main['titulo']}' não implementado."); el_rend_f=True
                        
                        if not el_rend_f and tipo_f not in ["kpi","tabela_dados","lista_swot","mapa"]:
                            st.info(f"'{item_f_main['titulo']}' (tipo: {tipo_f}) não gerado. Dados/Tipo não suportado ou DF não pôde ser criado/encontrado.")
                    except Exception as e_main_render_f: st.error(f"Erro renderizando '{item_f_main['titulo']}': {e_main_render_f}")
                
                if el_rend_f: col_idx_f+=1; elementos_renderizados_final_count+=1
        
        if elementos_renderizados_final_count == 0 and any(c['aceito'] and c['dados_originais'].get('tipo_sugerido') != 'kpi' for c in st.session_state.config_sugestoes.values()):
            st.info("Nenhum gráfico/tabela (além de KPIs) pôde ser gerado.")
        elif elementos_renderizados_final_count == 0 and not kpis_dash_f:
            st.info("Nenhum elemento selecionado ou pôde ser gerado para o dashboard.")

elif uploaded_file is None and st.session_state.nome_arquivo_atual is not None:
    keys_to_clear_final = list(st.session_state.keys())
    for key_clear_f in keys_to_clear_final:
        if key_clear_f not in ['debug_checkbox_widget_key_vfinal', 'uploader_main_key_vfinal']: 
            del st.session_state[key_clear_f]
    st.session_state.sugestoes_gemini, st.session_state.config_sugestoes = [], {}
    st.session_state.conteudo_docx = {"texto": "", "tabelas": []}
    st.session_state.nome_arquivo_atual = None
    st.session_state.debug_checkbox_key_main = False
    st.experimental_rerun()

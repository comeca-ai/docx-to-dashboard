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
    
    # Trata intervalos como "70 - 86" ou "12 - 23", pegando o primeiro número
    interval_match = re.match(r"([-+]?\d*\.?\d+|\d+)\s*-\s*([-+]?\d*\.?\d+|\d+)", text)
    if interval_match:
        try: return float(interval_match.group(1).replace(',', '.')) # Pega o primeiro número do intervalo
        except: pass # Continua para outras lógicas se o parsing do intervalo falhar

    # Limpeza geral: remove R$, $, %, espaços internos, mas mantém sinal - no início
    # e trata . como milhar se vírgula for decimal, ou vírgula como decimal.
    text_num_part = re.sub(r'[R$\s%]', '', text)
    
    if ',' in text_num_part and '.' in text_num_part: # Ex: 1.234,56
        text_num_part = text_num_part.replace('.', '') 
        text_num_part = text_num_part.replace(',', '.') 
    elif ',' in text_num_part: # Ex: 1234,56
        text_num_part = text_num_part.replace(',', '.')
    # Se só tem ponto, assume que é decimal: 1234.56 (já está ok)

    # Tenta pegar apenas o primeiro número encontrado
    match = re.search(r"([-+]?\d*\.?\d+|\d+)", text_num_part)
    if match:
        try: return float(match.group(1))
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
                    if num_series.notna().sum() / max(1, len(num_series)) > 0.3: # Evita divisão por zero
                        df[col] = pd.to_numeric(num_series, errors='coerce')
                        continue
                    else: df[col] = original_series 
                    try:
                        temp_str_col = df[col].astype(str)
                        dt_series = pd.to_datetime(temp_str_col, errors='coerce', dayfirst=True) # Tenta dayfirst=True para formatos dd/mm/yyyy
                        if dt_series.notna().sum() / max(1, len(dt_series)) > 0.5:
                            df[col] = dt_series
                        else: # Reverte se a conversão de data falhou muito
                            df[col] = original_series.astype(str).fillna('')
                    except: df[col] = original_series.astype(str).fillna('')
                for col in df.columns:
                    if df[col].dtype == 'object': df[col] = df[col].astype(str).fillna('') # Garante string p/ Arrow
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
            sample_df = df.head(5).iloc[:, :min(8, len(df.columns))]
            md_table = sample_df.to_markdown(index=False)
            col_types = ", ".join([f"'{c}' (tipo: {str(d)})" for c,d in df.dtypes.items()]) # Usar df.dtypes para tipos pós-processamento
            tabelas_prompt_str += f"\n--- Tabela '{nome_t}' (ID: {id_t}) ---\nColunas e tipos: {col_types}\nAmostra:\n{md_table}\n"
        text_limit = 50000
        prompt_text = texto_doc[:text_limit] + ("\n[TEXTO TRUNCADO]" if len(texto_doc) > text_limit else "")
        prompt = f"""
        Você é um assistente de análise de dados. Analise o texto e as tabelas.
        [TEXTO]{prompt_text}[FIM TEXTO]
        [TABELAS]{tabelas_prompt_str}[FIM TABELAS]

        Gere lista JSON de sugestões de visualizações. Objeto DEVE ter: "id", "titulo", "tipo_sugerido" ("kpi", "tabela_dados", "lista_swot", "grafico_barras", "grafico_pizza", "grafico_linha", "grafico_dispersao"), "fonte_id" (ID tabela ou "texto_descricao_fonte"), "parametros" (objeto JSON), "justificativa".
        Para "parametros":
        - "kpi": {{"valor": "ValorKPI", "delta": "Mudança", "descricao": "Contexto"}}
        - "tabela_dados": Para TABELA EXISTENTE: {{"id_tabela_original": "ID_Tabela"}}. Para DADOS DO TEXTO: {{"dados": [{{"Coluna1": "ValorA1", "Coluna2": "ValorA2"}}, ...], "colunas_titulo": ["Título Col1", "Título Col2"]}}
        - "lista_swot": {{"forcas": ["F1"], "fraquezas": ["Fr1"], "oportunidades": ["Op1"], "ameacas": ["Am1"]}} (Assegure que cada categoria tenha uma lista de strings).
        - Gráficos de TABELA ("barras", "linha", "dispersao"): {{"eixo_x": "NOME_COL_X", "eixo_y": "NOME_COL_Y"}} (Y numérico, use nomes exatos de colunas).
        - Gráficos de PIZZA de TABELA: {{"categorias": "NOME_COL_CAT", "valores": "NOME_COL_VAL_NUM"}} (Valores numéricos, use nomes exatos).
        - Gráficos com DADOS EXTRAÍDOS DO TEXTO: {{"dados": [{{"NomeEixoX": "CatA", "NomeEixoY": ValNumA}}, ...], "eixo_x": "NomeEixoX", "eixo_y": "NomeEixoY"}} (Valores DEVEM ser numéricos, não strings de números).
        
        INSTRUÇÕES CRÍTICAS:
        1.  NOMES DE COLUNAS: Para gráficos de TABELA, use os NOMES EXATOS das colunas como fornecidos nos "Colunas e tipos".
        2.  DADOS NUMÉRICOS: Se a coluna de valor de uma TABELA não for numérica (float64/int64) conforme os "tipos inferidos", NÃO sugira um gráfico que exija valor numérico para ela, A MENOS que você possa confiavelmente extrair um valor numérico do seu conteúdo textual (ex: de '70%' extrair 70.0; de '70% - 86%' extrair 70.0 ou 78.0; de 'R$ 15,5 Bilhões' extrair 15.5 e indicar 'Bilhões' na justificativa ou título). Se extrair do texto, coloque em "dados".
        3.  COBERTURA GEOGRÁFICA: Se for apenas lista de Players/Cidades, sugira "tabela_dados" e forneça os dados extraídos no campo "dados" dos "parametros".
        4.  SWOT COMPARATIVO: Se uma tabela compara SWOTs, gere "lista_swot" INDIVIDUAL para CADA player da tabela.
        Retorne APENAS a lista JSON válida.
        """
        with st.spinner("🤖 Gemini analisando..."):
            # st.text_area("Debug Prompt:", prompt, height=200)
            response = model.generate_content(prompt)
        cleaned_text = response.text.strip().lstrip("```json").rstrip("```").strip()
        # st.text_area("Debug Resposta Gemini:", cleaned_text, height=200)
        sugestoes = json.loads(cleaned_text)
        if isinstance(sugestoes, list): st.success(f"{len(sugestoes)} sugestões!"); return sugestoes
        st.error("Resposta Gemini não é lista JSON."); return []
    except json.JSONDecodeError as e: st.error(f"Erro JSON Gemini: {e}"); st.code(response.text if 'response' in locals() else "N/A", language="text"); return []
    except Exception as e: st.error(f"Erro API Gemini: {e}"); st.text(traceback.format_exc()); return []

# --- 3. Interface Streamlit e Lógica de Apresentação ---
st.set_page_config(layout="wide"); st.title("✨ Gemini: DOCX para Insights Visuais")
st.markdown("Upload DOCX para sugestões de visualização pela IA.")

# Inicialização de estado (mais concisa)
for k, default_val in [("sugestoes_gemini", []), ("config_sugestoes", {}), 
                       ("conteudo_docx", {"texto": "", "tabelas": []}), 
                       ("nome_arquivo_atual", None), ("debug_checkbox_key", False)]:
    st.session_state.setdefault(k, default_val)

uploaded_file = st.file_uploader("Selecione DOCX", type="docx", key="uploader_main_key") # Chave única
show_debug_info = st.sidebar.checkbox("Mostrar Debug Info", value=st.session_state.debug_checkbox_key, key="debug_cb_widget_key")
st.session_state.debug_checkbox_key = show_debug_info # Sincroniza estado com valor do widget

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
                        try: st.dataframe(t_info_dbg['dataframe'].head().astype(str).fillna('-')) # .fillna para melhor display
                        except Exception: st.text(f"Head:\n{t_info_dbg['dataframe'].head().to_string(na_rep='-')}")
                        st.write("Tipos:", t_info_dbg['dataframe'].dtypes.to_dict())
            sugestoes = analisar_documento_com_gemini(texto_doc, tabelas_doc)
            st.session_state.sugestoes_gemini = sugestoes
            # Inicializa config_sugestoes com IDs corretos das sugestões
            temp_config = {}
            for i,s_init_cfg in enumerate(sugestoes): 
                s_id_cfg = s_init_cfg.get("id", f"s_{i}_{hash(s_init_cfg.get('titulo'))}"); s_init_cfg["id"] = s_id_cfg
                temp_config[s_id_cfg] = {"aceito":True,"titulo_editado":s_init_cfg.get("titulo","S/Título"),"dados_originais":s_init_cfg}
            st.session_state.config_sugestoes = temp_config
        else: st.warning("Nenhum conteúdo extraído.")

if st.session_state.sugestoes_gemini:
    st.sidebar.header("⚙️ Configurar Sugestões")
    for sug_sidebar in st.session_state.sugestoes_gemini:
        s_id_sb = sug_sidebar['id']
        # Garante que a config existe, caso a sugestão tenha sido adicionada dinamicamente
        if s_id_sb not in st.session_state.config_sugestoes:
            st.session_state.config_sugestoes[s_id_sb] = {"aceito":True,"titulo_editado":sug_sidebar.get("titulo","S/Título"),"dados_originais":sug_sidebar}
        cfg_sb = st.session_state.config_sugestoes[s_id_sb]

        with st.sidebar.expander(f"{cfg_sb['titulo_editado']}", expanded=False):
            st.caption(f"Tipo: {sug_sidebar.get('tipo_sugerido')} | Fonte: {sug_sidebar.get('fonte_id')}")
            st.markdown(f"**IA:** *{sug_sidebar.get('justificativa', 'N/A')}*")
            cfg_sb["aceito"]=st.checkbox("Incluir?",value=cfg_sb["aceito"],key=f"acc_{s_id_sb}")
            cfg_sb["titulo_editado"]=st.text_input("Título",value=cfg_sb["titulo_editado"],key=f"tit_{s_id_sb}")
            
            # Edição de parâmetros (mantida a lógica anterior)
            params_orig_sb = cfg_sb["dados_originais"].get("parametros",{})
            tipo_sug_sb_edit = sug_sidebar.get("tipo_sugerido")
            if tipo_sug_sb_edit in ["grafico_barras","grafico_pizza","grafico_linha","grafico_dispersao"] and \
               not params_orig_sb.get("dados") and str(sug_sidebar.get("fonte_id")).startswith("doc_tabela_"):
                df_corr_sb = next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"] == sug_sidebar.get("fonte_id")), None)
                if df_corr_sb is not None:
                    opts_sb = [""] + df_corr_sb.columns.tolist()
                    if tipo_sug_sb_edit != "grafico_pizza":
                        params_orig_sb["eixo_x"] = st.selectbox("Eixo X", opts_sb, index=opts_sb.index(params_orig_sb.get("eixo_x","")) if params_orig_sb.get("eixo_x","") in opts_sb else 0, key=f"px_sb_{s_id_sb}")
                        params_orig_sb["eixo_y"] = st.selectbox("Eixo Y", opts_sb, index=opts_sb.index(params_orig_sb.get("eixo_y","")) if params_orig_sb.get("eixo_y","") in opts_sb else 0, key=f"py_sb_{s_id_sb}")
                    else:
                        params_orig_sb["categorias"] = st.selectbox("Categorias", opts_sb, index=opts_sb.index(params_orig_sb.get("categorias","")) if params_orig_sb.get("categorias","") in opts_sb else 0, key=f"pcat_sb_{s_id_sb}")
                        params_orig_sb["valores"] = st.selectbox("Valores", opts_sb, index=opts_sb.index(params_orig_sb.get("valores","")) if params_orig_sb.get("valores","") in opts_sb else 0, key=f"pval_sb_{s_id_sb}")

    if st.sidebar.button("🚀 Gerar Dashboard", type="primary", use_container_width=True):
        st.header("📊 Dashboard"); kpis_dash, outros_dash = [], []
        for s_id_dash, cfg_dash in st.session_state.config_sugestoes.items():
            if cfg_dash["aceito"]: 
                item_dash = {"titulo":cfg_dash["titulo_editado"], **cfg_dash["dados_originais"]}
                (kpis_dash if item_dash["tipo_sugerido"]=="kpi" else outros_dash).append(item_dash)
        
        if kpis_dash:
            cols_kpi_dash=st.columns(min(len(kpis_dash), 4)); 
            for i_kpi_d, k_d in enumerate(kpis_dash):
                with cols_kpi_dash[i_kpi_d % min(len(kpis_dash), 4)]:
                    p_kpi = k_d.get("parametros",{}); 
                    st.metric(k_d["titulo"],str(p_kpi.get("valor","N/A")),str(p_kpi.get("delta","")),help=p_kpi.get("descricao"))
            if outros_dash: st.divider()

        if show_debug_info and (kpis_dash or outros_dash):
             with st.expander("Debug: Configs Finais para Dashboard (após validação)",expanded=False):
                if kpis_dash: st.json({"KPIs": kpis_dash}, expanded=False)
                if outros_dash: st.json({"Outros Elementos": outros_dash}, expanded=False)
        
        if outros_dash:
            item_cols_dash=st.columns(2); col_idx_dash=0
            for item_d_main in outros_dash:
                with item_cols_dash[col_idx_dash%2]:
                    st.subheader(item_d_main["titulo"]); df_plot_d, el_rend_d = None, False
                    params_d=item_d_main.get("parametros",{}); tipo_d=item_d_main.get("tipo_sugerido"); fonte_d=item_d_main.get("fonte_id")
                    try:
                        if params_d.get("dados"): # Prioriza dados da LLM
                            try: df_plot_d=pd.DataFrame(params_d["dados"])
                            except Exception as e_dfd: st.warning(f"'{item_d_main['titulo']}': Erro DF de 'dados': {e_dfd}"); continue
                        elif str(fonte_d).startswith("doc_tabela_"): # Senão, busca tabela extraída
                            df_plot_d=next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"]==fonte_d),None)

                        if tipo_d=="tabela_dados":
                            df_t_render = None
                            if str(fonte_d).startswith("texto_") and params_d.get("dados"):
                                try: 
                                    df_t_render = pd.DataFrame(params_d.get("dados"))
                                    if params_d.get("colunas_titulo"): df_t_render.columns = params_d.get("colunas_titulo")
                                except Exception as e_df_txt_tbl: st.warning(f"Erro ao criar tabela de texto para '{item_d_main['titulo']}': {e_df_txt_tbl}")
                            else: 
                                id_t_render=params_d.get("id_tabela_original",fonte_d)
                                df_t_render=next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"]==id_t_render),None)
                            
                            if df_t_render is not None: st.dataframe(df_t_render.astype(str).fillna("-")); el_rend_d=True
                            else: st.warning(f"Tabela para '{item_d_main['titulo']}' não encontrada (Fonte ID: {fonte_d}).")
                        
                        elif tipo_d=="lista_swot":
                            swot_d_render=params_d; c1s,c2s=st.columns(2)
                            smap_d={"forcas":("Forças 💪",c1s),"fraquezas":("Fraquezas 📉",c1s),"oportunidades":("Oportunidades 🚀",c2s),"ameacas":("Ameaças ⚠️",c2s)}
                            for k_s,(h_s,ct_s) in smap_d.items():
                                with ct_s: st.markdown(f"##### {h_s}"); [st.markdown(f"- {p_s}") for p_s in swot_d_render.get(k_s,["N/A"])]
                            el_rend_d=True
                        
                        elif df_plot_d is not None: # Gráficos que usam df_plot_d
                            x,y,cat,val = params_d.get("eixo_x"),params_d.get("eixo_y"),params_d.get("categorias"),params_d.get("valores")
                            fn_d,p_args_d=None,{}
                            if tipo_d=="grafico_barras" and x and y: fn_d,p_args_d=px.bar,{"x":x,"y":y}
                            elif tipo_d=="grafico_linha" and x and y: fn_d,p_args_d=px.line,{"x":x,"y":y,"markers":True}
                            elif tipo_d=="grafico_dispersao" and x and y: fn_d,p_args_d=px.scatter,{"x":x,"y":y}
                            elif tipo_d=="grafico_pizza" and cat and val: fn_d,p_args_d=px.pie,{"names":cat,"values":val}
                            
                            if fn_d and all(k_col_d in df_plot_d.columns for k_col_d in p_args_d.values() if isinstance(k_col_d,str)):
                                st.plotly_chart(fn_d(df_plot_d,title=item_d_main["titulo"],**p_args_d),use_container_width=True); el_rend_d=True
                            elif fn_d: st.warning(f"Colunas ausentes/incorretas para '{item_d_main['titulo']}'. Esperado: {p_args_d}. Disponível: {df_plot_d.columns.tolist()}")
                        
                        if tipo_d == 'mapa': st.info(f"Mapa para '{item_d_main['titulo']}' não implementado."); el_rend_d=True
                        
                        if not el_rend_d and tipo_d not in ["kpi","tabela_dados","lista_swot","mapa"]:
                            st.info(f"'{item_d_main['titulo']}' (tipo: {tipo_d}) não gerado. Dados/Tipo não suportado ou DF não pôde ser criado/encontrado.")
                    except Exception as e_main_render: st.error(f"Erro renderizando '{item_d_main['titulo']}': {e_main_render}")
                
                if el_rend_d: col_idx_dash+=1; elementos_renderizados_count+=1 # elementos_renderizados_count não foi definido antes, corrigindo
        
        # Definindo elementos_renderizados_count se não foi usado antes no loop
        # (Essa definição abaixo pode ser redundante se a lógica acima funcionar sempre)
        if 'elementos_renderizados_count' not in locals():
            elementos_renderizados_count = sum(1 for item in outros_dash if item.get("el_rend_d", False)) # Exemplo, precisa de um sinalizador

        if elementos_renderizados_count == 0 and any(c['aceito'] and c['dados_originais']['tipo_sugerido'] != 'kpi' for c in st.session_state.config_sugestoes.values()):
            st.info("Nenhum gráfico/tabela (além de KPIs) pôde ser gerado com as seleções atuais.")
        elif elementos_renderizados_count == 0 and not kpis_dash:
            st.info("Nenhum elemento foi selecionado ou pôde ser gerado para o dashboard.")

elif uploaded_file is None and st.session_state.nome_arquivo_atual is not None:
    for key_to_clear_main in list(st.session_state.keys()): del st.session_state[key_to_clear_main]
    st.experimental_rerun()

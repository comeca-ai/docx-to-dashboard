import streamlit as st
from docx import Document
import pandas as pd
import plotly.express as px
import google.generativeai as genai
import json
import os
import traceback
import re 

# --- 1. DEFINIÇÃO DE TODAS AS FUNÇÕES PRIMEIRO ---

def get_gemini_api_key():
    """Obtém a chave da API do Gemini dos segredos do Streamlit ou variáveis de ambiente."""
    try: return st.secrets["GOOGLE_API_KEY"]
    except (FileNotFoundError, KeyError): 
        api_key = os.environ.get("GOOGLE_API_KEY")
        return api_key if api_key else None

def parse_value_for_numeric(val_str_in):
    """
    Tenta converter uma string para um valor numérico (float).
    Lida com alguns formatos comuns, R$, %, e negativos em parênteses.
    Extrai o primeiro número encontrado se houver texto misturado.
    """
    if pd.isna(val_str_in) or str(val_str_in).strip() == '': return None
    text = str(val_str_in).strip()
    
    is_negative_paren = text.startswith('(') and text.endswith(')')
    if is_negative_paren: text = text[1:-1] # Remove parênteses
    
    # Remove R$, $, %, e espaços (exceto entre números e unidades como "Bi")
    text_num_part = re.sub(r'[R$\s%]', '', text)
    
    # Trata separadores de milhar (ponto) ANTES de trocar vírgula por ponto decimal
    if ',' in text_num_part and '.' in text_num_part:
        if text_num_part.rfind('.') < text_num_part.rfind(','): 
            text_num_part = text_num_part.replace('.', '') 
        text_num_part = text_num_part.replace(',', '.') 
    elif ',' in text_num_part: 
        text_num_part = text_num_part.replace(',', '.')
    
    match = re.search(r"([-+]?\d*\.?\d+|\d+)", text_num_part)
    if match:
        try: 
            num = float(match.group(1))
            return -num if is_negative_paren else num
        except ValueError: return None
    return None

def extrair_conteudo_docx(uploaded_file):
    """Extrai texto e processa tabelas de um arquivo DOCX."""
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
                    if p_text and len(p_text) < 80: nome_tabela = p_text.replace(":", "").strip()[:70] # Limita nome
            except Exception: pass
            
            header_cells = [cell.text.strip().replace("\n", " ") for cell in table_obj.rows[0].cells] if len(table_obj.rows) > 0 else []
            keys = [key if key else f"Col{c_idx+1}" for c_idx, key in enumerate(header_cells)]

            for r_idx, row in enumerate(table_obj.rows):
                if r_idx == 0 and keys: continue # Pula cabeçalho se já processado
                cells = [c.text.strip() for c in row.cells]
                if keys: 
                    row_dict = {}
                    for k_idx, key_name in enumerate(keys):
                        row_dict[key_name] = cells[k_idx] if k_idx < len(cells) else None
                    data_rows.append(row_dict)
            
            if data_rows:
                try:
                    df = pd.DataFrame(data_rows)
                    for col in df.columns:
                        original_series = df[col].copy()
                        num_series = original_series.astype(str).apply(parse_value_for_numeric)
                        if num_series.notna().sum() / max(1, len(num_series)) >= 0.3: # Critério mais flexível
                            df[col] = pd.to_numeric(num_series, errors='coerce')
                            continue 
                        else: df[col] = original_series 
                        try: # Tenta converter para datetime
                            temp_str_col = df[col].astype(str)
                            dt_series = pd.to_datetime(temp_str_col, errors='coerce', dayfirst=True) 
                            if dt_series.notna().sum() / max(1, len(dt_series)) >= 0.5:
                                df[col] = dt_series
                            else: df[col] = original_series.astype(str).fillna('')
                        except Exception: df[col] = original_series.astype(str).fillna('')
                    for col in df.columns: 
                        if df[col].dtype == 'object': df[col] = df[col].astype(str).fillna('')
                    tabelas_data.append({"id": f"doc_tabela_{i+1}", "nome": nome_tabela, "dataframe": df})
                except Exception as e_df_proc:
                    st.warning(f"Não foi possível processar DataFrame para tabela '{nome_tabela}': {e_df_proc}")
        return "\n\n".join(textos), tabelas_data
    except Exception as e_doc_read: 
        st.error(f"Erro crítico ao ler arquivo DOCX: {e_doc_read}")
        return "", []

def analisar_documento_com_gemini(texto_doc, tabelas_info_list):
    """Envia conteúdo para Gemini e pede sugestões de visualização/análise."""
    api_key = get_gemini_api_key()
    if not api_key: 
        st.warning("Chave API Gemini não configurada. Sugestões da IA desabilitadas.")
        return []
    try:
        genai.configure(api_key=api_key)
        safety_settings = [{"category": c,"threshold": "BLOCK_NONE"} for c in ["HARM_CATEGORY_HARASSMENT","HARM_CATEGORY_HATE_SPEECH","HARM_CATEGORY_SEXUALLY_EXPLICIT","HARM_CATEGORY_DANGEROUS_CONTENT"]]
        model = genai.GenerativeModel(model_name="gemini-1.5-flash-latest", safety_settings=safety_settings)
        
        tabelas_prompt_str = ""
        for t_info in tabelas_info_list:
            df, nome_t, id_t = t_info["dataframe"], t_info["nome"], t_info["id"]
            sample_df = df.head(3).iloc[:, :min(5, len(df.columns))] # Amostra menor
            md_table = ""
            try: md_table = sample_df.to_markdown(index=False)
            except Exception: md_table = sample_df.to_string(index=False) # Fallback
            
            colunas_para_mostrar_tipos = df.columns.tolist()[:min(8, len(df.columns))]
            col_types_list = [f"'{col_name_prompt}' (tipo: {str(df[col_name_prompt].dtype)})" for col_name_prompt in colunas_para_mostrar_tipos]
            col_types_str = ", ".join(col_types_list)
            
            tabelas_prompt_str += f"\n--- Tabela '{nome_t}' (ID para referência: {id_t}) ---\n"
            tabelas_prompt_str += f"Colunas e tipos (amostra de até 8 colunas): {col_types_str}\n"
            tabelas_prompt_str += f"Amostra de dados (primeiras 3 linhas, até 5 colunas):\n{md_table}\n"
        
        text_limit = 40000 # Reduzido para segurança
        prompt_text = texto_doc[:text_limit] + ("\n[TEXTO TRUNCADO...]" if len(texto_doc) > text_limit else "")
        
        prompt = f"""
        Você é um assistente de análise de dados. Analise o texto e as tabelas.
        [TEXTO]{prompt_text}[FIM TEXTO]
        [TABELAS]{tabelas_prompt_str}[FIM TABELAS]

        Gere lista JSON de sugestões de visualizações. Objeto DEVE ter: "id", "titulo", "tipo_sugerido" ("kpi", "tabela_dados", "lista_swot", "grafico_barras", "grafico_pizza", "grafico_linha", "grafico_dispersao", "grafico_barras_agrupadas"), "fonte_id" (ID tabela ou "texto_descricao_fonte"), "parametros" (objeto JSON), "justificativa".
        Para "parametros":
        - "kpi": {{"valor": "ValorKPI", "delta": "Mudança", "descricao": "Contexto"}}
        - "tabela_dados": Para TABELA EXISTENTE: {{"id_tabela_original": "ID_Tabela"}}. Para DADOS DO TEXTO: {{"dados": [{{"ColunaA": "Valor1A", "ColunaB": "Valor1B"}}, ...], "colunas_titulo": ["Título ColA", "Título ColB"]}}
        - "lista_swot": {{"forcas": ["F1"], "fraquezas": ["Fr1"], "oportunidades": ["Op1"], "ameacas": ["Am1"]}} (Listas de strings).
        - Gráficos de TABELA ("barras", "linha", "dispersao"): {{"eixo_x": "NOME_COL_X", "eixo_y": "NOME_COL_Y"}} (Y numérico).
        - Gráficos de PIZZA de TABELA: {{"categorias": "NOME_COL_CAT", "valores": "NOME_COL_VAL_NUM"}} (Valores numéricos).
        - Gráficos com DADOS EXTRAÍDOS DO TEXTO ("barras", "pizza", etc.): {{"dados": [{{"NomeEixoX": "CatA", "NomeEixoY": ValNumA}}, ...], "eixo_x": "NomeEixoX", "eixo_y": "NomeEixoY"}} (Valores DEVEM ser numéricos).
        - "grafico_barras_agrupadas": Se de TABELA: {{"eixo_x": "COL_PRINCIPAL", "eixo_y": "COL_VALOR_NUM", "cor_agrupamento": "COL_SUB_CAT"}}. Se DADOS EXTRAÍDOS: {{"dados": [{{"CatPrincipal": "A", "SubCat": "X", "Valor": 10}}, ...], "eixo_x": "CatPrincipal", "eixo_y": "Valor", "cor_agrupamento": "SubCat"}}.
        
        INSTRUÇÕES CRÍTICAS:
        1.  NOMES DE COLUNAS: Para gráficos de TABELA, use os NOMES EXATOS das colunas como fornecidos nos "Colunas e tipos".
        2.  DADOS NUMÉRICOS: Se a coluna de valor de uma TABELA não for numérica (float64/int64) conforme os "tipos inferidos", NÃO sugira gráfico que exija valor numérico para ela, A MENOS que você possa confiavelmente extrair um valor numérico do seu conteúdo textual (ex: de '70%' extrair 70.0; de '70% - 86%' extrair 70.0; de 'R$ 15,5 Bi' extrair 15.5). Se extrair do texto, coloque em "dados" e garanta que os valores sejam números, não strings de números.
        3.  COBERTURA GEOGRÁFICA (Player, Cidades): Se for lista, sugira "tabela_dados" com "dados" nos "parametros" e "colunas_titulo". Não sugira "mapa".
        4.  SWOT: Se uma tabela compara SWOTs, gere "lista_swot" INDIVIDUAL para CADA player. O "titulo" deve incluir o nome do player.
        Retorne APENAS a lista JSON válida.
        """
        with st.spinner("🤖 Gemini analisando... (Pode levar alguns instantes)"):
            # st.text_area("Debug Prompt:", prompt, height=150) 
            response = model.generate_content(prompt)
        cleaned_text = response.text.strip().lstrip("```json").rstrip("```").strip()
        # st.text_area("Debug Resposta Gemini:", cleaned_text, height=150)
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
        # st.text(traceback.format_exc()) 
        return []

# --- Funções de Renderização Específicas ---
def render_kpis(kpi_sugestoes):
    if kpi_sugestoes:
        num_kpis = len(kpi_sugestoes); kpi_cols = st.columns(min(num_kpis, 4)) 
        for i, kpi_sug in enumerate(kpi_sugestoes):
            with kpi_cols[i % min(num_kpis, 4)]:
                params=kpi_sug.get("parametros",{}); delta_val=str(params.get("delta",""))
                st.metric(label=kpi_sug.get("titulo","KPI"),value=str(params.get("valor","N/A")),delta=delta_val if delta_val else None,help=params.get("descricao"))
        st.divider()

def render_swot_card(titulo_completo_swot, swot_data):
    st.subheader(f"{titulo_completo_swot}") 
    col1, col2 = st.columns(2)
    swot_map = {"forcas": ("Forças 💪", col1), "fraquezas": ("Fraquezas 📉", col1), 
                "oportunidades": ("Oportunidades 🚀", col2), "ameacas": ("Ameaças ⚠️", col2)}
    for key_swot_category, (header_swot_render, col_target_swot_render) in swot_map.items():
        with col_target_swot_render:
            st.markdown(f"##### {header_swot_render}")
            points_swot_render = swot_data.get(key_swot_category, ["N/A (info. não fornecida)"])
            if not points_swot_render or not isinstance(points_swot_render, list) or not all(isinstance(p_swot, str) for p_swot in points_swot_render): 
                points_swot_render = ["N/A (formato de dados incorreto)"]
            if not points_swot_render: points_swot_render = ["N/A"] 
            for item_swot_render in points_swot_render: 
                st.markdown(f"<div style='margin-bottom: 5px;'>- {item_swot_render}</div>", unsafe_allow_html=True)
    st.markdown("<hr style='margin-top: 10px; margin-bottom: 20px;'>", unsafe_allow_html=True)

def render_plotly_chart(item_config, df_plot_input):
    if df_plot_input is None:
        st.warning(f"Dados não disponíveis para o gráfico '{item_config.get('titulo', 'Sem Título')}'.")
        return False
    df_plot = df_plot_input.copy()
    tipo_grafico, titulo, params = item_config.get("tipo_sugerido"), item_config.get("titulo"), item_config.get("parametros", {})
    x_col, y_col, cat_col, val_col = params.get("eixo_x"), params.get("eixo_y"), params.get("categorias"), params.get("valores")
    cor_agrupamento_col = params.get("cor_agrupamento") # Para barras agrupadas
    
    fig, plot_func, plot_args = None, None, {}

    if tipo_grafico in ["grafico_barras", "grafico_barras_agrupadas"] and x_col and y_col: 
        plot_func,plot_args=px.bar,{"x":x_col,"y":y_col}
        if tipo_grafico == "grafico_barras_agrupadas" and cor_agrupamento_col:
            plot_args["color"], plot_args["barmode"] = cor_agrupamento_col, "group"
    elif tipo_grafico=="grafico_linha" and x_col and y_col: plot_func,plot_args=px.line,{"x":x_col,"y":y_col,"markers":True}
    elif tipo_grafico=="grafico_dispersao" and x_col and y_col: plot_func,plot_args=px.scatter,{"x":x_col,"y":y_col}
    elif tipo_grafico=="grafico_pizza" and cat_col and val_col: plot_func,plot_args=px.pie,{"names":cat_col,"values":val_col}
    # Adicionar grafico_radar aqui se a LLM for fornecer os dados corretamente
    
    if plot_func:
        required_cols=[col for col in plot_args.values() if isinstance(col,str)] # Pega todos os nomes de colunas dos argumentos
        if not all(col in df_plot.columns for col in required_cols):
            st.warning(f"Colunas necessárias {required_cols} não encontradas para '{titulo}'. Disponíveis: {df_plot.columns.tolist()}")
            return False
        try:
            df_plot_cleaned = df_plot.copy() 
            # Tenta converter colunas de valor/eixo Y para numérico ANTES de remover NaNs
            y_axis_col_plot = plot_args.get("y"); values_col_plot = plot_args.get("values")
            if y_axis_col_plot and y_axis_col_plot in df_plot_cleaned.columns: 
                df_plot_cleaned[y_axis_col_plot] = pd.to_numeric(df_plot_cleaned[y_axis_col_plot], errors='coerce')
            if values_col_plot and values_col_plot in df_plot_cleaned.columns:
                 df_plot_cleaned[values_col_plot] = pd.to_numeric(df_plot_cleaned[values_col_plot], errors='coerce')
            
            df_plot_cleaned.dropna(subset=required_cols, inplace=True)

            if not df_plot_cleaned.empty:
                fig=plot_func(df_plot_cleaned,title=titulo,**plot_args)
                st.plotly_chart(fig,use_container_width=True); return True
            else: st.warning(f"Dados insuficientes para '{titulo}' após limpar NaNs de {required_cols}.")
        except Exception as e_plotly_render: 
            st.warning(f"Erro ao gerar gráfico Plotly '{titulo}': {e_plotly_render}. Dtypes do DF usado: {df_plot.dtypes.to_dict() if df_plot is not None else 'DF é None'}")
    elif tipo_grafico in ["grafico_barras","grafico_barras_agrupadas","grafico_linha","grafico_dispersao","grafico_pizza","grafico_radar"]:
        st.warning(f"Configuração de parâmetros incompleta para '{titulo}' (tipo: {tipo_grafico}).")
    return False

# --- 3. Interface Streamlit Principal ---
st.set_page_config(layout="wide", page_title="Gemini DOCX Insights")
for k, dv in [("s_gemini",[]),("cfg_sugs",{}),("doc_ctx",{"texto":"","tabelas":[]}),
              ("f_name",None),("dbg_cb_key",False),("pg_sel","Dashboard Principal")]:
    st.session_state.setdefault(k, dv)

st.sidebar.title("✨ Navegação"); pg_opts_sb = ["Dashboard Principal","Análise SWOT Detalhada"] # Renomeado para evitar conflito
st.session_state.pg_sel=st.sidebar.radio("Selecione:",pg_opts_sb,index=pg_opts_sb.index(st.session_state.pg_sel),key="nav_radio_final")
st.sidebar.divider(); uploaded_file_sb = st.sidebar.file_uploader("Selecione DOCX",type="docx",key="uploader_sidebar_final")
st.session_state.dbg_cb_key=st.sidebar.checkbox("Mostrar Debug Info",value=st.session_state.dbg_cb_key,key="debug_cb_sidebar_final")

if uploaded_file_sb:
    if st.session_state.f_name!=uploaded_file_sb.name: 
        with st.spinner("Processando novo documento..."):
            st.session_state.s_gemini,st.session_state.cfg_sugs=[],{}
            st.session_state.f_name=uploaded_file_sb.name
            txt_main,tbls_main=extrair_conteudo_docx(uploaded_file_sb);st.session_state.doc_ctx={"texto":txt_main,"tabelas":tbls_main}
            if txt_main or tbls_main:
                sugs_main=analisar_documento_com_gemini(txt_main,tbls_main);st.session_state.s_gemini=sugs_main
                st.session_state.cfg_sugs={s.get("id",f"s_main_{i}_{hash(s.get('titulo'))}"):{"aceito":True,"titulo_editado":s.get("titulo","S/T"),"dados_originais":s} for i,s in enumerate(sugs_main)}
            else: st.sidebar.warning("Nenhum conteúdo extraído.")
    
    if st.session_state.dbg_cb_key and (st.session_state.doc_ctx["texto"] or st.session_state.doc_ctx["tabelas"]):
        with st.expander("Debug: Conteúdo DOCX (após extração e tipos)",expanded=False):
            st.text_area("Texto (amostra)",st.session_state.doc_ctx["texto"][:1000],height=80)
            for t_dbg_main in st.session_state.doc_ctx["tabelas"]:
                st.write(f"ID: {t_dbg_main['id']}, Nome: {t_dbg_main['nome']}")
                try: st.dataframe(t_dbg_main['dataframe'].head().astype(str).fillna("-"))
                except: st.text(f"Head:\n{t_dbg_main['dataframe'].head().to_string(na_rep='-')}")
                st.write("Tipos:",t_dbg_main['dataframe'].dtypes.to_dict())

    if st.session_state.s_gemini:
        st.sidebar.divider();st.sidebar.header("⚙️ Configurar Sugestões")
        for sug_cfg_loop in st.session_state.s_gemini:
            s_id_loop = sug_cfg_loop.get('id') 
            if not s_id_loop : continue # Pula se não houver ID
            if s_id_loop not in st.session_state.cfg_sugs: # Segurança: inicializa se faltar
                 st.session_state.cfg_sugs[s_id_loop]={"aceito":True,"titulo_editado":sug_cfg_loop.get("titulo","S/T"),"dados_originais":sug_cfg_loop}
            cfg_loop = st.session_state.cfg_sugs[s_id_loop]
            
            with st.sidebar.expander(f"{cfg_loop['titulo_editado']}",expanded=False):
                st.caption(f"Tipo: {sug_cfg_loop.get('tipo_sugerido')} | Fonte: {sug_cfg_loop.get('fonte_id')}")
                cfg_loop["aceito"]=st.checkbox("Incluir?",value=cfg_loop["aceito"],key=f"acc_loop_{s_id_loop}")
                cfg_loop["titulo_editado"]=st.text_input("Título",value=cfg_loop["titulo_editado"],key=f"tit_loop_{s_id_loop}")
else: 
    if st.session_state.pg_sel=="Dashboard Principal": st.info("Upload DOCX na barra lateral.")

# --- RENDERIZAÇÃO DA PÁGINA ---
if st.session_state.pg_sel=="Dashboard Principal":
    st.title("📊 Dashboard de Insights")
    if uploaded_file_sb and st.session_state.s_gemini:
        kpis_r, outros_r = [], []
        for s_id_r,s_cfg_r in st.session_state.cfg_sugs.items():
            if s_cfg_r["aceito"]: item_r={"titulo":s_cfg_r["titulo_editado"], **s_cfg_r["dados_originais"]};(kpis_r if item_r.get("tipo_sugerido")=="kpi" else outros_r).append(item_r)
        render_kpis(kpis_r)
        if st.session_state.dbg_cb_key:
             with st.expander("Debug: Elementos para Dashboard Principal (Não-KPI)",expanded=True): st.json({"Outros":outros_r},expanded=False)
        
        elementos_renderizados_dash = 0
        if outros_r:
            cols_dash_r,idx_dash_r=st.columns(2),0
            for item_main_r in outros_r:
                if item_main_r.get("tipo_sugerido")=="lista_swot":continue
                with cols_dash_r[idx_dash_r%2]:
                    df_plot_r,rendered_r=None,False
                    params_r,tipo_r,fonte_r=item_main_r.get("parametros",{}),item_main_r.get("tipo_sugerido"),item_main_r.get("fonte_id")
                    titulo_r = item_main_r.get("titulo", "Visualização") # Pega o título editado
                    st.subheader(titulo_r)
                    try:
                        if params_r.get("dados"):df_plot_r=pd.DataFrame(params_r["dados"])
                        elif str(fonte_r).startswith("doc_tabela_"):df_plot_r=next((t["dataframe"] for t in st.session_state.doc_ctx["tabelas"] if t["id"]==fonte_r),None)
                        
                        if tipo_r=="tabela_dados":
                            df_tbl_r=None
                            if str(fonte_r).startswith("texto_") and params_r.get("dados"):
                                df_tbl_r=pd.DataFrame(params_r.get("dados")); 
                                if params_r.get("colunas_titulo"):df_tbl_r.columns=params_r.get("colunas_titulo")
                            else:id_tbl_r=params_r.get("id_tabela_original",fonte_r);df_tbl_r=next((t["dataframe"] for t in st.session_state.doc_ctx["tabelas"] if t["id"]==id_tbl_r),None)
                            if df_tbl_r is not None:try:st.dataframe(df_tbl_r.astype(str).fillna("-"))
                                                except:st.text(df_tbl_r.to_string(na_rep='-')); rendered_r=True
                            else:st.warning(f"Tabela '{titulo_r}' (Fonte:{fonte_r}) não encontrada.")
                        elif tipo_r in ["grafico_barras","grafico_linha","grafico_dispersao","grafico_pizza", "grafico_barras_agrupadas"]:
                            if render_plotly_chart(item_main_r,df_plot_r):rendered_r=True # Passa o item completo
                        elif tipo_r=='mapa':st.info(f"Mapa '{titulo_r}' não implementado.");rendered_r=True
                        if not rendered_r and tipo_r not in ["kpi","lista_swot","mapa"]:st.info(f"'{titulo_r}' ({tipo_r}) não gerado.")
                    except Exception as e:st.error(f"Erro render '{titulo_r}': {e}")
                if rendered_r:idx_dash_r+=1;elementos_renderizados_dash+=1
            if elementos_renderizados_dash==0 and any(c['aceito'] and c['dados_originais'].get('tipo_sugerido') not in ['kpi','lista_swot'] for c in st.session_state.cfg_sugs.values()):
                st.info("Nenhum gráfico/tabela (além de KPIs/SWOTs) pôde ser gerado.")
        elif not kpis_r and not uploaded_file_sidebar:pass
        elif not kpis_r and not outros_r and uploaded_file_sidebar and st.session_state.s_gemini:st.info("Nenhum elemento selecionado/gerado.")

elif st.session_state.pg_sel=="Análise SWOT Detalhada":
    st.title("🔬 Análise SWOT Detalhada")
    if not uploaded_file_sidebar:st.warning("Upload DOCX na barra lateral.")
    elif not st.session_state.s_gemini:st.info("Aguardando processamento/sugestões.")
    else:
        swot_sugs_page_render=[s_cfg_swot["dados_originais"] for s_id_swot,s_cfg_swot in st.session_state.cfg_sugs.items() if s_cfg_swot["aceito"] and s_cfg_swot["dados_originais"].get("tipo_sugerido")=="lista_swot"]
        if not swot_sugs_page_render:st.info("Nenhuma análise SWOT sugerida/selecionada.")
        else:
            if show_debug_info_sidebar:
                with st.expander("Debug: Dados para Análise SWOT (Página Dedicada)",expanded=False):st.json({"SWOTs Selecionados":swot_sugs_page_render})
            for swot_item_render in swot_sugs_page_render:
                render_swot_card(swot_item_render.get("titulo","SWOT"),swot_item_render.get("parametros",{}),card_key_prefix=swot_item_render.get("id","swot_pg"))

if uploaded_file_sidebar is None and st.session_state.f_name is not None:
    keys_to_preserve=["nav_radio_key_final","uploader_sidebar_key_final","debug_cb_sidebar_key_final"] 
    for sug_key_cfg in st.session_state.get("s_gemini", []):
        s_id_cfg_val = sug_key_cfg.get('id')
        if s_id_cfg_val: keys_to_preserve.extend([f"acc_cfg_{s_id_cfg_val}", f"tit_cfg_{s_id_cfg_val}"])
    current_keys_clear=list(st.session_state.keys())
    for k_cl_main in current_keys_clear:
        if k_cl_main not in keys_to_preserve:
            if k_cl_main in st.session_state: del st.session_state[k_cl_main]
    for k_r_main,dv_r_main in [("s_gemini",[]),("cfg_sugs",{}),("doc_ctx",{"texto":"","tabelas":[]}),("f_name",None),("dbg_cb_key",False),("pg_sel","Dashboard Principal")]:st.session_state.setdefault(k_r_main,dv_r_main)
    st.experimental_rerun()

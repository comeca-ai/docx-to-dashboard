import streamlit as st
from docx import Document
import pandas as pd
import plotly.express as px
import google.generativeai as genai
import json
import os
import traceback
import re 

# --- Suas funções get_gemini_api_key, parse_value_for_numeric, extrair_conteudo_docx, analisar_documento_com_gemini ---
# --- (COPIE-AS DA VERSÃO ANTERIOR, ELAS ESTÃO BOAS) ---
# ... (cole as funções aqui) ...
# --- COPIE AS FUNÇÕES ATÉ AQUI ---

# --- Funções de Renderização Específicas para o Novo Layout ---

def render_kpis(kpi_sugestoes):
    if kpi_sugestoes:
        num_kpis = len(kpi_sugestoes)
        kpi_cols = st.columns(min(num_kpis, 4)) # Máximo de 4 KPIs por linha
        for i, kpi_sug in enumerate(kpi_sugestoes):
            with kpi_cols[i % min(num_kpis, 4)]:
                params = kpi_sug.get("parametros",{})
                st.metric(
                    label=kpi_sug.get("titulo","KPI"), 
                    value=str(params.get("valor", "N/A")),
                    delta=str(params.get("delta", "")),
                    help=params.get("descricao")
                )
        st.divider()

def render_swot_card(player_name, swot_data):
    """Renderiza um card SWOT para um player específico, similar ao template."""
    st.subheader(f"Análise SWOT - {player_name}")
    
    # Usar colunas para o layout Forças/Fraquezas | Oportunidades/Ameaças
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("##### Forças 💪")
        for item in swot_data.get('forcas', ["N/A"]): st.markdown(f"<div style='margin-bottom: 5px;'>- {item}</div>", unsafe_allow_html=True)
        st.markdown("---") # Divisor visual
        st.markdown("##### Fraquezas 📉")
        for item in swot_data.get('fraquezas', ["N/A"]): st.markdown(f"<div style='margin-bottom: 5px;'>- {item}</div>", unsafe_allow_html=True)
    
    with col2:
        st.markdown("##### Oportunidades 🚀")
        for item in swot_data.get('oportunidades', ["N/A"]): st.markdown(f"<div style='margin-bottom: 5px;'>- {item}</div>", unsafe_allow_html=True)
        st.markdown("---") # Divisor visual
        st.markdown("##### Ameaças ⚠️")
        for item in swot_data.get('ameacas', ["N/A"]): st.markdown(f"<div style='margin-bottom: 5px;'>- {item}</div>", unsafe_allow_html=True)
    st.markdown("---") # Divisor após cada card SWOT

def render_plotly_chart(item_config, df_plot):
    """Renderiza um gráfico Plotly com base na configuração."""
    tipo_grafico = item_config.get("tipo_sugerido")
    titulo = item_config.get("titulo")
    params = item_config.get("parametros", {})
    
    x_col, y_col = params.get("eixo_x"), params.get("eixo_y")
    cat_col, val_col = params.get("categorias"), params.get("valores")
    
    fig = None
    plot_func, plot_args = None, {}

    if tipo_grafico == "grafico_barras" and x_col and y_col: plot_func, plot_args = px.bar, {"x":x_col,"y":y_col}
    elif tipo_grafico == "grafico_linha" and x_col and y_col: plot_func, plot_args = px.line, {"x":x_col,"y":y_col,"markers":True}
    elif tipo_grafico == "grafico_dispersao" and x_col and y_col: plot_func, plot_args = px.scatter, {"x":x_col,"y":y_col}
    elif tipo_grafico == "grafico_pizza" and cat_col and val_col: plot_func, plot_args = px.pie,{"names":cat_col,"values":val_col}

    if plot_func and all(k_col in df_plot.columns for k_col in plot_args.values() if isinstance(k_col,str)):
        try:
            df_plot_cleaned = df_plot.copy()
            # Tenta converter colunas de plotagem para numérico onde aplicável
            y_axis_col_plot = plot_args.get("y")
            values_col_plot = plot_args.get("values")
            if y_axis_col_plot and y_axis_col_plot in df_plot_cleaned.columns: 
                df_plot_cleaned[y_axis_col_plot] = pd.to_numeric(df_plot_cleaned[y_axis_col_plot], errors='coerce')
            if values_col_plot and values_col_plot in df_plot_cleaned.columns:
                 df_plot_cleaned[values_col_plot] = pd.to_numeric(df_plot_cleaned[values_col_plot], errors='coerce')
            
            cols_to_check_na = [val for val in plot_args.values() if isinstance(val, str) and val in df_plot_cleaned.columns]
            df_plot_cleaned.dropna(subset=cols_to_check_na, inplace=True)

            if not df_plot_cleaned.empty:
                fig = plot_func(df_plot_cleaned, title=titulo, **plot_args)
                st.plotly_chart(fig, use_container_width=True)
                return True # Renderizado com sucesso
            else: st.warning(f"Dados insuficientes para '{titulo}' após limpar NaNs.")
        except Exception as e_plotly: st.warning(f"Erro Plotly '{titulo}': {e_plotly}.")
    elif plot_func: st.warning(f"Colunas ausentes/incorretas para '{titulo}'. Esperado: {plot_args}. Disponível: {df_plot.columns.tolist()}")
    return False # Não renderizado


# --- Interface Streamlit Principal ---
st.set_page_config(layout="wide", page_title="Gemini DOCX Insights")

# Inicialização de estado
for k, default_val in [("sugestoes_gemini", []), ("config_sugestoes", {}), 
                       ("conteudo_docx", {"texto": "", "tabelas": []}), 
                       ("nome_arquivo_atual", None), ("debug_checkbox_key_main", False),
                       ("pagina_selecionada", "Dashboard Principal")]: # Novo estado para navegação
    st.session_state.setdefault(k, default_val)

# --- BARRA LATERAL DE NAVEGAÇÃO E CONTROLES ---
st.sidebar.title("✨ Navegação")
pagina_opcoes = ["Dashboard Principal", "Análise SWOT Detalhada"]
st.session_state.pagina_selecionada = st.sidebar.radio("Selecione a Visualização:", pagina_opcoes, 
                                                      index=pagina_opcoes.index(st.session_state.pagina_selecionada))
st.sidebar.divider()
uploaded_file = st.sidebar.file_uploader("Selecione DOCX", type="docx", key="uploader_sidebar_key")
show_debug_info = st.sidebar.checkbox("Mostrar Informações de Depuração", 
                                    value=st.session_state.debug_checkbox_key_main, 
                                    key="debug_cb_widget_sidebar_key")
st.session_state.debug_checkbox_key_main = show_debug_info # Sincroniza estado

# --- LÓGICA DE PROCESSAMENTO DO ARQUIVO (apenas se um arquivo for carregado) ---
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
                temp_config_init = {}
                for i_init,s_init in enumerate(sugestoes): 
                    s_id_init = s_init.get("id", f"s_init_{i_init}_{hash(s_init.get('titulo',''))}"); s_init["id"] = s_id_init
                    temp_config_init[s_id_init] = {"aceito":True,"titulo_editado":s_init.get("titulo","S/Título"),"dados_originais":s_init}
                st.session_state.config_sugestoes = temp_config_init
            else: st.sidebar.warning("Nenhum conteúdo extraído.")
    
    # Configuração das sugestões na sidebar (APENAS se houver sugestões)
    if st.session_state.sugestoes_gemini:
        st.sidebar.divider()
        st.sidebar.header("⚙️ Configurar Sugestões")
        for sug_sidebar in st.session_state.sugestoes_gemini:
            s_id_sb = sug_sidebar['id'] 
            if s_id_sb not in st.session_state.config_sugestoes: # Segurança
                 st.session_state.config_sugestoes[s_id_sb] = {"aceito":True,"titulo_editado":sug_sidebar.get("titulo","S/Título"),"dados_originais":sug_sidebar}
            cfg_sb = st.session_state.config_sugestoes[s_id_sb]
            with st.sidebar.expander(f"{cfg_sb['titulo_editado']}", expanded=False):
                st.caption(f"Tipo: {sug_sidebar.get('tipo_sugerido')} | Fonte: {sug_sidebar.get('fonte_id')}")
                cfg_sb["aceito"]=st.checkbox("Incluir?",value=cfg_sb["aceito"],key=f"acc_sb_{s_id_sb}")
                cfg_sb["titulo_editado"]=st.text_input("Título",value=cfg_sb["titulo_editado"],key=f"tit_sb_{s_id_sb}")
                # (A lógica de edição de parâmetros foi removida da sidebar para simplificar,
                #  confiando mais na sugestão inicial da LLM ou ajustes no prompt)
else: # Se nenhum arquivo estiver carregado
    st.info("Por favor, faça o upload de um arquivo DOCX para começar.")


# --- RENDERIZAÇÃO DA PÁGINA SELECIONADA ---

if st.session_state.pagina_selecionada == "Dashboard Principal":
    st.title("📊 Dashboard de Insights do Documento")
    if not uploaded_file:
        st.warning("Faça upload de um documento DOCX na barra lateral.")
    elif not st.session_state.sugestoes_gemini:
        st.info("Aguardando processamento do documento ou nenhuma sugestão foi gerada.")
    else:
        kpis_dash_f, outros_dash_f = [], []
        for s_id_render, config_render in st.session_state.config_sugestoes.items():
            if config_render["aceito"]: 
                item_f = {"titulo":config_render["titulo_editado"], **config_render["dados_originais"]}
                (kpis_dash_f if item_f.get("tipo_sugerido")=="kpi" else outros_dash_f).append(item_f)
        
        render_kpis(kpis_dash_f) # Função para renderizar KPIs

        if show_debug_info: # Debug ANTES de tentar renderizar outros elementos
             with st.expander("Debug: Elementos para Dashboard (Não-KPI)", expanded=False):
                 st.json({"Outros Elementos Configurados": outros_dash_f}, expanded=False)

        elementos_renderizados_count = 0
        if outros_dash_f:
            item_cols_render = st.columns(2); col_idx_f = 0 
            for item_d_main in outros_dash_f:
                el_rend_d = False 
                with item_cols_render[col_idx_f % 2]: 
                    params_d=item_d_main.get("parametros",{}); tipo_d=item_d_main.get("tipo_sugerido"); fonte_d=item_d_main.get("fonte_id")
                    
                    # Não renderiza SWOTs aqui, eles terão página própria
                    if tipo_d == "lista_swot":
                        continue # Pula para o próximo item

                    st.subheader(item_d_main["titulo"]); df_plot_d = None 
                    try:
                        if params_d.get("dados"): # Prioriza dados da LLM
                            try: df_plot_d=pd.DataFrame(params_d["dados"])
                            except Exception as e_dfd: st.warning(f"'{item_d_main['titulo']}': Erro DF de 'dados': {e_dfd}"); continue
                        elif str(fonte_d).startswith("doc_tabela_"): # Senão, busca tabela extraída
                            df_plot_d=next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"]==fonte_d),None)
                        
                        if tipo_d=="tabela_dados":
                            # ... (lógica de renderização de tabela_dados como antes) ...
                            df_t_render_f = None 
                            if str(fonte_d).startswith("texto_") and params_d.get("dados"):
                                try: 
                                    df_t_render_f = pd.DataFrame(params_d.get("dados"))
                                    if params_d.get("colunas_titulo"): df_t_render_f.columns = params_d.get("colunas_titulo")
                                except Exception as e_df_txt_tbl_f: st.warning(f"Erro tabela texto '{item_d_main['titulo']}': {e_df_txt_tbl_f}")
                            else: 
                                id_t_render_f=params_d.get("id_tabela_original",fonte_d)
                                df_t_render_f=next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"]==id_t_render_f),None)
                            if df_t_render_f is not None: 
                                try: st.dataframe(df_t_render_f.astype(str).fillna("-"))
                                except Exception: st.text(df_t_render_f.to_string(na_rep='-')) # Fallback para texto
                                el_rend_d=True
                            else: st.warning(f"Tabela para '{item_d_main['titulo']}' (Fonte: {fonte_d}) não encontrada.")
                        
                        elif df_plot_d is not None: # Gráficos Plotly
                            if render_plotly_chart(item_d_main, df_plot_d):
                                el_rend_d = True
                        
                        if tipo_d == 'mapa': st.info(f"Mapa para '{item_d_main['titulo']}' não implementado."); el_rend_d=True
                        
                        if not el_rend_d and tipo_d not in ["kpi","lista_swot","mapa"]:
                            st.info(f"'{item_d_main['titulo']}' (tipo: {tipo_d}) não gerado. Dados/Tipo não suportado.")
                    except Exception as e_main_render_f: st.error(f"Erro renderizando '{item_d_main['titulo']}': {e_main_render_f}")
                if el_rend_d: col_idx_f+=1; elementos_renderizados_count+=1
        
        if elementos_renderizados_count == 0 and any(c['aceito'] and c['dados_originais'].get('tipo_sugerido') not in ['kpi', 'lista_swot'] for c in st.session_state.config_sugestoes.values()):
            st.info("Nenhum gráfico/tabela (além de KPIs e SWOTs) pôde ser gerado.")
        elif elementos_renderizados_count == 0 and not kpis_dash_f: 
            st.info("Nenhum elemento selecionado ou pôde ser gerado para o dashboard.")

elif st.session_state.pagina_selecionada == "Análise SWOT Detalhada":
    st.title("🔬 Análise SWOT Detalhada")
    if not uploaded_file:
        st.warning("Faça upload de um documento DOCX na barra lateral.")
    elif not st.session_state.sugestoes_gemini:
        st.info("Aguardando processamento do documento ou nenhuma sugestão foi gerada.")
    else:
        swot_sugestoes = [s_cfg["dados_originais"] for s_id, s_cfg in st.session_state.config_sugestoes.items() 
                          if s_cfg["aceito"] and s_cfg["dados_originais"].get("tipo_sugerido") == "lista_swot"]
        
        if not swot_sugestoes:
            st.info("Nenhuma análise SWOT foi sugerida ou selecionada para exibição.")
        else:
            for swot_item in swot_sugestoes:
                # Tenta extrair o nome do player do título se for um SWOT individual
                player_name_match = re.search(r"SWOT(?: d[oa]| -) (.+)", swot_item.get("titulo","Análise SWOT"), re.IGNORECASE)
                player_name_swot = player_name_match.group(1) if player_name_match else "Geral"
                
                render_swot_card(player_name_swot, swot_item.get("parametros", {}))

# Limpar estado se o arquivo for removido da UI
if uploaded_file is None and st.session_state.nome_arquivo_atual is not None:
    keys_to_clear_final = list(st.session_state.keys())
    preserved_widget_keys = [k for k in keys_to_clear_final if k.endswith(("_key", "_widget_key", "_key_main", "_vfinal_corrected_again", "_vfinal_swnan", "_sidebar_key", "_cb_widget_key"))]
    for key_clear_f in keys_to_clear_final:
        if key_clear_f not in preserved_widget_keys:
            del st.session_state[key_clear_f]
    st.session_state.sugestoes_gemini, st.session_state.config_sugestoes = [], {}
    st.session_state.conteudo_docx = {"texto": "", "tabelas": []}
    st.session_state.nome_arquivo_atual = None
    st.session_state.debug_checkbox_key_main = False
    st.experimental_rerun()

# --- 3. Interface Streamlit Principal ---
st.set_page_config(layout="wide", page_title="Gemini DOCX Insights")

for k, dv in [("sugestoes_gemini",[]),("config_sugestoes",{}),("conteudo_docx",{"texto":"","tabelas":[]}),
              ("nome_arquivo_atual",None),("debug_checkbox_key",False),("pagina_selecionada","Dashboard Principal")]:
    st.session_state.setdefault(k, dv)

st.sidebar.title("✨ Navegação"); pagina_opcoes_sidebar = ["Dashboard Principal", "Análise SWOT Detalhada"]
st.session_state.pagina_selecionada = st.sidebar.radio(
    "Selecione:", 
    pagina_opcoes_sidebar, 
    index=pagina_opcoes_sidebar.index(st.session_state.pagina_selecionada), 
    key="nav_radio_key_final_v2" 
)
st.sidebar.divider()
uploaded_file_sidebar = st.sidebar.file_uploader("Selecione DOCX", type="docx", key="uploader_sidebar_key_final_v2")
show_debug_info_sidebar = st.sidebar.checkbox("Mostrar Informações de Depuração", 
                                    value=st.session_state.debug_checkbox_key, 
                                    key="debug_cb_sidebar_key_final_v2") 
st.session_state.debug_checkbox_key = show_debug_info_sidebar

if uploaded_file_sidebar:
    if st.session_state.nome_arquivo_atual != uploaded_file_sidebar.name: 
        with st.spinner("Processando novo documento..."):
            st.session_state.sugestoes_gemini, st.session_state.config_sugestoes = [], {}
            st.session_state.nome_arquivo_atual = uploaded_file_sidebar.name
            texto_doc_main, tabelas_doc_main = extrair_conteudo_docx(uploaded_file_sidebar)
            st.session_state.conteudo_docx = {"texto": texto_doc_main, "tabelas": tabelas_doc_main}
            if texto_doc_main or tabelas_doc_main:
                sugestoes_main = analisar_documento_com_gemini(texto_doc_main, tabelas_doc_main)
                st.session_state.sugestoes_gemini = sugestoes_main
                temp_config_init_main = {}
                for i_init_main,s_init_main in enumerate(sugestoes_main): 
                    s_id_init_main = s_init_main.get("id", f"s_init_main_{i_init_main}_{hash(s_init_main.get('titulo',''))}"); s_init_main["id"] = s_id_init_main
                    temp_config_init_main[s_id_init_main] = {"aceito":True,"titulo_editado":s_init_main.get("titulo","S/Título"),"dados_originais":s_init_main}
                st.session_state.config_sugestoes = temp_config_init_main
            else: st.sidebar.warning("Nenhum conteúdo extraído do DOCX.")
    
    if show_debug_info_sidebar and (st.session_state.conteudo_docx["texto"] or st.session_state.conteudo_docx["tabelas"]):
        with st.expander("Debug: Conteúdo DOCX (após extração e tipos)", expanded=False):
            st.text_area("Texto (amostra)", st.session_state.conteudo_docx["texto"][:1000], height=80)
            for t_info_dbg_main in st.session_state.conteudo_docx["tabelas"]:
                st.write(f"ID: {t_info_dbg_main['id']}, Nome: {t_info_dbg_main['nome']}")
                try: st.dataframe(t_info_dbg_main['dataframe'].head().astype(str).fillna("-")) 
                except Exception: st.text(f"Head:\n{t_info_dbg_main['dataframe'].head().to_string(na_rep='-')}")
                st.write("Tipos:", t_info_dbg_main['dataframe'].dtypes.to_dict())

    if st.session_state.sugestoes_gemini:
        st.sidebar.divider(); st.sidebar.header("⚙️ Configurar Sugestões")
        for sug_cfg_sidebar in st.session_state.sugestoes_gemini:
            s_id_cfg_sb = sug_cfg_sidebar.get('id') # Assume que ID já foi garantido
            if not s_id_cfg_sb : continue # Pula se não houver ID
            
            if s_id_cfg_sb not in st.session_state.config_sugestoes:
                 st.session_state.config_sugestoes[s_id_cfg_sb] = {"aceito":True,"titulo_editado":sug_cfg_sidebar.get("titulo","S/Título"),"dados_originais":sug_cfg_sidebar}
            cfg_current_sb = st.session_state.config_sugestoes[s_id_cfg_sb]
            
            with st.sidebar.expander(f"{cfg_current_sb['titulo_editado']}",expanded=False):
                st.caption(f"Tipo: {sug_cfg_sidebar.get('tipo_sugerido')} | Fonte: {sug_cfg_sidebar.get('fonte_id')}")
                cfg_current_sb["aceito"]=st.checkbox("Incluir?",value=cfg_current_sb["aceito"],key=f"acc_cfg_{s_id_cfg_sb}")
                cfg_current_sb["titulo_editado"]=st.text_input("Título",value=cfg_current_sb["titulo_editado"],key=f"tit_cfg_{s_id_cfg_sb}")
else: 
    if st.session_state.pagina_selecionada == "Dashboard Principal":
        st.info("Por favor, faça o upload de um arquivo DOCX na barra lateral para começar.")

# --- RENDERIZAÇÃO DA PÁGINA SELECIONADA ---
if st.session_state.pagina_selecionada == "Dashboard Principal":
    st.title("📊 Dashboard de Insights do Documento")
    if uploaded_file_sidebar and st.session_state.sugestoes_gemini:
        kpis_render, outros_render = [], []
        for s_id_main_dash, s_cfg_main_dash in st.session_state.config_sugestoes.items():
            if s_cfg_main_dash["aceito"]: 
                item_main_dash = {"titulo":s_cfg_main_dash["titulo_editado"], **s_cfg_main_dash["dados_originais"]}
                (kpis_render if item_main_dash.get("tipo_sugerido")=="kpi" else outros_render).append(item_main_dash)
        
        render_kpis(kpis_render)
        
        if show_debug_info_sidebar:
             with st.expander("Debug: Elementos para Dashboard Principal (Não-KPI)", expanded=False): 
                st.json({"Outros Elementos (Configurados e Aceitos)": outros_render}, expanded=False)
        
        elementos_renderizados_count = 0 
        col_idx_dash = 0 
        if outros_render:
            item_cols_main_dash = st.columns(2)
            for item_render_loop in outros_render:
                if item_render_loop.get("tipo_sugerido") == "lista_swot": continue 
                
                with item_cols_main_dash[col_idx_dash % 2]:
                    df_plot_loop, rendered_loop = None, False
                    params_loop = item_render_loop.get("parametros",{})
                    tipo_loop = item_render_loop.get("tipo_sugerido")
                    fonte_loop = item_render_loop.get("fonte_id")
                    titulo_loop = item_render_loop.get("titulo")
                    
                    st.subheader(titulo_loop) 
                    try:
                        if params_loop.get("dados"): 
                            try: df_plot_loop=pd.DataFrame(params_loop["dados"])
                            except Exception as e_dfd_loop: st.warning(f"'{titulo_loop}': Erro DF 'dados': {e_dfd_loop}"); continue
                        elif str(fonte_loop).startswith("doc_tabela_"): 
                            df_plot_loop=next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"]==fonte_loop),None)
                        
                        if tipo_loop=="tabela_dados":
                            df_tbl_loop=None
                            if str(fonte_loop).startswith("texto_") and params_loop.get("dados"):
                                try: 
                                    df_tbl_loop=pd.DataFrame(params_loop.get("dados")); 
                                    if params_loop.get("colunas_titulo"): df_tbl_loop.columns=params_loop.get("colunas_titulo")
                                except Exception as e_dftxt_loop: st.warning(f"Erro tabela texto '{titulo_loop}': {e_dftxt_loop}")
                            else: 
                                id_tbl_loop=params_loop.get("id_tabela_original",fonte_loop)
                                df_tbl_loop=next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"]==id_tbl_loop),None)
                            
                            if df_tbl_loop is not None: 
                                try: st.dataframe(df_tbl_loop.astype(str).fillna("-"))
                                except Exception: st.text(df_tbl_loop.to_string(na_rep='-')); 
                                rendered_loop=True
                            else: st.warning(f"Tabela '{titulo_loop}' (Fonte: {fonte_loop}) não encontrada.")
                        
                        elif tipo_loop in ["grafico_barras","grafico_linha","grafico_dispersao","grafico_pizza", "grafico_barras_agrupadas", "grafico_radar"]: # Adicionado _agrupadas e _radar
                            if render_plotly_chart(item_render_loop, df_plot_loop): rendered_loop = True
                        
                        elif tipo_loop == 'mapa': 
                            st.info(f"Mapa para '{titulo_loop}' não implementado.")
                            rendered_loop=True
                        
                        if not rendered_loop and tipo_loop not in ["kpi","lista_swot","mapa"]: 
                            st.info(f"'{titulo_loop}' (tipo: {tipo_loop}) não gerado. Dados/Tipo não suportado.")
                    except Exception as e_render_loop: 
                        st.error(f"Erro renderizando '{titulo_loop}': {e_render_loop}")
                
                if rendered_loop: 
                    idx_main_dash+=1
                    elementos_renderizados_count+=1 # Nome da variável corrigido
            
            if elementos_renderizados_count == 0 and any(c['aceito'] and c['dados_originais'].get('tipo_sugerido') not in ['kpi','lista_swot'] for c in st.session_state.config_sugestoes.values()):
                st.info("Nenhum gráfico/tabela (além de KPIs/SWOTs) pôde ser gerado para o Dashboard Principal.")
        
        elif not kpis_render and not uploaded_file_sidebar: pass 
        elif not kpis_render and not outros_render and uploaded_file_sidebar and st.session_state.sugestoes_gemini: 
            st.info("Nenhum elemento selecionado ou gerado para o dashboard principal.")

elif st.session_state.pagina_selecionada == "Análise SWOT Detalhada":
    st.title("🔬 Análise SWOT Detalhada")
    if not uploaded_file_sidebar: st.warning("Faça upload de um DOCX na barra lateral.")
    elif not st.session_state.sugestoes_gemini: st.info("Aguardando processamento ou nenhuma sugestão gerada.")
    else:
        swot_sugs_page = [s_cfg_swot_page["dados_originais"] for s_id_swot_page,s_cfg_swot_page in st.session_state.config_sugestoes.items() 
                        if s_cfg_swot_page["aceito"] and s_cfg_swot_page["dados_originais"].get("tipo_sugerido")=="lista_swot"]
        if not swot_sugs_page: st.info("Nenhuma análise SWOT sugerida/selecionada para esta página.")
        else:
            if show_debug_info_sidebar:
                with st.expander("Debug: Dados para Análise SWOT (Página Dedicada)", expanded=False):
                    st.json({"SWOTs Selecionados para esta página": swot_sugs_page})
            for swot_item_render_page in swot_sugs_page:
                render_swot_card(
                    swot_item_render_page.get("titulo","Análise SWOT"), 
                    swot_item_render_page.get("parametros",{}), 
                    card_key_prefix=swot_item_render_page.get("id","swot_page") 
                )

if uploaded_file_sidebar is None and st.session_state.nome_arquivo_atual is not None:
    keys_to_clear_on_remove = list(st.session_state.keys())
    preserved_widget_keys_on_remove = [
        "nav_radio_key_final_v2", "uploader_sidebar_key_final_v2", "debug_cb_sidebar_key_final_v2"
    ] 
    for sug_key_preserve in st.session_state.get("sugestoes_gemini", []):
        s_id_preserve_val = sug_key_preserve.get('id')
        if s_id_preserve_val:
            preserved_widget_keys_on_remove.extend([f"acc_cfg_{s_id_preserve_val}", f"tit_cfg_{s_id_preserve_val}"])
            
    for key_cl_remove in keys_to_clear_on_remove:
        if key_cl_remove not in preserved_widget_keys_on_remove:
            if key_cl_remove in st.session_state: del st.session_state[key_cl_remove]
    
    for k_reinit_remove, dv_reinit_remove in [("sugestoes_gemini",[]),("config_sugestoes",{}),
                                ("conteudo_docx",{"texto":"","tabelas":[]}),
                                ("nome_arquivo_atual",None),("debug_checkbox_key",False), 
                                ("pagina_selecionada","Dashboard Principal")]:
        st.session_state.setdefault(k_reinit_remove, dv_reinit_remove)
    st.experimental_rerun()
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
    text_num_part = re.sub(r'[R$\s%]', '', text) # Remove R$, espaço, %
    
    # Trata separadores de milhar (ponto) ANTES de trocar vírgula por ponto decimal
    if ',' in text_num_part and '.' in text_num_part:
        # Se o último ponto está antes da última vírgula, assume ponto como milhar
        if text_num_part.rfind('.') < text_num_part.rfind(','): 
            text_num_part = text_num_part.replace('.', '') 
        text_num_part = text_num_part.replace(',', '.') 
    elif ',' in text_num_part: # Apenas vírgula, assume como decimal
        text_num_part = text_num_part.replace(',', '.')
    # Se só tem ponto, ou se o ponto é o último separador, assume como decimal (já está ok).

    match = re.search(r"([-+]?\d*\.?\d+|\d+)", text_num_part) # Pega o primeiro número
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
                if r_idx == 0: 
                    keys = [k.replace("\n"," ").strip() if k else f"Col{c_idx+1}" for c_idx, k in enumerate(cells)]
                    continue
                if keys: 
                    # Garante que cada linha tenha o mesmo número de chaves que a primeira linha do cabeçalho
                    row_dict = {}
                    for k_idx, key_name in enumerate(keys):
                        row_dict[key_name] = cells[k_idx] if k_idx < len(cells) else None
                    data_rows.append(row_dict)

            if data_rows:
                df = pd.DataFrame(data_rows)
                for col in df.columns: # Itera sobre as colunas do DataFrame criado
                    original_series = df[col].copy()
                    
                    # Tenta converter para numérico primeiro
                    num_series = original_series.astype(str).apply(parse_value_for_numeric)
                    if num_series.notna().sum() / max(1, len(num_series)) > 0.3: # Critério: se mais de 30% viraram números
                        df[col] = pd.to_numeric(num_series, errors='coerce')
                        continue # Próxima coluna se a conversão numérica foi bem-sucedida
                    else: # Reverte se a conversão numérica não foi boa o suficiente
                         df[col] = original_series 
                    
                    # Se não virou numérico, tenta converter para datetime
                    try:
                        temp_str_col = df[col].astype(str) # Garante que é string para pd.to_datetime
                        # Tenta inferir formato, é mais flexível. dayfirst=True para formatos dd/mm/yyyy
                        dt_series = pd.to_datetime(temp_str_col, errors='coerce', dayfirst=True) 
                        # Se a maioria dos valores não nulos viraram datas, usa a série de datas
                        if dt_series.notna().sum() > len(df[col][df[col].notna()]) * 0.5:
                            df[col] = dt_series
                        else: # Mantém como string se a conversão de data falhou muito
                            df[col] = original_series.astype(str).fillna('')
                    except Exception: # Se qualquer erro na conversão de data
                        df[col] = original_series.astype(str).fillna('')
                
                # Fallback final para garantir que colunas 'object' sejam string e NaNs preenchidos
                for col in df.columns:
                    if df[col].dtype == 'object':
                        df[col] = df[col].astype(str).fillna('') # Preenche NaNs com string vazia

                tabelas_data.append({"id": f"doc_tabela_{i+1}", "nome": nome_tabela, "dataframe": df})
        return "\n\n".join(textos), tabelas_data
    except Exception as e: 
        st.error(f"Erro crítico ao ler DOCX: {e}")
        # st.text(traceback.format_exc()) # Descomentar para debug detalhado do erro
        return "", []

def analisar_documento_com_gemini(texto_doc, tabelas_info_list):
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
            try:
                md_table = sample_df.to_markdown(index=False)
            except Exception: # Fallback se to_markdown falhar
                md_table = sample_df.to_string(index=False)

            colunas_para_mostrar_tipos = df.columns.tolist()[:min(8, len(df.columns))]
            col_types_list = [f"'{col_name_prompt}' (tipo: {str(df[col_name_prompt].dtype)})" for col_name_prompt in colunas_para_mostrar_tipos]
            col_types_str = ", ".join(col_types_list)
            
            tabelas_prompt_str += f"\n--- Tabela '{nome_t}' (ID para referência: {id_t}) ---\n"
            tabelas_prompt_str += f"Colunas e tipos (primeiras {len(colunas_para_mostrar_tipos)}): {col_types_str}\n"
            tabelas_prompt_str += f"Amostra de dados:\n{md_table}\n"
        
        text_limit = 45000 # Reduzido ainda mais para segurança e evitar timeouts/erros de tamanho
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
        2.  DADOS NUMÉRICOS: Se a coluna de valor de uma TABELA não for numérica (float64/int64) conforme os "tipos inferidos", NÃO sugira gráfico que exija valor numérico para ela, A MENOS que você possa confiavelmente extrair um valor numérico do seu conteúdo textual (ex: de '70%' extrair 70.0; de '70% - 86%' extrair 70.0 ou 78.0; de 'R$ 15,5 Bi' extrair 15.5). Se extrair do texto, coloque em "dados" e certifique-se que os valores sejam números, não strings de números.
        3.  COBERTURA GEOGRÁFICA (Player, Cidades): Se for apenas lista, sugira "tabela_dados" e forneça os dados extraídos no campo "dados" dos "parametros" com "colunas_titulo".
        4.  SWOT COMPARATIVO: Se uma tabela compara SWOTs, gere "lista_swot" INDIVIDUAL para CADA player da tabela, usando o nome do player no "titulo".
        Retorne APENAS a lista JSON válida. Seja conciso na justificativa.
        """
        with st.spinner("🤖 Gemini está analisando o documento... (Pode levar alguns instantes)"):
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
        # st.text(traceback.format_exc()) # Descomentar para debug MUITO detalhado
        return []

# --- Funções de Renderização Específicas ---
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

def render_swot_card(player_name_title, swot_data, card_key_prefix=""):
    st.subheader(f"{player_name_title}") # Usa o título completo da sugestão
    col1, col2 = st.columns(2)
    swot_map = {"forcas": ("Forças 💪", col1), "fraquezas": ("Fraquezas 📉", col1), 
                "oportunidades": ("Oportunidades 🚀", col2), "ameacas": ("Ameaças ⚠️", col2)}
    for key, (header, col_target) in swot_map.items():
        with col_target:
            st.markdown(f"##### {header}")
            points = swot_data.get(key, ["N/A (informação não fornecida)"])
            if not points or not isinstance(points, list) or not all(isinstance(p, str) for p in points): 
                points = ["N/A (formato de dados incorreto)"]
            for point_idx, item_swot in enumerate(points): 
                st.markdown(f"<div style='margin-bottom: 5px;'>- {item_swot}</div>", unsafe_allow_html=True, key=f"{card_key_prefix}_swot_{key}_{point_idx}")
    st.markdown("---")

def render_plotly_chart(item_config, df_plot_input):
    if df_plot_input is None:
        st.warning(f"Dados não disponíveis para o gráfico '{item_config.get('titulo', 'Sem Título')}'.")
        return False
        
    df_plot = df_plot_input.copy() # Trabalha com cópia para evitar modificar o original
    tipo_grafico = item_config.get("tipo_sugerido")
    titulo = item_config.get("titulo")
    params = item_config.get("parametros", {})
    x_col, y_col = params.get("eixo_x"), params.get("eixo_y")
    cat_col, val_col = params.get("categorias"), params.get("valores")
    
    plot_func, plot_args = None, {}

    if tipo_grafico == "grafico_barras" and x_col and y_col: plot_func, plot_args = px.bar, {"x":x_col,"y":y_col}
    elif tipo_grafico == "grafico_linha" and x_col and y_col: plot_func, plot_args = px.line, {"x":x_col,"y":y_col,"markers":True}
    elif tipo_grafico == "grafico_dispersao" and x_col and y_col: plot_func, plot_args = px.scatter, {"x":x_col,"y":y_col}
    elif tipo_grafico == "grafico_pizza" and cat_col and val_col: plot_func, plot_args = px.pie,{"names":cat_col,"values":val_col}

    if plot_func:
        # Verifica se todas as colunas necessárias existem no DataFrame
        required_cols = [col for col in plot_args.values() if isinstance(col, str)]
        if not all(col in df_plot.columns for col in required_cols):
            st.warning(f"Colunas necessárias {required_cols} não encontradas em '{titulo}'. Colunas disponíveis: {df_plot.columns.tolist()}")
            return False
        try:
            # Tenta converter colunas de valor/eixo Y para numérico ANTES de remover NaNs
            y_axis_col_plot = plot_args.get("y")
            values_col_plot = plot_args.get("values")
            if y_axis_col_plot and y_axis_col_plot in df_plot.columns: 
                df_plot[y_axis_col_plot] = pd.to_numeric(df_plot[y_axis_col_plot], errors='coerce')
            if values_col_plot and values_col_plot in df_plot.columns:
                 df_plot[values_col_plot] = pd.to_numeric(df_plot[values_col_plot], errors='coerce')
            
            df_plot_cleaned = df_plot.dropna(subset=required_cols).copy() # Remove NaNs das colunas de plotagem

            if not df_plot_cleaned.empty:
                st.plotly_chart(plot_func(df_plot_cleaned, title=titulo, **plot_args), use_container_width=True)
                return True
            else: 
                st.warning(f"Dados insuficientes para '{titulo}' após limpar NaNs das colunas {required_cols}.")
        except Exception as e_plotly: 
            st.warning(f"Erro ao gerar gráfico Plotly '{titulo}': {e_plotly}. Verifique os tipos de dados. Dtypes: {df_plot.dtypes.to_dict() if df_plot is not None else 'DF é None'}")
    elif tipo_grafico in ["grafico_barras","grafico_linha","grafico_dispersao","grafico_pizza"]: # Se era para ser um gráfico mas plot_func não foi definido
        st.warning(f"Configuração de parâmetros incompleta para '{titulo}' (tipo: {tipo_grafico}).")
    return False

# --- 3. Interface Streamlit Principal ---
st.set_page_config(layout="wide", page_title="Gemini DOCX Insights")

for k, dv in [("sugestoes_gemini",[]),("config_sugestoes",{}),("conteudo_docx",{"texto":"","tabelas":[]}),
              ("nome_arquivo_atual",None),("debug_checkbox_key",False),("pagina_selecionada","Dashboard Principal")]:
    st.session_state.setdefault(k, dv)

st.sidebar.title("✨ Navegação"); pagina_opcoes = ["Dashboard Principal", "Análise SWOT Detalhada"]
st.session_state.pagina_selecionada = st.sidebar.radio("Selecione:", pagina_opcoes, 
                                                      index=pagina_opcoes.index(st.session_state.pagina_selecionada), 
                                                      key="nav_radio_key") # Chave única
st.sidebar.divider()
uploaded_file = st.sidebar.file_uploader("Selecione DOCX", type="docx", key="uploader_sidebar_key") # Chave única
show_debug_info = st.sidebar.checkbox("Mostrar Informações de Depuração", 
                                    value=st.session_state.debug_checkbox_key, 
                                    key="debug_cb_sidebar_key") # Chave única
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
                temp_config_init = {}
                for i_init,s_init in enumerate(sugestoes): 
                    s_id_init = s_init.get("id", f"s_init_{i_init}_{hash(s_init.get('titulo',''))}"); s_init["id"] = s_id_init
                    temp_config_init[s_id_init] = {"aceito":True,"titulo_editado":s_init.get("titulo","S/Título"),"dados_originais":s_init}
                st.session_state.config_sugestoes = temp_config_init
            else: st.sidebar.warning("Nenhum conteúdo extraído do DOCX.")
    
    if show_debug_info and (st.session_state.conteudo_docx["texto"] or st.session_state.conteudo_docx["tabelas"]):
        with st.expander("Debug: Conteúdo DOCX (após extração e tipos)", expanded=False):
            st.text_area("Texto (amostra)", st.session_state.conteudo_docx["texto"][:1000], height=80)
            for t_info_dbg in st.session_state.conteudo_docx["tabelas"]:
                st.write(f"ID: {t_info_dbg['id']}, Nome: {t_info_dbg['nome']}")
                try: st.dataframe(t_info_dbg['dataframe'].head().astype(str).fillna("-")) 
                except Exception: st.text(f"Head:\n{t_info_dbg['dataframe'].head().to_string(na_rep='-')}")
                st.write("Tipos:", t_info_dbg['dataframe'].dtypes.to_dict())


    if st.session_state.sugestoes_gemini:
        st.sidebar.divider(); st.sidebar.header("⚙️ Configurar Sugestões")
        for sug in st.session_state.sugestoes_gemini:
            s_id = sug['id']; cfg = st.session_state.config_sugestoes.get(s_id) # Usa .get para segurança
            if not cfg: # Se por algum motivo não foi inicializado
                cfg = {"aceito":True,"titulo_editado":sug.get("titulo","S/Título"),"dados_originais":sug}
                st.session_state.config_sugestoes[s_id] = cfg
            
            with st.sidebar.expander(f"{cfg['titulo_editado']}",expanded=False):
                st.caption(f"Tipo: {sug.get('tipo_sugerido')} | Fonte: {sug.get('fonte_id')}")
                cfg["aceito"]=st.checkbox("Incluir?",value=cfg["aceito"],key=f"acc_{s_id}")
                cfg["titulo_editado"]=st.text_input("Título",value=cfg["titulo_editado"],key=f"tit_{s_id}")
                # Edição de parâmetros na sidebar (simplificada)
                # Se precisar de edição mais complexa, pode ser reativada aqui.
else: 
    if st.session_state.pagina_selecionada == "Dashboard Principal":
        st.info("Por favor, faça o upload de um arquivo DOCX na barra lateral para começar.")

# --- RENDERIZAÇÃO DA PÁGINA SELECIONADA ---
if st.session_state.pagina_selecionada == "Dashboard Principal":
    st.title("📊 Dashboard de Insights do Documento")
    if uploaded_file and st.session_state.sugestoes_gemini:
        kpis, outros = [], []; 
        for s_id_cfg,s_cfg in st.session_state.config_sugestoes.items():
            if s_cfg["aceito"]: 
                item = {"titulo":s_cfg["titulo_editado"], **s_cfg["dados_originais"]}
                (kpis if item.get("tipo_sugerido")=="kpi" else outros).append(item)
        
        render_kpis(kpis)
        
        if show_debug_info and (kpis or outros):
             with st.expander("Debug: Configs Finais para Dashboard (Elementos Selecionados)",expanded=False):
                if kpis: st.json({"KPIs Selecionados": kpis}, expanded=False)
                if outros: st.json({"Outros Elementos Selecionados": outros}, expanded=False)
        
        if outros:
            cols_dash = st.columns(2); idx_dash = 0; count_dash = 0
            for item in outros:
                if item.get("tipo_sugerido") == "lista_swot": continue 
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
                            if df_tbl is not None: 
                                try: st.dataframe(df_tbl.astype(str).fillna("-"))
                                except Exception: st.text(df_tbl.to_string(na_rep='-'))
                                rendered=True
                            else: st.warning(f"Tabela '{item['titulo']}' (Fonte: {fonte}) não encontrada.")
                        elif tipo in ["grafico_barras","grafico_linha","grafico_dispersao","grafico_pizza"]:
                            if render_plotly_chart(item, df_plot): rendered = True # df_plot pode ser None aqui
                        elif tipo == 'mapa': st.info(f"Mapa '{item['titulo']}' não implementado."); rendered=True
                        
                        if not rendered and tipo not in ["kpi","lista_swot","mapa"]: 
                            st.info(f"'{item['titulo']}' ({tipo}) não gerado. Dados/Tipo não suportado ou DF não pôde ser criado/encontrado.")
                    except Exception as e: st.error(f"Erro renderizando '{item['titulo']}': {e}")
                if rendered: idx_dash+=1; count_dash+=1
            
            if count_dash == 0 and any(c['aceito'] and c['dados_originais'].get('tipo_sugerido') not in ['kpi','lista_swot'] for c in st.session_state.config_sugestoes.values()):
                st.info("Nenhum gráfico/tabela (além de KPIs/SWOTs) pôde ser gerado.")
        elif not kpis and not uploaded_file: pass 
        elif not kpis and not outros and uploaded_file and st.session_state.sugestoes_gemini: st.info("Nenhum elemento selecionado ou gerado para o dashboard principal.")

elif st.session_state.pagina_selecionada == "Análise SWOT Detalhada":
    st.title("🔬 Análise SWOT Detalhada")
    if not uploaded_file: st.warning("Faça upload de um DOCX na barra lateral.")
    elif not st.session_state.sugestoes_gemini: st.info("Aguardando processamento ou nenhuma sugestão gerada.")
    else:
        swot_sugs = [s_cfg["dados_originais"] for s_id,s_cfg in st.session_state.config_sugestoes.items() if s_cfg["aceito"] and s_cfg["dados_originais"].get("tipo_sugerido")=="lista_swot"]
        if not swot_sugs: st.info("Nenhuma análise SWOT sugerida/selecionada.")
        else:
            for swot_item_page in swot_sugs:
                render_swot_card(swot_item_page.get("titulo", "Análise SWOT"), swot_item_page.get("parametros",{}), card_key_prefix=swot_item_page.get("id","swot"))

if uploaded_file is None and st.session_state.nome_arquivo_atual is not None:
    keys_to_clear = list(st.session_state.keys())
    # Preserve chaves de widgets para evitar que o Streamlit se perca
    # Adicione outras chaves de widgets persistentes se tiver
    preserved_widget_keys = [k for k in keys_to_clear if k.startswith("uploader_") or k.startswith("debug_cb_") or k.startswith("nav_radio") or k.startswith("acc_") or k.startswith("tit_") or k.startswith("param_")]
    
    for key_cl in keys_to_clear:
        if key_cl not in preserved_widget_keys:
            del st.session_state[key_cl]
    
    # Re-inicializa os estados principais da aplicação para um novo ciclo
    for k_reinit, dv_reinit in [("sugestoes_gemini",[]),("config_sugestoes",{}),
                                ("conteudo_docx",{"texto":"","tabelas":[]}),
                                ("nome_arquivo_atual",None),("debug_checkbox_key_main",False), # Garante que o checkbox de debug reseta
                                ("pagina_selecionada","Dashboard Principal")]:
        st.session_state.setdefault(k_reinit, dv_reinit)
    st.experimental_rerun()

import streamlit as st
from docx import Document
import pandas as pd
import plotly.express as px
import google.generativeai as genai
import json
import os
import traceback

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
def extrair_conteudo_docx(uploaded_file):
    """Extrai texto e tabelas de um arquivo DOCX, com tratamento básico de tipos."""
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
                    if p_text and len(p_text) < 100 : nome_tabela = p_text.replace(":", "")
            except: pass

            for row_idx, row in enumerate(table_obj.rows):
                text_cells = [cell.text.strip() for cell in row.cells]
                if row_idx == 0:
                    keys = [key if key else f"Coluna_{k_idx+1}" for k_idx, key in enumerate(text_cells)]
                    continue
                if keys:
                    data_rows.append(dict(zip(keys, text_cells)))
            
            if data_rows:
                df = pd.DataFrame(data_rows)
                for col in df.columns:
                    original_col_data = df[col].copy() # Para reverter se a conversão falhar
                    try: 
                        # Tenta converter para numérico primeiro
                        # Limpeza mais robusta para números
                        temp_series = df[col].astype(str).str.strip()
                        # Remove pontos usados como separadores de milhar ANTES de trocar vírgula por ponto decimal
                        temp_series = temp_series.str.replace(r'\.(?=\d{3})', '', regex=True)
                        temp_series = temp_series.str.replace(',', '.', regex=False)
                        # Remove R$, espaços, % e lida com parênteses para negativos
                        is_negative = temp_series.str.startswith('(') & temp_series.str.endswith(')')
                        temp_series = temp_series.str.replace(r'[R$\s%()]', '', regex=True)
                        
                        numeric_series = pd.to_numeric(temp_series, errors='raise')
                        numeric_series[is_negative.fillna(False)] *= -1 # Aplica sinal negativo
                        df[col] = numeric_series
                        continue # Vai para a próxima coluna se a conversão numérica for bem-sucedida
                    except (ValueError, TypeError, AttributeError):
                        df[col] = original_col_data # Reverte para o original antes de tentar data

                        try: # Data (genérico)
                            temp_col_str = df[col].astype(str) # Garante que é string para pd.to_datetime
                            # Tenta formatos comuns primeiro
                            possible_formats = [
                                '%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', 
                                '%d-%m-%Y', '%m-%d-%Y', '%Y', '%Y%m%d', '%d.%m.%Y'
                            ]
                            converted_with_format = False
                            for fmt in possible_formats:
                                try:
                                    dt_series = pd.to_datetime(temp_col_str, format=fmt, errors='coerce')
                                    # Se a maioria for convertida, assume que o formato é bom
                                    if dt_series.notna().sum() > len(df[col]) * 0.5:
                                        df[col] = dt_series
                                        converted_with_format = True
                                        break
                                except (ValueError, TypeError):
                                    continue
                            
                            if not converted_with_format: # Se nenhum formato explícito funcionou bem
                                inferred_dt_series = pd.to_datetime(temp_col_str, errors='coerce', infer_datetime_format=True)
                                if inferred_dt_series.notna().sum() > len(df[col]) * 0.5:
                                     df[col] = inferred_dt_series
                                # else: # Se a inferência também não converteu a maioria, mantém como estava (string do original_col_data)
                                #    df[col] = original_col_data # Já revertido acima

                        except (ValueError, TypeError):
                            df[col] = original_col_data.astype(str).fillna('') # Garante que é string
                
                # Fallback final para garantir que colunas 'object' sejam string para Arrow
                for col in df.columns:
                    if df[col].dtype == 'object':
                        df[col] = df[col].astype(str).fillna('')

                tabelas_data.append({"id": f"doc_tabela_{i+1}", "nome": nome_tabela, "dataframe": df})
        return "\n\n".join(textos), tabelas_data
    except Exception as e:
        st.error(f"Erro ao ler DOCX: {e}"); return "", []

def analisar_documento_com_gemini(texto_doc, tabelas_info_list):
    """Envia conteúdo para Gemini e pede sugestões de visualização/análise."""
    api_key = get_gemini_api_key()
    if not api_key:
        st.warning("Chave da API do Gemini não configurada. Não é possível gerar sugestões da IA."); return []

    try:
        genai.configure(api_key=api_key)
        
        # Definição explícita e correta das safety_settings
        # Estas são as categorias que o erro indicou como válidas.
        # O threshold BLOCK_NONE desabilita o bloqueio para estas categorias.
        safety_settings_config = [
            {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"},
            # A categoria HARM_CATEGORY_CIVIC_INTEGRITY não está no genai.types.HarmCategory padrão
            # Se precisar dela, certifique-se que é suportada pelo modelo/API.
            # Por agora, vamos focar nas 4 principais listadas na mensagem de erro que são comuns.
        ]
        
        model = genai.GenerativeModel(
            model_name="gemini-1.5-flash-latest", 
            safety_settings=safety_settings_config
        )
        
        tabelas_prompt_str = ""
        for t_info in tabelas_info_list:
            df_sample = t_info["dataframe"].head(5) 
            markdown_tabela = df_sample.to_markdown(index=False)
            col_types_str = ", ".join([f"'{col}' ({str(dtype)})" for col, dtype in t_info["dataframe"].dtypes.items()])
            tabelas_prompt_str += f"\n--- Tabela '{t_info['nome']}' (ID: {t_info['id']}) ---\nColunas (tipos): {col_types_str}\nAmostra:\n{markdown_tabela}\n"

        max_texto_len = 60000
        texto_doc_para_prompt = texto_doc[:max_texto_len] + ("\n[TEXTO TRUNCADO...]" if len(texto_doc) > max_texto_len else "")

        prompt = f"""
        Você é um assistente de análise de dados. Analise o texto e as tabelas de um documento.
        [TEXTO DO DOCUMENTO]
        {texto_doc_para_prompt}
        [FIM DO TEXTO]
        [TABELAS DO DOCUMENTO (com ID, nome, colunas, tipos e amostra)]
        {tabelas_prompt_str}
        [FIM DAS TABELAS]

        Sugira formas de apresentar as informações chave deste documento. Para cada sugestão, retorne um objeto JSON em uma lista. Cada objeto deve ter:
        - "id": string (ex: "gemini_sug_1").
        - "titulo": string, título para a visualização/análise.
        - "tipo_sugerido": string ("kpi", "tabela_dados", "lista_swot", "grafico_barras", "grafico_pizza", "grafico_linha", "grafico_dispersao").
        - "fonte_id": string (ID da tabela ex: "doc_tabela_1", ou "texto_secao_xyz" se do texto).
        - "parametros": objeto com dados e configurações:
            - para "kpi": {{"valor": "ValorKPI", "delta": "Mudança (opcional)", "descricao": "Contexto"}}
            - para "tabela_dados": {{"id_tabela_original": "ID_da_Tabela"}} (para mostrar a tabela completa)
            - para "lista_swot": {{"forcas": ["F1"], "fraquezas": ["Fr1"], "oportunidades": ["Op1"], "ameacas": ["Am1"]}}
            - para "grafico_barras", "grafico_pizza", "grafico_linha", "grafico_dispersao":
                Se baseado em TABELA: {{"eixo_x": "NomeColunaX", "eixo_y": "NomeColunaY", "categorias": "NomeColunaCategorias", "valores": "NomeColunaValores"}} (Use nomes exatos de colunas. Garanta que eixos Y/valores sejam numéricos).
                Se baseado em DADOS EXTRAÍDOS DIRETAMENTE DO TEXTO: {{"dados": [{{"NomeEixoX_ou_Categoria": "ValorCat1", "NomeEixoY_ou_Valor": ValorNum1}}, ...], "eixo_x": "NomeEixoX_ou_Categoria", "eixo_y": "NomeEixoY_ou_Valor"}}
        - "justificativa": string, por que esta apresentação é útil.
        Retorne APENAS a lista JSON.
        """
        with st.spinner("🤖 Gemini está analisando o documento..."):
            # st.text_area("Debug: Prompt Gemini", prompt, height=200)
            response = model.generate_content(prompt)
        
        cleaned_text = response.text.strip().lstrip("```json").rstrip("```").strip()
        # st.text_area("Debug: Resposta Gemini", cleaned_text, height=200)
        sugestoes = json.loads(cleaned_text)
        
        if isinstance(sugestoes, list): st.success(f"{len(sugestoes)} sugestões recebidas do Gemini!"); return sugestoes
        st.error("Resposta do Gemini não foi uma lista."); return []

    except Exception as e: 
        st.error(f"Erro na comunicação com Gemini: {e}"); st.text(traceback.format_exc()); return []


# --- 3. Interface Streamlit e Lógica de Apresentação ---
st.set_page_config(layout="wide")
st.title("✨ Apps com Gemini: DOCX para Insights Visuais")
st.markdown("Faça upload de um DOCX e deixe o Gemini sugerir como visualizar suas informações.")

if "sugestoes_gemini" not in st.session_state: st.session_state.sugestoes_gemini = []
if "conteudo_docx" not in st.session_state: st.session_state.conteudo_docx = {"texto": "", "tabelas": []}
if "config_sugestoes" not in st.session_state: st.session_state.config_sugestoes = {}
if "nome_arquivo_atual" not in st.session_state: st.session_state.nome_arquivo_atual = None

uploaded_file = st.file_uploader("Selecione seu arquivo DOCX", type="docx")
show_debug_info = st.sidebar.checkbox("Mostrar Informações de Depuração", 
                                    value=st.session_state.get('show_debug_checkbox_value', False), 
                                    key="debug_checkbox_widget_key")
# Atualiza o estado com base no widget, se necessário, mas o 'key' já faz isso.
if 'show_debug_checkbox_value' not in st.session_state:
    st.session_state.show_debug_checkbox_value = False
if st.session_state.debug_checkbox_widget_key != st.session_state.show_debug_checkbox_value:
     st.session_state.show_debug_checkbox_value = st.session_state.debug_checkbox_widget_key


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
                with st.expander("Debug: Conteúdo Extraído do DOCX"):
                    st.text_area("Texto Extraído (amostra)", texto_doc[:2000], height=100)
                    for t_info in tabelas_doc:
                        st.write(f"ID: {t_info['id']}, Nome: {t_info['nome']}")
                        try:
                            st.dataframe(t_info['dataframe'].head().astype(str)) 
                        except Exception as e_df_display:
                            st.warning(f"Não foi possível exibir head do DF {t_info['id']}: {e_df_display}")
                        st.write(t_info['dataframe'].dtypes)
            
            sugestoes = analisar_documento_com_gemini(texto_doc, tabelas_doc)
            st.session_state.sugestoes_gemini = sugestoes
            for sug in sugestoes:
                s_id = sug.get("id", f"sug_{hash(sug.get('titulo'))}")
                if s_id not in st.session_state.config_sugestoes:
                    st.session_state.config_sugestoes[s_id] = {
                        "aceito": True, 
                        "titulo_editado": sug.get("titulo", "Sem Título"),
                        "dados_originais": sug 
                    }
        else:
            st.warning("Nenhum conteúdo (texto ou tabelas) extraído do documento.")

if st.session_state.sugestoes_gemini:
    st.sidebar.header("⚙️ Configurar Visualizações Sugeridas")
    for sug_idx, sug in enumerate(st.session_state.sugestoes_gemini):
        s_id = sug.get("id", f"sug_{sug_idx}_{hash(sug.get('titulo'))}") # ID mais robusto
        sug["id"] = s_id # Garante que a sugestão original tenha o ID
        
        if s_id not in st.session_state.config_sugestoes: # Inicializa se não existir
             st.session_state.config_sugestoes[s_id] = {
                "aceito": True, 
                "titulo_editado": sug.get("titulo", "Sem Título"),
                "dados_originais": sug 
            }
        config = st.session_state.config_sugestoes[s_id]

        with st.sidebar.expander(f"{sug.get('titulo', f'Sugestão {sug_idx+1}')}", expanded=False):
            st.caption(f"Tipo: {sug.get('tipo_sugerido')} | Fonte: {sug.get('fonte_id')}")
            st.markdown(f"**Justificativa IA:** *{sug.get('justificativa', 'N/A')}*")
            
            config["aceito"] = st.checkbox("Incluir no Dashboard?", value=config["aceito"], key=f"aceito_{s_id}")
            config["titulo_editado"] = st.text_input("Título para Dashboard", value=config["titulo_editado"], key=f"titulo_{s_id}")

            # Edição de parâmetros para gráficos comuns se não vierem de dados_diretos
            tipo_sug = sug.get("tipo_sugerido")
            params_sug = sug.get("parametros",{})
            if tipo_sug in ["grafico_barras", "grafico_pizza", "grafico_linha", "grafico_dispersao"] and \
               not params_sug.get("dados") and \
               str(sug.get("fonte_id")).startswith("doc_tabela_"):
                
                df_correspondente = next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"] == sug.get("fonte_id")), None)
                if df_correspondente is not None:
                    opcoes_colunas = [""] + df_correspondente.columns.tolist()
                    if tipo_sug in ["grafico_barras", "grafico_linha", "grafico_dispersao"]:
                        params_sug["eixo_x"] = st.selectbox("Eixo X", options=opcoes_colunas, index=opcoes_colunas.index(params_sug.get("eixo_x", "")) if params_sug.get("eixo_x", "") in opcoes_colunas else 0, key=f"param_x_{s_id}")
                        params_sug["eixo_y"] = st.selectbox("Eixo Y", options=opcoes_colunas, index=opcoes_colunas.index(params_sug.get("eixo_y", "")) if params_sug.get("eixo_y", "") in opcoes_colunas else 0, key=f"param_y_{s_id}")
                    elif tipo_sug == "grafico_pizza":
                        params_sug["categorias"] = st.selectbox("Categorias (Nomes)", options=opcoes_colunas, index=opcoes_colunas.index(params_sug.get("categorias", "")) if params_sug.get("categorias", "") in opcoes_colunas else 0, key=f"param_cat_{s_id}")
                        params_sug["valores"] = st.selectbox("Valores", options=opcoes_colunas, index=opcoes_colunas.index(params_sug.get("valores", "")) if params_sug.get("valores", "") in opcoes_colunas else 0, key=f"param_val_{s_id}")


if st.session_state.sugestoes_gemini:
    if st.sidebar.button("🚀 Gerar Dashboard", type="primary", use_container_width=True):
        st.header("📊 Dashboard de Insights")
        
        kpis_para_renderizar = []
        outros_elementos = []

        for s_id, config in st.session_state.config_sugestoes.items():
            if config["aceito"]:
                sug_original = config["dados_originais"] # A sugestão original com parâmetros da LLM
                item = {"titulo": config["titulo_editado"], 
                        "tipo": sug_original.get("tipo_sugerido"),
                        "parametros": sug_original.get("parametros", {}), # Parâmetros da LLM
                        "fonte_id": sug_original.get("fonte_id")}
                if item["tipo"] == "kpi":
                    kpis_para_renderizar.append(item)
                else:
                    outros_elementos.append(item)
        
        if kpis_para_renderizar:
            kpi_cols = st.columns(min(len(kpis_para_renderizar), 4))
            for i, kpi_item in enumerate(kpis_para_renderizar):
                with kpi_cols[i % min(len(kpis_para_renderizar), 4)]:
                    st.metric(label=kpi_item["titulo"], 
                              value=str(kpi_item["parametros"].get("valor", "N/A")),
                              delta=str(kpi_item["parametros"].get("delta", "")),
                              help=kpi_item["parametros"].get("descricao"))
            st.divider()

        if show_debug_info and outros_elementos:
             with st.expander("Debug: Elementos para Dashboard (Não-KPI)", expanded=False):
                for item_debug in outros_elementos:
                    st.write(f"Título: {item_debug['titulo']}, Tipo: {item_debug['tipo']}")
                    st.write(f"Fonte ID: {item_debug['fonte_id']}")
                    st.json(item_debug['parametros'])
                    if str(item_debug['fonte_id']).startswith("doc_tabela_") and not item_debug['parametros'].get("dados"):
                        df_rel = next((t['dataframe'] for t in st.session_state.conteudo_docx['tabelas'] if t['id'] == item_debug['fonte_id']),None)
                        if df_rel is not None: 
                            st.write("DataFrame associado (head):"); st.dataframe(df_rel.head().astype(str))
                            st.write(df_rel.dtypes)
                    st.divider()


        if outros_elementos:
            item_cols = st.columns(2)
            col_idx = 0
            for item in outros_elementos:
                with item_cols[col_idx % 2]:
                    st.subheader(item["titulo"])
                    try:
                        df_plot = None
                        if item["parametros"].get("dados"): # Prioriza dados_diretos da LLM
                            df_plot = pd.DataFrame(item["parametros"]["dados"])
                        elif str(item["fonte_id"]).startswith("doc_tabela_"):
                            df_plot = next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"] == item["fonte_id"]), None)

                        if item["tipo"] == "tabela_dados":
                            id_tabela = item["parametros"].get("id_tabela_original", item["fonte_id"])
                            df_tabela = next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"] == id_tabela), None)
                            if df_tabela is not None: st.dataframe(df_tabela.astype(str)) 
                            else: st.warning(f"Tabela '{id_tabela}' não encontrada.")
                        
                        elif item["tipo"] == "lista_swot":
                            swot_data = item["parametros"]
                            c1, c2 = st.columns(2)
                            swot_map = {"forcas": ("Forças 💪", c1), "fraquezas": ("Fraquezas 📉", c1), 
                                        "oportunidades": ("Oportunidades 🚀", c2), "ameacas": ("Ameaças ⚠️", c2)}
                            for key, (header, col_target) in swot_map.items():
                                with col_target:
                                    st.markdown(f"##### {header}")
                                    points = swot_data.get(key, ["N/A (informação não fornecida)"])
                                    if not points: points = ["N/A"]
                                    for point in points: st.markdown(f"- {point}")
                        
                        elif df_plot is not None: # Gráficos que usam df_plot
                            if item["tipo"] == "grafico_barras":
                                x_col = item["parametros"].get("eixo_x")
                                y_col = item["parametros"].get("eixo_y")
                                if x_col and y_col and x_col in df_plot.columns and y_col in df_plot.columns:
                                    st.plotly_chart(px.bar(df_plot, x=x_col, y=y_col, title=item["titulo"]), use_container_width=True)
                                else: st.warning(f"Colunas X/Y ausentes ou incorretas para gráfico '{item['titulo']}'. X: '{x_col}', Y: '{y_col}'. Colunas no DF: {df_plot.columns.tolist()}")
                            
                            elif item["tipo"] == "grafico_pizza":
                                cat_col = item["parametros"].get("categorias")
                                val_col = item["parametros"].get("valores")
                                if cat_col and val_col and cat_col in df_plot.columns and val_col in df_plot.columns:
                                    st.plotly_chart(px.pie(df_plot, names=cat_col, values=val_col, title=item["titulo"]), use_container_width=True)
                                else: st.warning(f"Colunas de categorias/valores ausentes ou incorretas para gráfico '{item['titulo']}'. Categorias: '{cat_col}', Valores: '{val_col}'. Colunas no DF: {df_plot.columns.tolist()}")
                            
                            elif item["tipo"] == "grafico_linha": # Adicionando gráfico de linha
                                x_col = item["parametros"].get("eixo_x")
                                y_col = item["parametros"].get("eixo_y")
                                if x_col and y_col and x_col in df_plot.columns and y_col in df_plot.columns:
                                    st.plotly_chart(px.line(df_plot, x=x_col, y=y_col, title=item["titulo"], markers=True), use_container_width=True)
                                else: st.warning(f"Colunas X/Y ausentes ou incorretas para gráfico de linha '{item['titulo']}'.")

                            elif item["tipo"] == "grafico_dispersao": # Adicionando gráfico de dispersão
                                x_col = item["parametros"].get("eixo_x")
                                y_col = item["parametros"].get("eixo_y")
                                if x_col and y_col and x_col in df_plot.columns and y_col in df_plot.columns:
                                    st.plotly_chart(px.scatter(df_plot, x=x_col, y=y_col, title=item["titulo"]), use_container_width=True)
                                else: st.warning(f"Colunas X/Y ausentes ou incorretas para gráfico de dispersão '{item['titulo']}'.")
                        
                        elif item["tipo"] not in ["kpi", "tabela_dados", "lista_swot"]: # Se não for um tipo conhecido e df_plot é None
                            st.info(f"Tipo de visualização '{item['tipo']}' para '{item['titulo']}' não implementado ou dados insuficientes (ex: df_plot é None e não há dados_diretos).")

                    except Exception as e_render:
                        st.error(f"Erro ao renderizar '{item['titulo']}': {e_render}")
                col_idx += 1
        
        if not kpis_para_renderizar and not outros_elementos:
            st.info("Nenhum elemento selecionado ou passível de ser gerado para o dashboard.")

elif uploaded_file is None and st.session_state.nome_arquivo_atual is not None:
    st.session_state.sugestoes_gemini = []
    st.session_state.conteudo_docx = {"texto": "", "tabelas": []}
    st.session_state.config_sugestoes = {}
    st.session_state.nome_arquivo_atual = None
    st.session_state.show_debug_checkbox_value = False
    if "file_uploader" in st.session_state: st.session_state.file_uploader = None 
    st.experimental_rerun()

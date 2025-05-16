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
# ... (suas funções extrair_conteudo_docx e analisar_documento_com_gemini permanecem as mesmas) ...
def extrair_conteudo_docx(uploaded_file):
    """Extrai texto e tabelas de um arquivo DOCX, com tratamento básico de tipos."""
    try:
        document = Document(uploaded_file)
        textos = [p.text for p in document.paragraphs if p.text.strip()]
        tabelas_data = [] # Lista de dicionários com {"id", "nome", "dataframe"}

        for i, table_obj in enumerate(document.tables):
            data_rows = []
            keys = None
            nome_tabela = f"Tabela Documento {i+1}" # Nome padrão
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
                    try: 
                        df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', '.', regex=False).str.replace(r'[R$\s%()]', '', regex=True))
                    except (ValueError, TypeError):
                        try: 
                            df[col] = pd.to_datetime(df[col].astype(str), errors='coerce', infer_datetime_format=True)
                        except (ValueError, TypeError):
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
        model = genai.GenerativeModel(model_name="gemini-1.5-flash-latest",
                                      safety_settings=[{"category": c, "threshold": "BLOCK_NONE"} for c in genai.types.HarmCategory])
        
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
        - "tipo_sugerido": string ("kpi", "tabela_dados", "lista_swot", "grafico_barras", "grafico_pizza").
        - "fonte_id": string (ID da tabela ex: "doc_tabela_1", ou "texto_secao_xyz" se do texto).
        - "parametros": objeto com dados e configurações:
            - para "kpi": {{"valor": "ValorKPI", "delta": "Mudança (opcional)", "descricao": "Contexto"}}
            - para "tabela_dados": {{"id_tabela_original": "ID_da_Tabela"}} (para mostrar a tabela completa)
            - para "lista_swot": {{"forcas": ["F1"], "fraquezas": ["Fr1"], "oportunidades": ["Op1"], "ameacas": ["Am1"]}}
            - para "grafico_barras" ou "grafico_pizza" (idealmente de dados extraídos do texto ou tabelas simples): {{"dados": [{{"categoria": "A", "valor": 10}}, {{"categoria": "B", "valor": 20}}], "eixo_categoria": "categoria", "eixo_valor": "valor"}}
        - "justificativa": string, por que esta apresentação é útil.
        Retorne APENAS a lista JSON. Use nomes exatos de colunas das tabelas fornecidas se referenciá-las.
        """
        with st.spinner("🤖 Gemini está analisando o documento..."):
            response = model.generate_content(prompt)
        
        cleaned_text = response.text.strip().lstrip("```json").rstrip("```").strip()
        sugestoes = json.loads(cleaned_text)
        
        if isinstance(sugestoes, list): st.success(f"{len(sugestoes)} sugestões recebidas do Gemini!"); return sugestoes
        st.error("Resposta do Gemini não foi uma lista."); return []

    except Exception as e: 
        st.error(f"Erro na comunicação com Gemini: {e}"); st.text(traceback.format_exc()); return []


# --- 3. Interface Streamlit e Lógica de Apresentação ---
st.set_page_config(layout="wide")
st.title("✨ Apps com Gemini: DOCX para Insights Visuais")
st.markdown("Faça upload de um DOCX e deixe o Gemini sugerir como visualizar suas informações.")

# Gerenciamento de estado
if "sugestoes_gemini" not in st.session_state: st.session_state.sugestoes_gemini = []
if "conteudo_docx" not in st.session_state: st.session_state.conteudo_docx = {"texto": "", "tabelas": []}
if "config_sugestoes" not in st.session_state: st.session_state.config_sugestoes = {}
if "nome_arquivo_atual" not in st.session_state: st.session_state.nome_arquivo_atual = None

uploaded_file = st.file_uploader("Selecione seu arquivo DOCX", type="docx")

# O valor do checkbox é diretamente acessado de st.session_state.show_debug_checkbox
# O 'key' no widget garante que o session_state seja atualizado quando o usuário interage.
show_debug_info = st.sidebar.checkbox(
    "Mostrar Informações de Depuração", 
    value=st.session_state.get('show_debug_checkbox_state', False), # Usar um nome de chave diferente para o estado inicial
    key="show_debug_checkbox_state" # Chave para o widget, que atualiza este item no session_state
)
# A linha abaixo causava o erro e é desnecessária por causa do parâmetro 'key' no checkbox:
# st.session_state.show_debug_checkbox = show_debug_info 


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
            if show_debug_info: # Agora show_debug_info reflete o estado atual do checkbox
                with st.expander("Debug: Conteúdo Extraído do DOCX"):
                    st.text_area("Texto Extraído (amostra)", texto_doc[:2000], height=100)
                    for t_info in tabelas_doc:
                        st.write(f"ID: {t_info['id']}, Nome: {t_info['nome']}")
                        st.dataframe(t_info['dataframe'].head().astype(str)) 
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
    for sug in st.session_state.sugestoes_gemini:
        s_id = sug.get("id", f"sug_{hash(sug.get('titulo'))}")
        config = st.session_state.config_sugestoes.get(s_id)
        if not config: continue

        with st.sidebar.expander(f"{sug.get('titulo', 'Sugestão')}", expanded=False):
            st.caption(f"Tipo: {sug.get('tipo_sugerido')} | Fonte: {sug.get('fonte_id')}")
            st.markdown(f"**Justificativa IA:** *{sug.get('justificativa', 'N/A')}*")
            # A modificação do config["aceito"] e config["titulo_editado"] abaixo
            # já atualiza o st.session_state.config_sugestoes[s_id] diretamente
            # porque 'config' é uma referência a esse dicionário.
            is_aceito = st.checkbox("Incluir no Dashboard?", value=config["aceito"], key=f"aceito_{s_id}")
            titulo_edit = st.text_input("Título para Dashboard", value=config["titulo_editado"], key=f"titulo_{s_id}")
            
            # Atualiza o session_state se houve mudança (importante para evitar re-atribuição desnecessária)
            if config["aceito"] != is_aceito:
                config["aceito"] = is_aceito
            if config["titulo_editado"] != titulo_edit:
                config["titulo_editado"] = titulo_edit


if st.session_state.sugestoes_gemini:
    if st.sidebar.button("🚀 Gerar Dashboard", type="primary", use_container_width=True):
        st.header("📊 Dashboard de Insights")
        
        kpis_para_renderizar = []
        outros_elementos = []

        for s_id, config in st.session_state.config_sugestoes.items():
            if config["aceito"]:
                sug_original = config["dados_originais"]
                item = {"titulo": config["titulo_editado"], 
                        "tipo": sug_original.get("tipo_sugerido"),
                        "parametros": sug_original.get("parametros", {}),
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

        if outros_elementos:
            item_cols = st.columns(2)
            col_idx = 0
            for item in outros_elementos:
                with item_cols[col_idx % 2]:
                    st.subheader(item["titulo"])
                    try:
                        if item["tipo"] == "tabela_dados":
                            id_tabela = item["parametros"].get("id_tabela_original")
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
                                    for point in swot_data.get(key, ["N/A"]): st.markdown(f"- {point}")
                        
                        elif item["tipo"] in ["grafico_barras", "grafico_pizza"] and item["parametros"].get("dados"):
                            df_plot = pd.DataFrame(item["parametros"]["dados"])
                            cat_col = item["parametros"].get("eixo_categoria", df_plot.columns[0] if len(df_plot.columns)>0 else None)
                            val_col = item["parametros"].get("eixo_valor", df_plot.columns[1] if len(df_plot.columns)>1 else None)

                            if cat_col and val_col and cat_col in df_plot.columns and val_col in df_plot.columns:
                                if item["tipo"] == "grafico_barras":
                                    st.plotly_chart(px.bar(df_plot, x=cat_col, y=val_col, title=item["titulo"]), use_container_width=True)
                                elif item["tipo"] == "grafico_pizza":
                                    st.plotly_chart(px.pie(df_plot, names=cat_col, values=val_col, title=item["titulo"]), use_container_width=True)
                            else: st.warning(f"Dados ou colunas ausentes para gráfico '{item['titulo']}'.")
                        
                        else:
                            st.info(f"Tipo de visualização '{item['tipo']}' para '{item['titulo']}' não implementado ou dados insuficientes.")
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
    if "file_uploader" in st.session_state: st.session_state.file_uploader = None 
    st.experimental_rerun()

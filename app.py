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
def parse_value_for_numeric(val_str):
    if pd.isna(val_str) or str(val_str).strip() == '': return None
    text = str(val_str).strip()
    # Tenta extrair o primeiro número, lidando com . e , como decimais/milhares
    # Remove R$, $, %, e espaços internos de números, mas mantém o sinal - no início
    text_num_part = re.sub(r'[R$\s%]', '', text)
    # Heurística para separadores: se tem ',' e não '.', trata ',' como decimal.
    # Se tem '.' e não ',', trata '.' como decimal.
    # Se tem ambos, assume '.' como milhar e ',' como decimal.
    if ',' in text_num_part and '.' in text_num_part:
        text_num_part = text_num_part.replace('.', '') # Remove separador de milhar
        text_num_part = text_num_part.replace(',', '.') # Converte vírgula decimal
    elif ',' in text_num_part:
        text_num_part = text_num_part.replace(',', '.')
    
    # Tenta pegar apenas o primeiro número encontrado (inteiro ou decimal)
    match = re.search(r"([-+]?\d*\.\d+|\d+)", text_num_part)
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
            try: # Tenta nomear a tabela
                prev_el = table_obj._element.getprevious()
                if prev_el is not None and prev_el.tag.endswith('p'):
                    p_text = "".join(node.text for node in prev_el.xpath('.//w:t')).strip()
                    if p_text and len(p_text) < 80: nome_tabela = p_text.replace(":", "").strip()
            except: pass
            for r_idx, row in enumerate(table_obj.rows):
                cells = [c.text.strip() for c in row.cells]
                if r_idx == 0: keys = [k if k else f"Col{c_idx+1}" for c_idx, k in enumerate(cells)]; continue
                if keys: data_rows.append(dict(zip(keys, cells + [None]*(len(keys)-len(cells))))) # Preenche se células faltarem
            if data_rows:
                df = pd.DataFrame(data_rows)
                for col in df.columns:
                    original_series = df[col].copy()
                    # Tenta converter para numérico
                    num_series = original_series.astype(str).apply(parse_value_for_numeric)
                    if num_series.notna().sum() / len(num_series) > 0.3: # Se >30% são números
                        df[col] = pd.to_numeric(num_series, errors='coerce')
                        continue # Próxima coluna
                    else: # Reverte se a conversão numérica não foi boa
                        df[col] = original_series 
                    # Tenta converter para datetime
                    try:
                        temp_str_col = df[col].astype(str)
                        dt_series = pd.to_datetime(temp_str_col, errors='coerce') # Deixa o pandas inferir
                        if dt_series.notna().sum() / len(dt_series) > 0.5: # Se >50% são datas
                            df[col] = dt_series
                        else: # Mantém como string se a conversão de data falhou muito
                            df[col] = original_series.astype(str).fillna('')
                    except: df[col] = original_series.astype(str).fillna('')
                # Fallback final para string
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
            sample_df = df.head(5).iloc[:, :min(8, len(df.columns))] # Amostra menor
            md_table = sample_df.to_markdown(index=False)
            col_types = ", ".join([f"'{c}' ({str(d)})" for c,d in df.dtypes.items()])
            tabelas_prompt_str += f"\n--- Tabela '{nome_t}' (ID: {id_t}) ---\nColunas/Tipos: {col_types}\nAmostra:\n{md_table}\n"
        
        text_limit = 50000
        prompt_text = texto_doc[:text_limit] + ("\n[TEXTO TRUNCADO]" if len(texto_doc) > text_limit else "")
        
        prompt = f"""
        Você é um assistente de análise de dados. Analise o texto e as tabelas de um documento.
        [TEXTO]{prompt_text}[FIM TEXTO]
        [TABELAS]{tabelas_prompt_str}[FIM TABELAS]

        Gere uma lista JSON de sugestões de visualizações. CADA objeto DEVE ter: "id", "titulo", "tipo_sugerido" ("kpi", "tabela_dados", "lista_swot", "grafico_barras", "grafico_pizza", "grafico_linha", "grafico_dispersao"), "fonte_id" (ID tabela ou "texto_descricao_fonte"), "parametros" (objeto JSON), "justificativa".
        Para "parametros":
        - "kpi": {{"valor": "ValorKPI", "delta": "Mudança", "descricao": "Contexto"}}
        - "tabela_dados": Para TABELA EXISTENTE: {{"id_tabela_original": "ID_Tabela"}}. Para DADOS DO TEXTO: {{"dados": [{{"Coluna1": "ValorA1", "Coluna2": "ValorA2"}}, ...], "colunas_titulo": ["Título Col1", "Título Col2"]}}
        - "lista_swot": {{"forcas": ["F1"], "fraquezas": ["Fr1"], "oportunidades": ["Op1"], "ameacas": ["Am1"]}}
        - Gráficos de TABELA ("barras", "linha", "dispersao"): {{"eixo_x": "NOME_COL_X", "eixo_y": "NOME_COL_Y"}} (Y numérico).
        - Gráficos de PIZZA de TABELA: {{"categorias": "NOME_COL_CAT", "valores": "NOME_COL_VAL_NUM"}} (Valores numéricos).
        - Gráficos com DADOS EXTRAÍDOS DO TEXTO: {{"dados": [{{"NomeEixoX": "CatA", "NomeEixoY": ValNumA}}, ...], "eixo_x": "NomeEixoX", "eixo_y": "NomeEixoY"}} (Valores DEVEM ser numéricos).
        Use NOMES EXATOS de colunas. Se coluna de valor não for numérica (ex: '70% - 80%'), EXTRAIA valor numérico (ex: 70.0 ou média). Para '17,35 Bilhões', extraia 17.35. Se não tratável, NÃO SUGIRA gráfico que exija número. Para Cobertura Geográfica (Player, Cidades), sugira "tabela_dados" com campo "dados" se extrair do texto. Para SWOTs comparativos, gere "lista_swot" INDIVIDUAL por player.
        Retorne APENAS a lista JSON.
        """
        with st.spinner("🤖 Gemini analisando..."):
            # st.text_area("Debug Prompt:", prompt, height=200) # Para depuração
            response = model.generate_content(prompt)
        cleaned_text = response.text.strip().lstrip("```json").rstrip("```").strip()
        # st.text_area("Debug Resposta Gemini:", cleaned_text, height=200) # Para depuração
        sugestoes = json.loads(cleaned_text)
        if isinstance(sugestoes, list): st.success(f"{len(sugestoes)} sugestões!"); return sugestoes
        st.error("Resposta Gemini não é lista JSON."); return []
    except json.JSONDecodeError as e: st.error(f"Erro JSON Gemini: {e}"); st.code(response.text if 'response' in locals() else "N/A", language="text"); return []
    except Exception as e: st.error(f"Erro API Gemini: {e}"); st.text(traceback.format_exc()); return []

# --- 3. Interface Streamlit e Lógica de Apresentação ---
st.set_page_config(layout="wide"); st.title("✨ Gemini: DOCX para Insights Visuais")
st.markdown("Upload DOCX para sugestões de visualização pela IA.")

# Inicialização de estado
for k in ["sugestoes_gemini", "config_sugestoes"]: st.session_state.setdefault(k, [] if k == "sugestoes_gemini" else {})
st.session_state.setdefault("conteudo_docx", {"texto": "", "tabelas": []})
st.session_state.setdefault("nome_arquivo_atual", None)
st.session_state.setdefault("debug_checkbox", False)

uploaded_file = st.file_uploader("Selecione DOCX", type="docx", key="uploader")
st.session_state.debug_checkbox = st.sidebar.checkbox("Mostrar Debug Info", value=st.session_state.debug_checkbox, key="debug_cb")

if uploaded_file:
    if st.session_state.nome_arquivo_atual != uploaded_file.name: 
        st.session_state.sugestoes_gemini, st.session_state.config_sugestoes = [], {}
        st.session_state.nome_arquivo_atual = uploaded_file.name
    if not st.session_state.sugestoes_gemini: 
        texto_doc, tabelas_doc = extrair_conteudo_docx(uploaded_file)
        st.session_state.conteudo_docx = {"texto": texto_doc, "tabelas": tabelas_doc}
        if texto_doc or tabelas_doc:
            st.success(f"'{uploaded_file.name}' lido.")
            if st.session_state.debug_checkbox:
                with st.expander("Debug: Conteúdo DOCX (após extração e tipos)", expanded=False):
                    st.text_area("Texto (amostra)", texto_doc[:1000], height=80)
                    for t_info_dbg in tabelas_doc:
                        st.write(f"ID: {t_info_dbg['id']}, Nome: {t_info_dbg['nome']}")
                        try: st.dataframe(t_info_dbg['dataframe'].head().astype(str))
                        except: st.text(f"Head:\n{t_info_dbg['dataframe'].head().to_string()}")
                        st.write("Tipos:", t_info_dbg['dataframe'].dtypes.to_dict())
            sugestoes = analisar_documento_com_gemini(texto_doc, tabelas_doc)
            st.session_state.sugestoes_gemini = sugestoes
            for i,s in enumerate(sugestoes): s["id"]=s.get("id",f"s_{i}_{hash(s.get('titulo'))}") # Garante ID
            st.session_state.config_sugestoes={s['id']:{"aceito":True,"titulo_editado":s.get("titulo","S/Título"),"dados_originais":s} for s in sugestoes}
        else: st.warning("Nenhum conteúdo extraído.")

if st.session_state.sugestoes_gemini:
    st.sidebar.header("⚙️ Configurar Sugestões")
    for sug in st.session_state.sugestoes_gemini:
        s_id, cfg = sug['id'], st.session_state.config_sugestoes[sug['id']]
        with st.sidebar.expander(f"{cfg['titulo_editado']}", expanded=False):
            st.caption(f"Tipo: {sug.get('tipo_sugerido')} | Fonte: {sug.get('fonte_id')}")
            st.markdown(f"**IA:** *{sug.get('justificativa', 'N/A')}*")
            cfg["aceito"]=st.checkbox("Incluir?",value=cfg["aceito"],key=f"acc_{s_id}")
            cfg["titulo_editado"]=st.text_input("Título",value=cfg["titulo_editado"],key=f"tit_{s_id}")
            # Edição de parâmetros se gráfico de tabela e sem dados_diretos
            # (Lógica de edição de parâmetros simplificada para brevidade, pode ser expandida)
    if st.sidebar.button("🚀 Gerar Dashboard", type="primary", use_container_width=True):
        st.header("📊 Dashboard"); kpis, outros = [], []
        for s_id, cfg in st.session_state.config_sugestoes.items():
            if cfg["aceito"]: item = {"titulo":cfg["titulo_editado"], **cfg["dados_originais"]}; (kpis if item["tipo_sugerido"]=="kpi" else outros).append(item)
        if kpis:
            cols=st.columns(min(len(kpis),4)); [c.metric(k["titulo"],str(k.get("parametros",{}).get("valor","N/A")),str(k.get("parametros",{}).get("delta","")),help=k.get("parametros",{}).get("descricao")) for i,k in enumerate(kpis) for c in [cols[i%min(len(kpis),4)]]]
            if outros: st.divider()
        if st.session_state.debug_checkbox and (kpis or outros):
             with st.expander("Debug: Configs Finais para Dashboard",expanded=False): st.json({"KPIs":kpis,"Outros":outros}) # Mostra parâmetros finais
        if outros:
            item_cols=st.columns(2); col_idx=0
            for item in outros:
                with item_cols[col_idx%2]:
                    st.subheader(item["titulo"]); df_plot, el_rend = None, False
                    params=item.get("parametros",{}); tipo=item.get("tipo_sugerido"); fonte=item.get("fonte_id")
                    try:
                        if params.get("dados"): df_plot=pd.DataFrame(params["dados"])
                        elif str(fonte).startswith("doc_tabela_"): df_plot=next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"]==fonte),None)
                        
                        if tipo=="tabela_dados":
                            id_t=params.get("id_tabela_original",fonte)
                            # Se fonte é textual e tem dados, usa eles. Senão, busca tabela por ID.
                            if str(fonte).startswith("texto_") and params.get("dados"):
                                df_t = pd.DataFrame(params.get("dados"))
                                if params.get("colunas_titulo"): df_t.columns = params.get("colunas_titulo")
                            else:
                                df_t=next((t["dataframe"] for t in st.session_state.conteudo_docx["tabelas"] if t["id"]==id_t),None)
                            if df_t is not None: st.dataframe(df_t.astype(str).fillna("-")); el_rend=True
                            else: st.warning(f"Tabela '{id_t}' não encontrada para '{item['titulo']}'.")
                        elif tipo=="lista_swot":
                            c1,c2=st.columns(2); smap={"forcas":("Forças 💪",c1),"fraquezas":("Fraquezas 📉",c1),"oportunidades":("Oportunidades 🚀",c2),"ameacas":("Ameaças ⚠️",c2)}
                            for k,(h,ct) in smap.items():
                                with ct: st.markdown(f"##### {h}"); [st.markdown(f"- {p}") for p in params.get(k,["N/A"])]
                            el_rend=True
                        elif df_plot is not None:
                            x,y,cat,val=params.get("eixo_x"),params.get("eixo_y"),params.get("categorias"),params.get("valores")
                            fn,p_args=None,{}
                            if tipo=="grafico_barras" and x and y: fn,p_args=px.bar,{"x":x,"y":y}
                            elif tipo=="grafico_linha" and x and y: fn,p_args=px.line,{"x":x,"y":y,"markers":True}
                            elif tipo=="grafico_dispersao" and x and y: fn,p_args=px.scatter,{"x":x,"y":y}
                            elif tipo=="grafico_pizza" and cat and val: fn,p_args=px.pie,{"names":cat,"values":val}
                            if fn and all(k_col in df_plot.columns for k_col in p_args.values() if isinstance(k_col,str)):
                                st.plotly_chart(fn(df_plot,title=item["titulo"],**p_args),use_container_width=True); el_rend=True
                            elif fn: st.warning(f"Colunas X/Y ou Cat/Val ausentes/incorretas para '{item['titulo']}'. Params: {p_args}. Cols DF: {df_plot.columns.tolist()}")
                        if tipo=='mapa': st.info(f"Mapa para '{item['titulo']}' não implementado."); el_rend=True
                        if not el_rend and tipo not in ["kpi","tabela_dados","lista_swot","mapa"]:
                            st.info(f"'{item['titulo']}' (tipo: {tipo}) não gerado. Dados/Tipo não suportado.")
                    except Exception as e: st.error(f"Erro renderizando '{item['titulo']}': {e}")
                if el_rend: col_idx+=1
        if not kpis and not outros: st.info("Nenhum elemento para o dashboard.")
elif uploaded_file is None and st.session_state.nome_arquivo_atual is not None:
    for key_to_clear in list(st.session_state.keys()): # Limpa todo o session_state
        del st.session_state[key_to_clear]
    st.experimental_rerun()

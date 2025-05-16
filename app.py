import streamlit as st
from docx import Document
import pandas as pd
import plotly.express as px
import google.generativeai as genai
import json
import os
import traceback
import re # Para limpeza numérica mais avançada

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
def clean_and_convert_to_numeric(series):
    """Tenta limpar e converter uma série Pandas para numérico."""
    if series.dtype == 'object' or isinstance(series.dtype, pd.StringDtype):
        # Converte para string para garantir que métodos .str funcionem
        s = series.astype(str).str.strip()
        
        # Guardar informação de 'Bilhões' ou 'Milhões' para multiplicar depois, se necessário (simplificado)
        # Esta parte pode ser expandida para maior precisão
        multiplier = pd.Series([1.0] * len(s), index=s.index)
        multiplier[s.str.contains("bilh(ão|ões)", case=False, na=False)] = 1_000_000_000
        multiplier[s.str.contains("milh(ão|ões)", case=False, na=False)] = 1_000_000
        
        # Remove palavras como Bilhões/Milhões, R$, % etc.
        s = s.str.replace(r"(R\$|\$|Bilhões|Bilhão|Milhões|Milhão|%|\s)", "", regex=True)
        
        # Trata separador de milhar (ponto) ANTES de trocar vírgula por ponto decimal
        s = s.str.replace(r'\.(?=\d{3})', '', regex=True) 
        s = s.str.replace(',', '.', regex=False)
        
        # Lida com negativos em parênteses (ex: (123.45) -> -123.45)
        is_negative_paren = s.str.startswith('(') & s.str.endswith(')')
        s_num = s.str.replace(r'[()]', '', regex=True) # Remove parênteses
        
        numeric_col = pd.to_numeric(s_num, errors='coerce')
        numeric_col[is_negative_paren.fillna(False)] *= -1
        
        # Aplica multiplicador (simplificado, assume que o número já foi extraído)
        # numeric_col *= multiplier 
        # Nota: a multiplicação por bilhão/milhão aqui é complexa se o texto "Bilhão" já foi removido.
        # Seria melhor tratar isso antes, ou deixar a LLM interpretar o texto completo.
        # Por agora, focaremos na conversão do número em si.
        
        return numeric_col
    return pd.to_numeric(series, errors='coerce') # Tenta converter diretamente se não for objeto/string

def extrair_conteudo_docx(uploaded_file):
    """Extrai texto e tabelas de um arquivo DOCX, com tratamento de tipos aprimorado."""
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
                    if p_text and len(p_text) < 100 : nome_tabela = p_text.replace(":", "").strip()
            except: pass

            for row_idx, row in enumerate(table_obj.rows):
                text_cells = [cell.text.strip() for cell in row.cells]
                if row_idx == 0:
                    keys = [key.replace("\n", " ").strip() if key else f"Coluna_{k_idx+1}" for k_idx, key in enumerate(text_cells)]
                    continue
                if keys: # Garante que keys foi definido (primeira linha era cabeçalho)
                    # Assegura que o número de chaves corresponde ao número de células, preenchendo se necessário
                    row_data = {}
                    for k_idx, key_name in enumerate(keys):
                         row_data[key_name] = text_cells[k_idx] if k_idx < len(text_cells) else None
                    data_rows.append(row_data)
            
            if data_rows:
                df = pd.DataFrame(data_rows)
                for col in df.columns:
                    original_col_data = df[col].copy()
                    
                    # 1. Tentar converter para numérico usando a função de limpeza
                    converted_numeric = clean_and_convert_to_numeric(df[col])
                    if converted_numeric.notna().sum() > len(df[col]) * 0.3: # Se pelo menos 30% virou número
                        df[col] = converted_numeric
                        continue
                    else: # Reverte se a conversão numérica não foi muito bem-sucedida
                         df[col] = original_col_data.copy()

                    # 2. Tentar converter para datetime (se não for numérico)
                    try:
                        temp_col_str = df[col].astype(str)
                        possible_formats = [
                            '%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', 
                            '%d-%m-%Y', '%m-%d-%Y', '%Y', '%Y%m%d', '%d.%m.%Y',
                            # Adicionar formatos que podem incluir apenas ano-mês ou ano
                             '%Y-%m', '%m-%Y' 
                        ]
                        converted_with_format = False
                        for fmt in possible_formats:
                            try:
                                dt_series = pd.to_datetime(temp_col_str, format=fmt, errors='coerce')
                                if dt_series.notna().sum() > len(df[col]) * 0.5:
                                    df[col] = dt_series
                                    converted_with_format = True
                                    break
                            except (ValueError, TypeError):
                                continue
                        
                        if not converted_with_format:
                            inferred_dt_series = pd.to_datetime(temp_col_str, errors='coerce', infer_datetime_format=True) # infer_datetime_format é deprecated mas ainda funciona
                            if inferred_dt_series.notna().sum() > len(df[col]) * 0.5:
                                 df[col] = inferred_dt_series
                            else: # Mantém como string se a inferência falhou muito
                                df[col] = original_col_data.astype(str).fillna('')
                    except Exception:
                        df[col] = original_col_data.astype(str).fillna('')
                
                # Fallback final para garantir que colunas 'object' sejam string
                for col in df.columns:
                    if df[col].dtype == 'object':
                        df[col] = df[col].astype(str).fillna('')

                tabelas_data.append({"id": f"doc_tabela_{i+1}", "nome": nome_tabela, "dataframe": df})
        return "\n\n".join(textos), tabelas_data
    except Exception as e:
        st.error(f"Erro crítico ao ler DOCX: {e}"); return "", []

def analisar_documento_com_gemini(texto_doc, tabelas_info_list):
    api_key = get_gemini_api_key()
    if not api_key:
        st.warning("Chave da API do Gemini não configurada. Não é possível gerar sugestões da IA."); return []
    try:
        genai.configure(api_key=api_key)
        safety_settings_config = [
            {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"},
        ]
        model = genai.GenerativeModel(model_name="gemini-1.5-flash-latest", safety_settings=safety_settings_config)
        
        tabelas_prompt_str = ""
        for t_info in tabelas_info_list:
            df = t_info["dataframe"]
            df_sample = df.head(7) 
            if len(df.columns) > 10: df_sample = df_sample.iloc[:, :10]
            markdown_tabela = df_sample.to_markdown(index=False)
            tabelas_prompt_str += f"\n--- Tabela '{t_info['nome']}' (ID para referência: {t_info['id']}) ---\n"
            col_types_str = ", ".join([f"'{col}' (tipo inferido: {str(dtype)})" for col, dtype in df.dtypes.items()])
            tabelas_prompt_str += f"Colunas e tipos: {col_types_str}\nAmostra:\n{markdown_tabela}\n"

        max_texto_len = 60000
        texto_doc_para_prompt = texto_doc[:max_texto_len] + ("\n[TEXTO TRUNCADO...]" if len(texto_doc) > max_texto_len else "")

        prompt = f"""
        Você é um assistente de análise de dados e visualização. Analise o texto e as tabelas de um documento.
        [TEXTO DO DOCUMENTO]
        {texto_doc_para_prompt}
        [FIM DO TEXTO]
        [TABELAS DO DOCUMENTO (com ID, nome, colunas, tipos e amostra)]
        {tabelas_prompt_str}
        [FIM DAS TABELAS]

        Gere uma lista JSON de sugestões de visualizações. CADA objeto na lista DEVE ter:
        - "id": string (ex: "gemini_sug_1").
        - "titulo": string (título para a visualização).
        - "tipo_sugerido": string ("kpi", "tabela_dados", "lista_swot", "grafico_barras", "grafico_pizza", "grafico_linha", "grafico_dispersao").
        - "fonte_id": string (ID da tabela ex: "doc_tabela_1", ou descrição textual da fonte ex: "texto_sumario_executivo").
        - "parametros": objeto JSON com dados e configurações. É CRUCIAL usar NOMES EXATOS de colunas das tabelas como fornecidos.
            - Para "kpi": {{"valor": "ValorDoKPI", "delta": "Mudança (opcional)", "descricao": "Contexto do KPI"}}
            - Para "tabela_dados": {{"id_tabela_original": "ID_da_Tabela_Referenciada"}}
            - Para "lista_swot": {{"forcas": ["Ponto Força 1"], "fraquezas": ["Ponto Fraqueza 1"], "oportunidades": ["Ponto Oportunidade 1"], "ameacas": ["Ponto Ameaça 1"]}}
            - Para "grafico_barras", "grafico_linha", "grafico_dispersao":
                Se baseado em TABELA (use o "fonte_id" da tabela): {{"eixo_x": "NOME_EXATO_COLUNA_X", "eixo_y": "NOME_EXATO_COLUNA_Y"}} (eixo_y geralmente numérico).
                Se DADOS EXTRAÍDOS DO TEXTO: {{"dados": [{{"NomeQualquerParaEixoX": "CategoriaA", "NomeQualquerParaEixoY": ValorNumericoA}}, ...], "eixo_x": "NomeQualquerParaEixoX", "eixo_y": "NomeQualquerParaEixoY"}}
            - Para "grafico_pizza":
                Se baseado em TABELA: {{"categorias": "NOME_EXATO_COLUNA_CATEGORIAS", "valores": "NOME_EXATO_COLUNA_VALORES_NUMERICOS"}}
                Se DADOS EXTRAÍDOS DO TEXTO: {{"dados": [{{"NomeQualquerParaCategoria": "CategoriaA", "NomeQualquerParaValor": ValorNumericoA}}, ...], "categorias": "NomeQualquerParaCategoria", "valores": "NomeQualquerParaValor"}}
        - "justificativa": string (breve explicação da utilidade).

        Exemplos de `parametros` para GRÁFICOS DE TABELAS (assumindo que a tabela 'doc_tabela_1' tem colunas 'Ano' e 'Vendas'):
        - Gráfico de Linha: {{"eixo_x": "Ano", "eixo_y": "Vendas"}}
        - Gráfico de Barras: {{"eixo_x": "Ano", "eixo_y": "Vendas"}}

        Exemplo de `parametros` para GRÁFICO DE PIZZA COM DADOS EXTRAÍDOS DO TEXTO:
        {{"dados": [{{"Região": "Norte", "Faturamento": 50000}}, {{"Região": "Sul", "Faturamento": 75000}}], "categorias": "Região", "valores": "Faturamento"}}

        Certifique-se de que os NOMES DE COLUNAS nos 'parametros' correspondam EXATAMENTE aos nomes de colunas e tipos de dados fornecidos na descrição das tabelas. Se uma coluna de valor não for numérica (ex: 'object' contendo '70% - 80%'), instrua para extrair um valor numérico representativo (ex: média, ou o primeiro número) se possível, ou não sugira o gráfico se não for viável tratar como numérico.
        Retorne APENAS a lista JSON válida.
        """
        with st.spinner("🤖 Gemini está analisando o documento..."):
            # st.text_area("Debug: Prompt Gemini", prompt, height=200)
            response = model.generate_content(prompt)
        cleaned_text = response.text.strip().lstrip("```json").rstrip("```").strip()
        # st.text_area("Debug: Resposta Gemini", cleaned_text, height=200)
        sugestoes = json.loads(cleaned_text)
        if isinstance(sugestoes, list): st.success(f"{len(sugestoes)} sugestões recebidas do Gemini!"); return sugestoes
        st.error("Resposta do Gemini não foi uma lista."); return []
    except json.JSONDecodeError as e:
        st.error(f"Erro ao decodificar JSON do Gemini: {e}")
        if 'response' in locals() and hasattr(response, 'text'): st.code(response.text, language="text")
        return []
    except Exception as e: st.error(f"Erro na comunicação com Gemini: {e}"); st.text(traceback.format_exc()); return []

# --- 3. Interface Streamlit e Lógica de Apresentação ---
# ... (O restante do código da interface do Streamlit permanece o mesmo da versão anterior) ...
# COPIE O RESTANTE DO CÓDIGO A PARTIR DA LINHA:
# st.set_page_config(layout="wide")
# ATÉ O FINAL DO ARQUIVO DA ÚLTIMA VERSÃO COMPLETA QUE ENVIEI.
# (As funções get_gemini_api_key, extrair_conteudo_docx, e analisar_documento_com_gemini foram as únicas modificadas nesta resposta)

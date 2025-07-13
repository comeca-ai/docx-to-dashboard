# Docx to Dashboard - Análise Profunda Multi-Agente

Este aplicativo Streamlit permite fazer upload de documentos DOCX e gerar dashboards automatizados com insights através da API do Google Gemini.

## ✨ Novas Funcionalidades - Sistema Multi-Agente

### Análise Profunda Multi-Agente
O sistema agora utiliza múltiplos agentes especializados para fornecer análises mais profundas e abrangentes:

#### 🔢 Agente de Análise de Dados
- Identifica KPIs críticos com níveis de criticidade
- Detecta tendências e padrões nos dados
- Encontra correlações entre variáveis
- Identifica outliers e anomalias
- Gera insights quantitativos profundos

#### 🎯 Agente de Análise Estratégica
- Análise SWOT detalhada e contextualizada
- Recomendações acionáveis com priorização
- Identificação de cenários futuros
- Fatores críticos de sucesso
- Análise de riscos e oportunidades

#### 🧠 Agente Sintetizador
- Combina insights de todos os agentes
- Cria conexões entre análises quantitativas e estratégicas
- Prioriza ações baseadas em dados
- Gera roadmap de implementação (imediato, 30 dias, 90 dias)

## 📊 Funcionalidades da Interface

### 1. Dashboard Principal
- KPIs visuais interativos
- Gráficos e tabelas dinâmicas
- Indicador de análise profunda disponível

### 2. Análise SWOT Detalhada
- Análises SWOT estruturadas
- Visualização organizada por categorias

### 3. **NOVO** Análise Profunda Multi-Agente
- Síntese executiva dos principais insights
- KPIs críticos com indicadores de prioridade
- Análise de tendências e correlações
- Recomendações acionáveis priorizadas
- Roadmap de implementação estruturado

## 🚀 Como Usar

1. **Upload**: Carregue um documento DOCX na barra lateral
2. **Processamento**: O sistema executa automaticamente:
   - Extração de texto e tabelas
   - Análise com Gemini para visualizações
   - **NOVO**: Análise profunda multi-agente
3. **Navegação**: Use a barra lateral para alternar entre:
   - Dashboard Principal
   - Análise SWOT Detalhada
   - **NOVO**: Análise Profunda Multi-Agente

## 🔧 Configuração

- Configure a variável de ambiente `GOOGLE_API_KEY` com sua chave da API do Google Gemini
- Ou adicione a chave em `st.secrets["GOOGLE_API_KEY"]`

## 🎨 Melhorias da Interface

- **Indicadores visuais** de criticidade e prioridade
- **Códigos de cores** para diferentes tipos de insights
- **Layout organizado** em colunas para melhor visualização
- **Roadmap temporal** estruturado
- **Métricas interativas** com tooltips explicativos

## 🧪 Sistema Multi-Agente

O novo sistema executa análises especializadas em paralelo:

1. **Agente de Dados** → Análise quantitativa profunda
2. **Agente Estratégico** → Insights de negócio e estratégia
3. **Agente Sintetizador** → Integração e priorização

Cada agente utiliza prompts especializados para maximizar a qualidade da análise em sua área de expertise.
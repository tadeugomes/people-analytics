# Sistema de Análise de Diversidade Dinâmica

## Sumário
- Visão geral e fluxo
- Estrutura de abas e arquivos
- Como gerar e atualizar (pipeline Python)
- Como usar a versão Excel‑only (sem Python)
- Macro e caminhos do Python
- Limiar/cores e interpretações
- Troubleshooting (erros comuns)

Este sistema foi desenvolvido para analisar automaticamente dados de diversidade em CSV e gerar relatórios Excel com estatísticas compreensíveis para usuários leigos.

## 🔄 Novo Fluxo de Trabalho

### Estrutura do Excel
1. **DADOS_BRUTOS** (Primeira aba): Seus dados originais ficam aqui
2. **Abas automatizadas**: Todas as análises são geradas automaticamente

### Como Funciona
- **Adicione novos dados** diretamente na aba `DADOS_BRUTOS`
- **Execute o atualizador** para gerar todas as análises automaticamente
- **Sem configuração manual**: O sistema detecta automaticamente as mudanças

## Funcionalidades Principais

### 🔄 Análise Dinâmica Automática
- **Detecção automática de tipos de dados**: O sistema identifica automaticamente se cada coluna é categórica ou numérica
- **Atualização dinâmica de campos**: Quando novos dados são inseridos, o sistema se adapta automaticamente
- **Sem configuração manual necessária**: Funciona com qualquer CSV que tenha estrutura similar

### 📊 Análises Estatísticas Completas

#### 1. Estatísticas Descritivas
- Para variáveis categóricas: frequências, percentagens, valores mais comuns
- Para variáveis numéricas: média, mediana, desvio padrão, quartis

#### 2. Testes Estatísticos com Explicações
- **Teste Qui-quadrado**: Verifica associação entre variáveis categóricas
- **Interpretação em linguagem simples**: Cada teste vem com explicação clara do que significa

#### 3. Índices de Diversidade
- **Índice de Simpson**: Mede a diversidade (0-1, onde 1 é máxima diversidade)
- **Índice de Shannon**: Outra medida de diversidade
- **Índice de Dominância**: Mostra quanto uma categoria domina sobre as outras

#### 4. Visualizações Automáticas no Excel
- Gráficos de barras integrados diretamente na planilha Excel
- Percentagens incluídas nos gráficos e em tabelas
- Formatação profissional com rótulos de dados
- Gráficos interativos dentro do próprio Excel

### 📋 Relatório Excel Estruturado

O relatório gerado contém as seguintes abas:

0. **0_HOME**: capa com instruções, atalho para abas e "Última atualização"
1. **DADOS_BRUTOS**: seus dados originais em formato de Tabela do Excel (não é apagada)
2. **1_VISAO_GERAL**: informações sobre cada variável
3. (Removida) Estatísticas descritivas resumidas — incorporamos métricas onde necessário
4. **3_TESTES_ESTATISTICOS**: resultados dos testes com interpretações
5. **4_INDICES_DIVERSIDADE**: medidas de diversidade por variável
6. **5_TABELA_[VARIAVEL]**: tabelas de frequência para cada variável
7. **6_VISUALIZACOES**: gráficos de barras integrados no Excel com tabelas de dados
8. **LISTAS**: listas-mestres para validação de dados (categorias por coluna)
9. **AJUDA_GLOSSARIO**: explicações dos índices e testes, boas práticas

## Como Usar

### Pré-requisitos
- Python 3.7 ou superior
- Pacotes necessários: pandas, openpyxl, matplotlib, seaborn, scipy, statsmodels

### Instalação
```bash
# Criar ambiente virtual
python3 -m venv venv
source venv/bin/activate  # Linux/Mac
# ou
venv\Scripts\activate  # Windows

# Instalar pacotes
pip install pandas openpyxl matplotlib seaborn scipy statsmodels
```

### Passo 1: Criar o Excel inicial
```bash
# Gerar o Excel inicial com dados brutos na primeira aba
python3 python_pipeline/diversity_analyzer.py "caminho/do/arquivo.csv" "relatorio.xlsx"
```

Para gerar diretamente em um arquivo `.xlsm` com macros preservadas, aponte para um `.xlsm` já existente (modelo/base):
```bash
# Gera as abas no .xlsm existente mantendo o VBA
python3 python_pipeline/diversity_analyzer.py "caminho/do/arquivo.csv" "analise_diversidade_completa.xlsm"
```
Observação: criar macros novas do zero não é suportado pelo Python; é necessário usar um `.xlsm` existente para preservação do VBA.

### Passo 2: Atualizar análises quando dados mudam
```bash
# Quando você adicionar novos dados na aba DADOS_BRUTOS, execute:
python3 python_pipeline/update_excel_analysis.py "relatorio.xlsx"
```

Ao atualizar, o sistema também:
- Recria a aba 0_HOME com links de navegação e carimbo de "Última atualização";
- Converte `DADOS_BRUTOS` em uma Tabela do Excel (filtros e listras de linhas);
- Gera `LISTAS` com categorias únicas e aplica validação (listas suspensas) em colunas categóricas;
- Destaca em vermelho valores fora da lista (possíveis inconsistências);
- Adiciona `AJUDA_GLOSSARIO` com notas de interpretação dos indicadores.

### Suporte a .xlsm (com macros)
- Se seu arquivo for `.xlsm`, as macros são preservadas automaticamente durante a atualização.
- Use o mesmo comando, apontando para o `.xlsm`:
  ```bash
  python3 python_pipeline/update_excel_analysis.py "analise_diversidade_completa.xlsm"
  ```
  Observação: o script carrega o workbook com `keep_vba=True`, remove/recria somente as abas de análise e mantém `DADOS_BRUTOS` e o projeto VBA intactos.

## 📦 Estrutura de Pastas e Fluxos

### 1) Fluxo com Python (pipeline principal)
- Pastas/arquivos: `python_pipeline/` (scripts), `.xlsm` com macros
- Uso típico:
  - Gerar a partir de CSV:
    ```bash
    python3 python_pipeline/diversity_analyzer.py "caminho/do/dataset.csv" "analise_diversidade_completa.xlsm"
    ```
  - Atualizar um relatório existente (.xlsx/.xlsm):
    ```bash
    python3 python_pipeline/update_excel_analysis.py "analise_diversidade_completa.xlsm"
    ```
- Macro (VBA): o botão “Atualizar Análises (Python)” chama `update_excel_analysis.py`:
  - Procura primeiro na pasta do workbook
  - Se não encontrar, procura em `python_pipeline/`
  - Você pode definir `DIVERSITY_PYTHON` com o caminho do Python

### 2) Fluxo Excel‑only (sem Python para recalcular)
- Pastas/arquivos: `excel_only/` (scripts geradores/injetores)
- Opções:
  1. Gerar um arquivo novo “somente Excel” (recalcula via fórmulas dinâmicas):
     ```bash
     python3 excel_only/create_excel_only_report.py \
       "caminho/do/dataset.csv" \
       "analise_diversidade_excel_only.xlsx"
     ```
  2. Injetar abas Excel‑only (_XL) no .xlsm de produção (mantendo macros):
     ```bash
     python3 excel_only/inject_excel_only_sheets.py "analise_diversidade_completa.xlsm"
     ```
     Abas adicionadas: 3A_XL_*, 3C_XL_*, 3B_XL_*, 3D_XL_*, 5_XL_VISUALIZACOES.

## 🧰 Macro e Caminhos do Python
- O módulo VBA (`python_pipeline/excel_macro_auto_update.bas`) procura `update_excel_analysis.py` em:
  1) mesma pasta do workbook; 2) subpasta `python_pipeline/`
- macOS/Windows: detecta o sistema; usa `python3`/`python`, ou `DIVERSITY_PYTHON` se definido.
- Para trocar o caminho do Python sem editar a macro:
  - macOS (shell atual): `export DIVERSITY_PYTHON=/opt/anaconda3/bin/python`
  - Windows (persistente): `setx DIVERSITY_PYTHON "C:\\Users\\voce\\Anaconda3\\python.exe"`

## 🚀 Deploy no Google Cloud Run (somente Excel via API)

Pré‑requisitos
- gcloud SDK instalado e autenticado (`gcloud init`)
- Projeto GCP selecionado (`gcloud config set project SEU_PROJETO`)
- Docker disponível (ou use Cloud Build)

1) Build e Push da imagem
```bash
# Opção A: Cloud Build (recomendado)
gcloud builds submit --tag gcr.io/SEU_PROJETO/diversity-excel:latest .

# Opção B: Docker local
docker build -t gcr.io/SEU_PROJETO/diversity-excel:latest -f api_service/Dockerfile .
docker push gcr.io/SEU_PROJETO/diversity-excel:latest
```

2) Deploy no Cloud Run
```bash
gcloud run deploy diversity-excel \
  --image gcr.io/SEU_PROJETO/diversity-excel:latest \
  --platform managed \
  --region us-central1 \
  --port 8080 \
  --allow-unauthenticated
```

3) Testes com curl
```bash
# Health
curl https://SEU_ENDPOINT/ -s | jq

# Gerar relatório (pipeline Python → Excel completo)
curl -X POST https://SEU_ENDPOINT/process \
  -F "file=@Diversidade/Caso 1 - Variáveis categóricas/dataset.csv" \
  -F "excel_only=false" \
  --output relatorio_diversidade.xlsx

# Gerar relatório Excel-only (fórmulas dinâmicas)
curl -X POST https://SEU_ENDPOINT/process \
  -F "file=@Diversidade/Caso 1 - Variáveis categóricas/dataset.csv" \
  -F "excel_only=true" \
  --output relatorio_excel_only.xlsx

# Injetar abas Excel-only (_XL) no workbook (.xlsm/.xlsx)
curl -X POST https://SEU_ENDPOINT/inject_excel_only \
  -F "workbook=@analise_diversidade_completa.xlsm" \
  --output workbook_com_xl_abas.xlsm
```

Notas de Produção
- Tamanho de upload: Cloud Run tem limites (ajuste conforme o seu dataset). Para arquivos maiores, prefira upload em Cloud Storage e leitura via GCS.
- Variáveis de ambiente: thresholds e liderança via `DIVERSITY_CONFIG` (ex.: `--set-env-vars DIVERSITY_CONFIG=/app/docs/config_diversidade.json`).
- Autenticação: em produção, remova `--allow-unauthenticated` e controle o acesso via IAM.
- Região/Conta: ajuste `--region` e a service account.

### Exemplo de Uso
```bash
# Criar relatório inicial
python3 python_pipeline/diversity_analyzer.py "Diversidade/Caso 1 - Variáveis categóricas/diversity_expanded_dataset_100.csv" "analise_diversidade.xlsx"

# Após adicionar novos dados na aba DADOS_BRUTOS, atualizar:
python3 python_pipeline/update_excel_analysis.py "analise_diversidade.xlsx"
```

## Exemplo de Saída

### Interpretação de Testes Estatísticos
O sistema fornece explicações como:
- "Há uma associação estatisticamente significativa entre Gênero e Cargo (p = 0.0234). Isso sugere que estas variáveis não são independentes."
- "Não há evidência suficiente para afirmar que existe uma associação entre Departamento e Tipo_Contrato (p = 0.4567). As variáveis parecem ser independentes."

### Interpretação de Índices de Diversidade
- **Alta diversidade** (índice ≥ 0.8): "A distribuição é bem equilibrada entre diferentes categorias."
- **Diversidade moderada** (índice ≥ 0.6): "Há uma boa distribuição, mas com algumas categorias predominantes."
- **Baixa diversidade** (índice < 0.6): "Algumas categorias são claramente predominantes."

## 🧩 Troubleshooting (erros comuns)

- Funções dinâmicas não reconhecidas (#NAME?)
  - Seu Excel precisa suportar `LET`, `FILTER`, `UNIQUE`, `SORTBY`. Use Excel 365 atualizado.

- Fórmulas não atualizam na versão Excel‑only
  - Verifique se DADOS_BRUTOS está como Tabela e se o nome é `TBL_DADOS`.
  - Cálculo em “Automático” (Fórmulas → Opções de Cálculo → Automático).
  - Os cabeçalhos usados nas referências (ex.: `TBL_DADOS[Gender]`) precisam existir exatamente com esse nome.

- Índice de Shannon com erro (LN de zero)
  - As fórmulas Excel‑only já descartam `p=0` (via `FILTER(p,p>0)`). Se ainda aparecer erro, verifique se há categorias totalmente vazias.

- Cores/semáforo não aparecem
  - As abas 3A/3C/3B/3D e _XL usam formatação condicional baseada nas colunas de índice. Confirme que os índices estão numéricos (não texto).

- Macro “Atualizar Análises (Python)” não encontra o Python
  - Defina a variável de ambiente `DIVERSITY_PYTHON` com o caminho do Python.
  - macOS: pode precisar colocar o caminho completo (ex.: `/opt/anaconda3/bin/python`).

- Macro não encontra `update_excel_analysis.py`
  - A macro procura na pasta do workbook e em `python_pipeline/`. Certifique‑se que o script esteja em um desses lugares.

- Atualização “só pelo Excel” dentro do .xlsm
  - Injete abas Excel‑only com `excel_only/inject_excel_only_sheets.py`. Elas recalculam ao editar `DADOS_BRUTOS`.

- Performance com muitos dados
  - Prefira o pipeline Python (mais rápido e escalável). A versão Excel‑only pode ficar lenta com filtros dinâmicos extensos.

## Características Técnicas

### Flexibilidade
- Funciona com qualquer número de colunas
- Adapta-se automaticamente a diferentes tipos de dados
- Lida com valores nulos e missing data

### Performance
- Processamento eficiente de grandes conjuntos de dados
- Geração rápida de relatórios
- Uso otimizado de memória

### Qualidade
- Formatação profissional do Excel
- Validação de dados e tratamento de erros
- Código bem documentado e extensível

## Estrutura do Código

```python
diversity_analyzer.py
├── DiversityAnalyzer (classe principal)
│   ├── load_data() - Carrega e valida CSV
│   ├── detect_data_types() - Identifica tipos de colunas
│   ├── generate_descriptive_stats() - Estatísticas descritivas
│   ├── perform_statistical_tests() - Testes estatísticos
│   ├── generate_diversity_indices() - Índices de diversidade
│   ├── create_visualizations() - Gráficos
│   └── create_excel_report() - Gera relatório Excel
```

## 🔄 Fluxo de Trabalho Contínuo

### Quando Novos Dados Chegarem

#### Opção 1: Atualizar arquivo CSV existente
1. **Adicione novos dados** ao seu arquivo CSV
2. **Regenere o Excel** (se quiser começar do zero):
   ```bash
   python3 diversity_analyzer.py "novo_arquivo.csv" "relatorio_atualizado.xlsx"
   ```

#### Opção 2: Atualizar diretamente no Excel (Recomendado)
1. **Abra o Excel** gerado
2. **Adicione novos dados** diretamente na aba `DADOS_BRUTOS`
3. **Execute o atualizador**:
   ```bash
   python3 update_excel_analysis.py "relatorio.xlsx"
   ```

### Vantagens do Novo Sistema

- **Dados preservados**: A aba `DADOS_BRUTOS` nunca é apagada
- **Atualização rápida**: Apenas as abas de análise são regeneradas
- **Flexibilidade**: Você pode editar dados diretamente no Excel
- **Segurança**: Seus dados originais estão sempre seguros

### Exemplo Prático

```bash
# 1. Criar relatório inicial com gráficos integrados
python3 diversity_analyzer.py "dados_iniciais.csv" "analise_diversidade.xlsx"

# 2. Abrir o Excel e adicionar novas linhas na aba DADOS_BRUTOS

# 3. Atualizar análises (gráficos são regerados automaticamente)
python3 update_excel_analysis.py "analise_diversidade.xlsx"

# 4. Repetir o passo 2 e 3 sempre que houver novos dados
```

## 📁 Arquivos do Sistema

- `python_pipeline/` (fluxo com Python + macros)
  - `diversity_analyzer.py`: gera o relatório inicial (a partir de CSV)
  - `update_excel_analysis.py`: atualiza o relatório (a partir do DADOS_BRUTOS)
  - `excel_macro_auto_update.bas`: módulo VBA com o botão “Atualizar Análises (Python)”
- `excel_only/` (fluxo 100% Excel com fórmulas dinâmicas)
  - `create_excel_only_report.py`: cria um relatório que recalcula apenas com Excel
  - `inject_excel_only_sheets.py`: injeta abas “Excel‑only” (_XL) em um workbook existente
- `README_Analise_Dinamica.md`: Documentação completa

## 💡 Dicas de Uso

1. **Faça backup** do seu Excel antes de grandes atualizações
2. **Mantenha a estrutura** das colunas ao adicionar novos dados
3. **Verifique os resultados** após cada atualização
4. **Use o atualizador** sempre que modificar dados na aba `DADOS_BRUTOS`

O sistema é totalmente dinâmico e não requer modificações manuais quando novos dados são adicionados, desde que a estrutura das colunas seja mantida.

## Benefícios

- **Automação completa**: Não requer intervenção manual
- **Resultados compreensíveis**: Explicações claras para não-técnicos
- **Atualização dinâmica**: Adapta-se a novos dados automaticamente
- **Análise completa**: Cobertura estatística abrangente
- **Formato profissional**: Relatórios Excel bem formatados

Este sistema transforma dados brutos em insights acionáveis sobre diversidade, tornando a análise estatística acessível a todos os níveis de usuários.

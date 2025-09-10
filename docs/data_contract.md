# Data Contract — Relatório de Diversidade

Este documento define o schema mínimo esperado para geração/atualização dos relatórios de diversidade.

## Colunas Obrigatórias/Recomendadas

- Gênero (obrigatório): uma das colunas a seguir (apenas uma necessária)
  - `genero`, `gênero`, `sexo`, `gender`
- Raça/Cor (recomendado):
  - `raca`, `raça`, `raca_cor`, `cor`, `race`, `etnia`, `ethnicity`
- Variáveis categóricas (uma ou mais):
  - `departamento`, `area`, `cargo`, `funcao`, `nivel`, `localidade`, etc.
- Idade (opcional):
  - `idade` (numérica). Se presente, gera-se automaticamente `faixa_etaria`.

Observações:
- Nomes de colunas são normalizados (acentos removidos, minúsculas, espaços → `_`).
- Para melhor legibilidade, mantenha os títulos coerentes e estáveis.

## Mapeamentos

- Gênero → `Feminino`, `Masculino`, `Outro/NS`
- Raça/Cor (IBGE) → `Branca`, `Preta`, `Parda`, `Amarela`, `Indígena`, `Não informado`

## Exemplo de CSV

```
genero,raca,departamento,cargo,idade
Feminino,Parda,Comercial,Analista,29
Masculino,Branca,TI,Dev Pleno,34
Feminino,Preta,Financeiro,Assistente,22
Masculino,Parda,Comercial,Analista,31
```

## Configuração Opcional (config_diversidade.json)

```
{
  "thresholds": { "low": 0.6, "high": 0.8 },
  "leadership_keywords": ["gerente","diretor","coordenador","supervisor","líder","gestor","head","chief","c-level","vp","presidente","manager","director","lead"],
  "leadership_column_hints": ["cargo","função","job","title","posição","nível","senioridade","lead","gestão","role","position","level","seniority"],
  "leadership_columns": ["Cargo","Título do Cargo"],
  "company_name": "Minha Empresa",
  "company_logo": "./docs/logo.png"
}
```

Locais suportados:
- Variável de ambiente `DIVERSITY_CONFIG` apontando para o arquivo JSON
- `./config_diversidade.json`
- `./docs/config_diversidade.json`


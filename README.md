# Matriz de Causa e Efeito (MCE) — Referência e Plano de Comparação

## Resumo do arquivo (tabela)

| Item | Informação |
|---|---|
| Tipo de documento | Matriz de Causa e Efeito (MCE) |
| Início efetivo do conteúdo | Página 16 |
| Identificação no nome do arquivo | Projeto + Sistema + Disciplina + Versão |
| Versão | A letra no final do nome do arquivo (ex.: `_U`, `_P`) representa a versão |
| Planilhas relevantes | `PG_016`, `PG_017`, `PG_018`, ... |

---

## Mapeamentos de códigos (nome do arquivo)

### Projeto

| Código | Projeto |
|---|---|
| 3010.2J | P-80 |
| 3010.2P | P-83 |

### Sistema

| Código | Sistema |
|---|---|
| 1200 | Topside |
| 1351 | Hull |

### Disciplina

| Código | Disciplina |
|---|---|
| KES-001 | Shutdown |
| KES-002 | Fogo e Gás |

---

# Estrutura da Planilha PG_xxx

Referência estrutural baseada na `PG_016`.

## 1) Região de Causas

| Sub-região | Range | Observação |
|------------|-------|------------|
| Cabeçalho das causas | B59:AR59 | REF DOC, INTERFACE, DESCRIPTION, VOTING, TAG NUMBER, DELAY |
| Corpo das causas | B64:AR100 | Linhas reais das causas |
| Índice lógico das causas | AU61:AU110 | Numeração 1..50 usada na matriz |

---

## 2) Região de Efeitos

| Sub-região | Range | Observação |
|------------|-------|------------|
| Régua de efeitos (1..50) | AW59:CT59 | Índice horizontal dos efeitos |
| Labels verticais fixos | AU2:AU56 | SYS, NOTES, INTERFACE, REF DOC, DESCRIPTION, ACTION, TAG NUMBER, DELAY |
| Cabeçalho vertical dos efeitos | AW2:CT56 | Metadados completos de cada efeito |

---

## 3) Região de Cruzamento Causa × Efeito

| Sub-região | Range | Observação |
|------------|-------|------------|
| Matriz lógica | AW61:CT110 | Interseção causa × efeito (X ou vazio) |

---

# Reconhecimento Estrutural no C# (.NET 8)

O parser deverá identificar três grandes grupos:

1. **Causas**
   - Detectar linha âncora pelo cabeçalho (`REF DOC`).
   - Capturar range horizontal até `DELAY`.
   - Identificar última linha com dados em `REF DOC`.

2. **Efeitos**
   - Detectar régua 1..N.
   - Determinar range das colunas de efeitos.
   - Capturar cabeçalho vertical completo.

3. **Cruzamento**
   - Range delimitado por:
     - Colunas = mesmas da régua de efeitos.
     - Linhas = mesmas do índice lógico das causas.
   - Comparação célula a célula.

---

# Estratégia de Comparação (Macro → Micro)

- Macro:
  - Número de planilhas `PG_xxx`
  - Nome das planilhas
  - Dimensão da matriz (NxM)

- Micro:
  - Hash por linha de causa
  - Hash por cabeçalho de efeito
  - Comparação célula a célula no cruzamento

---

# Modelo Lógico Sugerido

```
Workbook
 └── Worksheet (PG_xxx)
      ├── CausesRegion
      │     └── CauseRow
      ├── EffectsRegion
      │     └── EffectColumn
      └── CauseEffectMatrix
            └── Cell
```

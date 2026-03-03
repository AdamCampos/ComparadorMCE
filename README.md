# Matriz de Causa e Efeito (MCE) — Referência e Plano de Comparação

## Resumo do arquivo (tabela)

| Item | Informação |
|---|---|
| Tipo de documento | Matriz de Causa e Efeito (MCE) |
| Início efetivo do conteúdo | Página 16 |
| Identificação no nome do arquivo | Projeto + Sistema + Disciplina + Versão |
| Versão | A **letra no final** do nome do arquivo (ex.: `_U`, `_P`) representa a versão |

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

### Exemplo de interpretação do nome do arquivo

`I-DE-3010.2J-1351-847-KES-001_U.xlsm`  
→ Projeto: P-80 (3010.2J)  
→ Sistema: Hull (1351)  
→ Disciplina: Shutdown (KES-001)  
→ Versão: U

---

## Mapa mental (comparador .NET 8 — macro → micro)

- **Objetivo**
  - Comparar arquivos de MCE **página a página** e **linha a linha**
  - Destacar apenas **diferenças** (igualdades não interessam)
- **Entrada**
  - 2 arquivos (ex.: versões diferentes: `_U` vs `_P`)
  - Metadados extraídos do **nome do arquivo**
    - Projeto / Sistema / Disciplina / Versão
- **Pipeline de comparação**
  - **1) Macro (visão estrutural)**
    - Quantidade de planilhas (sheets)
    - Nomes das planilhas
    - (Opcional) presença/ausência de áreas esperadas
  - **2) Meso (visão por “página”/bloco)**
    - Definir um **range fixo** por planilha (ex.: linhas/colunas relevantes)
    - Quebrar o range em “páginas” lógicas (ex.: janelas por N linhas)
  - **3) Micro (linha a linha)**
    - Ler cada linha do range fixo
    - Normalizar linha (ex.: trim, colunas relevantes, formato consistente)
    - Gerar **HASH por linha**
    - Comparar hashes entre arquivos/planilhas
      - Hash igual → ignorar
      - Hash diferente → registrar dif
- **Representação de diferenças (em árvore)**
  - Arquivo A vs Arquivo B
    - Planilha
      - “Página”/Bloco
        - Linha
          - Colunas divergentes (detalhamento opcional)
- **Saídas**
  - Relatório de diffs
    - Resumo macro (sheets: adicionadas/removidas/renomeadas)
    - Diferenças por planilha (quantidade de linhas divergentes)
    - Drill-down até linha/coluna quando necessário
- **Notas de implementação (.NET 8)**
  - Núcleo do comparador: modelos de domínio
    - `WorkbookDiff → WorksheetDiff → BlockDiff → RowDiff`
  - Estratégia de hash
    - Hash estável (mesma normalização = mesmo hash)
    - Guardar também o “conteúdo normalizado” para explicar a diferença

# ComparadorMCE --- Matriz de Causa e Efeito (MCE)

## 1. Contexto

Projeto para leitura, normalização e futura comparação de Matrizes de
Causa e Efeito (MCE) em arquivos Excel (P-80 / P-83), utilizando .NET 8
e EPPlus 8+.

O conteúdo efetivo da matriz inicia na página 16.\
A versão do documento é definida pela letra final no nome do arquivo
(ex.: `_U`, `_P`, `_E`).

------------------------------------------------------------------------

# 2. Estrutura da Solution

**Solution:** ComparadorMCE

Projetos:

-   ComparadorMCE (Console / CLI)
-   ComparadorMCE.Core (Regra de negócio)
-   GUI (WinForms --- ainda não utilizado)

Biblioteca utilizada: - EPPlus 8+

------------------------------------------------------------------------

# 3. Diretório padrão de arquivos

Raiz default:

    C:\Projetos\VisualStudio\ComparadorMCE\resources

Subpastas: - P80 - P83 - OLDS

O CLI pode receber caminho por argumento, mas por padrão trabalha nessa
estrutura.

------------------------------------------------------------------------

# 4. Escopo Atual Implementado

## 4.1 Acesso ao Excel

✔ Abertura de arquivo via EPPlus 8+\
✔ Configuração correta de licença (SetNonCommercialPersonal)\
✔ Listagem segura das planilhas

------------------------------------------------------------------------

## 4.2 Filtro de Planilhas (Atualizado)

A captura é aplicada **somente** às planilhas cujo nome atende ao
padrão:

    PG_xxx

Regras:

-   `xxx` deve ser numérico
-   `xxx >= 16`
-   Pode conter sufixos após o número

Exemplos válidos:

-   PG_016\
-   PG_016 (3)\
-   PG_017 aaaaa\
-   PG_201\
-   PG_201_REV1

Exemplos ignorados:

-   PG_001_COVER\
-   PG_002\
-   Qualquer PG com número \< 16\
-   Qualquer aba que não comece com PG\_

A validação é feita via Regex + parsing numérico.

------------------------------------------------------------------------

# 5. Estrutura da Planilha PG_xxx

Referência baseada em PG_016.

## 5.1 Região de Título da Matriz

-   Faixa central mesclada
-   Normalmente entre linhas \~50--58
-   Texto como: "INPUTS FROM INTEGRATOR"
-   Diferente do cabeçalho de processo

Extração:

-   Busca restrita à faixa da matriz
-   Prioriza células mescladas
-   Trata corretamente quebras de linha

------------------------------------------------------------------------

## 5.2 Região de Causas (Implementado)

A região de causas agora é extraída e normalizada corretamente.

### 5.2.1 Detecção Dinâmica de Cabeçalho

O sistema identifica automaticamente o cabeçalho da tabela mesmo quando:

-   Há quebra de linha dentro da célula
-   Há texto adicional após o rótulo (ex.: "REF DOC I-DE-3010.2P-")

Os seguintes rótulos são detectados via `Contains()`:

-   V
-   REF DOC
-   INTERFACE
-   DESCRIPTION
-   VOTING
-   TAG NUMBER
-   DELAY

------------------------------------------------------------------------

### 5.2.2 Colunas Extraídas

Cada linha válida de causa gera um objeto estruturado contendo:

-   V (valor da legenda, ex.: P)
-   RefDoc
-   Interface
-   Description
-   Voting
-   TagNumber
-   Delay
-   RowIndex (linha original da planilha)

Critério de linha válida:

-   Coluna V preenchida
-   Linha contém conteúdo relevante

------------------------------------------------------------------------

## 5.3 Planilha RESUMO (Expandida)

Ao processar um arquivo:

-   Cria ou reutiliza a planilha "RESUMO"
-   Limpa antes de reprocessar
-   Lista somente PG_xxx com número \>= 16

### Estrutura Atual do RESUMO

  Coluna   Conteúdo
  -------- ------------------
  A        Nome da planilha
  B        Título da Matriz
  C        V
  D        RefDoc
  E        Interface
  F        Description
  G        Voting
  H        TagNumber
  I        Delay

Para cada causa encontrada:

-   Colunas A e B são repetidas
-   Cada causa gera uma nova linha

Se a planilha não possuir causas detectadas, ainda assim é registrada
uma linha base com A e B.

------------------------------------------------------------------------

# 6. Estado Atual do Projeto

✔ Extração correta de título\
✔ Filtro correto de PG_xxx \>= 16\
✔ Extração estruturada da região de causas\
✔ Normalização robusta de cabeçalhos\
✔ Geração automática e expandida da planilha RESUMO

------------------------------------------------------------------------

# 7. Próximos Passos

-   Identificação formal de Voting Groups\
-   Extração estruturada da matriz causa × efeito\
-   Modelagem completa de efeitos\
-   Geração de hash por causa\
-   Estrutura de comparação entre arquivos

------------------------------------------------------------------------

Projeto em fase de consolidação da captura estrutural.

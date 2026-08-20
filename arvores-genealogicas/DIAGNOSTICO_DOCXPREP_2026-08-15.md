# Diagnóstico de continuidade — DocxPrep, "História da Família Pires"

> **Modo somente leitura.** Nenhum arquivo dentro de `DocxPrep\saida\` foi
> escrito, movido ou alterado nesta tarefa — só lido, via `python-docx`
> aberto em modo leitura e comandos de listagem/leitura de arquivo. Nada foi
> normalizado, corrigido ou inventado; onde a resposta não pôde ser
> determinada com os artefatos existentes, está marcada **NÃO DETERMINÁVEL**,
> com o que faltaria para determinar.
>
> Data: 2026-08-15. Metodologia: leitura direta dos dois `_limpo.docx` e dos
> dois `originais/*.docx` correspondentes via `python-docx` (scripts
> descartáveis rodados no scratchpad da sessão, sem gravar nada em
> `DocxPrep\`), leitura do código-fonte de `docxprep/` (todos os `.py`),
> leitura de `CHECKPOINT.md`, `README.md`, `git log`, e leitura de
> `WordFluxAI\CLAUDE.md` + do projeto irmão `scripts-indesign\` (fora de
> `WordFluxAI`) para as seções 5 e 6. Todas as contagens abaixo vêm de
> execução direta contra os arquivos reais, não de documentação copiada.

---

## 1. O que o DocxPrep já fez

### 1.1 Artefatos gerados por documento

**Primeira parte** — `saida/História da Família Pires - primeira parte do livro final/`

| Artefato | Caminho (relativo à pasta acima) | Tamanho | Modificado |
|---|---|---:|---|
| DOCX limpo | `História da Família Pires - primeira parte do livro final_limpo.docx` | 6.005.511 bytes (5,7 MB) | 2026-08-12 01:34 |
| Relatório de auditoria | `relatorio_auditoria.md` | 1.857 bytes | 2026-08-12 01:34 |
| Relatório de inventário | `relatorio_inventario.md` | 23.283 bytes | 2026-08-12 01:34 |
| Imagens extraídas | `imagens_extraidas/*.jpeg`, `*.png` (116 arquivos) | 5,6 MB no total | 2026-08-12 01:34 |
| JSON | — | — | **não existe** |
| Log | — | — | **não existe** |

**Segunda parte** — `saida/História da Família Pires  - segunda parte do livro/` (espaço duplo no nome)

| Artefato | Caminho (relativo à pasta acima) | Tamanho | Modificado |
|---|---|---:|---|
| DOCX limpo | `História da Família Pires  - segunda parte do livro_limpo.docx` | 15.170.984 bytes (14,5 MB) | 2026-08-12 02:31 |
| Relatório de auditoria | `relatorio_auditoria.md` | 1.918 bytes | 2026-08-12 02:31 |
| Relatório de inventário | `relatorio_inventario.md` | 43.773 bytes | 2026-08-12 02:31 |
| Imagens extraídas | `imagens_extraidas/*.jpeg` (297 arquivos) | 16 MB no total | 2026-08-12 02:31 |
| JSON | — | — | **não existe** |
| Log | — | — | **não existe** |

Confirmado por busca (`find saida -iname "*.json" -o -iname "*.log"`): **nenhum JSON ou log existe em `saida/`, em nenhum dos dois documentos.** Não há nada além desses 4 tipos de artefato por documento.

### 1.2 Etapas do pipeline — o que rodou

Lendo `docxprep/pipeline.py` e `docxprep/cli.py`, o pipeline tem 4 fases:

| Fase | O que faz | Rodou nos dois arquivos? |
|---|---|---|
| 1 — Sanitização | Colapsa espaços múltiplos, remove tabs soltos, divide quebras manuais, remove parágrafos vazios | **Sim** |
| 2 — Tipografia | Aspas retas → tipográficas, reticências, hífen de diálogo | **Sim** |
| 3 — Estilos | Detecta/aplica títulos (Capítulo = alta confiança automática; Grupos A/B/C de baixa confiança = decisão humana já registrada em `CHECKPOINT.md`), converte negrito/itálico direto em estilo de caractere nomeado | **Sim** |
| 4 — Fotos | Substitui legendas/imagens por marcador `[FOTO N — individual]`/`[FOTOS N-M — grupo de X]`, extrai toda imagem embutida para disco antes de remover do DOCX | **Sim** |

Confirmado por `git log --oneline` (10 commits, topo `748a280`) e por `CHECKPOINT.md`, cuja seção "ESTADO FINAL" declara as 4 fases completas nos dois arquivos com 0 divergências de integridade. Reconfirmado agora, ao vivo, nesta tarefa (ver §3).

### 1.3 O que o código PODE gerar hoje vs. o que gerou

Busca no código-fonte inteiro (`grep -rniE "json|icml|\.txt|infografic|genealog|arvore" docxprep/`) retornou **zero ocorrências**. O DocxPrep, como está hoje:

- **PODE gerar:** exatamente os 4 tipos de artefato acima (DOCX limpo, 2 relatórios Markdown, imagens extraídas). É só isso — `cli.py` só tem o comando `process`, sem nenhuma opção de exportação alternativa.
- **NÃO PODE gerar hoje:** JSON, TXT de texto puro, ICML, nem qualquer marcação de infográfico/árvore genealógica. Essas capacidades **não existem no código**, não é uma questão de terem sido desativadas ou não usadas — não foram escritas.
- O que gerou para estes dois arquivos é **exatamente** o que o código pode gerar hoje — não há descompasso entre capacidade e execução.

---

## 2. Marcações de foto e infográfico — existem?

**Pergunta central, resposta direta: existe marcação de FOTO. Não existe, em lugar nenhum, marcação de infográfico/árvore genealógica.**

### 2.1 Marcadores de foto — existem, e são consistentes

Abrindo os dois `_limpo.docx` e contando parágrafos por estilo:

| | Primeira parte | Segunda parte |
|---|---:|---:|
| Parágrafos totais | 1.473 | 3.003 |
| Parágrafos com estilo `MARCADOR_FOTO` | 53 | 92 |
| `w:drawing` (imagem embutida) ainda no DOCX | **0** | **0** |

Os marcadores seguem o padrão `[FOTO N — individual]` ou `[FOTOS N-M — grupo de X]` (ex.: `[FOTO 0 — individual]`, `[FOTO 2A — individual]`, `[FOTOS 122-123 — grupo de 2]`). Busquei por qualquer outro padrão de marcador (`\[FOTO|\{\{|Figura\s+\d|<<|>>|\[IMG|\[ARVORE|\[TREE|\[INFOGRAFIC`) fora do estilo `MARCADOR_FOTO`: **zero ocorrências** nos dois arquivos. Não existe "Figura X" nem legenda isolada fora do padrão `MARCADOR_FOTO` — o padrão é único e consistente.

### 2.2 Marcador de infográfico/árvore genealógica — **não existe, sem rodeios**

Nenhum parágrafo em nenhum dos dois `_limpo.docx` contém qualquer marcação relacionada a árvore genealógica, tronco, infográfico ou equivalente. O DocxPrep nunca teve essa funcionalidade (confirmado em §1.3) e nunca a aplicou. Isso é esperado — o DocxPrep e o módulo de extração genealógica (`WordFluxAI/infograficos/arvores-genealogicas/`) são ferramentas **separadas**: o segundo lê o DOCX limpo do primeiro, mas o primeiro nunca escreve marcação genealógica.

O que existe, e que **pode ser confundido** com marcação genealógica se lido superficialmente, são os próprios cabeçalhos do texto original do autor — ver §4 (isso não é marcação do DocxPrep, é estrutura pré-existente no manuscrito).

### 2.3 As imagens foram extraídas — mas a correspondência com o marcador tem uma pegadinha real

Sim, as 116 (primeira) + 297 (segunda) = **413 imagens** foram extraídas, cada uma com nome `<arquivo>_p<NNNN>_<contador>.<ext>` (ex.: `..._p0131_003.jpeg`).

**Verifiquei diretamente se o índice `NNNN` do nome do arquivo bate com a posição real do marcador no `_limpo.docx` final — e não bate, por um motivo estrutural, não um bug:**

O número `NNNN` é a posição do parágrafo **no momento em que a Fase 4 processou aquele trecho**, antes de qualquer parágrafo seguinte ser removido pelo próprio merge de grupos de foto (cada grupo de N parágrafos vira 1 parágrafo-marcador, deslocando os índices de tudo que vem depois). Exemplo verificado agora, ao vivo, no `_limpo.docx` da primeira parte:

| Fonte | Índice apontado |
|---|---:|
| Nome do arquivo extraído (`..._p0131_003.jpeg`) | 131 |
| `relatorio_inventario.md` (linha do marcador `[FOTO 2]`) | #130-#131 |
| **Posição real do parágrafo `[FOTO 2 — individual]` no `_limpo.docx`, agora, verificada diretamente** | **128** |

Ou seja: **nem o nome do arquivo nem o relatório apontam para a posição real no DOCX final** — ambos registram a posição no espaço de índices de quando a Fase 4 rodou, que só coincide com o DOCX final até o primeiro grupo de foto ser mesclado; depois disso, diverge cada vez mais.

**A correspondência que continua 100% confiável é textual, não numérica:** `relatorio_inventario.md` lista, na mesma linha, o texto do marcador e o(s) nome(s) de arquivo daquele grupo — isso é seguro porque é uma associação direta, não uma aritmética de índice. Qualquer script futuro que precise ligar "imagem X" a "posição no `_limpo.docx`" deve localizar o parágrafo **pelo texto do marcador** (`p.text == "[FOTO 2 — individual]"`), nunca pelo número do nome do arquivo.

### 2.4 Onde as imagens estavam no original, para eventual recriação de marcadores

Todas as imagens embutidas do original (verificado: 116 na primeira parte, 297 na segunda — contagem direta de `w:drawing` nos dois `originais/*.docx`, batendo exatamente com o total extraído) estavam **embutidas em parágrafos ao lado das legendas textuais** ("Foto N", "Fotos N,M,...", etc.) ou sozinhas em parágrafos vazios logo após a legenda. `relatorio_inventario.md` de cada arquivo já lista, por grupo, o intervalo de parágrafos do **original** (`#i-#j`, coluna "Parágrafos (orig.)") onde cada bloco estava — essa é a referência para reconstrução, não os nomes de arquivo.

---

## 3. Fidelidade ao original

Os dois originais estão acessíveis (`originais/*.docx`, read-only). Comparação feita agora, ao vivo.

### 3.1 Contagem de parágrafos

| | Primeira parte | Segunda parte |
|---|---:|---:|
| Original | 2.017 | 3.891 |
| Após Fases 1+2 (antes de título/foto) | 1.530 | 3.095 |
| `_limpo.docx` final (após Fase 4) | 1.473 | 3.003 |

A queda de 2.017→1.530 e 3.891→3.095 é majoritariamente remoção de parágrafo vazio (Fase 1) parcialmente compensada por divisão de quebra manual (também Fase 1) — valores batem exatamente com os contadores já registrados em `CHECKPOINT.md` (662/839 removidos, 175/43 criados por divisão). A segunda queda (1.530→1.473, 3.095→3.003) é a fusão de grupos de foto em um único parágrafo-marcador (Fase 4) — comportamento intencional, documentado em `relatorio_inventario.md`.

### 3.2 Contagem de imagens

| | Primeira parte | Segunda parte |
|---|---:|---:|
| `w:drawing` no original | 116 | 297 |
| `w:drawing` no `_limpo.docx` | **0** | **0** |
| Arquivos em `imagens_extraidas/` | 116 | 297 |

**116 = 116 e 297 = 297 nos dois arquivos: nenhuma imagem perdida, nenhuma duplicada.** Toda imagem embutida foi extraída para disco antes de ser removida do DOCX de saída — confirma o "princípio inviolável" que `README.md` e `CHECKPOINT.md` declaram.

### 3.3 Verificação de integridade textual — o que ela cobre e o que não cobre

`docxprep/integridade.py` só simula Fases 1+2 (texto puro), por desenho — comentário explícito em `pipeline.py`: *"a verificação de integridade só faz sentido quando fase1 E fase2 rodaram juntas"*. Reproduzi essa checagem **agora, ao vivo**, no ponto exato do pipeline em que ela é válida (logo após Fase 1+2, antes de Fase 3/4):

- **Primeira parte: 0 divergências.**
- **Segunda parte: 0 divergências.**

Isso confirma, de forma independente e reexecutada nesta sessão (não copiada de `CHECKPOINT.md`), que nenhuma palavra, pontuação, acento ou caractere do autor foi alterado pelas Fases 1-2 — só formatação mecânica, exatamente como `README.md` promete.

**Testei também rodar essa mesma checagem contra o `_limpo.docx` FINAL (pós Fase 3+4)** — não por achar que deveria bater, mas para não deixar a pergunta em aberto: dá **1.231 divergências (primeira)** e **243 (segunda)**. Isso **não é perda de conteúdo** — é o comparador (desenhado só para Fases 1+2) não sabendo interpretar a Fase 4, que **deliberadamente** substitui o parágrafo de legenda pelo texto do marcador (ex.: esperado `'FOTO 02'`, encontrado `'[FOTO 2 — individual]'`). A partir da primeira substituição, os índices de parágrafo também desalinham em cascata (mesmo fenômeno do §2.3), inflando artificialmente a contagem. Preciso deixar registrado que **não existe hoje, no DocxPrep, um verificador automático de integridade que cubra as Fases 3 e 4** — a garantia dessas duas fases é: (a) Fase 3 só muda estilo, nunca o texto de um `<w:t>` (verificável lendo `fase3_estilos.py` — só chama `set_pstyle`/`r.style`, nunca edita texto); (b) Fase 4 documenta, parágrafo a parágrafo, cada substituição em `relatorio_inventario.md`.

### 3.4 Conteúdo presente no original e ausente no `_limpo.docx`, item a item

Com o que os artefatos atuais permitem determinar, a única categoria de conteúdo que **não sobrevive verbatim** ao `_limpo.docx` é:

**O texto literal das legendas de foto** (ex.: `"FOTO 02"`, `"Fotos 109, 110,111, 112, 113"`, `"Foto 105 e 106"`) — cada uma é apagada e substituída pelo marcador gerado. Isso é uma transformação **documentada e intencional** da Fase 4, não uma perda silenciosa: o parágrafo de origem (`#i` a `#j`) e o marcador que o substituiu estão listados em `relatorio_inventario.md`.

Dito isso, **a redação exata da legenda original não é retida em nenhum artefato do DocxPrep** — nem no `_limpo.docx` (foi apagada), nem em `relatorio_inventario.md` (só guarda o intervalo de parágrafo e o marcador, não o texto da legenda). Para recuperar o texto exato de uma legenda específica, é preciso voltar ao `originais/*.docx` e ler o parágrafo pelo índice **do momento em que a Fase 4 rodou** (o mesmo espaço de índices do §2.3 — não o índice do `_limpo.docx` final, que já mudou). Isso é relevante porque a sessão anterior documentou (em `CHECKPOINT.md`) casos onde a legenda tinha mais informação que o marcador capturou — ex.: `"Fotos 134,135,...,145"` (12 números) virou `[FOTO 134 — individual]` (parece 1 foto, mas a coluna "Imagens extraídas" do relatório mostra corretamente 12 arquivos). **Isso não é perda de imagem** (as 12 imagens estão extraídas e listadas), **é perda do texto da legenda original**, que só existe hoje no `originais/*.docx`.

Não encontrei, na comparação de contagens acima, nenhuma outra categoria de conteúdo ausente — mas não fiz uma comparação exaustiva parágrafo-a-parágrafo do corpo narrativo inteiro (2.017 + 3.891 parágrafos), o que estaria além do escopo proporcional deste diagnóstico. **NÃO DETERMINÁVEL sem esforço adicional:** se existe alguma outra perda pontual de conteúdo textual fora do padrão de legenda de foto — para fechar essa lacuna com certeza total, seria preciso um comparador de conjunto (não posicional) entre os parágrafos não-vazios do original e do `_limpo.docx`, tolerante às transformações legítimas de Fase 1/2 (aspas, espaços) — ferramenta que não existe hoje e não construí nesta tarefa por ser leitura, não escrita de código.

---

## 4. Mapa estrutural

### 4.1 O que é 100% confiável: os 2 Patriarcas

O DocxPrep aplicou um estilo de nível 3 (`Título 31`, `style_id='Ttulo31'`, **não** o "Heading 3" nativo do Word — é um estilo próprio do documento) a **exatamente 2 parágrafos em todo o corpus**, um por arquivo:

| Arquivo | Índice no `_limpo.docx` | Texto |
|---|---:|---|
| Primeira parte | #384 | "Prudêncio de Sousa Pires (Irmão do patriarca Benedito de Sousa Pires)" |
| Segunda parte | #2181 | "Benedito de Sousa Pires (irmão do patriarca Prudêncio de Souza Pires)" |

Isso bate exatamente com o fato, já confirmado independentemente em `infograficos/arvores-genealogicas/DIAGNOSTICO_2026-08-15.md` (lido, não alterado), de que o livro tem **2 patriarcas**. É o sinal mais limpo e confiável que existe no `_limpo.docx` para localizar o topo de cada metade da árvore.

**Importante — isso NÃO foi criado pelo DocxPrep.** `docxprep/docx_package.py` só chama `find_or_create_heading_style` para os níveis 1 e 2 (nunca 3) — busquei no código inteiro por qualquer referência a nível 3, não existe. O estilo `Título 31` já existia no documento **antes** do DocxPrep processá-lo — é estrutura que o autor (ou um editor anterior) já tinha aplicado no Word.

### 4.2 O que é parcialmente confiável, com ressalva importante: "Título 11" e "Título 21"

Estes SÃO os estilos que o DocxPrep usa para título de alta confiança (Capítulo, nível 1) e baixa confiança (Grupos A/B/C — "Filhos de X:", "Tronco", etc., nível 2, decisão já registrada em `CHECKPOINT.md`) — **mas eles já existiam no documento original antes do DocxPrep tocar nele.**

Prova: `docx_package.find_or_create_heading_style` **reaproveita** um estilo `Título 1\d*`/`Título 2\d*` já existente no documento, só criando um novo (`Titulo1`, sem espaço/acento) se nenhum existir. O nome observado — `"Título 11"`, com espaço e acento — só aparece se **já existia** assim no arquivo de origem. Confirma-se pela contagem:

| | Total de parágrafos com o estilo, no `_limpo.docx` | Aplicados pelo DocxPrep (`CHECKPOINT.md`) | Diferença |
|---|---:|---:|---:|
| Título 11, primeira parte | 117 | 10 (só "Capítulo N") | 107 já eram do original |
| Título 11, segunda parte | 393 | 0 | 393 já eram do original |
| Título 21, primeira parte | 115 | 113 | 2 — **origem não determinada** |
| Título 21, segunda parte | 215 | 203 | 12 — **origem não determinada** |

Para "Título 21", a diferença (2 e 12) é pequena o bastante para não ser conclusiva por aritmética simples — **NÃO DETERMINÁVEL com certeza** se são parágrafos que já tinham esse estilo no original (mesmo padrão do "Título 11") ou algum efeito colateral do próprio pipeline que eu não instrumentei para isolar. O que faltaria para determinar: rodar Fase 3 isoladamente e comparar o estilo de cada parágrafo *antes* de qualquer aplicação, parágrafo a parágrafo — não fiz isso por ser instrumentação de código, fora do escopo de leitura desta tarefa.

**Ressalva crítica, verificada por leitura direta do texto de cada parágrafo:** "Título 11" está longe de ser um marcador estrutural limpo. Inspecionei os 510 parágrafos (117+393) com esse estilo nos dois arquivos: além de "Capítulo N", "Tronco", "Tronco-raiz" e "Patriarca" (cabeçalhos genuínos), uma fração muito grande são **listas de nomes de pessoas, um nome por parágrafo**, aparentemente estilizadas assim pelo autor por algum motivo de formatação visual, não como cabeçalho de seção — ex., primeira parte, parágrafos #1116 a #1168 (53 parágrafos seguidos, cada um um único nome como "Bruno Henrique," ou "Fabrício Pires,"), e dezenas de blocos semelhantes na segunda parte (ex. #991-#1028, #1058-#1071, etc.). **Não tentei filtrar isso automaticamente por não ter um critério textual confiável e determinístico para separar "nome-que-é-cabeçalho" de "nome-que-é-item-de-lista"** sem arriscar inventar uma regra não verificada — fica como **NÃO DETERMINÁVEL sem revisão humana** qual desses ~510 parágrafos é estrutural e qual é ruído de formatação do autor.

### 4.3 Mapa dos marcadores estruturais **confiáveis** (filtrados por texto, não só por estilo)

Usei correspondência de texto exata/prefixo (`^Capítulo`, `== "Tronco"`, `== "Tronco-raiz"`, `== "Patriarca"`) sobre os parágrafos com estilo Título 1x/2x/3x, não a lista bruta por estilo. Tabela completa (índices no `_limpo.docx` de cada arquivo, hoje):

**Primeira parte** — Patriarca #384; blocos "Tronco-raiz" abrindo em #398, #783, #1205; blocos "Tronco" abrindo em #457, #495, #548, #588, #618, #691, #740, #752, #796, #948, #1169, #1225, #1412 (lista truncada no dump bruto — ver arquivo completo do dump em anexo à sessão, não reproduzido aqui por extensão). "Capítulo": bloco índice/TOC em #82-#101 (7 entradas em sequência — parece um sumário, não abertura real de seção), e capítulos "reais" (espaçados ao longo do texto) em #111, #133, #168, #247 ("Capítulo - IV", grafia com hífen preservada tal qual está), #302.

**Segunda parte** — Patriarca #2181; blocos "Tronco-raiz" abrindo em #172, #918 ("Tronco- raiz", grafia com espaço antes do hífen preservada), #1849, #2069, #2139, #2153, #2162, #2196, #2207, #2333, #2430, #2564, #2663, #2667; muitos blocos "Tronco" (não listados individualmente aqui por volume — ~20 ocorrências). Único "Capítulo" real: #2821 "Capítulo VII" (continua a numeração da primeira parte). A partir de #2840 a estrutura muda para "I Encontro da Família Pires" ... "IX Encontro da Família Pires" (não é mais tronco genealógico, é histórico de encontros de família) — esse trecho final é a mesma "zona de apêndice/galeria" que o `CHECKPOINT.md` do próprio DocxPrep já sinalizou como estruturalmente diferente do resto do livro (muitos marcadores de foto sem número, muitas legendas sem imagem).

Essa tabela é fiel ao que está no texto — não tentei corrigir, completar ou inferir quais dos "Tronco"/"Tronco-raiz" sem nome próprio ao lado correspondem a qual tronco-raiz numerado do `dados.json` do módulo de genealogia; essa ligação teria que ser feita cruzando texto, não é automática.

### 4.4 Onde cada árvore genealógica deveria entrar no fluxo do texto

**NÃO DETERMINÁVEL com precisão de página** — DOCX não guarda número de página fixo (só existe após paginação em tempo de diagramação/impressão); isso já está corretamente registrado assim em `infograficos/arvores-genealogicas/DIAGNOSTICO_2026-08-15.md`, e concordo com essa conclusão.

**O que é determinável, com o critério proposto:** cada árvore (1 por tronco familiar, conforme já definido em `Familia_Pires.md`) deveria entrar **logo depois do bloco de texto narrativo daquele tronco e antes do início do próximo bloco "Tronco"/"Tronco-raiz"/"Capítulo"** — ou seja, o parágrafo que antecede a árvore é o último parágrafo de conteúdo do tronco (a listagem final de trinetos/tetranetos daquele ramo, ou a nota "Não tiveram filhos", etc.), e o parágrafo que sucede é o cabeçalho "Tronco"/"Tronco-raiz" seguinte (§4.3 já lista essas âncoras). Esse critério usa exatamente o mesmo mecanismo — "âncora anterior mais próxima por índice de parágrafo" — que `infograficos/arvores-genealogicas/DIAGNOSTICO_2026-08-15.md` já propõe (§5 daquele arquivo) para o cruzamento capítulo↔tronco, só que aqui aplicado à posição da própria árvore em vez de um metadado.

**Bloqueio real para implementar isso:** o `paragrafo_idx` usado por `dados.json` (do extrator de genealogia) **não é o mesmo espaço de índices** dos números que estou reportando aqui (que são do `_limpo.docx` lido agora, diretamente) — confirmado em `infograficos/arvores-genealogicas/DIAGNOSTICO_2026-08-15.md`, que já registra esse aviso. Uma ponte entre os dois espaços de índice ainda não existe (nem eu construí uma nesta tarefa — seria escrita de código novo).

---

## 5. Rota de importação para o InDesign

**Distinção importante antes de avaliar:** esta pergunta é sobre como o **texto narrativo do livro** (o `_limpo.docx` do DocxPrep — capítulos, troncos, corpo de texto) entra no InDesign. As **árvores genealógicas em si** já têm rota decidida e documentada em `WordFluxAI/CLAUDE.md`: são geradas como HTML/CSS/JS e exportadas a PDF via `window.print()` (mesmo padrão de `tabelas/indicadores-resultados/`), para depois serem inseridas como imagem/PDF estático no layout — **não** é Place nativo de conteúdo estruturado. As três rotas abaixo não se aplicam às árvores, só ao texto do livro.

### (a) Place nativo do `.docx` + `.jsx` de pós-processamento

- **O que já existe pronto:** o `_limpo.docx` em si (títulos em estilos nomeados e distintos — `Título 11`/`21`/`31` —, ênfase em estilos de caractere nomeados — `Enfase-Bold`/`Italico`/`BoldItalico` —, marcadores de foto isolados em `MARCADOR_FOTO`). Isso é exatamente o que essa rota precisa: um Word "limpo" o bastante para o Place nativo do InDesign não se perder, com uma paleta de estilo pequena e nomeada (9-10 estilos de parágrafo, 3 de caractere, por arquivo — ver §4 para a lista completa de nomes/`style_id`). **Também já existe, no projeto irmão `scripts-indesign\` (fora de `WordFluxAI`, confirmado por leitura direta), exatamente o tipo de ferramenta que essa rota precisa para o pós-processamento**: `AuditarEstilos.jsx` (audita estilos aplicados pós-Place), `RelatorioDialetoWord.jsx` (lista toda ocorrência de um estilo "dialeto do Word" para decidir o mapeamento), `AplicarEstiloPorGrupo.jsx` e `AplicarMapeamentoCondicional.jsx` (aplicam o remapeamento, com modo prévia/teste antes de aplicar de verdade, e undo atômico). Foram construídos para outro livro ("Extensão Universitária") mas resolvem **o mesmo problema estrutural**: traduzir nomes de estilo "Word" para os estilos reais do template InDesign. `AplicarMapeamentoCondicional.jsx` em particular já implementa split-por-prefixo-de-texto (ex.: separar "Figura"/"Fonte:" de corpo de texto comum quando os dois compartilham um único estilo Word) — é literalmente o mesmo mecanismo que separaria `MARCADOR_FOTO` de texto comum, caso a Família Pires também misturasse os dois num único estilo (não é o caso aqui — `MARCADOR_FOTO` já é um estilo próprio — mas o padrão é diretamente reaproveitável se aparecer outro caso assim).
- **Esforço:** médio. O `_limpo.docx` está pronto; falta (1) mapear os ~10 nomes de estilo do DocxPrep para os nomes reais do template InDesign do livro (trabalho de decisão humana + preenchimento de tabela, não código novo), e (2) adaptar (não reescrever do zero) os scripts de `scripts-indesign\` para esse mapeamento específico — a estrutura de script já existe e é genérica o bastante (o `MAPEAMENTO`/`MAPEAMENTO_CONDICIONAL` é uma tabela de configuração, não lógica hardcoded para o outro livro).
- **Risco de perda de informação:** baixo, **se** o InDesign preservar os nomes de estilo do Word no Place nativo (ele preserva, por padrão, criando estilos com o mesmo nome — comportamento documentado do próprio InDesign, não precisei verificar experimentalmente para confiar nisso, é um fato de produto amplamente conhecido). O risco real está em fotos/marcadores: `MARCADOR_FOTO` vira só texto colocado no fluxo — colocar a imagem física correspondente no lugar exato continua sendo um passo manual ou de outro script, não resolvido por nenhuma das 3 rotas sozinha.
- **Reaproveitamento para livros futuros:** **alto** — é a mesma ferramenta que já está sendo generalizada para "Extensão Universitária"; um terceiro livro no mesmo formato reaproveitaria tudo.

### (b) DocxPrep exporta TXT/JSON com texto + marcadores; `.jsx` monta o fluxo

- **O que já existe pronto:** nada — confirmado em §1.3, o DocxPrep não tem nenhuma capacidade de exportação além do próprio `.docx`.
- **O que teria que ser construído:** (1) um exportador novo em `docxprep/` (TXT ou JSON com texto + qual estilo cada trecho tem — trabalho de código, provavelmente pequeno já que os dados já estão estruturados internamente em `RelatorioDados`/`document.paragraphs`), (2) um `.jsx` novo do zero que leia esse formato e **construa o fluxo de texto no InDesign programaticamente** (criar text frames, aplicar estilo por trecho, inserir imagem nos pontos certos) — isso é ordens de grandeza mais trabalho de ExtendScript do que os scripts de pós-processamento de (a), que só leem/reescrevem estilo de parágrafos que **já existem** no documento.
- **Risco de perda de informação:** médio-alto — qualquer formatação que o exportador não pense em serializar (nota de rodapé, hyperlink, formatação de caractere não coberta pelos 3 estilos de ênfase nomeados) se perde silenciosamente, a menos que o exportador seja tão completo quanto o próprio `.docx` — nesse ponto, por que não usar o `.docx` diretamente?
- **Reaproveitamento:** médio — o exportador poderia servir de base para o mesmo módulo de genealogia (que já faz algo parecido, ler DOCX → JSON, só que sem o vínculo com InDesign), mas o `.jsx` de montagem de fluxo teria que ser praticamente reescrito por livro, já que a estrutura de cada livro (capítulo/tronco/o que for) é diferente.

### (c) DocxPrep gera ICML; InDesign faz Place direto

- **O que já existe pronto:** nada, mesma situação de (b) — confirmado, nenhum código de geração ICML existe.
- **O que teria que ser construído:** um gerador de ICML (XML de InCopy) em `docxprep/` — formato bem documentado pela Adobe, mas verboso e com regras próprias de nós/estilo que não têm equivalente hoje no código (o pipeline manipula `python-docx`/`lxml` no dialeto OOXML do Word, não no dialeto ICML — seria essencialmente escrever um serializador novo, não adaptar o existente).
- **Risco de perda de informação:** baixo, **se** o gerador for bem feito — ICML é pensado justamente para carregar estilo nomeado + texto sem depender de interpretação heurística do InDesign (ao contrário de (a), onde o InDesign decide sozinho como interpretar o Word). Essa é a vantagem real desta rota: elimina a ambiguidade de "como o InDesign vai interpretar esse estilo Word" que (a) tem.
- **O que dispensa:** um `.jsx` de importação (o Place de ICML já entra com os estilos mapeados, conforme a premissa do próprio item) — mas **não dispensa** um `.jsx` ou processo equivalente para posicionar as imagens extraídas nos pontos certos, isso é comum às três rotas.
- **Reaproveitamento:** alto, uma vez construído — um serializador ICML genérico serve para qualquer livro que passe pelo DocxPrep, sem depender de replicar `MAPEAMENTO`/`MAPEAMENTO_CONDICIONAL` por livro como em (a).

### Recomendação: **(a)**

Concordo com a premissa do pedido (ExtendScript não lê `.docx` nativamente — confirmado, é limitação conhecida e real do InDesign/ExtendScript, não vou contestar isso). Entre as três, recomendo **(a)** porque é a única onde a maior parte do trabalho pesado **já está feita ou já tem um molde funcional testado em produção** (o `_limpo.docx` de um lado, os scripts de `scripts-indesign\` do outro) — (b) e (c) partem de zero em pelo menos uma ponta significativa (exportador novo + montador de fluxo do zero, ou serializador ICML novo do zero). Com o prazo de entrega já mencionado em sessões anteriores do DocxPrep (`CHECKPOINT.md` registra "amanhã" como deadline em 2026-08-12, ou seja, o livro já pode ter sido ou estar prestes a ser diagramado por Place nativo, sem essa decisão de rota ter sido tomada formalmente) — **NÃO DETERMINÁVEL nesta tarefa se o Place já aconteceu**: não há artefato de InDesign (`.indd`) acessível nesta árvore de diretórios para confirmar se o texto já foi colocado manualmente ou não. Isso é uma pergunta a fazer ao usuário, não algo que os arquivos respondem.

---

## 6. O script: existe, ou criar?

### 6.1 Confirmação — não há `.jsx` para Família Pires

Busquei por `.jsx` em todo `projetosweb/` (fora de `node_modules`/`venv`/`.git`): existem 7 arquivos, **todos** em `D:\IABRACADABRA\projetosweb\scripts-indesign\` (`AplicarEstiloPorGrupo.jsx`, `AplicarMapeamentoCondicional.jsx`, `AuditarEstilos.jsx`, `DiagnosticoTabelas.jsx`, `InventarioEstilos.jsx`, `IrParaTabela.jsx`, `RelatorioDialetoWord.jsx`). Busquei "Pires"/"Família" no conteúdo de todos: as únicas 3 ocorrências são a função `nomeFamiliaFonte()` (**família de fonte tipográfica**, termo técnico de InDesign, nada a ver com a Família Pires) — **falso positivo, confirmado lendo o contexto de cada ocorrência**. `CLAUDE.md` desse projeto (linha 101) nomeia o alvo explicitamente: `01 - EXTENSÃO UNIVERSITÁRIA E TRANSFORMAÇÃO SOCIAL-8-1ªRodadaExtendScript Debugger.indd`. **Confirmo, de forma independente, exatamente o que o pedido já afirmava**: `scripts-indesign` é do livro "Extensão Universitária", e não existe nenhum `.jsx` para Família Pires em lugar nenhum do `projetosweb/`.

Achado colateral relevante: esses 7 scripts fazem **correção de estilo pós-diagramação** (o `.indd` de "Extensão Universitária" já existe, já foi Place, os scripts corrigem tabela/estilo depois) — não fazem importação inicial de conteúdo. Não são um script "pronto para copiar", mas são o molde mais próximo do que a rota (a) precisa (ver §5).

### 6.2 Responsabilidades do script, na rota recomendada (a)

**O que o script (ou conjunto de scripts, seguindo o padrão de `scripts-indesign\` de dividir por responsabilidade) DEVE fazer:**
1. Rodar **depois** do Place nativo do `.docx` no InDesign (não faz o Place em si — isso é ação humana, `File > Place`).
2. **Auditar** (modo só-leitura, como `AuditarEstilos.jsx`) quais estilos de parágrafo/caractere vieram do Word e quantos parágrafos usam cada um — para confirmar que o Place preservou os nomes esperados (`Título 11`, `Título 21`, `Título 31`, `MARCADOR_FOTO`, `Enfase-Bold`, `Enfase-Italico`, `Enfase-BoldItalico`, `Body Text`, `Normal`).
3. **Mapear**, com aprovação humana antes de aplicar (modo prévia obrigatório, como os scripts existentes já fazem), cada estilo "dialeto DocxPrep" para o estilo real do template InDesign do livro Família Pires.
4. **Aplicar** o mapeamento em lote, com undo atômico (`UndoModes.ENTIRE_SCRIPT`) e validação prévia completa antes de aplicar qualquer coisa (mesmo padrão de segurança de `AplicarEstiloPorGrupo.jsx`).
5. **Localizar** os parágrafos `MARCADOR_FOTO` (por texto, não por índice — ver §2.3) e gerar um relatório de onde cada um está na página, para o posicionamento manual (ou semi-automático, se decidido depois) das imagens.

**O que o script NÃO deve fazer:**
1. Não deve tentar interpretar/corrigir o texto do autor — mesmo princípio inviolável do próprio DocxPrep (`README.md`), reaproveitado aqui.
2. Não deve tentar resolver sozinho os casos ambíguos já documentados em `CHECKPOINT.md` (ex.: marcador `[FOTO 134 — individual]` que na verdade tem 12 imagens) — isso é decisão humana, não algo para um script "adivinhar".
3. Não deve fazer o Place em si nem criar o documento InDesign do zero — pressupõe um `.indd` já existente com o template do livro.
4. Não deve gerar nem posicionar as árvores genealógicas — isso é outro fluxo (HTML→PDF→inserção manual, §5), fora do escopo deste script.

### 6.3 O que precisa existir antes da primeira linha

1. **O `.indd` do livro Família Pires, com o `.docx` já colocado via Place nativo** (ou pelo menos um `.indd` de teste com Place feito) — sem isso, não há o que auditar/mapear; é pré-requisito literal, os scripts de `scripts-indesign\` só rodam com `app.activeDocument` de um InDesign aberto.
2. **O template de estilos do InDesign do livro** — nomes reais dos estilos de destino (capítulo, tronco, corpo, legenda de foto, ênfase). Sem isso não dá para preencher a tabela de mapeamento — é decisão editorial/de design, não técnica.
3. **Decisão humana sobre como tratar os `MARCADOR_FOTO`** — viram texto substituído por imagem manualmente? Um script insere a imagem automaticamente na posição do marcador? Isso muda o escopo do passo 5 de §6.2.
4. Confirmação de que o **prazo de entrega já mencionado nas sessões do DocxPrep** ainda está em vigor ou já passou — muda a urgência real de escrever isso agora.

---

## 7. Bloqueios e primeiro passo

### (a) Depende de decisão do usuário
1. **Se o texto do livro já foi colocado (Place) no InDesign ou não** — não há como eu determinar isso pelos arquivos acessíveis nesta árvore (nenhum `.indd` do livro está acessível aqui). Impacto alto: muda se o próximo passo é "fazer o Place" ou "auditar o que já foi colocado".
2. **Template de estilos do InDesign do livro** (§6.3, item 2) — sem ele, a tabela de mapeamento da rota (a) não pode ser preenchida.
3. **Tratamento dos `MARCADOR_FOTO`** — inserção manual de imagem vs. script — muda o escopo do script de pós-processamento.
4. **Confirmar rota (a) vs. (b) vs. (c)** — dei uma recomendação (§5), mas é decisão do usuário, não minha.
5. As pendências de qualidade **já documentadas em `CHECKPOINT.md`** (marcadores tipo "individual" com várias imagens, `[FOTO 361]` duplicado, zona de apêndice com estrutura diferente) — já foram conscientemente aceitas pelo usuário antes desta tarefa; só relisto aqui porque afetam diretamente o §6.2 item 5 (onde posicionar cada imagem).

### (b) Resolvível sozinho, com autorização para escrever
- Construir a ponte de índice `paragrafo_idx` (genealogia) ⇄ posição real no `_limpo.docx` (§2.3, §4.4) — trabalho de código, sem incerteza técnica, mas depende de decisão de arquitetura (onde esse crosswalk mora: no módulo de genealogia, no DocxPrep, ou em um script à parte).
- Adaptar `AuditarEstilos.jsx`/`RelatorioDialetoWord.jsx` de `scripts-indesign\` para o vocabulário de estilos da Família Pires — é preenchimento de configuração sobre uma ferramenta que já funciona, não escrita do zero.
- Escrever o exportador TXT/JSON da rota (b), ou o serializador ICML da rota (c), **se** o usuário preferir uma dessas em vez da (a) recomendada.

### Primeiro passo concreto

**Perguntar ao usuário se o `.docx` do livro já foi colocado (Place) no InDesign.** Essa resposta determina se o próximo passo técnico é "abrir o `.indd` e rodar `AuditarEstilos.jsx` para ver o que realmente chegou" (se já foi colocado) ou "preparar a tabela de mapeamento de estilo antes do primeiro Place" (se ainda não foi) — e nenhuma das duas frentes tem sentido começar sem essa resposta primeiro.

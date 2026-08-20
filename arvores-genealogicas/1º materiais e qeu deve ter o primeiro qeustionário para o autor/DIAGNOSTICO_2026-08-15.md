# Diagnóstico completo — Árvores Genealógicas (Família Pires)

> Modo somente leitura. Nenhum arquivo de produção foi criado, alterado ou
> apagado nesta tarefa, exceto este próprio relatório. `DocxPrep` não foi
> tocado (apenas listado). `infograficos/quantidade-alunos/` e `tabelas/` não
> foram tocados. Nenhum dado foi inventado, normalizado ou corrigido; onde a
> resposta não pôde ser determinada com os artefatos existentes, está
> marcada **NÃO DETERMINÁVEL**.
>
> Data do diagnóstico: 2026-08-15. Metodologia: leitura direta de
> `Familia_Pires.md`, `extrator.py`, `tests/test_extrator.py`,
> `relatorio_extracao.md`, `relatorio_ambiguidades.md`, execução da suíte de
> testes (`pytest`, somente leitura, sem gerar/alterar arquivos), e um script
> Python de contagem programática sobre `dados.json` (rodado fora do
> projeto, no scratchpad da sessão — não deixou artefato aqui). Todas as
> contagens abaixo vêm dessa análise direta do JSON, não da documentação.

> ## ⚠️ Atualização (2026-08-15, mesmo dia — sessão seguinte)
> As seções **1–8 abaixo descrevem o estado ANTES da 2ª parte ser
> processada** e citam números/planos que **mudaram depois**:
> - A 2ª parte **já foi extraída** (modo combinado). Números atuais: 2
>   patriarcas, 17 troncos-raiz, 70 troncos familiares, 97 casos em
>   `relatorio_ambiguidades.md` (não mais 1/3/15/47 — ver `Familia_Pires.md`
>   seção 4, que é a fonte viva desses números).
> - **O plano de Etapa 2 mudou**: não é mais HTML/CSS/JS + `window.print()`
>   + Adobe Garamond Pro. É um **gerador de SVG em Python** (EB Garamond,
>   SIL OFL, saída para Illustrator/InDesign). A seção 6 abaixo ("Etapa 2 —
>   o que falta") e o bloqueio de fonte na seção 7 estão **desatualizados**
>   nesse ponto — ver `Familia_Pires.md` seção 5 para o plano atual.
> - Números de linha citados na seção 8 (ex.: "linhas 66–123 de
>   `relatorio_ambiguidades.md`") **não valem mais** — o arquivo foi
>   regenerado e cresceu.
> - **O que continua válido:** as seções 1–3 (inventário, rastreabilidade,
>   fidelidade da 1ª parte) e o método de trabalho. Novas seções **9 e 10**,
>   no fim deste arquivo, continuam a investigação a partir daqui.

---

## 1. Inventário real

| Arquivo | Modificado | Tamanho | Papel | Estado |
|---|---|---|---|---|
| `originais/primeira_parte_limpo.docx` | 2026-08-12 04:48 | 6.005.511 bytes | Cópia read-only da 1ª parte sanitizada (vinda do DocxPrep) | Completo |
| `extrator.py` | 2026-08-12 05:01 | 34.934 bytes | DOCX → `dados.json` + relatórios | Completo e funcional (ver §1.1) |
| `dados.json` | 2026-08-12 05:01 | 317.770 bytes | Saída estruturada da Etapa 1 | Completo **só para a 1ª parte** (por desenho — ver §2) |
| `relatorio_extracao.md` | 2026-08-12 05:01 | 18.096 bytes | Cobertura extraído vs. declarado pelo autor | Completo |
| `relatorio_ambiguidades.md` | 2026-08-12 05:01 | 28.795 bytes | 47 casos estruturais + 68 de confiança média | Completo, **nenhum caso resolvido** (ver §3.2) |
| `requirements.txt` | 2026-08-12 04:48 | 45 bytes | `python-docx`, `lxml`, `pytest` | Completo |
| `tests/test_extrator.py` | 2026-08-12 05:00 | 7.521 bytes | 26 testes unitários das funções de parsing | Completo *para o que cobre* — ver ressalva abaixo |
| `venv/` | — | 2.313 arquivos | Ambiente Python isolado | Presente, funcional |
| `.pytest_cache/`, `tests/__pycache__/`, `__pycache__/` | — | — | Cache de execução (gerado, não versionado) | Transiente, sem conteúdo de projeto |
| `Familia_Pires.md` | 2026-08-12 05:08 | 8.814 bytes | Documento de retomada | Completo |

**Verificação ativa (não apenas leitura de docs):**
- `pytest tests/ -q` → **26 passed** (confirma a alegação do `Familia_Pires.md`, testado agora, não apenas lido).
- `dados.json` tem 897 registros de pessoa/casal com `texto_original`+`paragrafo_idx`, todos preenchidos (100%) — ver §3.1.

### O que o `Familia_Pires.md` afirma existir vs. o que existe em disco
**Nenhuma divergência.** Toda a estrutura descrita no diagrama da seção 1 do
`Familia_Pires.md` bate exatamente com o que está em disco — mesmos arquivos,
mesma organização.

### O que existe em disco e não está documentado
Nada de conteúdo de projeto. Os únicos itens não mencionados explicitamente
são diretórios de cache (`__pycache__/`, `.pytest_cache/`) — artefatos
transientes de execução, já cobertos pelo `.gitignore` da raiz, sem relevância
documental.

### §1.1 Ressalva sobre os testes
Os 26 testes cobrem as **funções de parsing isoladas** (`parse_pessoa`,
`detectar_cabecalho`, `parse_resumo`, classificadores de ambiguidade). **Não
existe teste de integração/golden-file** que trave o `dados.json` gerado
inteiro (ex.: "a extração da 1ª parte deve sempre produzir exatamente 3
troncos-raiz, 15 troncos familiares, 678 pessoas"). Isso significa que, ao
generalizar o extrator para a 2ª parte, uma regressão silenciosa na 1ª parte
não seria pega automaticamente pela suíte atual.

---

## 2. Cobertura dos dados

### Patriarcas, troncos-raiz e troncos familiares hoje em `dados.json`

**1 patriarca extraído:** `patriarca-01` (par. 383) — identificação vazia no
JSON; ele é referenciado apenas pelo separador "Patriarca" e pela geração
"Filhos de Prudêncio e Fausta". Os 9 filhos listados nessa geração (todos com
`texto_original`+`paragrafo_idx`, grafia exatamente como no DOCX):

1. Leônidas de Sousa Pires (Pai Velho) — par. 389
2. Manoela de Sousa Pires — par. 390
3. Josina Pires de Castro (Vó Josina) — par. 391
4. Joséde Sousa Pires (Cabeça de Mambira) — par. 392 *(sem espaço entre "José" e "de" — preservado fielmente, confiança média)*
5. Florênça de Sousa Pires Araújo (Tia Flor) — par. 393
6. Maria de Sousa Pires (Mariquinha) — par. 394
7. João Montana de Sousa Pires (Advogado) — par. 395
8. Benedito de Sousa Pires Barros — par. 396
9. Melquiades de Sousa Pires — par. 397

Desses 9, **apenas 3 têm seção "Tronco-raiz" elaborada** nesta parte:

| Tronco-raiz | Casal (como o autor escreveu) | Filhos declarados (resumo do autor, livro inteiro) | Troncos familiares extraídos |
|---|---|---|---|
| `tronco-raiz-01` | Leônidas de Sousa Pires e Luiza Gomes de Castro | 8 | 9 |
| `tronco-raiz-02` | Manoela de Sousa Pires e Manoel Gomes de Castro | 3 | 3 |
| `tronco-raiz-03` | Josina de Sousa Pires e Manoel Gomes de Castro | 7 | 3 |

Os outros 6 filhos do patriarca (José/Cabeça de Mambira, Florênça, Maria,
João Montana, Benedito, Melquiades) **aparecem apenas nesta lista de 9
nomes** — não têm seção "Tronco-raiz" própria nos dados atuais. Consistente
com a 1ª parte cobrir só um subconjunto do livro.

**Total de troncos familiares na 1ª parte: 15** (9 + 3 + 3), somando **678
pessoas** dentro deles/geração-de-filhos (660 dentro dos 15 troncos
familiares + 18 nas listas de "1ª geração" dos 3 troncos-raiz — bate
exatamente com o total do `relatorio_extracao.md`, verificado pela contagem
programática, não copiado do relatório).

**Anomalia notável, já documentada e preservada:** `tronco-raiz-01` declara 8
filhos mas tem **9** troncos familiares — José Pires de Castro aparece duas
vezes (`tf-02`, 1º casamento, 18 pessoas; `tf-03`, 2º casamento, 34 pessoas).
Confirmado no dado bruto, não apenas na documentação.

**Achado adicional (não estava explícito no `Familia_Pires.md`):** em
`tronco-raiz-03`, **apenas 3 dos 7 filhos declarados têm seção "Tronco"
elaborada** (Conceição, Pedro, Galdina). Os outros 4 filhos listados na "1ª
geração" (par. 1215) não têm tronco familiar próprio nos dados. Diferente do
caso José Pires (duplicação, já explicada), aqui é uma **ausência**: não há,
nos artefatos atuais, como determinar se é porque (a) esses 4 filhos não
tiveram descendência registrada pelo autor, (b) a elaboração deles está mais
adiante no texto e não foi capturada, ou (c) está na 2ª parte. **NÃO
DETERMINÁVEL** sem reler o trecho relevante do DOCX diretamente — recomendo
conferência humana antes de tratar a cobertura da 1ª parte como fechada.

### Comparação contra os totais do livro (2 patriarcas / 17 troncos-raiz / 70 troncos familiares)

| Nível | Declarado (livro inteiro) | Extraído (1ª parte, `dados.json`) | Faltando |
|---|---|---|---|
| Patriarcas | 2 | 1 | 1 patriarca inteiro (0% desta métrica coberto além do 1º) |
| Troncos-raiz | 17 | 3 | 14 troncos-raiz — inclui os 6 filhos do patriarca-01 ainda não elaborados, mais todo o 2º patriarca |
| Troncos familiares | 70 | 15 | 55 |

Isso é o **esperado pelo desenho do projeto** (1ª parte primeiro, 2ª depois),
não um erro de extração — confirmado tanto pela documentação quanto pela
inspeção direta de `dados.json` (`meta.arquivo_origem` = apenas
`"primeira_parte_limpo.docx"`, nenhuma referência à 2ª parte em nenhum
registro).

### A 2ª parte já está sanitizada mas não extraída — confirmado por 2 fontes independentes
1. **`dados.json`:** `meta.arquivo_origem = "primeira_parte_limpo.docx"` —
   único arquivo de origem referenciado em todo o JSON (897 registros
   checados, nenhum aponta para outro arquivo).
2. **`DocxPrep\saida\` (lido, não alterado):** confirmei em disco que existe
   `História da Família Pires  - segunda parte do livro\História da Família
   Pires  - segunda parte do livro_limpo.docx`, já sanitizado, com
   `relatorio_auditoria.md` e `relatorio_inventario.md` próprios — mas
   **nenhuma cópia dela existe em `infograficos/arvores-genealogicas/originais/`**
   (só a 1ª parte foi copiada para dentro do módulo). Ou seja: a 2ª parte
   está pronta na origem, mas o módulo isolado ainda não a recebeu nem a
   processou.

---

## 3. Fidelidade (não "correção")

**Aviso obrigatório:** o que segue mede *fidelidade à fonte* (se o dado no
JSON corresponde ao que está escrito no DOCX sanitizado, com rastro para
conferir). Isso **não é** e não pode ser lido como atestado de veracidade
genealógica — não há como este diagnóstico confirmar se os nomes, datas ou
relações estão factualmente corretos na vida real; apenas se o que o autor
escreveu foi capturado sem invenção.

### 3.1 Rastreabilidade à fonte

Contagem programática sobre `dados.json` (todos os 897 registros de
`pessoa` e `casal`, mais os 152 registros de `geracao`/cabeçalho, que usam o
campo `rotulo_original` em vez de `texto_original` por desenho — ambos os
campos preenchidos contam):

- **897 de 897 registros de pessoa/casal (100%) carregam `texto_original` E
  `paragrafo_idx` preenchidos.**
- **152 de 152 registros de geração (cabeçalhos "Filhos de…", "Netos de…"
  etc.) carregam `rotulo_original` E `paragrafo_idx` preenchidos.**
- **Registros que perderam rastreabilidade à fonte: nenhum.** Não há, na
  extração atual, nenhuma pessoa, casal ou cabeçalho sem `paragrafo_idx` ou
  sem texto de origem associado. O princípio inviolável do projeto está
  cumprido a 100% nos dados hoje existentes.

*(Nota metodológica: minha primeira passada do script marcou 152 "faltando"
por engano — eu havia checado o campo errado, `texto_original` em vez de
`rotulo_original`, em registros de cabeçalho de geração, que usam nome de
campo diferente por desenho. Corrigido e re-verificado antes deste relatório;
registro aqui para transparência do processo.)*

### 3.2 Estado do `relatorio_ambiguidades.md` — 47 + 68 casos, **nenhum resolvido**

O arquivo lista **47 casos estruturais** (que bloqueiam a Etapa 2 se não
revisados) mais **68 casos de confiança média** (cosméticos, contados à
parte dos 47):

| Categoria | Qtde | Bloqueia Etapa 2? |
|---|---|---|
| `possivel_multiplas_pessoas` | 28 | Sim — risco estrutural (parágrafo pode conter 2+ pessoas fundidas em 1 nome) |
| `item_vs_narrativa` | 10 | Sim — narrativa/legenda pode ter escondido uma pessoa real não capturada |
| `legenda_apos_foto` | 6 | Baixo — conferência rápida |
| `nota_nao_pessoa` | 2 | Baixo — conferência rápida |
| `pessoa_sem_nome` | 1 | Baixo — conferência rápida |
| **Subtotal "47 casos"** | **47** | — |
| `pessoa_confianca_media` (à parte) | 68 | Não — cosmético (apelido/cônjuge/observação incertos, mas a pessoa em si já está correta) |

**Busquei explicitamente por qualquer marcação de resolução** (`resolvido`,
`revisado`, `aprovado`, `validado`, "decisão humana") em
`relatorio_ambiguidades.md`: **nenhuma ocorrência encontrada.** Confirmo o
que o `Familia_Pires.md` já indicava: **os 47 casos seguem 100% abertos**, e
isso **bloqueia** um piloto visual confiável — os 28
`possivel_multiplas_pessoas` em particular, se renderizados como estão,
gerariam nós de árvore com nomes fundidos incorretamente (ex.: par. 511:
`"Maria das Gaças,José Correia,Dolores,Toscano,Oscar,Mactha, Marly,Luiza,
Etelvina,Fátima e Leônidas"` — claramente vários nomes em um só registro).

---

## 4. Quais árvores serão geradas

O `Familia_Pires.md` (§5) já define a unidade: **"um HTML por tronco"**. No
vocabulário do próprio livro, "Tronco" = filho de um tronco-raiz que teve
descendência — o que no `dados.json` corresponde a **tronco familiar**
(`troncos_raiz[].troncos_familiares[]`). Portanto: **1 infográfico por
tronco familiar.**

**Hoje, com os dados da 1ª parte, isso permite gerar até 15 unidades**
(nenhuma a mais — não há dados de outros troncos até a 2ª parte ser
processada). Tabela completa, ordenada por nº de pessoas (fato contado, não
estimado):

| Tronco familiar | Pessoas | Gerações | Par. | Cabeça (como o autor escreveu) |
|---|---:|---:|---:|---|
| `tronco-raiz-02-tf-11` | **190** | 18 | 948 | João Pires de Castro (Filho de Manoela e Manoel) |
| `tronco-raiz-02-tf-10` | **121** | 20 | 796 | América Pires de Castro (filha de Manoela e Manoel) |
| `tronco-raiz-03-tf-14` | 89 | 10 | 1303 | Pedro Pires de Castro (filho de Josina e Manoel) |
| `tronco-raiz-03-tf-13` | 43 | 14 | 1225 | Conceição Pires de Castro (Cunceição) (filho de Josina e Manoel) |
| `tronco-raiz-01-tf-06` | 41 | 14 | 618 | Fausta Pires de Castro (Tota) (Filha de Leônidas e Luiza) |
| `tronco-raiz-01-tf-03` | 34 | 13 | 495 | José Pires de Castro (Zé Pires) (2º casamento) |
| `tronco-raiz-01-tf-07` | 24 | 7 | 691 | Tomázia Pires de Castro (Tomazinha) |
| `tronco-raiz-01-tf-04` | 22 | 7 | 548 | Inêz Pires de Castro (Ineizinha) |
| `tronco-raiz-03-tf-15` | 22 | 13 | 1412 | Galdina Pires de Castro (Santa) |
| `tronco-raiz-01-tf-01` | 19 | 9 | 418 | Anibal Pires de Castro |
| `tronco-raiz-01-tf-02` | 18 | 5 | 457 | José Pires de Castro (Zé Pires) (1º casamento) |
| `tronco-raiz-01-tf-05` | 14 | 6 | 588 | Etelvina Pires de Castro |
| `tronco-raiz-02-tf-12` | 14 | 6 | 1169 | Anaídes Pires de Castro (Naídes) |
| `tronco-raiz-01-tf-09` | 9 | 5 | 752 | Ricardo Pires de Castro (Coronel Ricardo) |
| `tronco-raiz-01-tf-08` | 0 | 1 | 740 | Dolores Pires Reimbolt *(sem filhos — coerente com a nota "Não tiveram filhos" no par. 751, um caso `nota_nao_pessoa` já catalogado)* |

### Encaixe em 11,3 cm × 18 cm — o que dá para afirmar e o que não dá

**NÃO DETERMINÁVEL com precisão hoje.** Não existe nenhum CSS/tipografia
definida para a Etapa 2 (nenhum arquivo `.html`/`.css`/`.js` de árvore
existe — ver §6), então não há como calcular altura real em cm por pessoa/
geração. Qualquer número de "cabe" ou "não cabe" em milímetros seria
invenção, não medição.

**O que posso oferecer, deixando claro que é heurística e não fato medido:**
ranqueando pelos números acima (fatos), os candidatos com **risco relativo
alto de estouro dos 18 cm** são os do topo da tabela — `tf-11` (190
pessoas/18 gerações) e `tf-10` (121/20) destoam por ordem de grandeza de
todo o resto (o 3º maior, `tf-14`, já cai para 89). Esses dois muito
provavelmente vão exigir fragmentação em múltiplas páginas sequenciais
(mecanismo que o próprio `Familia_Pires.md` já prevê: "Fragmentar... nunca
comprimir para caber"). Os 12 menores (≤ 43 pessoas) são os candidatos mais
seguros para caber em bloco único, mas isso **só se confirma medindo o
piloto real**, não por esta contagem.

---

## 5. Onde entram no livro

**Não existe hoje, em nenhum artefato do projeto (`dados.json`,
`extrator.py`, relatórios), qualquer mapeamento entre um tronco/árvore e sua
posição no livro** (capítulo, página ou âncora). O único "endereço" que cada
registro carrega é `paragrafo_idx` — o índice de parágrafo dentro do DOCX
lido pelo `extrator.py`, que por design **não coincide** com os índices dos
relatórios do DocxPrep (aviso explícito já presente no `Familia_Pires.md` e
confirmado no cabeçalho de `relatorio_extracao.md`).

**Uma peça relevante existe, mas em outro espaço de índices:**
`DocxPrep\saida\...\relatorio_inventario.md` (lido, não alterado) tem uma
seção **"Estrutura de capítulos"** com títulos de alta confiança e seus
índices de parágrafo, ex.: `#83: Capítulo I`, `#86: Capítulo II` — mas esses
números `#83`/`#86` são índices **do DocxPrep**, não os `paragrafo_idx` do
`dados.json`. Não são diretamente comparáveis sem uma ponte.

### Como derivar esse mapeamento (proposta, não implementada)
1. Estender `extrator.py` (ou um script separado, só leitura) para também
   detectar, dentro do **mesmo espaço de índices que já usa** (o
   `paragrafo_idx` do python-docx lido por ele), os parágrafos que são
   títulos de capítulo (mesmo padrão textual — "Capítulo I", "Capítulo II"
   etc.) e registrar `{capitulo: "Capítulo I", paragrafo_idx_inicio: N}`.
2. Para cada tronco/pessoa, o capítulo é o do último título de capítulo com
   `paragrafo_idx` menor ou igual ao dele (busca pela âncora anterior mais
   próxima) — mesma técnica já usada implicitamente pelo classificador de
   ambiguidades para outras heurísticas.
3. Página exata **não é derivável de um DOCX** (DOCX não guarda número de
   página fixo — isso só existe após paginação em tempo de renderização/
   impressão). Se for necessário no futuro, só dá para obter renderizando o
   documento e inspecionando quebras de página, não lendo o XML.

Este trabalho ainda não foi feito — é só uma proposta de caminho, registrada
aqui porque o item 5 do pedido pediu explicitamente essa análise.

---

## 6. Etapa 2 (visual) — o que falta

**Estado atual: nada começado.** Confirmado por busca no projeto inteiro
(`infograficos/arvores-genealogicas/**/*.{html,css,js,woff,woff2,otf,ttf}`):
**zero arquivos** desse tipo existem no módulo.

- **HTML/CSS/JS de fluxograma:** não existe nenhum. Nenhum protótipo, nenhum
  rascunho.
- **Fonte Adobe Garamond Pro local:** **não está presente** em nenhum lugar
  do repositório (busquei "garamond" em todo o `WordFluxAI`, fora da
  `venv/`, zero resultados). Esta é uma fonte comercial da Adobe — mesmo que
  o autor tenha licença de uso (ex.: Creative Cloud/Adobe Fonts), os
  arquivos `.woff2`/`.otf` precisam ser exportados/obtidos e colocados
  localmente no módulo antes de qualquer `@font-face`. **Isso não é algo que
  eu deva resolver sozinho** — é uma dependência de licenciamento que exige
  a fonte já adquirida pelo usuário.

### O que precisa ser construído, em ordem de dependência
1. **Fonte Adobe Garamond Pro obtida e colocada localmente** no módulo
   (`.woff2`/`.otf`) — bloqueia tudo que depender de `@font-face`; sem
   decisão/ação humana, não avança.
2. **Resolver os 28 `possivel_multiplas_pessoas` + 10 `item_vs_narrativa`**
   (§3.2) — sem isso, qualquer árvore renderizada pode mostrar nós errados
   (nomes fundidos, pessoas perdidas). Bloqueia um piloto confiável.
3. **Decisão de escopo:** piloto visual agora (1ª parte só) vs. processar a
   2ª parte primeiro — decisão já identificada como pendente pelo próprio
   `Familia_Pires.md`, ainda sem resposta.
4. **Decisão de arquitetura visual:** o novo material reaproveita
   `assets/css/base.css`/`assets/js/base.js` da raiz (que já resolve
   impressão/PDF com `print-color-adjust` e `@page` dinâmico — ver
   `docs/padrao-impressao.md`), ou fica isolado como o
   `infograficos/quantidade-alunos/`? Hoje **não há decisão registrada**
   sobre isso em `Familia_Pires.md`. Tecnicamente o padrão de impressão de
   `base.js` (altura dinâmica) resolveria bem o problema de altura variável
   das árvores — mas as variáveis de card de `base.css` são 15,6×7,6cm, não
   11,3×18cm, então em qualquer caso precisaria de variáveis novas
   específicas deste módulo.
5. **CSS do diagrama vertical indentado** (recuo progressivo, filetes finos,
   versaletes, hierarquia por peso) — não iniciado.
6. **Gerador HTML por tronco** consumindo `dados.json` — não iniciado (nem
   como script Python de pré-render, nem como JS client-side; a abordagem
   ainda não foi escolhida).
7. **Piloto do primeiro tronco**, aprovação humana da estética antes de
   gerar os demais — como já previsto no `Familia_Pires.md`.
8. **Validação de fonte embutida no PDF** gerado pelo Chrome — só é
   testável depois que os passos 1–7 existirem.

---

## 7. Bloqueios

### (a) Dependem de decisão humana
1. **Fonte Adobe Garamond Pro** — obter os arquivos locais com licença
   válida. Maior impacto: bloqueia literalmente qualquer piloto visual.
2. **28 + 10 casos estruturais de ambiguidade** (§3.2) — decisão de como
   revisar (o autor original? o usuário, com o texto completo de cada caso
   já disponível em `relatorio_ambiguidades.md`?). Alto impacto: sem isso,
   um piloto pode ficar visualmente errado e precisar ser refeito.
3. **Caminho: piloto agora vs. 2ª parte primeiro** — decisão explícita
   pendente desde a pausa do projeto (`Familia_Pires.md` §6). Impacto médio-
   alto: muda a ordem de todo o trabalho seguinte.
4. **Qual tronco usar como piloto** — "o primeiro tronco" é ambíguo: por
   ordem estrita seria `tronco-raiz-01-tf-01` (Anibal, 19 pessoas/9
   gerações — tamanho moderado, bom candidato didático), mas não há
   confirmação registrada de que é esse o pretendido.
5. **Reaproveitar `base.css`/`base.js` da raiz ou manter isolado** (ver §6,
   item 4) — impacto médio: afeta a arquitetura de todos os HTMLs gerados.
6. **Tratamento do achado do §2** (4 dos 7 filhos de `tronco-raiz-03` sem
   tronco familiar elaborado) — impacto médio: pode indicar que a cobertura
   da 1ª parte não está tão fechada quanto documentado.

### (b) Resolvíveis sozinho, com código (depois de autorização para escrever)
- Não há bug técnico pendente em `extrator.py` — código funcional, 26/26
  testes passando, sem TODOs/FIXMEs no código.
- Construir o crosswalk `paragrafo_idx` ⇄ capítulo (§5) é trabalho de código
  novo, não um bloqueio — pode ser feito a qualquer momento, em paralelo,
  sem depender das decisões (a).
- Escrever o esqueleto HTML/CSS/JS de um tronco (uma vez a fonte e a
  ambiguidade estrutural resolvidas) é implementação direta, sem
  incerteza técnica.
- Adicionar teste de integração/golden-file para `dados.json` (§1.1) — não
  bloqueia nada, mas reduz risco antes de tocar a 2ª parte.

---

## 8. Plano proposto

A ordem que o próprio `Familia_Pires.md` já registra em "Próximo passo"
(validar os casos estruturais → decidir piloto vs. 2ª parte) **continua
correta e não está desatualizada** — nada neste diagnóstico contradiz essa
lógica; na verdade os achados de hoje (§2, o caso dos 4 filhos sem tronco em
`tronco-raiz-03`, e §3.2, confirmação de que nada foi resolvido) reforçam
que validar estrutura antes de gerar visual continua sendo o caminho certo.

**Uma adição a essa ordem:** o bloqueio de fonte (§7.a.1) não depende de
nada dos outros itens e pode ser resolvido **em paralelo, desde já**, sem
esperar a validação estrutural.

### Sequência recomendada
1. **Agora, em paralelo:** (i) usuário providencia os arquivos locais da
   Adobe Garamond Pro; (ii) revisão humana dos 28 casos
   `possivel_multiplas_pessoas` — começando pelo par. 411 e seguindo a lista
   completa em `relatorio_ambiguidades.md` (linhas 66–123), cada um já com
   texto integral e índice de parágrafo prontos para conferência.
2. Revisão dos 10 `item_vs_narrativa` (mesma lógica, lista já pronta,
   linhas 16–37 de `relatorio_ambiguidades.md`).
3. **Decisão do usuário:** liberar piloto visual da 1ª parte agora, ou
   processar a 2ª parte primeiro. Recomendo o piloto agora — ele testa a
   estética e o mecanismo de impressão/paginação sem depender de mais
   extração, e o achado do §2 (filhos sem tronco em `tronco-raiz-03`) pode
   ser investigado em paralelo sem bloquear o piloto.
4. Com fonte + ambiguidades resolvidas: decidir arquitetura (reaproveitar
   `base.css`/`base.js` com variáveis novas de card, ou isolar como
   `quantidade-alunos`), depois construir o piloto de
   `tronco-raiz-01-tf-01` (Anibal — tamanho moderado, bom primeiro teste)
   ou o tronco que o usuário preferir.
5. Validar visualmente + validar fonte embutida no PDF antes de generalizar
   para os outros 14 troncos da 1ª parte.

**Primeiro passo concreto, hoje:** revisar o par. 411 em
`relatorio_ambiguidades.md` (`possivel_multiplas_pessoas`) — é o primeiro
da lista e já aparece também em `pessoa_confianca_media`, então resolvê-lo
esclarece dois casos ao mesmo tempo.

---

## 9. Extração da 2ª parte (Passo 0 da Etapa 2) — investigação e correção

> A partir daqui o modo deixa de ser somente leitura: o `extrator.py` foi
> modificado (ver abaixo) e `dados.json`/`relatorio_extracao.md`/
> `relatorio_ambiguidades.md` foram regenerados. `DocxPrep` continuou sendo
> só lido (a cópia de `segunda_parte_limpo.docx` foi feita para
> `originais/`, dentro do módulo).

### Achado estrutural
A 2ª parte **não abre com "Patriarca" nem "Tronco-raiz"** — o parágrafo 0 já
é um separador `Tronco`, e a primeira pessoa é "Sebastião Pires de Castro
(Louro) (filho de Josina e Manoel)", **quarto filho** (ele mesmo diz isso no
texto). Josina e Manoel são o casal do `tronco-raiz-03` da 1ª parte, que
ficou com só 3 dos 7 filhos elaborados na 1ª parte. **A 2ª parte continua o
tronco-raiz-03 direto, sem repetir o cabeçalho "Tronco-raiz".**

### Por que isso quebra o extrator rodado arquivo-a-arquivo
`extrator.py` cria uma instância nova por arquivo processado; `self.raiz`
começa `None`. Numa rodada isolada da 2ª parte, os 3 primeiros `Tronco`
(Sebastião, Donatila, Maria — continuação do tronco-raiz-03) foram
rejeitados como `tronco_sem_raiz` — e isso **cascateou**: os cabeçalhos de
geração que vinham depois viraram `cabecalho_sem_contexto`, e as pessoas sob
eles **não viraram nem ambiguidade — foram perdidas silenciosamente**. Esse
resultado isolado foi descartado, não ficou no projeto.

### Validação da correção: fluxo único de parágrafos
Sem alterar a lógica de parsing, alimentei a mesma classe `Extrator` com
`paragrafos(parte1) + paragrafos(parte2)` numa só passada (script
descartável, só no scratchpad da sessão). Resultado — bate **exatamente**
com o declarado pelo autor:

| | Obtido (fluxo combinado) | Esperado (livro) |
|---|---|---|
| Patriarcas | 2 | 2 |
| Troncos-raiz | 17 | 17 |
| Troncos familiares | 70 | 70 |

`tronco-raiz-03` fecha em 6 troncos familiares (3 da parte 1 + Sebastião/
Donatila/Maria da parte 2), de 7 filhos declarados — ainda sobra 1 sem
tronco elaborado, mesmo padrão já visto em outros casos, não é bug. O
segundo patriarca é **"Benedito e Salviana"** (ver nota na seção 10 e em
`Familia_Pires.md` seção 4 — coincidência de nome com "Benedito de Sousa
Pires Barros", não assumida como a mesma pessoa).

Ambiguidades: **97 no total** (47 da parte 1 + 50 novos):
`possivel_multiplas_pessoas` 28→46 (+18), `item_vs_narrativa` 10→25 (+15),
`pessoa_sem_nome` 1→11 (+10), `legenda_apos_foto` 6→13 (+7),
`nota_nao_pessoa` 2→2 (+0). Zero `tronco_sem_raiz`/`cabecalho_sem_contexto`
no fluxo combinado — o problema estrutural desaparece completamente.

### Mudança aplicada em `extrator.py` (aprovada pelo usuário antes de aplicar)
1. **Todo registro com `paragrafo_idx`** (pessoa, casal, geração/cabeçalho,
   `resumo_autor`, foto, ambiguidade) **também carrega `arquivo_origem` +
   `paragrafo_idx_local`** — o nome do `.docx` de origem e o índice do
   parágrafo DENTRO daquele arquivo. `paragrafo_idx` sozinho, no modo
   combinado, é um índice GLOBAL contínuo entre os arquivos (não reinicia na
   2ª parte) — para abrir o Word certo e achar o parágrafo, usar
   `arquivo_origem` + `paragrafo_idx_local`, não `paragrafo_idx`.
   Retrocompatível: como a 1ª parte vem primeiro na concatenação, todos os
   `paragrafo_idx` já publicados (0–1472) continuam significando exatamente
   o mesmo parágrafo de antes.
2. **Novo modo `--combinado`**: `python extrator.py --combinado <docx1>
   <docx2> [dir_saida]` concatena os parágrafos de vários `.docx` num fluxo
   único antes de processar. O modo arquivo-único antigo continua existindo
   (mesmo schema de dados, `paragrafo_idx_local` == `paragrafo_idx` nesse
   caso).
3. `relatorio_extracao.md` ganhou uma seção "Contagem de registros
   (pessoa/casal)" gerada automaticamente a cada rodada — pessoa, casal,
   pessoa+casal, cabeçalhos de geração, e o total dos quatro.
4. Nenhuma função de detecção/parsing foi alterada; os 26 testes (que testam
   essas funções puras) continuam passando sem modificação.

Números de rastreabilidade (pessoa+casal), confirmados via a mesma lógica de
contagem: 1ª parte sozinha = 745 (+152 cabeçalhos = 897 no total, contando
cabeçalhos — esse foi o "897" citado antes). Combinado = 2904 (+485
cabeçalhos = 3389 no total). Ver `Familia_Pires.md` seção 4 e
`relatorio_extracao.md` para os números sempre atualizados.

---

## 10. Auditoria: a taxa de ambiguidade cai na parte 2 — detecção ou texto do autor?

**Pergunta:** a parte 2 gerou ~0,9 ambiguidade por tronco contra ~3,1 da
parte 1 — a diferença é do texto do autor ou de detecção que não disparou?

### Metodologia
Auditoria manual de 5 troncos da parte 2, lidos direto do
`segunda_parte_limpo.docx` linha a linha contra `dados.json`:
`tronco-raiz-03-tf-16` (Sebastião, 17 pessoas), `tronco-raiz-04-tf-20`
(Isaura, 260 pessoas — o maior da parte 2, bom teste de estresse),
`tronco-raiz-05-tf-36` (Prudêncio Pires Martins, 54 pessoas),
`tronco-raiz-11-tf-50` (Manoel Pires Campos, 32 pessoas),
`tronco-raiz-13-tf-60` (Abel Pires de Oliveira, 87 pessoas). Complementado
por 3 varreduras programáticas na íntegra dos dois `.docx` (não só nos 5
troncos), para não depender só de amostra pequena:

1. Nomes de pessoa com mais de 3 palavras (heurística p/ achar possíveis
   fusões de várias pessoas numa só) — quase todos eram nomes brasileiros
   legítimos de 4 partes (ex. "Aniceto de Araújo Pires"); nenhuma fusão real
   escapou nos 5 troncos.
2. Varredura de parágrafos logo após `[FOTO]` que **parecem** legenda solta
   (sem `/`, sem vírgula, sem conector "de/da/do", ≥3 tokens) mas **não**
   foram pegos por `eh_legenda_apos_foto` — rodada nos dois `.docx`
   inteiros, não só na amostra.
3. Pessoas com vírgula dentro do campo `conjuge` (não checado por
   `multiplas_pessoas_no_nome`, que só olha `nome`) — rodada no
   `dados.json` combinado inteiro.

### Achados

**(1) Ponto cego real em `eh_legenda_apos_foto`, mas simétrico entre as
partes.** Exemplo confirmado por leitura direta do DOCX
(`tronco-raiz-04-tf-20`, par. local 271–272 da parte 2):
```
271 | MARCADOR_FOTO | [FOTO 134 — individual]
272 | Body Text     | Aniceto Merita Anália José Juliana Manoel Niltácio
                       Jaime Maria Luziêta Marisa Benito Valmir
```
É claramente uma legenda (lista de primeiros nomes dos 12 filhos citados na
foto — mesmo padrão de outras legendas já corretamente capturadas), mas
`eh_legenda_apos_foto` não pega porque: sem "/", sem vírgula, sem conector
de/da/do, e sem nenhuma palavra-gatilho (data/local/"foto"/"legenda") nas
suas três condições — cai no branch de estilo, que só dispara se
`estilo != "Body Text"`, e aqui o estilo É "Body Text". Resultado: virou uma
"pessoa" falsa com nome de 12 palavras (capturada com `confianca=media`,
mas não como ambiguidade).

Contagem exata via varredura no arquivo inteiro:
- **Parte 2: 5 casos** desse tipo.
- **Parte 1: 6 casos** do mesmo tipo, mesmo motivo (varredura equivalente
  rodada na 1ª parte para comparar).

**A taxa é proporcionalmente parecida (até um pouco menor na parte 2:
5/3003 parágrafos ≈ 0,17% vs 6/1473 ≈ 0,41%) — não explica uma diferença de
~3,4× na taxa de ambiguidade por tronco.** É um bug real, pequeno (11 casos
no total, ambas as partes), que vale corrigir algum dia (adicionar
`legenda` como termo-gatilho, ou remover a exigência de estilo != "Body
Text" quando as outras condições já isolam bem o padrão), mas não é a causa
do que foi perguntado.

**(2) Achado isolado, confirma outro ponto cego menor:** 1 pessoa (parte 1,
par. local 716) com vírgula dentro do campo `conjuge` — não checada por
`multiplas_pessoas_no_nome`. Volume desprezível, não muda a conclusão.

**(3) Nos 5 troncos lidos à mão, todo caso de 2+ nomes num parágrafo só
(separados por vírgula) FOI corretamente capturado** como
`possivel_multiplas_pessoas` (ex.: par. local 1890 e 2472 da parte 2,
listas de nomes de netos/bisnetos separadas por vírgula — ambas viraram
ambiguidade normalmente).

### Conclusão
A diferença de taxa (~0,9 vs ~3,1 por tronco) é **majoritariamente do texto
do autor, não de detecção que falhou**. Confirmado por leitura direta: a
parte 2 usa, de forma muito mais consistente, **um parágrafo por pessoa**
("Nome/Cônjuge," em linha própria — visto em dezenas de parágrafos
consecutivos ao longo dos 5 troncos amostrados); a parte 1 tem mais
parágrafos com 2+ nomes espremidos numa linha só (que é exatamente o que
`possivel_multiplas_pessoas` mede) e mais trechos narrativos intercalados
entre as pessoas (`item_vs_narrativa`). Os dois pontos cegos de detecção
encontrados são reais, verificados por varredura completa (não só amostra),
mas pequenos (11 casos) e simétricos entre as partes — não são a causa da
diferença observada.

### ⚠️ Atualização (2026-08-15, mesmo dia — sessão seguinte): ponto cego corrigido
Dos 11 casos, **7 viravam "pessoa" falsa** (nome = lista inteira de nomes
soltos; às vezes `confianca=alta`, sem marcação nenhuma) — par. local 417 e
718 (`primeira_parte_limpo.docx`), 272, 850, 2020, 2026 e 2731
(`segunda_parte_limpo.docx`). Rastreei os 11 um a um contra o `dados.json`
antes de decidir: nenhum estava em `ambiguidades` sob nenhuma categoria.
Criada a função `eh_legenda_sem_conector()` (com testes) e a categoria
`legenda_sem_conector` em `relatorio_ambiguidades.md`; os 7 casos migraram
de "pessoa" falsa para essa ambiguidade nova. `dados.json` regenerado:
104 casos no total (era 97), 2579 pessoa (era 2586).

Os **outros 4** (par. local 332, 337, 353, 744, todos na 1ª parte) **ficaram
como estavam, por decisão explícita do usuário**: não são "pessoa" nem
ambiguidade — caem em `biografia_idx` porque estão fora de qualquer escopo
de geração (`self.escopo == "nenhum"`), em trechos de ensaio
narrativo/histórico entre seções, não em listagem genealógica. Por isso
`eh_legenda_apos_foto`/`eh_legenda_sem_conector` nunca chegam a ser chamadas
para eles no pipeline real (só rodam dentro do bloco `if self.escopo in
(...)`). Confirmado por leitura direta que não há indício de pessoa
genealógica nova perdida nesses 4 (3 citam gente que já tem registro em
outro tronco; o 4º repete o nome do próprio casal-tronco, já capturado).
Ver `Familia_Pires.md` seção 4 para o texto completo dessa decisão.

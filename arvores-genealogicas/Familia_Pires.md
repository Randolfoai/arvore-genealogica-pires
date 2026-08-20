# Família Pires — Árvores Genealógicas (estado do projeto)

> Documento de **retomada**. Objetivo: qualquer pessoa (ou IA) entender o
> estado do projeto sem reler o histórico do chat. Etapa 1 **concluída para
> o livro inteiro** (2 patriarcas / 17 troncos-raiz / 70 troncos familiares);
> Etapa 2 (gerador de SVG) **em andamento**. Última atualização: 2026-08-15.

---

## 1. Escopo e isolamento

- Todo este módulo vive **apenas** em `infograficos/arvores-genealogicas/`.
- O **Python é contido nesta subpasta**: `venv/` próprio + `requirements.txt`
  (`python-docx`, `lxml`, `pytest`). Rodar sempre com
  `./venv/Scripts/python.exe`.
- **NUNCA tocar** (fora do escopo, sob risco de quebrar o que já funciona):
  - `infograficos/quantidade-alunos/` — infográfico isolado por decisão
    deliberada do autor original.
  - `tabelas/` — outro material do WordFluxAI.
  - A **raiz do WordFluxAI**, que é **zero-dependência** (HTML/CSS/JS vanilla,
    sem `package.json`). Nenhuma dependência deste módulo pode subir para lá.
  - O projeto **DocxPrep** (`D:\IABRACADABRA\projetosweb\DocxPrep`): é a
    **origem read-only** dos DOCX sanitizados. Nunca escrever nele.

### Estrutura atual da subpasta
```
infograficos/arvores-genealogicas/
├── originais/
│   ├── primeira_parte_limpo.docx      (cópia read-only vinda do DocxPrep)
│   └── segunda_parte_limpo.docx       (cópia read-only vinda do DocxPrep)
├── extrator.py                        (DOCX → dados.json + relatórios)
├── dados.json                         (gerado — Etapa 1, COMBINADO: as duas partes)
├── relatorio_extracao.md             (gerado)
├── relatorio_ambiguidades.md         (gerado)
├── requirements.txt
├── venv/                              (isolado)
├── tests/
│   └── test_extrator.py              (26 testes, todos verdes)
├── DIAGNOSTICO_2026-08-15.md         (diagnóstico + auditoria; ver seções 9-10)
└── Familia_Pires.md                  (este arquivo)
```

### Como rodar
```
cd infograficos/arvores-genealogicas
./venv/Scripts/python.exe -m pytest tests/ -q      # testes

# combinado (RECOMENDADO — trata as duas partes como um so livro; necessario
# porque um tronco-raiz comeca numa parte e continua na outra):
./venv/Scripts/python.exe extrator.py --combinado originais/primeira_parte_limpo.docx originais/segunda_parte_limpo.docx .

# arquivo unico (legado — reprocessa so uma parte isolada; NAO fecha os
# totais do livro sozinho, ver aviso na secao 4):
./venv/Scripts/python.exe extrator.py originais/primeira_parte_limpo.docx
```
Em qualquer um dos dois modos, o último argumento (se não terminar em
`.docx`) é o diretório de saída; sem ele, usa o diretório atual.

---

## 2. Modelo de dados (decisões tomadas)

### Estrutura real do documento (4 níveis conceituais)
1. **Patriarca** (2 no livro; só 1 na primeira parte: Prudêncio de Sousa Pires).
2. **Tronco-raiz** (17 no livro; **3 na primeira parte**). Filhos do patriarca.
   Cada um abre com um parágrafo separador de texto "Tronco-raiz", seguido do
   casal-raiz e de uma frase-resumo do autor.
3. **Tronco familiar** (70 no livro). Cada filho de um tronco-raiz que teve
   descendência abre uma sub-seção com separador de texto "Tronco".
4. **Gerações** dentro de cada tronco familiar ("Filhos de…", "Netos de…",
   "Bisnetos de…", "Trinetos de…", "Tataranetos de…").

### Decisão: "aninhado e fiel" (NÃO achatar níveis)
`dados.json` preserva a hierarquia como o autor escreveu:
`patriarcas[]` · `troncos_raiz[]` → `{casal, resumo_autor, primeira_geracao,
troncos_familiares[]}` → cada tronco familiar → `{casal, biografia_idx,
geracoes[]}` → cada geração → `pessoas[]`.
**Motivo:** os cabeçalhos do autor são **relativos** ("Filhos de Conceição",
depois "Filhos de Leoiza" que é filha de Conceição) e **aninhados por
sub-grupos**, sem indentação. Achatar em níveis absolutos (netos/bisnetos da
raiz) exigiria reinterpretar/recalcular o nível de cada pessoa — ou seja,
"corrigir" o autor. Mantemos `rotulo_original` de cada cabeçalho e deixamos a
interpretação de nível absoluto para depois da validação humana.

### Princípio inviolável
- Toda pessoa/geração carrega **`texto_original`** e **`paragrafo_idx`**.
- Nada é inventado, normalizado ou corrigido: grafia, acentuação, apelidos,
  observações e ordem exatamente como no DOCX. Typos do autor são preservados
  (ex.: "Flhos", "Beatris", "eDevani", "Manel").
- ⚠️ Os **`paragrafo_idx` são os índices lidos pelo python-docx neste script**
  e **NÃO coincidem** com os índices dos relatórios do DocxPrep.
- ⚠️ **Desde o processamento combinado (2026-08-15), `paragrafo_idx` é um
  índice GLOBAL**, contínuo entre `primeira_parte_limpo.docx` e
  `segunda_parte_limpo.docx` (não reinicia em zero na segunda parte) — isso
  foi necessário porque um tronco-raiz (o de Josina e Manoel) começa na
  primeira parte e continua na segunda, sem repetir o separador
  "Tronco-raiz". **Para achar o parágrafo abrindo o Word de UM dos arquivos,
  use `arquivo_origem` + `paragrafo_idx_local` (índice DENTRO daquele
  arquivo), não `paragrafo_idx` sozinho.**

### Campos de `pessoa` (e contrato de dados para a Etapa 2)
`nome`, `apelido`, `conjuge`, `observacao`, `confianca` (`alta`|`media`),
`texto_original`, `paragrafo_idx`, **`arquivo_origem`**, **`paragrafo_idx_local`**.
Cônjuge vem por barra ("Nome/Cônjuge"); parênteses viram apelido (curto) ou
observação (status/datas/"faleceu…").

Os dois últimos campos (`arquivo_origem`, `paragrafo_idx_local`) também estão
presentes em **todo** registro com `paragrafo_idx`: casal (raiz e tronco),
cabeçalho de geração (que usa `rotulo_original` em vez de `texto_original`),
`resumo_autor`, marcador de foto e cada caso em
`relatorio_ambiguidades.md`. Isso é parte do **contrato de dados** que o
gerador de SVG da Etapa 2 deve respeitar: qualquer nó desenhado precisa
carregar `arquivo_origem` + `paragrafo_idx_local` (além do `paragrafo_idx`
global) para a revisão conseguir achar o parágrafo no Word original.
`biografia_idx` é a única exceção — segue sendo uma lista simples de índices
globais (não é um "registro" revisável, é só um marcador de onde há
narrativa/biografia).

---

## 3. Escopo de execução

- **Decisão original — primeira parte primeiro — cumprida e superada.**
  O extrator foi construído e validado só na primeira parte; a segunda foi
  processada em 2026-08-15, no **modo combinado** (ver seção 4). As duas
  partes juntas fecham exatamente com a narrativa mestra do livro.
- **Índice mestre de validação** ("Troncos familiares participantes do V
  Encontro"): não foi procurado nesta rodada — não foi necessário, porque os
  totais fecharam exatamente (2 patriarcas / 17 troncos-raiz / 70 troncos
  familiares) usando só a narrativa mestra do livro e as frases-resumo do
  autor. Se algum dia os totais não baterem, aí vale procurar essa lista.
- ⚠️ **As frases-resumo do autor (`resumo_autor`) só existem para os 3
  troncos-raiz da primeira parte** (Leônidas, Manoela, Josina) — nenhum dos
  14 troncos-raiz da segunda parte tem essa frase capturada. Não
  determinado se é porque o autor não repetiu esse padrão de texto na
  segunda parte ou porque a extração não achou o padrão lá; não investigado
  nesta rodada. Efeito prático: a tabela "extraído vs. declarado por nível"
  em `relatorio_extracao.md` só existe para os 3 primeiros troncos-raiz.

---

## 4. Status atual — Etapa 1 (extração) — CONCLUÍDA para o livro inteiro (2026-08-15)

Processamento combinado (`extrator.py --combinado`) rodado sobre as duas
partes como um fluxo único de parágrafos — necessário porque o tronco-raiz de
Josina e Manoel começa na primeira parte e continua na segunda sem repetir o
separador "Tronco-raiz" (ver `DIAGNOSTICO_2026-08-15.md`, seção 9, para a
investigação completa que levou a essa decisão).

### ⚠️ Propriedade estrutural do parser (não só um número que bateu — leia antes de rodar o extrator isolado em qualquer parte)
Os 3 troncos familiares de Sebastião (`tf-16`), Donatila (`tf-17`) e Maria
(`tf-18`) — filhos de Josina e Manoel (tronco-raiz-03) — **só são capturados
corretamente no modo `--combinado`**. Isso **não é conteúdo partido entre os
dois `.docx`**: confirmado por leitura direta, a 1ª parte termina (par. local
1472) com uma biografia completa, sem corte, e a 2ª parte começa (par. local
0) com um separador `Tronco` limpo. O casal, a biografia e os filhos de
Sebastião/Donatila/Maria estão **inteiramente dentro da 2ª parte** — nenhum
campo deles está fisicamente dividido entre arquivos.

**O que atravessa a fronteira é o vínculo com o tronco-raiz pai**, não o
conteúdo dos filhos. No livro isso é normal — como em qualquer hierarquia de
seções, o autor não repete "Tronco-raiz: Josina e Manoel" antes de cada
"Tronco" novo; o cabeçalho pai vale até o próximo cabeçalho do mesmo nível.

`extrator.py` modela isso com uma **variável de estado em memória**
(`self.raiz`, atributo de `Extrator`): é atribuída quando aparece um
separador "Tronco-raiz" e **permanece o mesmo objeto** enquanto vários
"Tronco" são processados em sequência — só muda quando um *novo*
"Tronco-raiz" aparece. Essa variável **não é recalculada a partir do texto
de cada parágrafo** — em particular, `_sep()` não lê a frase "(filho de
Josina e Manoel)", que está literalmente escrita no cabeçalho do próprio
Sebastião, para inferir o pai. A decisão é só por ordem/estado, mesmo com a
pista textual disponível ali do lado.

Rodando a 2ª parte **sozinha** (`Extrator()` novo, `self.raiz` começa
`None`), o "Tronco" do Sebastião (par. local 0 do arquivo) chega **antes de
qualquer "Tronco-raiz" existir nesse arquivo** — o "Tronco-raiz: Josina e
Manoel" que deveria valer está 1205 parágrafos atrás, na 1ª parte, que essa
rodada isolada nunca lê. `self.raiz` continua `None` → o "Tronco" vira
ambiguidade `tronco_sem_raiz` em vez de ser criado, e isso **cascateia**: os
cabeçalhos "Filhos de Sebastião e Justina..." também falham
(`cabecalho_sem_contexto`) e as pessoas sob eles nem chegam a virar
ambiguidade — caem em "escopo nenhum" e são absorvidas como biografia solta,
sem rastro.

No modo `--combinado`, `self.raiz` é atribuído a tronco-raiz-03 no parágrafo
global 1205 (1ª parte) e **continua o mesmo objeto em memória** até o
parágrafo global 1645 (início da 2ª parte, par. local 172), quando um novo
"Tronco-raiz" o substitui (tronco-raiz-04). Os 3 "Tronco" entre esses dois
pontos encontram `self.raiz` já apontando pro tronco certo, herdado da 1ª
parte, e são anexados corretamente.

**Consequência para qualquer revisão futura:** isso é uma **propriedade
estrutural do parser**, não uma coincidência desta rodada. Qualquer
tronco-raiz cujos filhos comecem a ser elaborados perto do fim de um arquivo
e continuem no próximo **só é capturado certo no modo combinado** — nunca
rodando os arquivos em separado, mesmo que o texto de cada um esteja
perfeitamente íntegro dentro do seu próprio arquivo. Vale também se o livro
um dia ganhar uma 3ª parte, ou se alguém rodar só um trecho isolado do meio.

- **2 patriarcas** — Prudêncio e Fausta (primeira parte); **Benedito e
  Salviana** (segunda parte, par. global 3653).
- **17 troncos-raiz** e **70 troncos familiares** — **fecha exatamente** com
  a narrativa mestra do livro (2/17/70). Não é coincidência isolada: os três
  números batem ao mesmo tempo, o que é uma validação forte de que a
  estrutura foi entendida corretamente.
- **Anomalia fiel confirmada e preservada:** José Pires aparece como **2
  troncos** (dois casamentos, 1º e 2º) — por isso o tronco-raiz de Leônidas
  tem 9 troncos vs 8 filhos declarados. Correto, não é bug.
- ⚠️ **Item estrutural para a revisão do autor, sem suposição da minha
  parte:** o patriarca 2 é identificado no texto como **"Benedito e
  Salviana"**. Há uma coincidência de nome com **"Benedito de Sousa Pires
  Barros"**, um dos 9 filhos listados do patriarca 1 (Prudêncio), que não tem
  tronco-raiz próprio elaborado nos dados. **Não assumo que sejam a mesma
  pessoa** — a extração não tem como decidir isso (nomes diferentes:
  "Benedito de Sousa Pires Barros" vs "Benedito" do casal "Benedito e
  Salviana"; sobrenomes diferentes no texto). Fica registrado aqui como algo
  para o autor confirmar ou descartar quando revisar.
- **Contagem de registros (rastreabilidade), pessoa+casal:**
  - Primeira parte sozinha (baseline anterior): 687 pessoa + 58 casal =
    **745** (+ 152 cabeçalhos de geração = **897** no total, contando
    cabeçalhos — esse foi o número "897" citado em diagnósticos anteriores;
    ele já incluía os cabeçalhos, não era só pessoa/casal).
  - **Combinado (as duas partes, atual, já com a correção de
    `legenda_sem_conector` abaixo):** 2579 pessoa + 318 casal = **2897**
    (+ 485 cabeçalhos de geração = **3382** no total). Ver
    `relatorio_extracao.md`, seção "Contagem de registros", gerada
    automaticamente a cada rodada do extrator.
- **Níveis absolutos (netos/bisnetos/…) no relatório são ESTIMATIVA**
  (mapeamento relativo→absoluto heurístico), **e só existem para os 3
  troncos-raiz da primeira parte** (sem `resumo_autor` nos 14 da segunda —
  ver seção 3). A contagem **fiel por rótulo** está no detalhamento por
  tronco familiar dentro de `relatorio_extracao.md`, para todos os 70.
- Testes: **26 passando** (inalterados pela mudança — testam funções puras de
  parsing, não o pipeline de arquivo/combinação).

### 104 casos abertos em `relatorio_ambiguidades.md` (era 47 só na parte 1, depois 97 no combinado)
Distribuição: `possivel_multiplas_pessoas` (46), `item_vs_narrativa` (25),
`legenda_apos_foto` (13), `pessoa_sem_nome` (11), `legenda_sem_conector` (7),
`nota_nao_pessoa` (2) — mais `pessoa_confianca_media` (contada à parte,
cosmética). Mesma priorização de antes continua válida (estrutural →
cosmético):
1. **`possivel_multiplas_pessoas` (46)** — parágrafos onde 2+ nomes ficaram
   fundidos numa linha. **Risco estrutural — revisar primeiro.**
2. **`item_vs_narrativa` (25)** — narrativas/legendas corretamente NÃO
   viradas em pessoa; conferir se nenhuma pessoa real foi perdida.
3. `legenda_apos_foto` (13), `legenda_sem_conector` (7), `nota_nao_pessoa`
   (2), `pessoa_sem_nome` (11) — exclusões corretas, conferência rápida.
4. **`pessoa_confianca_media`** — apelido/observação/cônjuge incertos.
   **Pode esperar** (cosmético, não estrutural).

### `legenda_sem_conector` — categoria nova, criada em 2026-08-15
Auditoria da diferença de taxa de ambiguidade entre partes (seção abaixo)
achou um ponto cego real em `eh_legenda_apos_foto`: legendas soltas
pós-`[FOTO]` (lista de primeiros-nomes, sem vírgula/barra/conector
"de/da/do") escapavam da detecção quando o parágrafo tinha estilo "Body
Text" — o primeiro ramo do detector só dispara com `estilo != "Body Text"`,
e nenhuma palavra-gatilho de data/local aparecia nesses casos. Isso fazia
a legenda virar uma "pessoa" falsa, com o nome sendo a lista inteira de
nomes soltos (ex.: `nome="Aniceto Merita Anália José Juliana Manoel
Niltácio Jaime Maria Luziêta Marisa Benito Valmir"`) — às vezes com
`confianca=alta` (nomes com ≤8 palavras), completamente sem marcação.

Nova função `eh_legenda_sem_conector()` cobre exatamente esse padrão (com
testes em `tests/test_extrator.py`). Resultado: **7 dos 11 casos
encontrados na auditoria agora estão em `legenda_sem_conector`** — 2 no
`primeira_parte_limpo.docx` (par. local 417, 718), 5 no
`segunda_parte_limpo.docx` (par. local 272, 850, 2020, 2026, 2731).

**Os outros 4 (par. local 332, 337, 353, 744, todos na 1ª parte) foram
deixados como estão, por decisão explícita** — não viram "pessoa" nem
ambiguidade; caem em `biografia_idx` porque estão fora de qualquer seção
"Filhos de..." (são trechos de ensaio narrativo/histórico entre seções, não
listagem genealógica). Confirmado por leitura direta que não é conteúdo
dividido entre arquivos nem pessoa nova perdida: 3 deles citam pessoas que,
pelo contexto, já têm registro próprio em seus troncos; o 4º
("Dolores Pires Reimbolt", par. 744) é legenda repetindo o nome do próprio
casal-tronco, já capturado em `casal`. Marcá-los exigiria uma heurística
adicional (detectar legenda mesmo fora de escopo de geração), mais arriscada
e sem indício de dado genealógico único perdido — descartado por ora.

### Por que a taxa de ambiguidade cai na parte 2 (~0,9/tronco vs ~3,1/tronco) — auditado
Pergunta: a diferença é do texto do autor ou de detecção que não disparou?
Auditoria manual de 5 troncos da parte 2 contra o `_limpo.docx`, mais
varredura programática das duas partes inteiras (evidência completa e
metodologia em `DIAGNOSTICO_2026-08-15.md`, seção 10):

- **Achado real, mas não era a causa principal:** existia um ponto cego no
  detector `eh_legenda_apos_foto` — legendas soltas pós-`[FOTO]` sem vírgula,
  sem "de/da/do" e sem palavra-gatilho (data/local) escapavam da detecção.
  **5 casos assim na parte 2, mas 6 na parte 1** — a taxa desse tipo
  específico de falha era **simétrica** entre as partes (proporcionalmente
  até um pouco menor na parte 2), então **não explicava** a diferença de
  ~3,4×. **Corrigido em 2026-08-15** (ver `legenda_sem_conector` acima) —
  7 dos 11 casos viram ambiguidade nova; os outros 4 ficaram como estavam,
  por decisão explícita (não são pessoa nem listagem genealógica perdida).
- **Achado isolado à parte:** 1 caso de vírgula dentro do campo `conjuge`
  (não checado por `multiplas_pessoas_no_nome`, que só olha `nome`) — parte
  1, par. local 716. Baixo volume, não muda a conclusão.
- **Conclusão:** a diferença de taxa é **majoritariamente do texto do
  autor** — a segunda parte tem, na leitura direta, uma formatação muito mais
  consistente de "uma pessoa por parágrafo", enquanto a primeira parte
  cramma com mais frequência 2+ nomes num parágrafo só (com vírgula) e tem
  mais trechos narrativos intercalados entre as pessoas. Os dois pontos
  cegos de detecção encontrados são reais e válidos para corrigir algum dia,
  mas são pequenos (11 casos no total) e não são a causa da diferença
  observada.

---

## 5. Etapa 2 (visual) — gerador de SVG — EM ANDAMENTO (decidido em 2026-08-15)

> **Plano anterior (HTML/CSS/JS + `window.print()` + Adobe Garamond Pro)
> CANCELADO e substituído por este.** Se algum documento antigo (ex.:
> `DIAGNOSTICO_2026-08-15.md`, escrito antes dessa decisão) mencionar
> HTML/CSS/JS, navegador ou Adobe Garamond Pro para a Etapa 2, está
> **desatualizado** — vale o que está aqui.

- **Ferramenta:** script **Python** (não HTML/CSS/JS, não navegador, sem
  Paged.js) que lê `dados.json` e escreve **um `.svg` por tronco familiar**.
- **Formato:** `width="113mm" height="182.03mm"` no elemento raiz + `viewBox`
  coerente. Tem que abrir no Illustrator/InDesign no tamanho físico exato,
  sem escala — conferir isso num arquivo real antes do lote.
- **Tipografia:** **EB Garamond** (SIL OFL 1.1, não Adobe Garamond Pro) — já
  em `fontes/`: `EBGaramond-Regular.ttf`, `EBGaramond-Italic.ttf`,
  `EBGaramond-SemiBold.ttf`, `EBGaramond-SemiBoldItalic.ttf`, `OFL.txt`.
  Regular = maioria dos nomes; SemiBold = nome do casal-tronco; Italic =
  datas/anotações; SemiBoldItalic = fallback destaque+data. **Medição de
  texto via `fontTools`** (métricas reais do arquivo de fonte do peso usado),
  nunca estimativa por contagem de caracteres — decide se o nome cabe e qual
  layout usar. Texto sai como `<text>` vivo (editável no Illustrator), nunca
  convertido em curvas.
- **Layout — dois modos sobre a mesma estrutura de dados:**
  - **Modo B (padrão):** ramificado esquerda→direita (casal à esquerda,
    filhos empilhados à direita).
  - **Modo A (fallback):** indentado vertical (gerações por recuo, crescendo
    para baixo), aceita qualquer profundidade/nº de irmãos.
  - Regra: calcular a largura necessária no modo B; se ultrapassar 113 mm,
    cai pra modo A automaticamente. Registrar no relatório qual modo foi
    usado em cada tronco e por quê. Se nem o modo A couber em 182,03 mm de
    altura, **não encolher a fonte** — gerar assim mesmo, marcar como
    excedente, listar no relatório (decisão caso a caso, não automática).
- **Nós sem descendência elaborada** (ex.: os filhos de tronco-raiz que não
  têm "Tronco" próprio) aparecem como nós terminais — omitir inventaria uma
  ausência que o autor não escreveu.
- **Marcação de revisão:** todo nó vindo de um caso de
  `relatorio_ambiguidades.md` (principalmente `possivel_multiplas_pessoas` e
  `item_vs_narrativa`) sai com contorno tracejado + marcador numerado,
  **tudo dentro de um único `<g id="revisao">`** (nada de marcação fora
  dele — permite esconder/apagar a camada inteira no Illustrator, nas 70 de
  uma vez). Cada marcador remete a legenda com `texto_original` +
  `paragrafo_idx` (e, pelo contrato de dados da seção 2, `arquivo_origem` +
  `paragrafo_idx_local`).
- **Nomenclatura:** `tronco-NN-NN.svg` (id estável, não nome da família) +
  **folha-índice** (arquivo, tronco-raiz, tronco familiar, casal, nº de nós,
  nº de nós marcados, modo de layout, se excedeu a mancha).
- **Cor:** escala de cinza; hierarquia por peso tipográfico, não por cor.
- **Princípio de manutenção:** o SVG é artefato **regerável**, não fonte —
  correção de conteúdo entra em `dados.json` (via `extrator.py`) e regera;
  nunca editar um SVG manualmente para "ficar correto".
- **Ordem de execução:** (1) ler este arquivo — feito; (2) extração da 2ª
  parte — feito (seção 4); (3) implementar o gerador, testar num tronco só
  com nome longo real (ex.: "Prudêncio de Sousa Pires"), mostrar o SVG antes
  de seguir; (4) só depois gerar o lote completo + folha-índice.

---

## 6. Próximo passo quando retomar

A opção (b) da versão anterior deste documento ("processar a segunda parte
primeiro") **já foi feita** (2026-08-15, livro inteiro fechado: 2/17/70).
Falta:

1. **Validar amostra** dos casos de risco estrutural: os 46
   `possivel_multiplas_pessoas` + os 25 `item_vs_narrativa` (ver
   `relatorio_ambiguidades.md`, cada caso traz texto integral + índice, agora
   também com `arquivo_origem` + `paragrafo_idx_local` para abrir o Word
   certo).
2. **Confirmar com o autor** a nota da seção 4 sobre "Benedito e Salviana"
   vs. "Benedito de Sousa Pires Barros".
3. Rodar o gerador de SVG (Etapa 2, seção 5) num tronco só, mostrar o
   resultado e só depois gerar o lote completo.

> Observação: existe uma frente paralela no WordFluxAI (correção de um bug de
> PDF na tabela) que não faz parte deste módulo e foi tratada separadamente.

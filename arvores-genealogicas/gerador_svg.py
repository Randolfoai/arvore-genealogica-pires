# -*- coding: utf-8 -*-
"""
Gerador de SVG das arvores genealogicas (Familia Pires) -- Etapa 2.

Le dados.json e escreve um .svg por tronco familiar, em tamanho fisico
exato (113mm x 182.03mm), para abrir no Illustrator/InDesign sem escala.
Nao usa navegador, nao usa Paged.js. Texto sai como <text> vivo (nao
convertido em curvas). Medicao de texto via fontTools, lendo as metricas
reais dos arquivos de fonte -- nunca estimativa por contagem de caracteres.

REGRA INVIOLAVEL: nada e inventado. Todo texto exibido (nome, apelido,
conjuge, observacao, rotulo_original) sai exatamente como esta em
dados.json. A UNICA inferencia feita aqui e ESTRUTURAL -- decidir qual
pessoa e "pai" de qual geracao para desenhar a arvore, ja que dados.json
guarda geracoes como lista relativa (nao arvore absoluta -- ver
Familia_Pires.md secao 2, "aninhado e fiel"). Essa inferencia NUNCA muda um
texto exibido, so a POSICAO de um no fielmente rotulado; usa o mesmo metodo
heuristico (interseccao de tokens de nome) ja usado por
extrator.contar_por_nivel().

Uso:
    python gerador_svg.py <tronco_id> [dir_saida]     # um tronco (piloto)
    python gerador_svg.py --lote [dir_saida]          # todos os 70 + indice

Sem dir_saida, escreve no diretorio atual.
"""
from __future__ import annotations

import base64
import json
import re
import sys
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path

from fontTools.ttLib import TTFont

import extrator as ex

AQUI = Path(__file__).resolve().parent

# --------------------------------------------------------------------------
# Constantes de layout (tudo em mm)
# --------------------------------------------------------------------------

PAGINA_L = 113.0
PAGINA_A = 182.03

MARGEM_ESQ = 4.0
MARGEM_DIR = 4.0
MARGEM_TOPO = 5.0
MARGEM_BASE = 4.0

TAM_TITULO = 3.6       # SemiBold, nome do casal-tronco
TAM_NOME = 2.5          # Regular, pessoas
TAM_OBS = 2.0            # Italic, observacao
TAM_ROTULO_GRUPO = 1.9   # Italic, subtitulo de geracao ("Netos de X:")
TAM_MARCADOR = 1.8       # numero do marcador de revisao

ALTURA_LINHA_FATOR = 1.4
INDENT_PASSO = 3.4       # por nivel de profundidade, modo A
INDENT_PASSO_B = 5.0     # deslocamento casal->filhos, modo B

COR_TEXTO = "#000000"
COR_LINHA = "#000000"
COR_REVISAO = "#c0392b"

PALETA_CINZA = ["#1e1e1e", "#3d3d3d", "#5c5c5c", "#7a7a7a", "#999999", "#b0b0b0"]

ESPESSURA_LINHA = 0.12   # mm (~0.34pt, dentro de 0,25-0,5pt pedido)

RAIO_CAIXA = 1.7         # mm, raio dos cantos das caixas (pedido: 1,5-2mm)
PAD_CAIXA_X = 1.0        # mm, folga horizontal entre texto e borda da caixa
PAD_CAIXA_Y = 0.7        # mm, folga vertical entre texto e borda da caixa
GAP_ENTRE_CAIXAS = 0.8   # mm, folga externa entre a borda de uma caixa e a proxima

# Fracoes de tamanho de fonte usadas p/ estimar o topo (acima do baseline)
# e o fundo (abaixo do baseline) do texto realmente desenhado, na hora de
# calcular os limites da caixa. EB Garamond Regular/Italic: capHeight real
# e 0.65em (tabela OS/2) e descent de linha e 0.298em (tabela hhea) --
# ASCENT_FRAC usa uma folga um pouco maior que capHeight (nao o ascent de
# linha inteiro, 1.007em, que sobra demais) pra caber acentos de maiusculas
# comuns em nomes portugueses (Ângela, Ênnio); DESCENT_FRAC ~ descent real,
# cobre descendentes (p, q, g, ç). Usadas tanto pelas pessoas quanto pelos
# titulos de casal-tronco, p/ toda caixa do arquivo ter a mesma folga.
ASCENT_FRAC = 0.80
DESCENT_FRAC = 0.30

# Paragrafos que o extrator capturou como "pessoa" (confianca=media) mas
# que na verdade sao legenda de foto ("Nome, Nome, Nome..." logo apos um
# paragrafo "[FOTO n - individual/coletiva]"), nao um novo membro da
# arvore. extrator.py ja reconhece esse padrao em outros pontos do
# documento (tipo "legenda_apos_foto" em ambiguidades) mas nao aqui.
# Confirmado par. a par. contra originais/primeira_parte_limpo.docx antes
# de excluir -- ver DIAGNOSTICO_2026-08-15.md. NAO mexe em extrator.py/
# dados.json (fora de escopo); o no soh e removido do desenho, a
# ambiguidade "possivel_multiplas_pessoas" ja registrada na legenda
# continua documentando o achado p/ conferencia.
NOS_LEGENDA_DE_FOTO = {
    ("primeira_parte_limpo.docx", 1255),  # "Mardonio, Leoiza, Ricardo,Santana, Beatrz" -- legenda do [FOTO 96]
}

NIVEL_PROFUNDIDADE = {
    "filhos": 1, "netos": 2, "bisnetos": 3,
    "trinetos": 4, "tataranetos": 5, "tetranetos": 6,
}

FONTES_DIR = AQUI / "fontes"
FONTES_INFO = {
    "regular": (FONTES_DIR / "EBGaramond-Regular.ttf", "EBGaramondPiresRegular"),
    "italic": (FONTES_DIR / "EBGaramond-Italic.ttf", "EBGaramondPiresItalic"),
    "semibold": (FONTES_DIR / "EBGaramond-SemiBold.ttf", "EBGaramondPiresSemiBold"),
    "semibold_italic": (FONTES_DIR / "EBGaramond-SemiBoldItalic.ttf", "EBGaramondPiresSemiBoldItalic"),
}


# --------------------------------------------------------------------------
# Medicao de texto via fontTools (metricas reais, nao estimativa)
# --------------------------------------------------------------------------

class Fonte:
    def __init__(self, caminho: Path, family_name: str):
        self.caminho = caminho
        self.family_name = family_name
        self._font = TTFont(str(caminho))
        self.units_per_em = self._font["head"].unitsPerEm
        self._hmtx = self._font["hmtx"]
        self._cmap = self._font.getBestCmap()
        self._fallback_glifo = self._cmap.get(ord("n")) or self._cmap.get(ord("o"))

    def largura_mm(self, texto: str, tamanho_mm: float) -> float:
        total_unidades = 0
        for ch in texto:
            nome_glifo = self._cmap.get(ord(ch), self._fallback_glifo)
            if nome_glifo is None:
                continue
            largura, _lsb = self._hmtx[nome_glifo]
            total_unidades += largura
        return (total_unidades / self.units_per_em) * tamanho_mm

    def base64_dados(self) -> str:
        return base64.b64encode(self.caminho.read_bytes()).decode("ascii")

    def base64_subconjunto(self, caracteres: set[str]) -> str:
        """Embute so os glifos realmente usados neste SVG (via
        fontTools.subset -- mesma dependencia ja instalada, sem lib nova).
        Metricas de medicao (largura_mm) continuam vindo do arquivo
        completo, carregado a parte; subsetting so afeta o que e embutido."""
        from fontTools import subset as ft_subset

        codepoints = {ord(c) for c in caracteres}
        seguro = " .,;:()/-'\"&0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
        codepoints |= {ord(c) for c in seguro}

        copia = TTFont(str(self.caminho))
        opcoes = ft_subset.Options()
        opcoes.layout_features = ["*"]  # mantem kerning/ligaduras se houver
        opcoes.notdef_outline = True
        subsetter = ft_subset.Subsetter(options=opcoes)
        subsetter.populate(unicodes=codepoints)
        subsetter.subset(copia)

        buf = BytesIO()
        copia.save(buf)
        return base64.b64encode(buf.getvalue()).decode("ascii")


def carregar_fontes() -> dict[str, Fonte]:
    fontes = {}
    for chave, (caminho, family) in FONTES_INFO.items():
        if not caminho.exists():
            raise FileNotFoundError(
                f"Fonte nao encontrada: {caminho} -- confira infograficos/"
                f"arvores-genealogicas/fontes/")
        fontes[chave] = Fonte(caminho, family)
    return fontes


def quebrar_linhas(fonte: Fonte, texto: str, tamanho_mm: float, largura_disp: float) -> list[str]:
    """Quebra 'texto' em linhas que cabem em largura_disp (mm), medindo com
    a fonte real. Nunca encolhe o tamanho -- so quebra por palavra."""
    if not texto:
        return []
    palavras = texto.split(" ")
    linhas = []
    atual = ""
    for palavra in palavras:
        candidata = f"{atual} {palavra}".strip()
        if fonte.largura_mm(candidata, tamanho_mm) <= largura_disp or not atual:
            atual = candidata
        else:
            linhas.append(atual)
            atual = palavra
    if atual:
        linhas.append(atual)
    return linhas


# --------------------------------------------------------------------------
# Construcao da arvore a partir das geracoes planas (heuristica estrutural)
# --------------------------------------------------------------------------

@dataclass
class NoArvore:
    tipo: str                      # "raiz" | "pessoa"
    pessoa: dict | None = None     # None p/ raiz
    nomes_chave: set = field(default_factory=set)
    grupos: list = field(default_factory=list)  # [{"rotulo":..,"nivel":..,"membros":[NoArvore,...]}]


def construir_arvore(tf: dict) -> NoArvore:
    raiz = NoArvore(tipo="raiz")
    for c in tf.get("casal", []):
        raiz.nomes_chave |= ex._nomes_set(c["texto_original"])

    candidatos = [raiz]  # ordem de insercao; busca linear, tronco e pequeno/medio

    for ger in tf.get("geracoes", []):
        pessoas = ger.get("pessoas") or []
        if not pessoas:
            continue  # geracao vazia (0 pessoas) nao produz conteudo visual
        ref_nomes = ex._nomes_set(ger.get("referencia") or "")
        pai = raiz
        melhor = -1
        for cand in candidatos:
            overlap = len(ref_nomes & cand.nomes_chave)
            if overlap > melhor:
                melhor = overlap
                pai = cand

        grupo = {
            "rotulo": ger.get("rotulo_original") or "",
            "nivel": ger.get("nivel"),
            "membros": [],
        }
        for p in pessoas:
            chave = (p.get("arquivo_origem"), p.get("paragrafo_idx_local"))
            if chave in NOS_LEGENDA_DE_FOTO:
                continue
            no = NoArvore(tipo="pessoa", pessoa=p)
            no.nomes_chave = ex._nomes_set(p.get("nome") or "", p.get("conjuge") or "")
            grupo["membros"].append(no)
            candidatos.append(no)
        pai.grupos.append(grupo)

    return raiz


def arvore_tem_profundidade(raiz: NoArvore) -> bool:
    """True se algum grupo esta pendurado numa pessoa (nao so na raiz) --
    isso descarta o Modo B (ramificado), que so serve p/ 1 nivel de fanout."""
    for grupo in raiz.grupos:
        for no in grupo["membros"]:
            if no.grupos:
                return True
    return False


# --------------------------------------------------------------------------
# Casal-raiz: extrai nomes limpos das linhas de texto de tf["casal"]
# --------------------------------------------------------------------------

_RE_ESPOSO = re.compile(r"^\s*espos[oa]\s*:\s*", re.IGNORECASE)


def nomes_do_casal(tf: dict) -> list[str]:
    """Extrai nomes de exibicao do casal-tronco a partir das linhas cruas em
    tf['casal']. Reaproveita parse_pessoa() do extrator (mesma logica ja
    usada/testada p/ separar nome de parenteses); so remove o prefixo
    'Esposo:'/'Esposa:' antes, que parse_pessoa nao trata."""
    nomes = []
    for c in tf.get("casal", []):
        txt = c["texto_original"].strip()
        sem_prefixo = _RE_ESPOSO.sub("", txt)
        d = ex.parse_pessoa(sem_prefixo)
        nome = (d.get("nome") or "").strip(" .,;")
        # filtra narrativa longa (parse_pessoa ja tirou os parenteses -- um
        # nome real fica curto mesmo com apelido/parentesco/nº casamento
        # dentro de parenteses; o que sobra longo aqui e prosa mesmo)
        if not nome or len(nome.split()) > 6:
            continue
        if nome not in nomes:
            nomes.append(nome)
    return nomes


# --------------------------------------------------------------------------
# Ambiguidades: indexar por (arquivo_origem, paragrafo_idx_local)
# --------------------------------------------------------------------------

def indexar_ambiguidades(dados: dict) -> dict[tuple[str, int], dict]:
    return {
        (a["arquivo_origem"], a["paragrafo_idx_local"]): a
        for a in dados.get("ambiguidades", [])
    }


def faixa_paragrafos_tronco(dados: dict, tf_id: str) -> tuple[int, int, str]:
    """Retorna (inicio_global, fim_global_exclusivo, arquivo_origem) do
    tronco tf_id, calculado a partir das posicoes de TODOS os separadores
    (patriarca/tronco-raiz/tronco) em ordem global."""
    marcos = []
    for p in dados["patriarcas"]:
        marcos.append((p["paragrafo_idx"], "patriarca", p["id"]))
    for r in dados["troncos_raiz"]:
        marcos.append((r["paragrafo_idx"], "tronco-raiz", r["id"]))
        for tf in r["troncos_familiares"]:
            marcos.append((tf["paragrafo_idx"], "tronco", tf["id"]))
    marcos.sort(key=lambda m: m[0])

    alvo_idx = next(i for i, m in enumerate(marcos) if m[2] == tf_id)
    inicio = marcos[alvo_idx][0]
    fim = marcos[alvo_idx + 1][0] if alvo_idx + 1 < len(marcos) else float("inf")
    arquivo = _achar_tronco(dados, tf_id)["arquivo_origem"]
    return inicio, fim, arquivo


def _achar_tronco(dados: dict, tf_id: str) -> dict:
    for r in dados["troncos_raiz"]:
        for tf in r["troncos_familiares"]:
            if tf["id"] == tf_id:
                return tf
    raise KeyError(f"tronco nao encontrado: {tf_id}")


def _achar_tronco_raiz_de(dados: dict, tf_id: str) -> dict:
    for r in dados["troncos_raiz"]:
        for tf in r["troncos_familiares"]:
            if tf["id"] == tf_id:
                return r
    raise KeyError(f"tronco-raiz nao encontrado p/ {tf_id}")


def ambiguidades_do_tronco(dados: dict, tf_id: str) -> list[dict]:
    """Ambiguidades cujo paragrafo cai dentro da faixa deste tronco (mesmo
    arquivo_origem, paragrafo_idx_local no intervalo)."""
    inicio, fim, arquivo = faixa_paragrafos_tronco(dados, tf_id)
    tf = _achar_tronco(dados, tf_id)
    local_inicio = tf["paragrafo_idx_local"]
    # fim em indice local: se o proximo marco for do mesmo arquivo, calcula
    # a diferenca; senao (proximo esta no outro arquivo, ou nao ha proximo),
    # a faixa vai ate o fim do arquivo atual.
    out = []
    for a in dados.get("ambiguidades", []):
        if a["arquivo_origem"] != arquivo:
            continue
        idx_global_equivalente = a["paragrafo_idx"]
        if inicio <= idx_global_equivalente < fim:
            out.append(a)
    return out


# --------------------------------------------------------------------------
# Runs de texto (nome/apelido/conjuge em Regular, observacao em Italic)
# --------------------------------------------------------------------------

def linha_principal_pessoa(p: dict) -> str:
    partes = p["nome"]
    if p.get("apelido"):
        partes += f" ({p['apelido']})"
    if p.get("conjuge"):
        partes += f" / {p['conjuge']}"
    return partes


class Escritor:
    """Acumula elementos SVG (texto/linhas/marcadores) em duas camadas:
    'miolo' (arte final) e 'revisao' (marcadores + legenda, removivel)."""

    def __init__(self, fontes: dict[str, Fonte]):
        self.fontes = fontes
        self.miolo: list[str] = []
        self.revisao: list[str] = []
        self.contador_marcador = 0
        self.legenda: list[dict] = []
        self.y_max_conteudo = 0.0
        self.caracteres_usados: dict[str, set] = {chave: set() for chave in fontes}

    def texto(self, x: float, y: float, texto: str, tamanho_mm: float,
              peso: str, cor: str = COR_TEXTO, ancora: str = "start"):
        family = self.fontes[peso].family_name
        self.caracteres_usados[peso].update(texto)
        texto_esc = (texto.replace("&", "&amp;").replace("<", "&lt;")
                     .replace(">", "&gt;"))
        self.miolo.append(
            f'<text x="{x:.3f}" y="{y:.3f}" font-family="{family}" '
            f'font-size="{tamanho_mm:.3f}" fill="{cor}" '
            f'text-anchor="{ancora}">{texto_esc}</text>'
        )
        self.y_max_conteudo = max(self.y_max_conteudo, y)

    def linha(self, x1, y1, x2, y2, cor=COR_LINHA, espessura=ESPESSURA_LINHA,
              tracejada=False, camada="miolo"):
        tracejo = ' stroke-dasharray="0.6,0.5"' if tracejada else ""
        el = (f'<line x1="{x1:.3f}" y1="{y1:.3f}" x2="{x2:.3f}" y2="{y2:.3f}" '
              f'stroke="{cor}" stroke-width="{espessura:.3f}"{tracejo}/>')
        (self.miolo if camada == "miolo" else self.revisao).append(el)

    def caixa(self, x, y, w, h, camada="miolo", fill="none"):
        stroke = "none" if fill != "none" else COR_TEXTO
        el = (f'<rect x="{x:.3f}" y="{y:.3f}" width="{w:.3f}" height="{h:.3f}" '
              f'rx="{RAIO_CAIXA:.2f}" ry="{RAIO_CAIXA:.2f}" fill="{fill}" '
              f'stroke="{stroke}" stroke-width="{ESPESSURA_LINHA:.3f}"/>')
        (self.miolo if camada == "miolo" else self.revisao).append(el)

    def retangulo(self, x, y, w, h, cor, espessura=0.1, tracejada=True, camada="revisao"):
        tracejo = ' stroke-dasharray="0.5,0.4"' if tracejada else ""
        el = (f'<rect x="{x:.3f}" y="{y:.3f}" width="{w:.3f}" height="{h:.3f}" '
              f'fill="none" stroke="{cor}" stroke-width="{espessura:.3f}"{tracejo}/>')
        (self.miolo if camada == "miolo" else self.revisao).append(el)

    def marca_revisao(self, x_ini: float, y_topo: float, largura: float, altura: float,
                       tipo: str, texto_original: str, paragrafo_idx: int,
                       arquivo_origem: str, paragrafo_idx_local: int):
        self.contador_marcador += 1
        n = self.contador_marcador
        self.retangulo(x_ini - 0.6, y_topo, largura + 1.2, altura,
                        COR_REVISAO, tracejada=True, camada="revisao")
        cx, cy = x_ini - 1.8, y_topo + altura / 2
        self.revisao.append(
            f'<circle cx="{cx:.3f}" cy="{cy:.3f}" r="1.4" fill="{COR_REVISAO}"/>')
        family = self.fontes["semibold"].family_name
        self.revisao.append(
            f'<text x="{cx:.3f}" y="{cy + 0.55:.3f}" font-family="{family}" '
            f'font-size="{TAM_MARCADOR:.3f}" fill="#ffffff" text-anchor="middle">'
            f'{n}</text>'
        )
        self.legenda.append({
            "n": n, "tipo": tipo, "texto_original": texto_original,
            "paragrafo_idx": paragrafo_idx, "arquivo_origem": arquivo_origem,
            "paragrafo_idx_local": paragrafo_idx_local,
        })


# --------------------------------------------------------------------------
# Auxiliar de tema
# --------------------------------------------------------------------------

def _cor_caixa(tema: str, profundidade: int) -> str:
    if tema != "escala-cinza":
        return "none"
    return PALETA_CINZA[min(profundidade, len(PALETA_CINZA) - 1)]


# --------------------------------------------------------------------------
# Modo A -- indentado vertical
# --------------------------------------------------------------------------

def layout_modo_a(esc: Escritor, raiz: NoArvore, x0: float, y0: float,
                  tema: str = "classico") -> float:
    """Desenha a arvore em modo indentado; retorna o y final usado."""
    y = y0

    for grupo in raiz.grupos:
        y = _desenhar_grupo(esc, grupo, profundidade=1, x_base=x0, y=y, tema=tema)
    return y


def _desenhar_grupo(esc: Escritor, grupo: dict, profundidade: int, x_base: float, y: float,
                    tema: str = "classico") -> float:
    x = x_base + (profundidade - 1) * INDENT_PASSO
    largura_disp = PAGINA_L - MARGEM_DIR - x

    fonte_rotulo = esc.fontes["italic"]
    altura_linha_rotulo = TAM_ROTULO_GRUPO * ALTURA_LINHA_FATOR
    if grupo["rotulo"]:
        for ln in quebrar_linhas(fonte_rotulo, grupo["rotulo"], TAM_ROTULO_GRUPO, largura_disp):
            y += altura_linha_rotulo
            esc.texto(x, y, ln, TAM_ROTULO_GRUPO, "italic", cor=COR_LINHA)

    x_membro = x + 1.2
    largura_membro = PAGINA_L - MARGEM_DIR - x_membro
    y_topo_grupo = y

    for no in grupo["membros"]:
        y, y_topo_no, altura_no, largura_no = _desenhar_pessoa(
            esc, no.pessoa, x_membro, y, largura_membro, tema=tema, profundidade=profundidade)

        # conector em L: trecho horizontal do tronco ate a borda da caixa
        esc.linha(x - 1.0, y_topo_no + (altura_no / 2) - (TAM_NOME * 0.3),
                  x_membro - PAD_CAIXA_X, y_topo_no + (altura_no / 2) - (TAM_NOME * 0.3))

        if no.pessoa.get("confianca") == "media":
            esc.marca_revisao(
                x_membro, y_topo_no, largura_no, altura_no,
                tipo="pessoa_confianca_media",
                texto_original=no.pessoa["texto_original"],
                paragrafo_idx=no.pessoa["paragrafo_idx"],
                arquivo_origem=no.pessoa["arquivo_origem"],
                paragrafo_idx_local=no.pessoa["paragrafo_idx_local"],
            )

        for sub_grupo in no.grupos:
            profundidade_filho = profundidade + NIVEL_PROFUNDIDADE.get(sub_grupo["nivel"], 1)
            y = _desenhar_grupo(esc, sub_grupo, profundidade_filho, x_base, y, tema=tema)

    # filete vertical ligando os membros deste grupo
    if grupo["membros"]:
        esc.linha(x - 1.0, y_topo_grupo, x - 1.0, y - (TAM_NOME * ALTURA_LINHA_FATOR * 0.4))

    return y


def _topo_altura_bloco(y_ini: float, y_fim_baseline: float, tamanho_mm: float) -> tuple[float, float]:
    """Topo/altura de um bloco de texto de 1 tamanho so (titulos de
    casal-tronco), usando ASCENT_FRAC/DESCENT_FRAC -- mesmas fracoes de
    _desenhar_pessoa, pra toda caixa do arquivo ter folga consistente.
    y_ini = cursor ANTES da 1a linha (mesma convencao de _desenhar_pessoa e
    dos rotulos de geracao: a 1a linha soma sua propria altura antes de
    desenhar). y_fim_baseline = baseline da ULTIMA linha ja desenhada."""
    altura_linha = tamanho_mm * ALTURA_LINHA_FATOR
    y_topo = (y_ini + altura_linha) - (ASCENT_FRAC * tamanho_mm)
    y_fundo = y_fim_baseline + (DESCENT_FRAC * tamanho_mm)
    return y_topo, y_fundo - y_topo


def _desenhar_pessoa(esc: Escritor, p: dict, x: float, y: float, largura_disp: float,
                     tema: str = "classico", profundidade: int = 1):
    """Desenha nome(+apelido+conjuge) em Regular e observacao em Italic
    (com quebra de linha real via fontTools), e a caixa ao redor de tudo.
    'y' e o cursor recebido do chamador, na mesma convencao usada pelos
    rotulos de geracao: a 1a linha soma sua propria altura ANTES de
    desenhar (ou seja, 'y' e o baseline do elemento anterior, nao o topo
    deste). Retorna (y_prox_cursor, y_topo_no, altura_no, largura_max):
    y_prox_cursor ja inclui a altura real da caixa (texto + padding) mais
    GAP_ENTRE_CAIXAS, pronto pra virar o 'y' do proximo irmao -- e isso
    que garante avanco >= altura da caixa, sem sobreposicao entre irmaos
    consecutivos (bug corrigido: antes o avanco usava so a altura de
    linha de texto, ignorando a caixa)."""
    cinza = tema == "escala-cinza"
    peso_nome = "semibold" if cinza else "regular"
    peso_obs  = "semibold_italic" if cinza else "italic"
    cor_txt   = "#ffffff" if cinza else COR_TEXTO
    fill_cx   = _cor_caixa(tema, profundidade)

    fonte_nome = esc.fontes[peso_nome]
    altura_linha = TAM_NOME * ALTURA_LINHA_FATOR
    y_topo_conteudo = (y + altura_linha) - (ASCENT_FRAC * TAM_NOME)

    linha_ppal = linha_principal_pessoa(p)
    linhas = quebrar_linhas(fonte_nome, linha_ppal, TAM_NOME, largura_disp) or [""]
    largura_max = 0.0
    for ln in linhas:
        y += altura_linha
        esc.texto(x, y, ln, TAM_NOME, peso_nome, cor=cor_txt)
        largura_max = max(largura_max, fonte_nome.largura_mm(ln, TAM_NOME))

    tam_ultima_linha = TAM_NOME
    if p.get("observacao"):
        fonte_obs = esc.fontes[peso_obs]
        for ln in quebrar_linhas(fonte_obs, p["observacao"], TAM_OBS, largura_disp):
            y += TAM_OBS * ALTURA_LINHA_FATOR
            esc.texto(x, y, ln, TAM_OBS, peso_obs, cor=cor_txt)
            largura_max = max(largura_max, fonte_obs.largura_mm(ln, TAM_OBS))
        tam_ultima_linha = TAM_OBS  # caixa precisa cobrir a ULTIMA linha
        # desenhada, seja ela nome ou observacao -- nao sempre TAM_NOME

    y_fundo_conteudo = y + (DESCENT_FRAC * tam_ultima_linha)

    y_topo_no = y_topo_conteudo
    altura_no = y_fundo_conteudo - y_topo_conteudo

    esc.caixa(x - PAD_CAIXA_X, y_topo_no - PAD_CAIXA_Y,
              largura_max + 2 * PAD_CAIXA_X, altura_no + 2 * PAD_CAIXA_Y,
              fill=fill_cx)

    # y_prox_cursor precisa embutir o padding de baixo desta caixa E o
    # padding de cima da PROXIMA (esc.caixa sempre desenha a partir de
    # y_topo_no - PAD_CAIXA_Y) -- os dois PAD_CAIXA_Y aqui nao sao o
    # mesmo, um fecha esta caixa e o outro abre a seguinte.
    y_fundo_caixa = y_topo_no + altura_no + PAD_CAIXA_Y
    y_prox_cursor = (y_fundo_caixa + PAD_CAIXA_Y + GAP_ENTRE_CAIXAS
                      - altura_linha + (ASCENT_FRAC * TAM_NOME))

    return y_prox_cursor, y_topo_no, altura_no, largura_max


def _medir_altura_pessoa(esc: Escritor, p: dict, largura_disp: float,
                         tema: str = "classico") -> float:
    """Mesmo avanco de cursor (topo->topo do proximo irmao) que
    _desenhar_pessoa produziria, sem desenhar nada -- usado para
    centralizar o titulo do casal no Modo B antes de saber onde os filhos
    vao parar. Precisa bater com _desenhar_pessoa, senao a centralizacao
    do titulo fica calculada contra uma altura diferente da real."""
    cinza = tema == "escala-cinza"
    fonte_nome = esc.fontes["semibold" if cinza else "regular"]
    altura_linha = TAM_NOME * ALTURA_LINHA_FATOR
    linha_ppal = linha_principal_pessoa(p)
    n_linhas = len(quebrar_linhas(fonte_nome, linha_ppal, TAM_NOME, largura_disp) or [""])
    avanco_baseline = n_linhas * altura_linha
    tam_ultima_linha = TAM_NOME
    if p.get("observacao"):
        fonte_obs = esc.fontes["semibold_italic" if cinza else "italic"]
        n_obs = len(quebrar_linhas(fonte_obs, p["observacao"], TAM_OBS, largura_disp))
        avanco_baseline += n_obs * (TAM_OBS * ALTURA_LINHA_FATOR)
        if n_obs:
            tam_ultima_linha = TAM_OBS
    return (avanco_baseline + (DESCENT_FRAC * tam_ultima_linha) + 2 * PAD_CAIXA_Y
            + GAP_ENTRE_CAIXAS - altura_linha + (ASCENT_FRAC * TAM_NOME))


# --------------------------------------------------------------------------
# Modo B -- ramificado esquerda->direita (so p/ arvores de 1 nivel: raiz + filhos)
# --------------------------------------------------------------------------

def largura_necessaria_modo_b(esc: Escritor, raiz: NoArvore, nomes_raiz: list[str],
                              tema: str = "classico") -> float:
    fonte_raiz = esc.fontes["semibold"]
    fonte_nome = esc.fontes["semibold" if tema == "escala-cinza" else "regular"]
    largura_raiz = max((fonte_raiz.largura_mm(n, TAM_TITULO) for n in nomes_raiz), default=0.0)
    largura_filhos = 0.0
    for grupo in raiz.grupos:
        for no in grupo["membros"]:
            largura_filhos = max(largura_filhos,
                                  fonte_nome.largura_mm(linha_principal_pessoa(no.pessoa), TAM_NOME))
    return MARGEM_ESQ + largura_raiz + INDENT_PASSO_B + largura_filhos + MARGEM_DIR


def layout_modo_b(esc: Escritor, raiz: NoArvore, nomes_raiz: list[str], y0: float,
                  tema: str = "classico") -> float:
    x_raiz = MARGEM_ESQ
    fonte_raiz = esc.fontes["semibold"]
    largura_raiz = max((fonte_raiz.largura_mm(n, TAM_TITULO) for n in nomes_raiz), default=0.0)
    x_filhos = x_raiz + largura_raiz + INDENT_PASSO_B
    largura_filhos_disp = PAGINA_L - MARGEM_DIR - x_filhos

    membros = [no for grupo in raiz.grupos for no in grupo["membros"]]

    # 1a passada (so medicao, nao desenha nada): altura total que os filhos
    # vao ocupar, pra centralizar o titulo do casal contra essa extensao.
    y_ini = y0 + TAM_NOME * 0.3
    y_medido = y_ini
    for no in membros:
        y_medido += _medir_altura_pessoa(esc, no.pessoa, largura_filhos_disp, tema=tema)
    altura_filhos_total = y_medido - y_ini

    # titulo do casal: a esquerda, empilhado, centralizado verticalmente
    # pela extensao dos filhos (nao no topo da pagina)
    altura_titulo_linha = TAM_TITULO * ALTURA_LINHA_FATOR
    altura_titulo_total = len(nomes_raiz) * altura_titulo_linha
    y_centro_filhos = y_ini + altura_filhos_total / 2
    y_titulo_ini = y_centro_filhos - altura_titulo_total / 2
    y_titulo = y_titulo_ini
    cor_raiz = "#ffffff" if tema == "escala-cinza" else COR_TEXTO
    for nome in nomes_raiz:
        y_titulo += altura_titulo_linha
        esc.texto(x_raiz, y_titulo, nome, TAM_TITULO, "semibold", cor=cor_raiz)
    y_fim_titulo = y_titulo
    y_topo_titulo, altura_titulo = _topo_altura_bloco(y_titulo_ini, y_titulo, TAM_TITULO)
    esc.caixa(x_raiz - PAD_CAIXA_X, y_topo_titulo - PAD_CAIXA_Y,
              largura_raiz + 2 * PAD_CAIXA_X, altura_titulo + 2 * PAD_CAIXA_Y,
              fill=_cor_caixa(tema, 0))
    x_borda_raiz = x_raiz + largura_raiz + PAD_CAIXA_X

    # 2a passada: desenha os filhos de verdade
    y_fh = y_ini
    for no in membros:
        y_fh, y_topo_no, altura_no, largura_no = _desenhar_pessoa(
            esc, no.pessoa, x_filhos, y_fh, largura_filhos_disp, tema=tema, profundidade=1)
        esc.linha(x_filhos - 2.0, y_topo_no + altura_no / 2,
                  x_filhos - PAD_CAIXA_X, y_topo_no + altura_no / 2)
        if no.pessoa.get("confianca") == "media":
            esc.marca_revisao(
                x_filhos, y_topo_no, largura_no, altura_no,
                tipo="pessoa_confianca_media",
                texto_original=no.pessoa["texto_original"],
                paragrafo_idx=no.pessoa["paragrafo_idx"],
                arquivo_origem=no.pessoa["arquivo_origem"],
                paragrafo_idx_local=no.pessoa["paragrafo_idx_local"],
            )
    if membros:
        esc.linha(x_filhos - 2.0, y_ini, x_filhos - 2.0, y_fh)
        esc.linha(x_borda_raiz, (y_ini + y_fh) / 2, x_filhos - 2.0, (y_ini + y_fh) / 2)

    return max(y_fim_titulo, y_fh)


# --------------------------------------------------------------------------
# Legenda (camada de revisao, desenhada ABAIXO da pagina 182.03mm -- arte
# fora do artboard, pratica padrao em fluxo Illustrator/InDesign; permite
# apagar/ocultar so o grupo id="revisao" sem tocar no miolo impresso)
# --------------------------------------------------------------------------

def desenhar_legenda(esc: Escritor, ambiguidades_extras: list[dict], y0: float) -> float:
    x = MARGEM_ESQ
    y = y0 + 6.0
    fam_tit = esc.fontes["semibold"].family_name
    fam_txt = esc.fontes["regular"].family_name

    esc.revisao.append(
        f'<line x1="0" y1="{y0:.3f}" x2="{PAGINA_L:.3f}" y2="{y0:.3f}" '
        f'stroke="{COR_REVISAO}" stroke-width="0.2" stroke-dasharray="1,0.6"/>')
    esc.revisao.append(
        f'<text x="{x:.3f}" y="{y:.3f}" font-family="{fam_tit}" '
        f'font-size="3" fill="{COR_REVISAO}">Camada de revisao '
        f'(apagar antes do miolo final)</text>')
    y += 5.0

    if esc.legenda:
        esc.revisao.append(
            f'<text x="{x:.3f}" y="{y:.3f}" font-family="{fam_tit}" '
            f'font-size="2.4" fill="{COR_REVISAO}">Nos marcados na arvore '
            f'(confianca=media):</text>')
        y += 4.0
        for item in esc.legenda:
            cab = (f"{item['n']}. ({item['tipo']}) `{item['arquivo_origem']}` "
                   f"par. local {item['paragrafo_idx_local']} "
                   f"(par. global {item['paragrafo_idx']})")
            esc.revisao.append(
                f'<text x="{x:.3f}" y="{y:.3f}" font-family="{fam_tit}" '
                f'font-size="2.2" fill="{COR_TEXTO}">{_esc(cab)}</text>')
            y += 3.4
            for ln in quebrar_linhas(esc.fontes["regular"], item["texto_original"],
                                      2.2, PAGINA_L - MARGEM_ESQ - MARGEM_DIR - 3):
                esc.revisao.append(
                    f'<text x="{x + 3:.3f}" y="{y:.3f}" font-family="{fam_txt}" '
                    f'font-size="2.2" fill="{COR_TEXTO}">{_esc(ln)}</text>')
                y += 3.2
            y += 1.5

    if ambiguidades_extras:
        y += 2.0
        esc.revisao.append(
            f'<text x="{x:.3f}" y="{y:.3f}" font-family="{fam_tit}" '
            f'font-size="2.4" fill="{COR_REVISAO}">Outros itens da faixa '
            f'de paragrafos deste tronco, nao capturados como pessoa '
            f'(conferir se nada foi perdido):</text>')
        y += 4.0
        for a in ambiguidades_extras:
            cab = (f"({a['tipo']}) `{a['arquivo_origem']}` par. local "
                   f"{a['paragrafo_idx_local']} (par. global {a['paragrafo_idx']})")
            esc.revisao.append(
                f'<text x="{x:.3f}" y="{y:.3f}" font-family="{fam_tit}" '
                f'font-size="2.2" fill="{COR_TEXTO}">{_esc(cab)}</text>')
            y += 3.4
            for ln in quebrar_linhas(esc.fontes["regular"], a["texto_original"],
                                      2.2, PAGINA_L - MARGEM_ESQ - MARGEM_DIR - 3):
                esc.revisao.append(
                    f'<text x="{x + 3:.3f}" y="{y:.3f}" font-family="{fam_txt}" '
                    f'font-size="2.2" fill="{COR_TEXTO}">{_esc(ln)}</text>')
                y += 3.2
            y += 1.5

    return y


def _esc(s: str) -> str:
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# --------------------------------------------------------------------------
# Montagem do SVG de um tronco
# --------------------------------------------------------------------------

def gerar_svg_tronco(dados: dict, fontes: dict[str, Fonte], tf_id: str,
                     tema: str = "classico") -> tuple[str, dict]:
    tf = _achar_tronco(dados, tf_id)
    tr = _achar_tronco_raiz_de(dados, tf_id)
    raiz = construir_arvore(tf)
    nomes_raiz = nomes_do_casal(tf) or ["(casal nao identificado)"]

    esc = Escritor(fontes)

    y = MARGEM_TOPO
    # floreio discreto no no-raiz: filete fino com pequeno losango central
    y_floreio = y
    esc.linha(MARGEM_ESQ, y_floreio, PAGINA_L - MARGEM_DIR, y_floreio, espessura=0.15)
    cx_floreio = PAGINA_L / 2
    esc.miolo.append(
        f'<rect x="{cx_floreio-0.5:.3f}" y="{y_floreio-0.5:.3f}" width="1" height="1" '
        f'fill="{COR_TEXTO}" transform="rotate(45 {cx_floreio:.3f} {y_floreio:.3f})"/>')
    y += 2.5

    # decide o modo ANTES de desenhar qualquer titulo -- Modo B desenha seu
    # proprio titulo (a esquerda, centralizado pelos filhos), Modo A usa o
    # titulo centralizado no topo da pagina.
    usa_modo_b_candidato = not arvore_tem_profundidade(raiz)
    largura_b = largura_necessaria_modo_b(esc, raiz, nomes_raiz, tema=tema) if usa_modo_b_candidato else None
    modo_usado = "B" if (usa_modo_b_candidato and largura_b <= PAGINA_L) else "A"

    if modo_usado == "B":
        y_inicio_arvore = y
        y_fim = layout_modo_b(esc, raiz, nomes_raiz, y, tema=tema)
        motivo = f"sem profundidade >1 nivel, largura calculada {largura_b:.1f}mm <= {PAGINA_L}mm"
    else:
        fonte_titulo = esc.fontes["semibold"]
        largura_titulo = max(
            (fonte_titulo.largura_mm(n, TAM_TITULO) for n in nomes_raiz), default=0.0)
        y_titulo_ini = y
        cor_raiz = "#ffffff" if tema == "escala-cinza" else COR_TEXTO
        for nome in nomes_raiz:
            y += TAM_TITULO * ALTURA_LINHA_FATOR
            esc.texto(PAGINA_L / 2, y, nome, TAM_TITULO, "semibold", cor=cor_raiz, ancora="middle")
        y_topo_titulo, altura_titulo = _topo_altura_bloco(y_titulo_ini, y, TAM_TITULO)
        esc.caixa(PAGINA_L / 2 - largura_titulo / 2 - PAD_CAIXA_X,
                  y_topo_titulo - PAD_CAIXA_Y,
                  largura_titulo + 2 * PAD_CAIXA_X, altura_titulo + 2 * PAD_CAIXA_Y,
                  fill=_cor_caixa(tema, 0))
        y += 1.5
        esc.linha(MARGEM_ESQ, y, PAGINA_L - MARGEM_DIR, y, espessura=0.15)
        y += 3.0
        y_inicio_arvore = y
        y_fim = layout_modo_a(esc, raiz, MARGEM_ESQ, y, tema=tema)
        if usa_modo_b_candidato:
            motivo = f"sem profundidade >1 nivel, mas largura calculada {largura_b:.1f}mm > {PAGINA_L}mm"
        else:
            motivo = "tem geracoes presas a pessoas especificas (profundidade >1 nivel)"

    altura_arvore_usada = y_fim - y_inicio_arvore
    excedeu = y_fim > (PAGINA_A - MARGEM_BASE)

    ambiguidades_extras = [
        a for a in ambiguidades_do_tronco(dados, tf_id)
        if a["tipo"] != "pessoa_confianca_media"
    ]

    # A legenda entra no MESMO fluxo de coordenadas do miolo, logo apos o
    # fim da arvore -- nao "fora da pagina": testado em Chrome headless e
    # conteudo alem de y=182.03 NAO aparece (o <svg> raiz corta ali mesmo
    # com overflow:visible computado), entao nao da pra confiar em arte
    # "fora do artboard" sem testar no Illustrator de verdade. Em vez disso,
    # reusa o MESMO mecanismo ja previsto na especificacao para arvore que
    # nao cabe: gera assim mesmo, altura do arquivo cresce alem de 182.03mm
    # so quando precisa (arvore + legenda), e isso fica marcado no indice.
    tem_legenda = bool(esc.legenda or ambiguidades_extras)
    if tem_legenda:
        y_legenda0 = y_fim + 3.0
        desenhar_legenda(esc, ambiguidades_extras, y_legenda0)
        altura_total_svg = max(PAGINA_A, esc.y_max_conteudo + 4.0)
    else:
        altura_total_svg = PAGINA_A

    n_pessoas = sum(len(g["membros"]) for g in _todos_grupos(raiz))
    n_marcados = len(esc.legenda)

    pesos_usados = {"regular", "semibold"}
    if any("italic" in ln for ln in esc.miolo):
        pesos_usados.add("italic")
    pesos_usados.add("italic")  # rotulos de grupo/observacao quase sempre usam
    if esc.legenda or ambiguidades_extras:
        pesos_usados.add("semibold")
    if esc.caracteres_usados.get("semibold_italic"):
        pesos_usados.add("semibold_italic")

    svg = montar_svg(esc, altura_total_svg, pesos_usados)

    info = {
        "arquivo": f"{_nome_arquivo(tf_id)}.svg",
        "tronco_raiz_id": tr["id"],
        "tronco_familiar_id": tf_id,
        "casal": " & ".join(nomes_raiz),
        "n_nos": n_pessoas,
        "n_nos_marcados": n_marcados,
        "n_itens_nao_capturados": len(ambiguidades_extras),
        "modo_layout": modo_usado,
        "motivo_modo": motivo,
        "altura_arvore_mm": round(altura_arvore_usada, 1),
        "excedeu_mancha": excedeu,
        "altura_arquivo_mm": round(altura_total_svg, 1),
        "arquivo_maior_que_pagina_nominal": altura_total_svg > PAGINA_A,
    }
    return svg, info


def _todos_grupos(no: NoArvore):
    for g in no.grupos:
        yield g
        for m in g["membros"]:
            yield from _todos_grupos(m)


def _nome_arquivo(tf_id: str) -> str:
    # tronco-raiz-01-tf-01 -> tronco-01-01
    m = re.match(r"tronco-raiz-(\d+)-tf-(\d+)", tf_id)
    if not m:
        return tf_id
    return f"tronco-{m.group(1)}-{m.group(2)}"


def montar_svg(esc: Escritor, altura_total: float, pesos_usados: set[str]) -> str:
    defs_fontes = []
    for chave in pesos_usados:
        fonte = esc.fontes[chave]
        caracteres = esc.caracteres_usados.get(chave, set())
        b64 = fonte.base64_subconjunto(caracteres)
        defs_fontes.append(
            f'@font-face {{ font-family: "{fonte.family_name}"; '
            f'src: url(data:font/ttf;base64,{b64}) format("truetype"); }}'
        )

    partes = []
    largura_str = f"{PAGINA_L:g}"
    altura_str = f"{altura_total:g}"
    partes.append(
        f'<svg xmlns="http://www.w3.org/2000/svg" '
        f'width="{largura_str}mm" height="{altura_str}mm" '
        f'viewBox="0 0 {largura_str} {altura_str}">'
    )
    partes.append(f'<style>{"".join(defs_fontes)}</style>')
    # sem fundo: nenhuma cor de pagina, para poder imprimir sobre papel de
    # qualquer cor. A moldura abaixo e so guia de trabalho (mancha nominal
    # de 182.03mm) -- grupo proprio "guia", removivel sem afetar miolo/
    # revisao; NAO usar sombreado alem dela, pois em troncos que excedem a
    # mancha o miolo de verdade continua alem desse limite (ver item 3).
    partes.append('<g id="guia">')
    partes.append(
        f'<rect x="0.15" y="0.15" width="{PAGINA_L-0.3:.3f}" height="{PAGINA_A-0.3:.3f}" '
        f'fill="none" stroke="#000000" stroke-width="0.1"/>'
    )
    partes.append('</g>')
    partes.append('<g id="miolo">')
    partes.extend(esc.miolo)
    partes.append('</g>')
    partes.append('<g id="revisao">')
    partes.extend(esc.revisao)
    partes.append('</g>')
    partes.append('</svg>')
    return "\n".join(partes)


# --------------------------------------------------------------------------
# Folha-indice
# --------------------------------------------------------------------------

def gerar_folha_indice(infos: list[dict]) -> str:
    L = ["# Indice de arvores geradas (Etapa 2)\n"]
    L.append(
        "> Duas colunas de estouro, propositalmente separadas: **arvore "
        "excede a mancha** e conteudo real (a arvore em si nao coube em "
        "182,03mm; decisao de layout/paginacao, caso a caso). **Arquivo "
        "cresceu so por causa da legenda** e cosmetico (a camada de "
        "revisao, que sera apagada antes do miolo final, empurrou o "
        "arquivo para alem de 182,03mm mesmo com a arvore cabendo "
        "sozinha) -- normalmente NAO exige decisao de paginacao.\n"
    )
    L.append(
        "| Arquivo | Tronco-raiz | Tronco familiar | Casal | Nos | "
        "Marcados | Modo | Arvore excede a mancha | Arquivo cresceu "
        "so pela legenda |"
    )
    L.append("|---|---|---|---|---:|---:|---|---|---|")
    for info in infos:
        L.append(
            f"| `{info['arquivo']}` | {info['tronco_raiz_id']} | "
            f"{info['tronco_familiar_id']} | {info['casal']} | "
            f"{info['n_nos']} | {info['n_nos_marcados']} | "
            f"{info['modo_layout']} | "
            f"{'SIM' if info['excedeu_mancha'] else 'nao'} | "
            f"{'sim' if info['arquivo_maior_que_pagina_nominal'] and not info['excedeu_mancha'] else ('--' if info['excedeu_mancha'] else 'nao')} |"
        )
    L.append("")

    n_excedeu = sum(1 for i in infos if i["excedeu_mancha"])
    n_so_legenda = sum(1 for i in infos
                        if i["arquivo_maior_que_pagina_nominal"] and not i["excedeu_mancha"])
    n_ambos_maior_pagina = sum(1 for i in infos if i["arquivo_maior_que_pagina_nominal"])
    L.append("## Resumo dos estouros\n")
    L.append(f"- Arvore excede a mancha (182,03mm), conteudo real: **{n_excedeu}**")
    L.append(f"- Arquivo cresceu alem de 182,03mm SO pela legenda de revisao "
             f"(arvore cabia sozinha): **{n_so_legenda}**")
    L.append(f"- Total de arquivos com altura final > 182,03mm (soma das duas "
             f"linhas acima, sem sobreposicao -- sao conjuntos disjuntos por "
             f"construcao): **{n_ambos_maior_pagina}**")
    L.append("")

    L.append("## Motivo do modo de layout, por tronco\n")
    for info in infos:
        L.append(f"- **{info['tronco_familiar_id']}** ({info['modo_layout']}): {info['motivo_modo']}")
    L.append("")
    return "\n".join(L)


# --------------------------------------------------------------------------
# main
# --------------------------------------------------------------------------

def main(argv=None):
    argv = list(argv or sys.argv[1:])

    tema = "classico"
    if "--tema" in argv:
        idx = argv.index("--tema")
        if idx + 1 < len(argv):
            tema = argv[idx + 1]
        del argv[idx:idx + 2]

    if not argv:
        print("Uso: python gerador_svg.py <tronco_id> [dir_saida] [--tema escala-cinza]")
        print("     python gerador_svg.py --lote [dir_saida] [--tema escala-cinza]")
        sys.exit(1)

    dados = json.loads((AQUI / "dados.json").read_text(encoding="utf-8"))
    fontes = carregar_fontes()

    dir_padrao = AQUI / "svg-cinza" if tema == "escala-cinza" else AQUI

    if argv[0] == "--lote":
        out_dir = Path(argv[1]) if len(argv) >= 2 else dir_padrao
        out_dir.mkdir(parents=True, exist_ok=True)
        infos = []
        for tr in dados["troncos_raiz"]:
            for tf in tr["troncos_familiares"]:
                svg, info = gerar_svg_tronco(dados, fontes, tf["id"], tema=tema)
                (out_dir / info["arquivo"]).write_text(svg, encoding="utf-8")
                infos.append(info)
                print(f"OK: {info['arquivo']} (modo {info['modo_layout']}, "
                      f"{info['n_nos']} nos, {info['n_nos_marcados']} marcados"
                      f"{', EXCEDEU' if info['excedeu_mancha'] else ''})")
        (out_dir / "indice.md").write_text(gerar_folha_indice(infos), encoding="utf-8")
        print(f"\n{len(infos)} SVGs gerados em {out_dir}. Indice: indice.md")
    else:
        tf_id = argv[0]
        out_dir = Path(argv[1]) if len(argv) >= 2 else dir_padrao
        out_dir.mkdir(parents=True, exist_ok=True)
        svg, info = gerar_svg_tronco(dados, fontes, tf_id, tema=tema)
        caminho = out_dir / info["arquivo"]
        caminho.write_text(svg, encoding="utf-8")
        print(f"OK: {caminho}")
        print(json.dumps(info, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

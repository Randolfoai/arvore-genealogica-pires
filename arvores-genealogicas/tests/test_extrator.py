# -*- coding: utf-8 -*-
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from extrator import (
    detectar_separador,
    detectar_cabecalho,
    parse_pessoa,
    parse_resumo,
    parece_pessoa,
    eh_marcador_foto,
    eh_nota_nao_pessoa,
    eh_legenda_apos_foto,
    eh_legenda_sem_conector,
    multiplas_pessoas_no_nome,
)


# ---------------- separadores ----------------

def test_separadores():
    assert detectar_separador("Tronco") == "tronco"
    assert detectar_separador("Tronco-raiz") == "tronco-raiz"
    assert detectar_separador(" Patriarca ") == "patriarca"
    assert detectar_separador("Filhos de X") is None
    assert detectar_separador("Troncos familiares") is None


# ---------------- cabecalhos ----------------

def test_cabecalho_basico_com_dois_pontos():
    c = detectar_cabecalho("Filhos de Conceição e Fausta e seus respetivos cônjuges:")
    assert c is not None
    assert c["nivel"] == "filhos"
    assert "Conceição" in c["referencia"] and "Fausta" in c["referencia"]


def test_cabecalho_sem_dois_pontos():
    assert detectar_cabecalho("Filhos de Lívia")["nivel"] == "filhos"
    assert detectar_cabecalho("Bisnetos de Raimundo e Januária")["nivel"] == "bisnetos"


def test_cabecalho_typos():
    assert detectar_cabecalho("Flhos de Uilame e Terezinha:")["nivel"] == "filhos"
    assert detectar_cabecalho("Netos de Beatris e Antônio:")["nivel"] == "netos"


def test_cabecalho_niveis():
    assert detectar_cabecalho("Trinetos de João e Leonilia:")["nivel"] == "trinetos"
    assert detectar_cabecalho("Tataranetos de Fulano:")["nivel"] == "tataranetos"
    assert detectar_cabecalho("Netos de X:")["nivel"] == "netos"


def test_cabecalho_denominados():
    c = detectar_cabecalho("Filhos de Josina e Manoel (denominados Tronco)")
    assert c["nivel"] == "filhos"
    assert c["denominados"] and "Tronco" in c["denominados"]
    c2 = detectar_cabecalho("Filhos de Prudêncio e Fausta: (Denominados Tronco-raiz)")
    assert c2["nivel"] == "filhos"


def test_cabecalho_umbrella_termina_ponto():
    # cabecalho guarda-chuva que termina em ponto e curto -> aceito
    c = detectar_cabecalho("Netos de Conceição e Fausta e os respectivos cônjuges.")
    assert c is not None and c["nivel"] == "netos"


def test_cabecalho_rejeita_narrativa():
    narr = ("Filho de Dom Vasco Francisco Pires e primo do capitão Luiz Pires, "
            "Antônio Pires fazia parte do grupo de missionários liderado por "
            "Manuel da Nóbrega e veio catequizar os povos nativos.")
    assert detectar_cabecalho(narr) is None


# ---------------- pessoas ----------------

def test_pessoa_simples():
    p = parse_pessoa("Conceição Pires de Castro,")
    assert p["nome"] == "Conceição Pires de Castro"
    assert p["apelido"] is None and p["conjuge"] is None
    assert p["confianca"] == "alta"


def test_pessoa_apelido():
    p = parse_pessoa("Galdina Pires de Castro (Santa),")
    assert p["nome"] == "Galdina Pires de Castro"
    assert p["apelido"] == "Santa"


def test_pessoa_conjuge_barra():
    p = parse_pessoa("Mardônio Pires de Castro/Divina Elza Dutra (1ºcasamento),")
    assert p["nome"] == "Mardônio Pires de Castro"
    assert p["conjuge"] == "Divina Elza Dutra"
    assert "casamento" in (p["observacao"] or "").lower()


def test_pessoa_barra_solteira_vira_observacao():
    p = parse_pessoa("Leoiza Pires de Castro/Solteira,")
    assert p["nome"] == "Leoiza Pires de Castro"
    assert p["conjuge"] is None
    assert "solteira" in (p["observacao"] or "").lower()


def test_pessoa_barra_faleceu():
    p = parse_pessoa("Mariza Pires de Castro/ faleceu ainda criança,")
    assert p["nome"] == "Mariza Pires de Castro"
    assert p["conjuge"] is None
    assert "faleceu" in (p["observacao"] or "").lower()


def test_pessoa_observacao_parentese_longo():
    p = parse_pessoa("José Pires de Castro (era gêmeo com Maria e faleceu ainda bebê)")
    assert p["nome"] == "José Pires de Castro"
    assert p["apelido"] is None
    assert p["observacao"] is not None
    assert p["confianca"] == "media"


def test_pessoa_terminador_e():
    p = parse_pessoa("Júlio Cesar P. Fernandes e")
    assert p["nome"] == "Júlio Cesar P. Fernandes"


def test_pessoa_conjuge_com_2o_casamento():
    p = parse_pessoa(
        "Clecyws Antônio de Castro Alves/Darque Ane Ribeiro de Castro Alves (2º casamento)")
    assert p["nome"] == "Clecyws Antônio de Castro Alves"
    assert p["conjuge"] == "Darque Ane Ribeiro de Castro Alves"
    assert "casamento" in (p["observacao"] or "").lower()


# ---------------- resumo ----------------

def test_resumo_josina():
    t = ("Consta neste trabalho, devidamente catalogados por gerações, um total de "
         "355 membros descendentes desse Tronco-raiz, Josina de Sousa Pires e Manoel "
         "Gomes de Castro, sendo sete filhos (denominados Tronco), 41 netos, 118 "
         "bisnetos, 175 trinetos e 14 tataranetos, identificados até junho de 2026.")
    r = parse_resumo(t)
    assert r["descendentes_total"] == 355
    assert r["filhos"] == 7
    assert r["netos"] == 41
    assert r["bisnetos"] == 118
    assert r["trinetos"] == 175
    assert r["tataranetos"] == 14
    assert "Josina" in r["casal"] and "Manoel" in r["casal"]


def test_resumo_leonidas_palavras():
    t = ("Constam neste trabalho, devidamente catalogados por gerações, 238 membros "
         "descendentes desse Tronco-raiz, Leônidas de Sousa Pires e Luiza Gomes de "
         "Castro, sendo oito filhos, denominados Troncos; 40 netos; 80 bisnetos; 102 "
         "trinetos; e oito tataranetos, identificados até junho de 2026.")
    r = parse_resumo(t)
    assert r["descendentes_total"] == 238
    assert r["filhos"] == 8
    assert r["netos"] == 40
    assert r["tataranetos"] == 8


def test_resumo_none_para_narrativa():
    assert parse_resumo("Conceição cresceu na labuta com gado.") is None


# ---------------- misc ----------------

def test_parece_pessoa():
    assert parece_pessoa("Conceição Pires de Castro,") is True
    assert parece_pessoa("Legenda 01: foto na fazenda") is False
    assert parece_pessoa("x " * 30) is False


def test_marcador_foto():
    assert eh_marcador_foto("MARCADOR_FOTO", "[FOTO 87 — individual]") is True
    assert eh_marcador_foto("Body Text", "Conceição Pires,") is False


# ---------------- guards novos ----------------

def test_nota_nao_pessoa():
    assert eh_nota_nao_pessoa("Orelha do livro") is True
    assert eh_nota_nao_pessoa("Não tiveram filhos") is True
    assert eh_nota_nao_pessoa("Conceição Pires de Castro,") is False


def test_legenda_apos_foto_nomes_soltos():
    # caption com primeiros-nomes soltos (sem 'de/da'), estilo de titulo
    assert eh_legenda_apos_foto(
        "Título 11", "Conceição Pedro Galdina Sebastião Donatila Maria") is True
    assert eh_legenda_apos_foto("Título 11", "América João Pires Anaídes") is True


def test_legenda_apos_foto_descritiva():
    assert eh_legenda_apos_foto(
        "Body Text",
        "Fausta e Conceição Pires, com os filhos, netos, em julho de 1995") is True


def test_legenda_nao_pega_pessoa_real():
    # nome real tem 'de' -> nao e caption de nomes soltos
    assert eh_legenda_apos_foto("Body Text", "Matias de Castro Lessa,") is False


# ---------------- legenda_sem_conector (ponto cego de eh_legenda_apos_foto) ----------------

def test_legenda_sem_conector_pega_casos_reais_da_auditoria():
    # casos reais achados na auditoria de 2026-08-15 (estilo "Body Text",
    # sem conector 'de/da/do', sem virgula/barra -- eh_legenda_apos_foto
    # nao pega estes por exigir estilo != "Body Text" no 1o ramo e nenhuma
    # palavra-gatilho de data/local aparece aqui)
    assert eh_legenda_sem_conector(
        "Body Text", "Anibal José Inês Etelvina Fausta Tomázia Dolores Ricardo") is True
    assert eh_legenda_sem_conector("Body Text", "Raimundo José Estevão") is True
    assert eh_legenda_sem_conector(
        "Body Text",
        "Aniceto Merita Anália José Juliana Manoel Niltácio Jaime Maria "
        "Luziêta Marisa Benito Valmir") is True


def test_legenda_sem_conector_ja_coberto_por_eh_legenda_apos_foto_nao_precisa():
    # eh_legenda_sem_conector so e chamada quando eh_legenda_apos_foto ja
    # retornou False; nao precisa reconhecer o que a outra ja pega.
    assert eh_legenda_apos_foto(
        "Título 11", "Conceição Pedro Galdina Sebastião Donatila Maria") is True


def test_legenda_sem_conector_nao_pega_pessoa_real():
    # tem conector 'de' -> nome real, nao e legenda solta
    assert eh_legenda_sem_conector("Body Text", "Anibal Pires de Castro") is False
    # tem barra -> pessoa com conjuge, nao e legenda
    assert eh_legenda_sem_conector("Body Text", "Nome Um/Nome Dois") is False
    # tem virgula -> caso de multiplas_pessoas_no_nome, nao este detector
    assert eh_legenda_sem_conector("Body Text", "Nome Um, Nome Dois") is False
    # menos de 3 tokens -> curto demais pra ser lista de nomes
    assert eh_legenda_sem_conector("Body Text", "Ana Maria") is False


def test_multiplas_pessoas():
    assert multiplas_pessoas_no_nome("Inês Pires de Castro , Etelvina Pires de Castro") is True
    assert multiplas_pessoas_no_nome("Conceição Pires de Castro") is False

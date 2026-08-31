#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Gera a apresentacao do portfolio PLI Hub em HTML e PPTX a partir de dados.json.

Uso:  python gerar_apresentacao.py

Saidas (na mesma pasta):
  - pli-hub-apresentacao.html  (deck navegavel no navegador)
  - pli-hub-apresentacao.pptx  (PowerPoint 16:9)

O conteudo vive APENAS em dados.json. Para mudar um texto, edite o JSON e
rode este script de novo; HTML e PPTX nunca divergem.
"""

import json
import pathlib

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

AQUI = pathlib.Path(__file__).parent
DADOS = json.loads((AQUI / "dados.json").read_text(encoding="utf-8"))

# ---------------------------------------------------------------- paleta

GROUND = RGBColor(0x06, 0x1A, 0x26)
SURFACE = RGBColor(0x0C, 0x29, 0x39)
RULE = RGBColor(0x1A, 0x46, 0x5F)
TEXT = RGBColor(0xE8, 0xF1, 0xF6)
MUTED = RGBColor(0x8F, 0xB0, 0xC2)
DIM = RGBColor(0x5D, 0x82, 0x96)
BLUE = RGBColor(0x3A, 0xA6, 0xD9)
GREEN = RGBColor(0x3E, 0xC2, 0x6E)
AMBER = RGBColor(0xE0, 0xA3, 0x3A)
DARKINK = RGBColor(0x0A, 0x2A, 0x16)

DISPLAY = "Segoe UI Semibold"
BODY = "Segoe UI"
MONO = "Consolas"

W, H = Inches(13.333), Inches(7.5)
MARGEM = Inches(0.72)
LARG = W - 2 * MARGEM


def hex_rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def cor_cat(cat):
    return hex_rgb(DADOS["categorias"][cat]["cor"])


# ---------------------------------------------------------------- helpers


def novo_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = GROUND
    return slide


def faixa(slide, cor):
    """Faixa de categoria no topo do slide."""
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, Pt(4))
    s.fill.solid()
    s.fill.fore_color.rgb = cor
    s.line.fill.background()
    s.shadow.inherit = False
    return s


def caixa(slide, x, y, w, h, preenche=None, borda=None, larg_borda=Pt(1)):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    if preenche is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = preenche
    if borda is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = borda
        s.line.width = larg_borda
    s.shadow.inherit = False
    return s


def texto(
    slide,
    x,
    y,
    w,
    h,
    conteudo,
    tam=14,
    cor=TEXT,
    fonte=BODY,
    negrito=False,
    espaco=1.0,
    alinha=PP_ALIGN.LEFT,
    espacamento_letras=None,
    anchor=MSO_ANCHOR.TOP,
):
    """Cria uma caixa de texto. `conteudo` pode ser str ou lista de paragrafos."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor

    linhas = conteudo if isinstance(conteudo, list) else [conteudo]
    for i, linha in enumerate(linhas):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = alinha
        p.line_spacing = espaco
        if i:
            p.space_before = Pt(6)
        r = p.add_run()
        r.text = linha
        r.font.size = Pt(tam)
        r.font.color.rgb = cor
        r.font.name = fonte
        r.font.bold = negrito
    return tb


def rodape(slide, esq, dir_, n, total):
    y = H - Inches(0.62)
    linha = caixa(slide, MARGEM, y - Inches(0.16), LARG, Pt(0.75), preenche=RULE)
    linha.line.fill.background()
    texto(slide, MARGEM, y, Inches(7), Inches(0.3), esq, tam=9, cor=DIM, fonte=MONO)
    texto(
        slide,
        W - MARGEM - Inches(5),
        y,
        Inches(5),
        Inches(0.3),
        f"{dir_}     {n:02d}/{total}",
        tam=9,
        cor=DIM,
        fonte=MONO,
        alinha=PP_ALIGN.RIGHT,
    )


def eyebrow(slide, x, y, txt, cor=BLUE):
    return texto(slide, x, y, Inches(6), Inches(0.26), txt.upper(), tam=9.5, cor=cor, fonte=MONO)


def pilula(slide, x, y, txt, fundo, cor_txt):
    largura = Inches(0.16) * len(txt) / 2 + Inches(0.4)
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, largura, Inches(0.26))
    s.fill.solid()
    s.fill.fore_color.rgb = fundo
    s.line.fill.background()
    s.shadow.inherit = False
    s.adjustments[0] = 0.15
    tf = s.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = txt.upper()
    r.font.size = Pt(9)
    r.font.name = MONO
    r.font.color.rgb = cor_txt
    return s


# ---------------------------------------------------------------- slides


def slide_capa(prs, n, total):
    s = novo_slide(prs)
    faixa(s, BLUE)
    caixa(s, W - Inches(5.4), 0, Inches(5.4), H, preenche=RGBColor(0x08, 0x22, 0x30))
    caixa(s, W - Inches(5.4), 0, Pt(1), H, preenche=RULE)

    eyebrow(s, MARGEM, Inches(0.95), f"{DADOS['org']} · {DADOS['data']}")
    texto(s, MARGEM, Inches(1.5), Inches(7.4), Inches(2.2), "PLI Hub", tam=88, negrito=True, fonte=DISPLAY, espaco=0.85)
    texto(s, MARGEM, Inches(3.75), Inches(6.2), Inches(1.4), DADOS["subtitulo"], tam=21, cor=MUTED, espaco=1.2)

    y = Inches(5.05)
    meta = [
        ("Aplicações", f"{len(DADOS['apps'])} sistemas em 3 categorias"),
        ("Central de acesso", DADOS["hub_url"]),
        ("Responsável técnico", DADOS["autor"]),
    ]
    for rotulo, valor in meta:
        texto(s, MARGEM, y, Inches(6.5), Inches(0.24), rotulo.upper(), tam=9, cor=DIM, fonte=MONO)
        texto(s, MARGEM, y + Inches(0.23), Inches(6.5), Inches(0.3), valor, tam=13, cor=TEXT)
        y += Inches(0.62)

    # coluna direita: indice das categorias
    x = W - Inches(5.4) + Inches(0.7)
    cw = Inches(4.0)
    yy = Inches(1.35)
    texto(s, x, yy, cw, Inches(0.26), "NESTE DOCUMENTO", tam=9.5, cor=DIM, fonte=MONO)
    yy += Inches(0.45)
    for cid, cat in DADOS["categorias"].items():
        lista = [a for a in DADOS["apps"] if a["categoria"] == cid]
        caixa(s, x, yy, cw, Pt(2.5), preenche=hex_rgb(cat["cor"]))
        texto(s, x, yy + Inches(0.12), cw, Inches(0.28), cat["nome"], tam=13.5, negrito=True, fonte=DISPLAY, cor=hex_rgb(cat["cor"]))
        yy += Inches(0.42)
        for a in lista:
            texto(s, x, yy, cw, Inches(0.24), a["nome"], tam=11, cor=MUTED)
            yy += Inches(0.26)
        yy += Inches(0.2)
    return s


def slide_panorama(prs, n, total):
    s = novo_slide(prs)
    faixa(s, BLUE)
    eyebrow(s, MARGEM, Inches(0.55), "Panorama")
    texto(s, MARGEM, Inches(0.95), Inches(9), Inches(0.7), "O portfólio em um relance", tam=36, negrito=True, fonte=DISPLAY)

    apps = DADOS["apps"]
    no_ar = sum(1 for a in apps if a["status"] == "no ar")
    numeros = [
        (str(len(apps)), "aplicações", TEXT),
        ("3", "categorias", TEXT),
        (str(no_ar), "no ar hoje", GREEN),
        (str(len(apps) - no_ar), "suspensas", AMBER),
    ]
    x = MARGEM
    for valor, rotulo, cor in numeros:
        texto(s, x, Inches(1.85), Inches(2.2), Inches(0.7), valor, tam=42, negrito=True, fonte=DISPLAY, cor=cor)
        texto(s, x, Inches(2.5), Inches(2.2), Inches(0.3), rotulo.upper(), tam=9, cor=DIM, fonte=MONO)
        x += Inches(2.5)
    caixa(s, MARGEM, Inches(2.95), LARG, Pt(0.75), preenche=RULE)

    col_w = (LARG - Inches(0.7)) / 3
    x = MARGEM
    for cid, cat in DADOS["categorias"].items():
        lista = [a for a in apps if a["categoria"] == cid]
        cor = hex_rgb(cat["cor"])
        caixa(s, x, Inches(3.3), col_w, Pt(3), preenche=cor)
        texto(s, x, Inches(3.5), col_w, Inches(0.6), cat["nome"], tam=16, negrito=True, fonte=DISPLAY, cor=cor, espaco=1.05)
        texto(s, x, Inches(4.1), col_w, Inches(0.25), f"{len(lista)} aplicações", tam=9, cor=DIM, fonte=MONO)
        y = Inches(4.5)
        for k, a in enumerate(lista, 1):
            texto(s, x, y, Inches(0.3), Inches(0.26), f"{k:02d}", tam=9, cor=DIM, fonte=MONO)
            texto(s, x + Inches(0.42), y - Inches(0.02), col_w - Inches(0.8), Inches(0.26), a["nome"], tam=12.5, cor=TEXT)
            ponto = caixa(s, x + col_w - Inches(0.22), y + Inches(0.05), Inches(0.1), Inches(0.1),
                          preenche=GREEN if a["status"] == "no ar" else AMBER)
            ponto.line.fill.background()
            caixa(s, x, y + Inches(0.3), col_w - Inches(0.1), Pt(0.75), preenche=RULE)
            y += Inches(0.44)
        x += col_w + Inches(0.35)

    rodape(s, DADOS["hub_url"], "Panorama", n, total)
    return s


def slide_app(prs, app, i, total_apps, n, total):
    s = novo_slide(prs)
    cor = cor_cat(app["categoria"])
    cat = DADOS["categorias"][app["categoria"]]["nome"]
    faixa(s, cor)

    eyebrow(s, MARGEM, Inches(0.55), f"Aplicação {i + 1:02d} / {total_apps}", cor=cor)
    texto(s, MARGEM + Inches(2.3), Inches(0.55), Inches(5), Inches(0.26), cat.upper(), tam=9.5, cor=DIM, fonte=MONO)
    if app["status"] == "no ar":
        pilula(s, W - MARGEM - Inches(1.05), Inches(0.5), "no ar", GREEN, DARKINK)
    else:
        pilula(s, W - MARGEM - Inches(1.3), Inches(0.5), "suspensa", AMBER, RGBColor(0x2A, 0x14, 0x08))

    # coluna esquerda -----------------------------------------------------
    esq_w = Inches(6.1)
    texto(s, MARGEM, Inches(1.15), esq_w, Inches(0.9), app["nome"], tam=40, negrito=True, fonte=DISPLAY, espaco=0.95)
    texto(s, MARGEM, Inches(2.05), esq_w, Inches(0.6), app["extenso"], tam=13, cor=MUTED, espaco=1.15)
    texto(s, MARGEM, Inches(2.95), esq_w, Inches(2.0), app["resumo"], tam=15, cor=TEXT, espaco=1.35)

    y_papel = Inches(5.35)
    caixa(s, MARGEM, y_papel, Pt(3), Inches(0.42), preenche=cor)
    texto(s, MARGEM + Inches(0.16), y_papel + Inches(0.05), esq_w, Inches(0.4), app["papel"], tam=15, negrito=True, fonte=DISPLAY, cor=cor)

    # painel direito ------------------------------------------------------
    px = MARGEM + esq_w + Inches(0.5)
    pw = W - MARGEM - px
    ph = Inches(5.02)
    caixa(s, px, Inches(1.15), pw, ph, preenche=SURFACE, borda=RULE)

    tx = px + Inches(0.42)
    tw = pw - Inches(0.84)
    texto(s, tx, Inches(1.48), tw, Inches(0.25), "O QUE ELA ENTREGA", tam=9, cor=DIM, fonte=MONO)

    # ~52 caracteres cabem por linha na largura util do painel, a 11,5 pt
    y = Inches(1.92)
    for d in app["destaques"]:
        linhas = max(1, -(-len(d) // 52))
        alt = Inches(0.235) * linhas
        tracinho = caixa(s, tx, y + Inches(0.1), Inches(0.16), Pt(1.2), preenche=cor)
        tracinho.line.fill.background()
        texto(s, tx + Inches(0.32), y, tw - Inches(0.32), alt, d, tam=11.5, cor=RGBColor(0xD3, 0xE3, 0xEC), espaco=1.2)
        y += alt + Inches(0.13)

    y_chips = max(y + Inches(0.28), Inches(1.15) + ph - Inches(0.86))
    caixa(s, tx, y_chips - Inches(0.2), tw, Pt(0.75), preenche=RULE)
    cx, cy = tx, y_chips
    for chip in app["stack"]:
        largura = Inches(0.075) * len(chip) + Inches(0.3)
        if cx + largura > tx + tw:
            cx = tx
            cy += Inches(0.36)
        c = caixa(s, cx, cy, largura, Inches(0.28), preenche=None, borda=RULE)
        tf = c.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = chip
        r.font.size = Pt(8.5)
        r.font.name = MONO
        r.font.color.rgb = MUTED
        cx += largura + Inches(0.1)

    rodape(s, app["url"], cat, n, total)
    return s


def slide_modulos(prs, n, total):
    s = novo_slide(prs)
    faixa(s, BLUE)
    eyebrow(s, MARGEM, Inches(0.55), "SIGMA-PLI · por dentro")
    texto(s, MARGEM, Inches(0.95), Inches(10), Inches(0.7), "A plataforma central, módulo a módulo", tam=34, negrito=True, fonte=DISPLAY)
    texto(
        s, MARGEM, Inches(1.72), Inches(9.5), Inches(0.5),
        "Cada módulo é uma área funcional própria. Juntos, formam a base sobre a qual as outras aplicações do PLI se apoiam.",
        tam=13, cor=MUTED, espaco=1.2,
    )

    mods = DADOS["modulos_sigma"]
    cols, gap = 4, Inches(0.28)
    cw = (LARG - gap * (cols - 1)) / cols
    ch = Inches(1.12)
    x0, y0 = MARGEM, Inches(2.5)
    for idx, (cod, nome, desc) in enumerate(mods):
        c, r = idx % cols, idx // cols
        x = x0 + c * (cw + gap)
        y = y0 + r * (ch + Inches(0.2))
        caixa(s, x, y, cw, ch, preenche=SURFACE, borda=RULE)
        caixa(s, x, y, Pt(3), ch, preenche=BLUE)
        texto(s, x + Inches(0.22), y + Inches(0.13), cw - Inches(0.4), Inches(0.22), cod, tam=9, cor=BLUE, fonte=MONO)
        texto(s, x + Inches(0.22), y + Inches(0.36), cw - Inches(0.4), Inches(0.28), nome, tam=13.5, negrito=True, fonte=DISPLAY)
        texto(s, x + Inches(0.22), y + Inches(0.64), cw - Inches(0.4), Inches(0.44), desc, tam=10, cor=MUTED, espaco=1.15)

    rodape(s, DADOS["apps"][0]["url"], "Sistemas Maiores", n, total)
    return s


def slide_integracoes(prs, n, total):
    s = novo_slide(prs)
    faixa(s, GREEN)
    eyebrow(s, MARGEM, Inches(0.55), "Arquitetura", cor=GREEN)
    texto(s, MARGEM, Inches(0.95), Inches(10), Inches(0.6), "Como as peças conversam", tam=34, negrito=True, fonte=DISPLAY)
    texto(
        s, MARGEM, Inches(1.65), Inches(9.8), Inches(0.5),
        "O portfólio não é um conjunto de sistemas isolados: quatro integrações já em produção ligam sensor, cadastro e decisão.",
        tam=13, cor=MUTED, espaco=1.2,
    )

    bw, bh = Inches(3.0), Inches(1.0)

    def bloco(x, y, cor, nome, sub):
        caixa(s, x, y, bw, bh, preenche=SURFACE, borda=cor)
        caixa(s, x, y, Pt(3.5), bh, preenche=cor)
        texto(s, x + Inches(0.24), y + Inches(0.2), bw - Inches(0.45), Inches(0.3), nome, tam=15, negrito=True, fonte=DISPLAY)
        texto(s, x + Inches(0.24), y + Inches(0.55), bw - Inches(0.45), Inches(0.35), sub, tam=10.5, cor=MUTED, espaco=1.1)

    def seta(x1, y1, x2, y2, rotulo):
        con = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Emu(int(x1)), Emu(int(y1)), Emu(int(x2)), Emu(int(y2)))
        con.line.color.rgb = DIM
        con.line.width = Pt(1.25)
        mx, my = (x1 + x2) / 2, (y1 + y2) / 2
        texto(s, Emu(int(mx - Inches(0.62))), Emu(int(my - Inches(0.34))), Inches(1.24), Inches(0.24),
              rotulo.upper(), tam=7.5, cor=DIM, fonte=MONO, alinha=PP_ALIGN.CENTER)

    col1, col2, col3 = MARGEM, MARGEM + Inches(4.3), MARGEM + Inches(8.6)

    texto(s, col1, Inches(2.35), Inches(3.0), Inches(0.22), "SENSORES E COLETA", tam=8.5, cor=DIM, fonte=MONO)
    bloco(col1, Inches(2.6), BLUE, "PLI-HazardTrack", "chuva MERGE/INPE → risco por trecho")
    bloco(col1, Inches(3.85), BLUE, "PLI Reporta", "relato do cidadão → incidente verificado")

    texto(s, col2, Inches(2.35), Inches(3.0), Inches(0.22), "OPERAÇÃO", tam=8.5, cor=DIM, fonte=MONO)
    bloco(col2, Inches(3.2), BLUE, "PLI Smart Router", "rota que desvia do risco")

    seta(col1 + bw, Inches(3.1), col2, Inches(3.5), "camadas de risco")
    seta(col1 + bw, Inches(4.35), col2, Inches(3.95), "geojson")

    texto(s, col3, Inches(2.35), Inches(3.0), Inches(0.22), "LEITURA DO TERRITÓRIO", tam=8.5, cor=DIM, fonte=MONO)
    bloco(col3, Inches(2.6), BLUE, "FAD-Stats 2.0", "IBGE · SECEX · MTE · ANTT")
    bloco(col3, Inches(3.85), AMBER, "Análises Exploratórias", "sinistralidade e densidade da malha")

    texto(s, col1, Inches(5.4), Inches(3.0), Inches(0.22), "PLATAFORMA CENTRAL", tam=8.5, cor=DIM, fonte=MONO)
    bloco(col1, Inches(5.65), GREEN, "SIGMA-PLI", "metadados, identidade e grafo")

    texto(s, col2, Inches(5.4), Inches(3.0), Inches(0.22), "DECISÃO", tam=8.5, cor=DIM, fonte=MONO)
    bloco(col2, Inches(5.65), BLUE, "SICARD", "carteira hierarquizada por AHP")

    texto(s, col3, Inches(5.4), Inches(3.0), Inches(0.22), "MÉTODO", tam=8.5, cor=DIM, fonte=MONO)
    bloco(col3, Inches(5.65), GREEN, "AHP Tool Calculator", "pesos e consistência")

    seta(col1 + bw, Inches(6.15), col2, Inches(6.15), "login e cadastros")
    seta(col3, Inches(6.15), col2 + bw, Inches(6.15), "método AHP")

    rodape(s, "4 integrações em produção", "Arquitetura", n, total)
    return s


def slide_situacao(prs, n, total):
    s = novo_slide(prs)
    faixa(s, BLUE)
    apps = DADOS["apps"]
    no_ar = sum(1 for a in apps if a["status"] == "no ar")
    fora = len(apps) - no_ar

    eyebrow(s, MARGEM, Inches(0.55), "Situação e próximos passos")
    texto(s, MARGEM, Inches(0.95), Inches(10), Inches(0.7), "Onde o portfólio está hoje", tam=36, negrito=True, fonte=DISPLAY)

    x = MARGEM
    for valor, rotulo, cor in [(str(no_ar), "no servidor próprio", GREEN), (str(fora), "suspensas no Render", AMBER), ("4", "integrações ativas", TEXT)]:
        texto(s, x, Inches(1.9), Inches(3), Inches(0.7), valor, tam=42, negrito=True, fonte=DISPLAY, cor=cor)
        texto(s, x, Inches(2.55), Inches(3), Inches(0.3), rotulo.upper(), tam=9, cor=DIM, fonte=MONO)
        x += Inches(3.4)
    caixa(s, MARGEM, Inches(3.0), LARG, Pt(0.75), preenche=RULE)

    blocos = [
        ("Infraestrutura própria",
         [f"As {no_ar} aplicações dos Sistemas Maiores rodam em containers Docker atrás de um Nginx na instância EC2 do PLI, cada uma com seu banco PostgreSQL/PostGIS.",
          "É a base madura do portfólio: disponível, versionada e integrada."]),
        ("Hospedagem externa",
         [f"As {fora} aplicações de Análises Exploratórias e Ferramentas estão no Render e hoje respondem Service Suspended.",
          "Retomá-las exige decidir entre reativar o plano ou migrá-las para a mesma infraestrutura dos sistemas maiores."]),
        ("Próximo passo",
         ["Concluir a análise aplicação por aplicação e classificar cada uma como produto, módulo, serviço compartilhado, biblioteca, ferramenta interna ou protótipo.",
          "Só então definir a arquitetura da empresa e a estratégia comercial."]),
    ]
    cw = (LARG - Inches(0.7)) / 3
    x = MARGEM
    for titulo, paragrafos in blocos:
        texto(s, x, Inches(3.45), cw, Inches(0.4), titulo, tam=17, negrito=True, fonte=DISPLAY)
        texto(s, x, Inches(4.0), cw, Inches(2.2), paragrafos, tam=12, cor=MUTED, espaco=1.35)
        x += cw + Inches(0.35)

    rodape(s, DADOS["hub_url"], f"{DADOS['org']} · {DADOS['data']}", n, total)
    return s


# ---------------------------------------------------------------- saidas


def gerar_pptx(destino):
    prs = Presentation()
    prs.slide_width, prs.slide_height = W, H

    apps = DADOS["apps"]
    total = 4 + (len(apps) - 1) + 2  # capa, panorama, sigma, modulos, demais, arquitetura, situacao

    n = 1
    slide_capa(prs, n, total); n += 1
    slide_panorama(prs, n, total); n += 1
    slide_app(prs, apps[0], 0, len(apps), n, total); n += 1
    slide_modulos(prs, n, total); n += 1
    for i, app in enumerate(apps[1:], start=1):
        slide_app(prs, app, i, len(apps), n, total); n += 1
    slide_integracoes(prs, n, total); n += 1
    slide_situacao(prs, n, total); n += 1

    prs.save(destino)
    return len(prs.slides._sldIdLst)


def gerar_html(destino):
    template = (AQUI / "template.html").read_text(encoding="utf-8")
    dados = json.dumps(DADOS, ensure_ascii=False, indent=2)
    if "/*DADOS*/" not in template:
        raise SystemExit("template.html nao tem o marcador /*DADOS*/")
    destino.write_text(template.replace("/*DADOS*/", dados), encoding="utf-8")


if __name__ == "__main__":
    html = AQUI / "pli-hub-apresentacao.html"
    pptx = AQUI / "pli-hub-apresentacao.pptx"
    gerar_html(html)
    n = gerar_pptx(pptx)
    print(f"HTML gerado : {html.name}")
    print(f"PPTX gerado : {pptx.name} ({n} slides)")

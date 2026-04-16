#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate an editable PPTX from the Anteriq group pitch content."""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

INK = RGBColor(0x1A, 0x1A, 0x2E)
INK_SOFT = RGBColor(0x3D, 0x3D, 0x54)
INK_MUTED = RGBColor(0x7A, 0x7A, 0x8C)
PAPER = RGBColor(0xF8, 0xF6, 0xF1)
PAPER_WARM = RGBColor(0xF0, 0xEC, 0xE3)
GOLD = RGBColor(0xC9, 0xA8, 0x4C)
BRONZE = RGBColor(0x7A, 0x68, 0x40)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
CARD_DARK = RGBColor(0x25, 0x25, 0x3E)
TEXT_LIGHT = RGBColor(0xA0, 0x9E, 0x96)
HIGHLIGHT_BG = RGBColor(0xE8, 0xE0, 0xCC)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
MARGIN = Inches(0.7)
CONTENT_W = SLIDE_W - 2 * MARGIN

DOT = "\u00b7"
ARROW = "\u2192"
DASH = "\u2014"
LAQUO = "\u00ab"
RAQUO = "\u00bb"

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H
blank = prs.slide_layouts[6]


def add_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def tb(slide, left, top, width, height):
    return slide.shapes.add_textbox(left, top, width, height)


def set_text(tf, text, size=14, bold=False, italic=False, color=INK_SOFT, align=PP_ALIGN.LEFT):
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return run


def add_para(tf, text, size=14, bold=False, italic=False, color=INK_SOFT, align=PP_ALIGN.LEFT, space_before=Pt(4)):
    p = tf.add_paragraph()
    p.alignment = align
    p.space_before = space_before
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return p


def header_block(slide, act, label, title, subtitle, dark=False):
    bg_color = INK if dark else PAPER
    title_color = PAPER if dark else INK
    sub_color = TEXT_LIGHT if dark else INK_SOFT
    label_color = GOLD if dark else BRONZE
    add_bg(slide, bg_color)
    y = Inches(0.4)
    if act:
        t = tb(slide, MARGIN, y, CONTENT_W, Inches(0.3))
        set_text(t.text_frame, act, size=9, bold=True, color=label_color)
        y += Inches(0.28)
    if label:
        t = tb(slide, MARGIN, y, CONTENT_W, Inches(0.3))
        set_text(t.text_frame, label.upper(), size=11, bold=True, color=label_color)
        y += Inches(0.35)
    t = tb(slide, MARGIN, y, CONTENT_W, Inches(0.6))
    set_text(t.text_frame, title, size=28, bold=True, color=title_color)
    y += Inches(0.65)
    if subtitle:
        t = tb(slide, MARGIN, y, CONTENT_W, Inches(0.6))
        set_text(t.text_frame, subtitle, size=16, color=sub_color)
        y += Inches(0.65)
    return y


def card_box(slide, left, top, width, height, bg=PAPER_WARM, border_color=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = Pt(2)
    else:
        shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def card_with_text(slide, left, top, width, height, title_text, body_text,
                   extra_line=None, bg=PAPER_WARM, border_color=None,
                   title_color=INK, body_color=INK_SOFT):
    card_box(slide, left, top, width, height, bg=bg, border_color=border_color)
    pad = Inches(0.15)
    t = tb(slide, left + pad, top + pad, width - 2 * pad, height - 2 * pad)
    tf = t.text_frame
    tf.word_wrap = True
    set_text(tf, title_text, size=13, bold=True, color=title_color)
    add_para(tf, body_text, size=11, color=body_color)
    if extra_line:
        add_para(tf, extra_line, size=10, bold=True, color=title_color)


def punchline(slide, y, text, dark=False):
    color = TEXT_LIGHT if dark else INK_SOFT
    t = tb(slide, MARGIN, y, CONTENT_W, Inches(0.6))
    set_text(t.text_frame, text, size=14, italic=True, color=color, align=PP_ALIGN.CENTER)


# ---------- S1 COVER ----------
s = prs.slides.add_slide(blank)
add_bg(s, INK)
t = tb(s, Inches(1.5), Inches(2.0), Inches(10), Inches(0.5))
set_text(t.text_frame, "ANTERIQ", size=20, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
t = tb(s, Inches(1.5), Inches(2.8), Inches(10), Inches(1.0))
set_text(t.text_frame, "La transformation IA\nde bout en bout", size=40, bold=True, color=PAPER, align=PP_ALIGN.CENTER)
t = tb(s, Inches(2.5), Inches(4.2), Inches(8), Inches(0.5))
set_text(t.text_frame, f"Strat{DOT}gie {DOT} Donn{DOT}es {DOT} IA {DOT} Infrastructure souveraine {DOT} Humain".replace(f"{DOT}gie", "\u00e9gie").replace(f"{DOT}es", "\u00e9es"),
         size=16, italic=True, color=GOLD, align=PP_ALIGN.CENTER)
t = tb(s, Inches(2), Inches(5.0), Inches(9), Inches(0.7))
set_text(t.text_frame,
         "Le premier groupe capable d'accompagner une organisation du diagnostic "
         "strat\u00e9gique au d\u00e9ploiement en production "
         f"{DASH} et de laisser le client propri\u00e9taire et autonome.",
         size=14, color=TEXT_LIGHT, align=PP_ALIGN.CENTER)
t = tb(s, Inches(4), Inches(6.2), Inches(5), Inches(0.4))
set_text(t.text_frame, f"Confidentiel {DASH} Avril 2026", size=12, color=INK_MUTED, align=PP_ALIGN.CENTER)

# ---------- S2 LE PROBL\u00c8ME ----------
s = prs.slides.add_slide(blank)
y = header_block(s, "Acte I", "Le probl\u00e8me",
                 "Trois options sur le march\u00e9. Toutes insatisfaisantes.",
                 "Les organisations veulent se transformer. Le march\u00e9 ne leur offre que des fragments.",
                 dark=True)
cw = (CONTENT_W - Inches(0.4)) / 3
cards_data = [
    ("Les ESN et int\u00e9grateurs",
     "Temps-homme, POC sans production. Client d\u00e9pendant. "
     "Donn\u00e9es dans le cloud. Personne ne s'occupe des \u00e9quipes."),
    ("Les \u00e9diteurs SaaS",
     "Lock-in maximal, donn\u00e9es hors de contr\u00f4le, "
     "personnalisation limit\u00e9e. Le client loue ce qu'il ne comprend pas."),
    ("Les cabinets de conseil",
     "Ils livrent des slides et partent. "
     "Pas de d\u00e9ploiement, pas d'infra, pas de suivi."),
]
for i, (ct, cb) in enumerate(cards_data):
    left = MARGIN + i * (cw + Inches(0.2))
    card_with_text(s, left, y, cw, Inches(1.5), ct, cb,
                   bg=CARD_DARK, title_color=PAPER, body_color=TEXT_LIGHT)

sy = y + Inches(1.7)
stats = [
    ("80%+", "des projets IA \u00e9chouent", "RAND Corporation, 2024"),
    ("85%", "des mod\u00e8les IA n'atteignent pas la production", "Gartner"),
    ("8%", "des grandes entreprises ont d\u00e9ploy\u00e9 l'IA \u00e0 l'\u00e9chelle", "McKinsey, 2025"),
]
for i, (num, lab, src) in enumerate(stats):
    left = MARGIN + i * (cw + Inches(0.2))
    t = tb(s, left, sy, cw, Inches(1.2))
    tf = t.text_frame
    tf.word_wrap = True
    set_text(tf, num, size=36, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
    add_para(tf, lab, size=11, color=TEXT_LIGHT, align=PP_ALIGN.CENTER)
    add_para(tf, src, size=8, color=INK_MUTED, align=PP_ALIGN.CENTER)

# ---------- S3 NOTRE R\u00c9PONSE ----------
s = prs.slides.add_slide(blank)
y = header_block(s, "Acte I", "Notre r\u00e9ponse",
                 "Un seul interlocuteur. La transformation compl\u00e8te.",
                 "Du diagnostic strat\u00e9gique au d\u00e9ploiement en production, "
                 "en passant par la structuration des donn\u00e9es et l'accompagnement des \u00e9quipes.",
                 dark=True)

flow = [
    ("D\u00e9couverte", "Mod\u00e8le cible"),
    None,
    ("Impl\u00e9mentation", "Donn\u00e9es & outils"),
    None,
    ("Impact humain", "Formation, reskilling"),
]
fx = MARGIN + Inches(0.5)
for item in flow:
    if item is None:
        t = tb(s, fx, y + Inches(0.15), Inches(0.5), Inches(0.4))
        set_text(t.text_frame, ARROW, size=24, color=INK_MUTED, align=PP_ALIGN.CENTER)
        fx += Inches(0.6)
    else:
        bw = Inches(2.5)
        card_box(s, fx, y, bw, Inches(0.7), bg=CARD_DARK)
        t = tb(s, fx, y + Inches(0.05), bw, Inches(0.6))
        tf = t.text_frame
        tf.word_wrap = True
        set_text(tf, item[0], size=13, bold=True, color=PAPER, align=PP_ALIGN.CENTER)
        add_para(tf, item[1], size=10, color=TEXT_LIGHT, align=PP_ALIGN.CENTER)
        fx += bw + Inches(0.1)

y += Inches(1.0)
convictions = [
    ("Souveraine",
     "Les donn\u00e9es restent chez le client. Pas de d\u00e9pendance cloud. EU AI Act, NIS2."),
    ("\u00c9mancipatrice",
     "Le client est propri\u00e9taire de tout \u2014 donn\u00e9es, mod\u00e8les, apps, code."),
    ("Scientifiquement rigoureuse",
     "Nos propres mod\u00e8les IA, open source, \u00e9valu\u00e9s par les pairs. 10 ans de R&D."),
    ("Humaine",
     "Chaque d\u00e9ploiement inclut un plan humain. Formation, reskilling, coaching."),
]
cw2 = (CONTENT_W - Inches(0.2)) / 2
for i, (ct, cb) in enumerate(convictions):
    col, row = i % 2, i // 2
    left = MARGIN + col * (cw2 + Inches(0.2))
    top = y + row * Inches(1.1)
    card_with_text(s, left, top, cw2, Inches(0.95), ct, cb,
                   bg=CARD_DARK, title_color=GOLD, body_color=TEXT_LIGHT)

# ---------- S4 DISCOVERY ----------
s = prs.slides.add_slide(blank)
y = header_block(s, "Acte II", "Phase 1 \u2014 Discovery",
                 "Construire la tour de contr\u00f4le de la transformation.",
                 "L'engagement commence au niveau du comit\u00e9 ex\u00e9cutif. "
                 "Le plan qu'on construit est directement outillable par la plateforme du groupe.")

cw2 = (CONTENT_W - Inches(0.3)) / 2
card_box(s, MARGIN, y, cw2, Inches(3.5), bg=PAPER_WARM)
t = tb(s, MARGIN + Inches(0.15), y + Inches(0.1), cw2 - Inches(0.3), Inches(3.3))
tf = t.text_frame
tf.word_wrap = True
set_text(tf, "Ce que nous faisons", size=13, bold=True, color=INK)
for b in [
    "Acculturation des dirigeants \u2014 lecture strat\u00e9gique IA",
    "Cartographie des opportunit\u00e9s \u2014 cas d'usage, impact quantifi\u00e9",
    "Feuille de route budg\u00e9t\u00e9e \u2014 plan prioris\u00e9, business cases",
    "Fondations de l'ontologie \u2014 premier mod\u00e8le de donn\u00e9es structur\u00e9",
]:
    add_para(tf, f"\u2022 {b}", size=11, color=INK_SOFT)

rx = MARGIN + cw2 + Inches(0.3)
card_box(s, rx, y, cw2, Inches(3.5), bg=PAPER_WARM, border_color=BRONZE)
t = tb(s, rx + Inches(0.15), y + Inches(0.1), cw2 - Inches(0.3), Inches(3.3))
tf = t.text_frame
tf.word_wrap = True
set_text(tf, "Ce qui rend cette phase diff\u00e9rente", size=13, bold=True, color=INK)
add_para(tf, "Pas de \u00ab passation \u00bb \u00e0 un int\u00e9grateur tiers \u2014 "
         "les m\u00eames personnes qui diagnostiquent accompagnent le d\u00e9ploiement.",
         size=11, color=INK_SOFT)
add_para(tf, "", size=6)
add_para(tf, "Briques activ\u00e9es", size=11, bold=True, color=BRONZE)
add_para(tf, "Conseil strat\u00e9gique Data & IA \u00b7 Ontologie & donn\u00e9es",
         size=11, color=INK_SOFT)
add_para(tf, "", size=6)
add_para(tf, "Revenu", size=11, bold=True, color=BRONZE)
add_para(tf, "One-shot \u00b7 100\u2013300K par programme", size=11, color=INK_SOFT)

# ---------- S5 BUILD ----------
s = prs.slides.add_slide(blank)
y = header_block(s, "Acte II", "Phase 2 \u2014 Build",
                 "Mettre en place le plan de cr\u00e9ation de valeur. Outiller l'organisation.",
                 "Le conseil et le build ne sont pas s\u00e9quentiels \u2014 ils sont int\u00e9gr\u00e9s. "
                 "R\u00e9sultats tangibles en semaines.")

cw4 = (CONTENT_W - Inches(0.6)) / 4
build = [
    ("Infrastructure",
     "Co-construction avec le DSI. Plateforme Hypsis on-prem. "
     "Pipelines de donn\u00e9es. Ontologie structur\u00e9e."),
    ("Applications m\u00e9tier",
     "Chaque cas d'usage = outil en production. "
     "Pas des POC, des outils au quotidien."),
    ("Acculturation DSI",
     "On construit avec lui, on transf\u00e8re la comp\u00e9tence. "
     "Il devient le gardien de l'\u00e9cosyst\u00e8me IA."),
    ("Formation & reskilling",
     "Plan humain d\u00e8s le premier jour. Formation, reconversion, "
     "coaching, conduite du changement."),
]
for i, (ct, cb) in enumerate(build):
    left = MARGIN + i * (cw4 + Inches(0.2))
    bc = BRONZE if i == 3 else None
    card_with_text(s, left, y, cw4, Inches(2.0), ct, cb, bg=PAPER_WARM, border_color=bc)

punchline(s, y + Inches(2.3),
          "Toutes les briques activ\u00e9es. Conseil \u00b7 Ontologie \u00b7 "
          "Infrastructure \u00b7 IA \u00b7 Humain. Revenu : one-shot \u00b7 100\u2013500K.")

# ---------- S6 RUN ----------
s = prs.slides.add_slide(blank)
y = header_block(s, "Acte II", "Phase 3 \u2014 Run",
                 "Le client est autonome. Le produit reste chez lui.",
                 "Le client peut partir avec tout ce qui a \u00e9t\u00e9 construit. "
                 "C'est pr\u00e9cis\u00e9ment parce qu'il peut partir qu'il choisit de rester.",
                 dark=True)

cw3 = (CONTENT_W - Inches(0.4)) / 3
run_cards = [
    ("Plateforme en place",
     "Hypsis reste d\u00e9ploy\u00e9e chez le client. Les \u00e9quipes internes "
     "construisent de nouvelles apps \u2014 sans d\u00e9pendre de nous."),
    ("Intelligence p\u00e9renne",
     "Mod\u00e8les IA, algorithmes, briques documentaires embarqu\u00e9s. "
     "Mises \u00e0 jour = nouvelles capacit\u00e9s sans r\u00e9int\u00e9gration."),
    ("Propri\u00e9t\u00e9 totale",
     "Donn\u00e9es, mod\u00e8les, applications, code \u2014 tout appartient au client. "
     "Rien pi\u00e9g\u00e9 dans un cloud."),
]
for i, (ct, cb) in enumerate(run_cards):
    left = MARGIN + i * (cw3 + Inches(0.2))
    bc = GOLD if i == 2 else None
    card_with_text(s, left, y, cw3, Inches(1.8), ct, cb,
                   bg=CARD_DARK, border_color=bc, title_color=PAPER, body_color=TEXT_LIGHT)

punchline(s, y + Inches(2.1),
          "Revenu r\u00e9current : plateforme Hypsis (h\u00e9bergement, support, "
          "mises \u00e0 jour) + licences mod\u00e8les IA. 3\u201315K/mois par client.",
          dark=True)

# ---------- S7 VERTICALES + BRIQUES ----------
s = prs.slides.add_slide(blank)
y = header_block(s, "Acte III", "Le groupe",
                 "Deux verticales. Cinq briques.",
                 "Le groupe s'organise en verticales sectorielles port\u00e9es par des experts "
                 "m\u00e9tier, rendues possibles par cinq capacit\u00e9s transverses.")

cw2 = (CONTENT_W - Inches(0.3)) / 2
for i, (ct, cb) in enumerate([
    ("Transformation ops industrielle",
     "A\u00e9ro, d\u00e9fense, industrie. Partenaire expert du secteur, int\u00e9gr\u00e9 au groupe."),
    ("Transformation ops tertiaire",
     "Services, finance, retail, luxe. Unit\u00e9 d\u00e9di\u00e9e, constitu\u00e9e par build-up."),
]):
    left = MARGIN + i * (cw2 + Inches(0.3))
    card_with_text(s, left, y, cw2, Inches(0.9), ct, cb, bg=PAPER_WARM, border_color=BRONZE)

y += Inches(1.1)
cw5 = (CONTENT_W - Inches(0.5)) / 5
bricks = [
    ("Strat\u00e9gie Data & IA",
     "COMEX, Master Plan, CDO Office", "Discovery + Build", ""),
    ("Ontologie & donn\u00e9es",
     "Cartographie, ing\u00e9nierie des donn\u00e9es", "Discovery + Build",
     "AI&D \u00b7 30 consultants \u00b7 6M CA"),
    ("Plateforme souveraine",
     "Cr\u00e9ation, d\u00e9ploiement, supervision d'apps IA on-prem", "Build + Run",
     "Hypsis"),
    ("Excellence scientifique",
     "Mod\u00e8les IA open source, \u00e9valu\u00e9s par les pairs", "Build + Run",
     "Jolibrain \u00b7 5 PhDs \u00b7 50+ GPUs"),
    ("Transformation humaine",
     "Formation, reskilling, conduite du changement", "Build",
     "Horizon d\u00e9but 2027"),
]
for i, (ct, cb, phase, sub) in enumerate(bricks):
    left = MARGIN + i * (cw5 + Inches(0.125))
    card_box(s, left, y, cw5, Inches(2.2), bg=PAPER_WARM)
    t = tb(s, left + Inches(0.1), y + Inches(0.08), cw5 - Inches(0.2), Inches(2.0))
    tf = t.text_frame
    tf.word_wrap = True
    set_text(tf, ct, size=11, bold=True, color=INK)
    add_para(tf, cb, size=9, color=INK_SOFT)
    if sub:
        add_para(tf, sub, size=8, color=INK_MUTED)
    add_para(tf, phase, size=9, bold=True, color=INK)

punchline(s, y + Inches(2.4),
          "Les verticales apportent l'expertise m\u00e9tier. Les briques apportent les "
          "capacit\u00e9s. C'est l'int\u00e9gration des deux qui produit la transformation.")

# ---------- S8 CHA\u00ceNON MANQUANT ----------
s = prs.slides.add_slide(blank)
y = header_block(s, "Acte III", "Le cha\u00eenon manquant",
                 "La brique strat\u00e9gique qui transforme le groupe.",
                 "La phase de Discovery est le point d'entr\u00e9e de toute la cha\u00eene de valeur. "
                 "C'est elle qui convainc le COMEX et g\u00e9n\u00e8re le pipeline.",
                 dark=True)

cw2 = (CONTENT_W - Inches(0.3)) / 2
card_box(s, MARGIN, y, cw2, Inches(3.2), bg=CARD_DARK, border_color=GOLD)
t = tb(s, MARGIN + Inches(0.15), y + Inches(0.1), cw2 - Inches(0.3), Inches(3.0))
tf = t.text_frame
tf.word_wrap = True
set_text(tf, "Le profil id\u00e9al", size=13, bold=True, color=PAPER)
for b in [
    "Cabinet fran\u00e7ais, ind\u00e9pendant, d\u00e9di\u00e9 \u00e0 la strat\u00e9gie Data & IA",
    "Positionn\u00e9 en amont \u2014 DG, CDO, responsables m\u00e9tier",
    "Cycle de vie complet : acculturation \u2192 autonomie des BU",
    "Culture de rigueur strat\u00e9gique (70% advisory / 30% build)",
    "30+ grands comptes, taux de renouvellement \u00e9lev\u00e9",
    "\u00c9cosyst\u00e8me propri\u00e9taire : universit\u00e9 interne, veille tech, capitalisation",
]:
    add_para(tf, f"\u2022 {b}", size=10, color=TEXT_LIGHT)

rx = MARGIN + cw2 + Inches(0.3)
card_box(s, rx, y, cw2, Inches(3.2), bg=CARD_DARK)
t = tb(s, rx + Inches(0.15), y + Inches(0.1), cw2 - Inches(0.3), Inches(3.0))
tf = t.text_frame
tf.word_wrap = True
set_text(tf, "La synergie", size=13, bold=True, color=PAPER)
add_para(tf,
         f"{LAQUO} Demain, le conseil strat\u00e9gique ne pourra plus se contenter de "
         f"recommander \u2014 il devra d\u00e9ployer. Pour d\u00e9ployer, il faut une plateforme, "
         f"des mod\u00e8les, et une infrastructure. {RAQUO}",
         size=11, italic=True, color=TEXT_LIGHT)
add_para(tf, "", size=6)
add_para(tf, "Ce que le groupe apporte", size=11, bold=True, color=GOLD)
add_para(tf, "Mod\u00e8les IA propri\u00e9taires, plateforme souveraine, infrastructure "
         "\u2014 ce que les cabinets conseil IA sous-traitent aujourd'hui \u00e0 des ESN.",
         size=10, color=TEXT_LIGHT)
add_para(tf, "", size=6)
add_para(tf, "Ce qu'un tel cabinet apporte", size=11, bold=True, color=GOLD)
add_para(tf, "Acc\u00e8s direct aux COMEX, 50+ consultants certifi\u00e9s, "
         "moteur commercial \u00e0 80% de r\u00e9currence.",
         size=10, color=TEXT_LIGHT)

# ---------- S9 STRUCTURE ----------
s = prs.slides.add_slide(blank)
y = header_block(s, "Acte IV", "Structure",
                 "Un groupe con\u00e7u pour le build-up.",
                 "Filiales autonomes, P&L s\u00e9par\u00e9s, gouvernance compatible fonds. "
                 "Chaque fondateur qui rejoint garde les r\u00eanes de ce qu'il a construit.")

sw = (CONTENT_W - Inches(0.6)) / 4
for i, (num, lab) in enumerate([
    ("17M\u20ac", "CA consolid\u00e9 ann\u00e9e 1"),
    ("95+", "effectifs totaux"),
    ("5", "briques de capacit\u00e9"),
    ("2", "verticales sectorielles"),
]):
    left = MARGIN + i * (sw + Inches(0.2))
    t = tb(s, left, y, sw, Inches(0.9))
    tf = t.text_frame
    set_text(tf, num, size=32, bold=True, color=BRONZE, align=PP_ALIGN.CENTER)
    add_para(tf, lab, size=11, color=INK_SOFT, align=PP_ALIGN.CENTER)

y += Inches(1.1)
cw3 = (CONTENT_W - Inches(0.4)) / 3
for i, (ct, cb) in enumerate([
    ("Gouvernance",
     "CEO groupe : Yvan Chabanne, ex-DG Scalian (~5 000 pers.)\n"
     "Chaque BU a son propre CEO et son P&L\n"
     "CEOs entrants obtiennent un si\u00e8ge au conseil"),
    ("Autonomie des BU",
     "Le groupe ajoute les synergies commerciales et les services partag\u00e9s. "
     "Il ne micro-manage pas. Un fondateur qui rejoint garde son CEO, "
     "son \u00e9quipe, ses clients."),
    ("BSPCE",
     "Loi de Finances 2026 : holding \u00e9met des BSPCE aux employ\u00e9s de toute "
     "filiale d\u00e9tenue \u00e0 75%+. Equity dans la holding, avantage fiscal, vesting 4 ans."),
]):
    left = MARGIN + i * (cw3 + Inches(0.2))
    bc = BRONZE if i == 2 else None
    card_with_text(s, left, y, cw3, Inches(2.0), ct, cb, bg=PAPER_WARM, border_color=bc)

# ---------- S10 MOD\u00c8LE \u00c9CONOMIQUE ----------
s = prs.slides.add_slide(blank)
y = header_block(s, "Acte IV", "Mod\u00e8le \u00e9conomique",
                 "Le conseil g\u00e9n\u00e8re le cash. La plateforme construit l'actif.",
                 "Trois flux de revenus. La part r\u00e9currente d\u00e9termine le multiple de valorisation.",
                 dark=True)

cw3 = (CONTENT_W - Inches(0.4)) / 3
for i, (ct, cb) in enumerate([
    ("\u2460 Conseil & transformation",
     "One-shot. Diagnostic, Master Plan, d\u00e9ploiement, formation. "
     "100\u2013300K par programme. Marge ~30%."),
    ("\u2461 Plateforme & infrastructure",
     "R\u00e9current. H\u00e9bergement, support, mises \u00e0 jour. "
     "3\u201315K/mois. Marge 70\u201380%."),
    ("\u2462 Intelligence embarqu\u00e9e",
     "R\u00e9current. Mod\u00e8les IA + algorithmes partenaires "
     "d\u00e9ploy\u00e9s sur la plateforme. Royalties aux cr\u00e9ateurs."),
]):
    left = MARGIN + i * (cw3 + Inches(0.2))
    card_with_text(s, left, y, cw3, Inches(1.6), ct, cb,
                   bg=CARD_DARK, title_color=PAPER, body_color=TEXT_LIGHT)

y += Inches(1.9)
cw2 = (CONTENT_W - Inches(0.3)) / 2
card_with_text(s, MARGIN, y, cw2, Inches(1.8),
               "Pourquoi le client reste",
               "La plateforme Hypsis est le socle de toutes les applications. "
               "La quitter = reconstruire tout. Le client reste pour les mises \u00e0 jour, "
               "les mod\u00e8les, le support \u2014 la plateforme est le canal de livraison "
               "de toute nouvelle intelligence.",
               bg=CARD_DARK, title_color=PAPER, body_color=TEXT_LIGHT)

rx = MARGIN + cw2 + Inches(0.3)
card_box(s, rx, y, cw2, Inches(1.8), bg=CARD_DARK, border_color=GOLD)
t = tb(s, rx + Inches(0.15), y + Inches(0.1), cw2 - Inches(0.3), Inches(1.6))
tf = t.text_frame
tf.word_wrap = True
set_text(tf, "Trajectoire du mix", size=13, bold=True, color=PAPER)
add_para(tf, "Aujourd'hui : ~95% one-shot (consulting)", size=11, color=TEXT_LIGHT)
add_para(tf, "Objectif ann\u00e9e 3 : ~73% services / 27% r\u00e9current", size=11, color=TEXT_LIGHT)
add_para(tf, "Objectif ann\u00e9e 5 : ~63% services / 37% r\u00e9current", size=11, color=TEXT_LIGHT)
add_para(tf, "Le r\u00e9current est le multiplicateur de valorisation du groupe.",
         size=10, italic=True, color=TEXT_LIGHT)

# ---------- S11 TRAJECTOIRE ----------
s = prs.slides.add_slide(blank)
y = header_block(s, "Acte IV", "Trajectoire",
                 "17M \u2192 80M en 5 ans. 15% de croissance organique.",
                 "Le saut ann\u00e9e 3 vient d'une acquisition mid-market. "
                 "La marge remonte gr\u00e2ce \u00e0 l'ARR plateforme.")

table_data = [
    ["", "Constitution", "Ann\u00e9e 1\u20132", "Ann\u00e9e 3", "Ann\u00e9e 5"],
    ["CA groupe", "17M", "22M", "55M", "80M"],
    ["dont ARR", "~0", "2M", "7M", "20M"],
    ["EBITDA", "2,5M", "3,5M", "7M", "16M"],
    ["Marge EBITDA", "15%", "16%", "13%", "20%"],
    ["\u00c9v\u00e9nement cl\u00e9", "4 acquisitions fondatrices",
     "Entr\u00e9e fonds PE + brique RH",
     "Acquisition mid-market (~30M CA)",
     "Croissance organique + ARR"],
]
rows, cols = len(table_data), len(table_data[0])
tbl_shape = s.shapes.add_table(rows, cols, MARGIN, y, CONTENT_W, Inches(2.5))
tbl = tbl_shape.table
for r in range(rows):
    for c in range(cols):
        cell = tbl.cell(r, c)
        cell.text = table_data[r][c]
        for p in cell.text_frame.paragraphs:
            for run in p.runs:
                run.font.size = Pt(11)
                run.font.color.rgb = INK
                if r == 0:
                    run.font.bold = True
                    run.font.color.rgb = INK_MUTED
                if c == 0 and r > 0:
                    run.font.bold = True
                if r == 2:
                    run.font.bold = True
                    run.font.color.rgb = BRONZE
        cell.fill.solid()
        if r == 2:
            cell.fill.fore_color.rgb = HIGHLIGHT_BG
        elif r == 0:
            cell.fill.fore_color.rgb = PAPER_WARM
        else:
            cell.fill.fore_color.rgb = WHITE

punchline(s, y + Inches(2.7),
          "La marge EBITDA baisse temporairement \u00e0 l'ann\u00e9e 3 "
          "(co\u00fbts d'int\u00e9gration) puis converge vers 20% gr\u00e2ce \u00e0 la "
          "mont\u00e9e en charge de l'ARR \u2014 la ligne qui fait la valorisation.")

# ---------- S12 L'\u00c9QUIPE ----------
s = prs.slides.add_slide(blank)
y = header_block(s, "Acte V", "L'\u00e9quipe",
                 "Une \u00e9quipe qui a fait le chemin \u2014 pas une qui le th\u00e9orise.",
                 "Construction de groupe \u00b7 Transformation industrielle \u00b7 "
                 "R&D IA en production.",
                 dark=True)

cw4 = (CONTENT_W - Inches(0.6)) / 4
team = [
    ("Yvan Chabanne", "CEO du groupe",
     "Ex-DG Scalian (ESN ~5 000 personnes). Construction et pilotage "
     "d'un groupe multi-filiales dans le conseil tech et l'ing\u00e9nierie."),
    ("William Le Ferrand", "CTO groupe \u00b7 CEO Hypsis",
     "Ex-CTO Sogeclair (ETI a\u00e9ronautique, 1 300 pers.). Transformation "
     "digitale d'une ETI industrielle cot\u00e9e. Cr\u00e9ateur de la plateforme Hypsis."),
    ("Emmanuel Benazera", "CSO groupe \u00b7 CEO Jolibrain",
     "10 ans de R&D IA en production. 5 PhDs, publications NeurIPS / ICML / AAAI. "
     "Clients : Airbus, Dassault, SNCF, CNES. 50+ GPUs."),
    ("Florian Lepage", "CRO groupe",
     "Ex-COO d'un \u00e9diteur de logiciel RH. Exp\u00e9rience du scale-up "
     "commercial et de la structuration d'une offre r\u00e9currente."),
]
for i, (name, role, bio) in enumerate(team):
    left = MARGIN + i * (cw4 + Inches(0.2))
    card_box(s, left, y, cw4, Inches(2.2), bg=CARD_DARK, border_color=GOLD)
    t = tb(s, left + Inches(0.12), y + Inches(0.1), cw4 - Inches(0.24), Inches(2.0))
    tf = t.text_frame
    tf.word_wrap = True
    set_text(tf, name, size=12, bold=True, color=PAPER)
    add_para(tf, role, size=10, bold=True, color=GOLD)
    add_para(tf, bio, size=9, color=TEXT_LIGHT)

y += Inches(2.5)
placeholders = [
    "CEO conseil strat\u00e9gique Data & IA",
    "CEO verticale industrielle",
    "CEO transformation humaine",
    "\u2026",
]
for i, ph in enumerate(placeholders):
    left = MARGIN + i * (cw4 + Inches(0.2))
    shape = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, y, cw4, Inches(0.8))
    shape.fill.solid()
    shape.fill.fore_color.rgb = CARD_DARK
    shape.line.color.rgb = INK_MUTED
    shape.line.width = Pt(1)
    shape.line.dash_style = 4
    t = tb(s, left, y + Inches(0.15), cw4, Inches(0.5))
    set_text(t.text_frame, ph, size=10, color=INK_MUTED, align=PP_ALIGN.CENTER)

punchline(s, y + Inches(1.0),
          "L'\u00e9quipe fondatrice couvre la construction de groupe, la transformation "
          "industrielle, la R&D IA en production et le scale-up commercial. "
          "Chaque fondateur qui rejoint le groupe compl\u00e8te le puzzle.",
          dark=True)

# ---------- S13 POSITIONNEMENT ----------
s = prs.slides.add_slide(blank)
y = header_block(s, "Acte V", "Positionnement",
                 "Ce que chaque acteur fait. Ce qui lui manque.",
                 "Nous assemblons les meilleures briques existantes "
                 "plut\u00f4t que de tout construire from scratch.")

pos_data = [
    ["Acteur", "Ce qu'il fait", "Ce qui lui manque"],
    ["Accenture / Capgemini", "Transformation IA grands groupes",
     "Trop cher pour ETI, pas souverain, lock-in"],
    ["ESN mid-market", "Int\u00e9gration tech, data",
     "Pas de vision strat\u00e9gique, pas de mod\u00e8les IA"],
    ["Cabinets conseil (McKinsey, BCG)", "Strat\u00e9gie, diagnostic",
     "Pas de d\u00e9ploiement, pas d'infra"],
    ["\u00c9diteurs IA (Dataiku, Palantir)", "Plateforme data/IA",
     "Lock-in, cloud, pas d'accompagnement humain"],
    ["Cabinets conseil Data & IA", "Strat\u00e9gie IA, Master Plan, CDO Office",
     "Pas de plateforme, pas de mod\u00e8les IA"],
    ["Le groupe", "Strat\u00e9gie + Data + IA + Infra souveraine + Humain",
     "Int\u00e9gr\u00e9, souverain, pas de lock-in, propri\u00e9t\u00e9 client"],
]
rows, cols = len(pos_data), len(pos_data[0])
tbl_shape = s.shapes.add_table(rows, cols, MARGIN, y, CONTENT_W, Inches(3.0))
tbl = tbl_shape.table
for r in range(rows):
    for c in range(cols):
        cell = tbl.cell(r, c)
        cell.text = pos_data[r][c]
        for p in cell.text_frame.paragraphs:
            for run in p.runs:
                run.font.size = Pt(11)
                run.font.color.rgb = INK
                if r == 0:
                    run.font.bold = True
                    run.font.color.rgb = INK_MUTED
                if r == rows - 1:
                    run.font.bold = True
        cell.fill.solid()
        if r == rows - 1:
            cell.fill.fore_color.rgb = HIGHLIGHT_BG
        elif r == 0:
            cell.fill.fore_color.rgb = PAPER_WARM
        else:
            cell.fill.fore_color.rgb = WHITE

# ---------- S14 POURQUOI NOUS REJOINDRE ----------
s = prs.slides.add_slide(blank)
y = header_block(s, "Acte VI", "Pourquoi nous rejoindre",
                 "Vous gardez votre entreprise. Vous acc\u00e9dez \u00e0 un groupe.",
                 "Ce que le groupe change pour un fondateur qui nous rejoint.",
                 dark=True)

cw3 = (CONTENT_W - Inches(0.4)) / 3
for i, (ct, cb) in enumerate([
    ("Votre identit\u00e9",
     "Votre CEO, votre P&L, votre \u00e9quipe, vos clients. "
     "Le groupe ne r\u00e9\u00e9crit pas ce qui fonctionne."),
    ("Votre front commercial",
     "Programmes de transformation \u00e0 100\u2013500K vendus au niveau COMEX. "
     "Vos comp\u00e9tences int\u00e9gr\u00e9es dans une offre plus large, vendue plus cher."),
    ("Du conseil \u00e0 l'impact",
     "Vous ne sous-traitez plus \u00e0 des ESN. Vous gardez le contr\u00f4le de la "
     "strat\u00e9gie ET vous d\u00e9ployez avec la plateforme et l'infra du groupe."),
]):
    left = MARGIN + i * (cw3 + Inches(0.2))
    bc = GOLD if i == 2 else None
    card_with_text(s, left, y, cw3, Inches(1.6), ct, cb,
                   bg=CARD_DARK, border_color=bc, title_color=PAPER, body_color=TEXT_LIGHT)

y += Inches(1.9)
cw2 = (CONTENT_W - Inches(0.3)) / 2
for i, (ct, cb) in enumerate([
    ("Participation \u00e0 la valeur",
     "BSPCE au niveau de la holding. Roll-over en equity. "
     "La croissance d'une BU b\u00e9n\u00e9ficie \u00e0 tous. Les int\u00e9r\u00eats sont align\u00e9s."),
    ("Protection, pas dilution",
     "P&L s\u00e9par\u00e9s = chaque entit\u00e9 valorisable ind\u00e9pendamment. "
     "Le groupe peut se vendre \u00ab \u00e0 l'appartement \u00bb ou en bloc. "
     "Les deux chemins restent ouverts."),
]):
    left = MARGIN + i * (cw2 + Inches(0.3))
    card_with_text(s, left, y, cw2, Inches(1.3), ct, cb,
                   bg=CARD_DARK, title_color=PAPER, body_color=TEXT_LIGHT)

# ---------- S15 CONTACT ----------
s = prs.slides.add_slide(blank)
add_bg(s, INK)
y = Inches(0.8)
t = tb(s, MARGIN, y, CONTENT_W, Inches(0.3))
set_text(t.text_frame, "Acte VI", size=9, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
y += Inches(0.3)
t = tb(s, MARGIN, y, CONTENT_W, Inches(0.3))
set_text(t.text_frame, "PROCHAINE \u00c9TAPE", size=11, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
y += Inches(0.4)
t = tb(s, Inches(2), y, Inches(9), Inches(0.7))
set_text(t.text_frame, "Construisons ensemble.", size=32, bold=True, color=PAPER, align=PP_ALIGN.CENTER)
y += Inches(0.8)
t = tb(s, Inches(2.5), y, Inches(8), Inches(0.6))
set_text(t.text_frame,
         "Le groupe se construit maintenant. Les briques techniques existent. "
         "L'\u00e9quipe est en place. La pi\u00e8ce manquante, c'est peut-\u00eatre vous.",
         size=14, color=TEXT_LIGHT, align=PP_ALIGN.CENTER)

y += Inches(0.9)
cw3 = (CONTENT_W - Inches(0.4)) / 3
for i, (ct, cb) in enumerate([
    ("Ce que nous avons",
     "\u2022 Plateforme souveraine en production\n"
     "\u2022 10 ans de R&D IA, mod\u00e8les d\u00e9ploy\u00e9s\n"
     "\u2022 30 consultants data, 6M CA\n"
     "\u2022 CEO ex-DG d'un groupe de 5 000 pers."),
    ("Ce que nous cherchons",
     "\u2022 La brique conseil strat\u00e9gique Data & IA\n"
     "\u2022 Un acc\u00e8s direct aux COMEX\n"
     "\u2022 Une m\u00e9thodologie de transformation \u00e9prouv\u00e9e\n"
     "\u2022 Un fondateur qui veut aller plus loin"),
    ("Prochaine \u00e9tape",
     "\u2022 \u00c9change informel entre fondateurs\n"
     "\u2022 D\u00e9monstration de la plateforme\n"
     "\u2022 Exploration des synergies"),
]):
    left = MARGIN + i * (cw3 + Inches(0.2))
    bc = GOLD if i == 2 else None
    card_with_text(s, left, y, cw3, Inches(2.0), ct, cb,
                   bg=CARD_DARK, border_color=bc, title_color=PAPER, body_color=TEXT_LIGHT)

y += Inches(2.3)
t = tb(s, Inches(4), y, Inches(5), Inches(0.4))
set_text(t.text_frame, "ANTERIQ", size=16, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
t = tb(s, Inches(4), y + Inches(0.4), Inches(5), Inches(0.3))
set_text(t.text_frame, f"Confidentiel {DASH} Avril 2026", size=10, color=INK_MUTED, align=PP_ALIGN.CENTER)

# ---------- SAVE ----------
out_path = "/home/wleferrand/dev/hypsis/hypsis/landing/group/anteriq-pitch.pptx"
prs.save(out_path)
print(f"Saved to {out_path}")

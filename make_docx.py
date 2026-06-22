#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# ── Page margins ──────────────────────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin   = Cm(2.2)
    section.right_margin  = Cm(2.2)

# ── Color palette ─────────────────────────────────────────────────────────────
GREEN       = RGBColor(0x2d, 0x6a, 0x4f)
GREEN_LIGHT = RGBColor(0x40, 0x91, 0x6c)
BROWN       = RGBColor(0x6b, 0x42, 0x26)
ORANGE      = RGBColor(0xe0, 0x7b, 0x39)
BLUE        = RGBColor(0x1a, 0x6e, 0xa8)
GOLD        = RGBColor(0xe6, 0xa8, 0x17)
RED         = RGBColor(0xc0, 0x39, 0x2b)
WHITE       = RGBColor(0xff, 0xff, 0xff)
LIGHT_GRAY  = RGBColor(0xf5, 0xf5, 0xf5)
DARK_GRAY   = RGBColor(0x44, 0x44, 0x44)

# ── Helpers ───────────────────────────────────────────────────────────────────

def set_cell_bg(cell, rgb: RGBColor):
    """Set background fill on a table cell."""
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement('w:shd')
    hex_color = f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  hex_color)
    tcPr.append(shd)

def set_table_borders(table, color="CCCCCC", size=4):
    """Add thin borders to every cell in a table."""
    tbl    = table._tbl
    tblPr  = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    tblBrd = OxmlElement('w:tblBorders')
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'),   'single')
        el.set(qn('w:sz'),    str(size))
        el.set(qn('w:color'), color)
        tblBrd.append(el)
    tblPr.append(tblBrd)

def heading(text, level=1, color=GREEN, space_before=14, space_after=4):
    p    = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    if level == 1:
        run = p.add_run(text)
        run.bold      = True
        run.font.size = Pt(17)
        run.font.color.rgb = color
        # green left border via paragraph shading (simulate with a run underline trick)
        # We'll use a page-wide bottom border instead
        pPr  = p._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bot  = OxmlElement('w:bottom')
        bot.set(qn('w:val'),   'single')
        bot.set(qn('w:sz'),    '8')
        bot.set(qn('w:color'), f"{color[0]:02X}{color[1]:02X}{color[2]:02X}")
        pBdr.append(bot)
        pPr.append(pBdr)
    elif level == 2:
        run = p.add_run(text)
        run.bold      = True
        run.font.size = Pt(13)
        run.font.color.rgb = color
    else:
        run = p.add_run(text)
        run.bold      = True
        run.font.size = Pt(11)
        run.font.color.rgb = color
    return p

def body(text, italic=False, color=None, bold=False, size=10.5, space_after=3):
    p    = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(space_after)
    run  = p.add_run(text)
    run.font.size   = Pt(size)
    run.italic      = italic
    run.bold        = bold
    if color:
        run.font.color.rgb = color
    return p

def info_box(text, bg: RGBColor, icon=""):
    """Single-cell shaded box."""
    t    = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    cell = t.cell(0, 0)
    set_cell_bg(cell, bg)
    cell.top_margin    = Pt(5)
    cell.bottom_margin = Pt(5)
    p    = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(2)
    run  = p.add_run(f"{icon}  {text}" if icon else text)
    run.font.size = Pt(9.5)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)

def tip_box(text):
    info_box(text, RGBColor(0xe8,0xf5,0xe9), "💡")

def warning_box(text):
    info_box(text, RGBColor(0xff,0xf8,0xe1), "⚠️")

def important_box(text):
    info_box(text, RGBColor(0xfd,0xec,0xea), "🚨")

def review_box(text):
    info_box(text, RGBColor(0xf0,0xf4,0xff), "💬")

def add_table(headers, rows, col_widths=None):
    n_cols = len(headers)
    t = doc.add_table(rows=1 + len(rows), cols=n_cols)
    set_table_borders(t)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT

    # Header row
    hdr_row = t.rows[0]
    for i, h in enumerate(headers):
        cell = hdr_row.cells[i]
        set_cell_bg(cell, GREEN)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p    = cell.paragraphs[0]
        run  = p.add_run(h)
        run.bold           = True
        run.font.color.rgb = WHITE
        run.font.size      = Pt(9)

    # Data rows
    for ri, row_data in enumerate(rows):
        row = t.rows[ri + 1]
        bg  = RGBColor(0xf9,0xf9,0xf9) if ri % 2 == 1 else WHITE
        for ci, cell_text in enumerate(row_data):
            cell = row.cells[ci]
            set_cell_bg(cell, bg)
            p    = cell.paragraphs[0]
            run  = p.add_run(str(cell_text))
            run.font.size = Pt(9)

    # Column widths
    if col_widths:
        for i, width in enumerate(col_widths):
            for row in t.rows:
                row.cells[i].width = Cm(width)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    return t

def section_divider():
    p   = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr= OxmlElement('w:pBdr')
    bot = OxmlElement('w:bottom')
    bot.set(qn('w:val'),   'single')
    bot.set(qn('w:sz'),    '4')
    bot.set(qn('w:color'), 'CCCCCC')
    pBdr.append(bot)
    pPr.append(pBdr)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(6)

# ═══════════════════════════════════════════════════════════════════════════════
#  TITLE PAGE
# ═══════════════════════════════════════════════════════════════════════════════

# Shaded title block
t = doc.add_table(rows=1, cols=1)
cell = t.cell(0,0)
set_cell_bg(cell, GREEN)
cell.top_margin    = Pt(14)
cell.bottom_margin = Pt(14)
p  = cell.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r  = p.add_run("🏔️  FILZMOOS 2026 – Průvodce výletem")
r.bold           = True
r.font.size      = Pt(20)
r.font.color.rgb = WHITE

doc.add_paragraph()

meta_lines = [
    "📍  Haus Dachstein, Nestlerweg 3, 5532 Filzmoos, Rakousko",
    "📅  1.–6. července 2026",
    "👨‍👩‍👧  5 dospělých (4 + 1 student) + dítě 1,5 roku",
    "🏍️  3 motorky + 1 auto",
    "🗺️  Salzburger Land, Pongau, Rakousko",
]
for line in meta_lines:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(line)
    r.font.size      = Pt(11)
    r.font.color.rgb = DARK_GRAY

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════════
#  1. SOUTĚSKY
# ═══════════════════════════════════════════════════════════════════════════════

heading("🌊  SOUTĚSKY V OKOLÍ – POROVNÁNÍ 5 MÍST")
tip_box("Přijeďte ráno (před 10h) nebo odpoledne (po 15h). Pláštěnku berte vždy – vlhkost a postřik jsou všude.")

# Liechtensteinklamm
heading("⭐⭐⭐⭐⭐  Liechtensteinklamm – NEJLEPŠÍ V OKOLÍ", level=2, color=GOLD)
add_table(
    ["Parametr", "Detail"],
    [
        ["Vzdálenost", "28 km | 30 min jízdy"],
        ["Vstupné", "€16 dospělý | €9 mládež do 18 let | do 6 let zdarma"],
        ["Vaše skupina", "4×€16 + 1×€9 (student) + dítě zdarma = €73  |  se SalzburgerLand Card: ZDARMA"],
        ["Kočárek / nosítko", "❌ Kočárek nelze (900 schodů)  |  ✅ Nosítko nutné"],
        ["Čas návštěvy", "1,5–2 hodiny"],
        ["Parkování", "Přímo u vstupu"],
        ["Web", "liechtensteinklamm.at"],
        ["Google Maps", 'maps.google.com/maps?q=Liechtensteinklamm+St+Johann+im+Pongau'],
    ],
    col_widths=[4.5, 11.5]
)
body("Nejhlubší soutěska v Alpách – stěny až 300 m, šířka na nejužším místě jen pár metrů. 1 km lávek a schodišť, na konci vodopád. 220 000 návštěvníků ročně – zvolena 2. nejkrásnějším místem Rakouska 2022.", space_after=4)
review_box('"Přijeli jsme v 9h ráno, skoro nikdo tam nebyl, krásné světlo."')
review_box('"S malým dítětem v šátku bez problému, jen vezměte pláštěnku."')
warning_box("Vzít: svetr nebo lehká bunda + pláštěnka – i za horkého dne je v rokli chlad a vlhko.")

section_divider()

# Salzachöfen
heading("⭐⭐⭐⭐  Salzachöfen – Zdarma, málo turistů", level=2, color=GREEN)
add_table(
    ["Parametr", "Detail"],
    [
        ["Vzdálenost", "35 km | 35 min jízdy"],
        ["Vstupné", "ZDARMA"],
        ["Kočárek / nosítko", "❌ Kočárek  |  ✅ Nosítko OK"],
        ["Čas návštěvy", "1–1,5 hodiny"],
        ["Google Maps", "maps.google.com/maps?q=Salzach%C3%B6fen+Golling"],
    ],
    col_widths=[4.5, 11.5]
)
body("Kaňon hloubky 90 m – řeka Salzach prořízla vápencové skály. Kolébka alpského raftingu (1931). Turistická stezka se schody. Autentičtější a bez front.", space_after=4)
warning_box("Po dešti ověřit stav vody před odjezdem.")

section_divider()

# Bluntautal
heading("⭐⭐⭐⭐  Bluntautal – Vhodné pro KOČÁREK! 🌟", level=2, color=BLUE)
add_table(
    ["Parametr", "Detail"],
    [
        ["Vzdálenost", "42 km | 40 min jízdy"],
        ["Vstupné", "Volný vstup (parkoviště ~€3)"],
        ["Kočárek / nosítko", "✅ Kočárek celou trasu!  |  ✅ Nosítko"],
        ["Čas návštěvy", "2 hodiny ke 2 jezírkům"],
        ["Google Maps", "maps.google.com/maps?q=Bluntautal+Golling"],
    ],
    col_widths=[4.5, 11.5]
)
body("Romantické 4km údolí, lesní cesta podél potoka, 2 horská jezírka, hostinec Bärenhütte. Minimální převýšení – ideální s kočárkem nebo malým dítětem.", space_after=4)
tip_box("Kombinovat s Gollinger Wasserfall (100m vodopád) – obojí je v Gollingu, celý den za minimum!")

section_divider()

# Gollinger Wasserfall
heading("⭐⭐⭐  Gollinger Wasserfall – 100m vodopád", level=2, color=GREEN_LIGHT)
add_table(
    ["Parametr", "Detail"],
    [
        ["Vzdálenost", "42 km | 40 min (stejné jako Bluntautal – kombinovat!)"],
        ["Vstupné", "~€4 dospělý"],
        ["Kočárek / nosítko", "❌ Kočárek obtížný  |  ✅ Nosítko OK"],
        ["Čas návštěvy", "1 hodina"],
    ],
    col_widths=[4.5, 11.5]
)
body("Přirozený krasový vodopád, cesta přes skalní průchody. Přístupný od roku 1805.")

section_divider()

# Srovnávací tabulka
heading("📊  Srovnání na první pohled", level=3, color=DARK_GRAY)
add_table(
    ["Soutěska", "Km", "Vstup", "Kočárek", "Náročnost", "Hodnocení"],
    [
        ["Liechtensteinklamm", "28", "€16/os.", "❌", "⭐⭐⭐", "⭐⭐⭐⭐⭐"],
        ["Salzachöfen",        "35", "zdarma",  "❌", "⭐⭐",   "⭐⭐⭐⭐"],
        ["Bluntautal",         "42", "volný",   "✅", "⭐",     "⭐⭐⭐⭐"],
        ["Gollinger Wasserfall","42", "~€4/os.", "❌", "⭐⭐",   "⭐⭐⭐"],
    ],
    col_widths=[4.5, 1.5, 2.5, 2.5, 2.5, 2.5]
)

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════════
#  2. PŘÍRODA – JESKYNĚ, VODOPÁDY, LEDOVCE
# ═══════════════════════════════════════════════════════════════════════════════

heading("🏔️  PŘÍRODA – JESKYNĚ, VODOPÁDY, LEDOVCE")

# Eisriesenwelt
heading("🧊  Eisriesenwelt – Největší ledová jeskyně světa", level=2, color=BLUE)
add_table(
    ["Parametr", "Detail"],
    [
        ["Vzdálenost", "30 km | 35 min jízdy"],
        ["Vstupné (Kombi = jeskyně + lanovka)", "€41 online / €45 na místě | student €30 | dítě do 5 let zdarma"],
        ["Vaše skupina (online)", "4×€41 + 1×€30 + dítě zdarma = €194  |  se SalzburgerLand Card: ZDARMA"],
        ["Kočárek / nosítko", "❌ Kočárek  |  ✅ Nosítko – ZIMNÍ OBLEČENÍ PRO DÍTĚ!"],
        ["Celkový čas", "~3 hodiny od parkoviště"],
        ["Trasa od parkoviště", "20 min pěšky + lanovka + 20 min pěšky + 1h+ v jeskyni"],
        ["Web + rezervace", "eisriesenwelt.at"],
    ],
    col_widths=[4.5, 11.5]
)
important_box("DÍTĚ MUSÍ MÍT ZIMNÍ OBLEČENÍ – uvnitř je pod 0°C celou hodinu! Kombinéza, čepice, rukavice.")
important_box("POVINNÁ ONLINE REZERVACE s konkrétním časovým slotem! Rezervujte předem – v červenci bývá plno.")
review_box('"Přijít na první ranní slot 9:30 – nejkrásnější světlo a málo lidí."')
review_box('"Dítě v šátku s zimní kombinézou – bez problému, uvnitř opravdu mrzne."')

section_divider()

# Dachstein
heading("🏔️  Dachstein Gletscher + Eispalast + Hängebrücke", level=2, color=GREEN)
add_table(
    ["Ticket", "Dospělý", "Student (mládež)", "Dítě"],
    [
        ["Kombi (gondola + atrakce)", "€73", "€55,50", "€37,50"],
        ["Pouze gondola tam+zpět",    "€61", "€46",    "€30,50"],
        ["Pouze atrakce (Hängebrücke+Eispalast+Treppe)", "€13,50", "€10,50", "€7,50"],
    ],
    col_widths=[6.5, 3, 3.5, 3]
)
body("Gondola na 2 700 m (6 min), Eispalast uvnitř ledovce (−2 až −5 °C), visutý most 400 m nad propastí, výhled na 30 třítisícovek. 15 km od Filzmoos.")
important_box("Fixplatzbuchung (rezervace gondoly) POVINNÁ online! + €2,90/os. Rezervovat předem na tickets.derdachstein.at")
tip_box("Mautstraße Dachstein je v ceně gondoly zdarma – nezapomenout orazítkovat lístek v bergstation!")

section_divider()

# Krimmler
heading("💧  Krimmler Wasserfälle – Nejvyšší vodopády Rakouska", level=2, color=GREEN_LIGHT)
add_table(
    ["Parametr", "Detail"],
    [
        ["Vzdálenost", "75 km | 75 min jízdy"],
        ["Vstupné", "~€5 dospělý | dítě ~€3"],
        ["Vaše skupina", "~€28 + parkování ~€5  |  se SalzburgerLand Card: ZDARMA"],
        ["Kočárek / nosítko", "✅ Kočárek ke spodní části (15 min)  |  ✅ Nosítko na celou trasu"],
        ["Celá trasa", "4 km, 500 m převýšení, 2–3 hodiny"],
        ["Web", "krimmler-wasserwelten.at"],
    ],
    col_widths=[4.5, 11.5]
)
warning_box("Pláštěnka POVINNÁ – vodní tříšť stříká desítky metrů i za slunečného počasí!")

section_divider()

# Bachlalm
heading("🦫  Bachlalm + Svišti (Murmeltiere)", level=2, color=GREEN_LIGHT)
body("5 km od ubytování | Shuttle Murmeltierexpress z Filzmoos | ✅ Kočárek na hlavní cestě OK")
body("Alpské jezírko s výhledem na Bischofsmütze, kolonie 30+ divokých svišťů. Ideální pro 1. nebo poslední den.")
tip_box("Perfektní výlet pro dítě 1,5 roku – svišti jsou blízko, nejsou plachí, jezírko je krásné.")

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════════
#  3. MOTORKÁŘSKÉ TRASY
# ═══════════════════════════════════════════════════════════════════════════════

heading("🏍️  MOTORKÁŘSKÉ TRASY")
warning_box("Ráno do 9h mohou být průsmyky uzavřené kvůli ledu (zejm. Grossglockner). Zkontrolovat grossglockner.at → Befahrbarkeit. Nejlepší čas odjezdu: 9–10h.")

routes = [
    ("A", "⭐⭐⭐⭐⭐", "Radstädter Tauernpass",   "~110 km",  "Bez mýtného",      "Filzmoos → Radstadt → Radstädter Tauernpass 1739m → Mauterndorf → Tamsweg → Untertauern → Filzmoos",         "Alpský průsmyk, serpentiny, Lungau = nejslunečnější kraj. Kombinovat s Burg Mauterndorf."),
    ("B", "⭐⭐⭐⭐",  "Hochkönig okruh",          "~100 km",  "Bez mýtného",      "Filzmoos → Bischofshofen → Werfen → Maria Alm → Dienten → Mühlbach → Filzmoos",                              "Výhledy na Hochkönig 2941m. Kombinovat s Eisriesenwelt nebo Hohenwerfen."),
    ("C", "⭐⭐⭐⭐⭐", "Grossglockner",            "~250 km",  "€36,50/moto",      "Filzmoos → Zell am See → Fusch (pokladna) → 48km Grossglockner → Kaiser-Franz-Josefs-Höhe → zpět",           "3 motorky + 1 auto = 3×€36,50 + €46,50 = €156 celkem. Celodenní výlet."),
    ("D", "⭐⭐⭐⭐",  "Gerlospass + Zillertal",   "~200 km",  "Gerlos ~€14/moto", "Filzmoos → Zell am See → Gerlospass → Zillertal → zpět nebo přes Brenner",                                   "Méně přetíženější než Grossglockner, krásné výhledy."),
    ("E", "⭐⭐⭐⭐",  "Obertauern okruh",         "~120 km",  "Bez mýtného",      "Filzmoos → Radstadt → Obertauern 1740m → Tweng → Mauterndorf → Radstadt → Filzmoos",                         "Klidnější průsmyk, méně turistů."),
]

for code, stars, name, km, mytne, trasa, pozn in routes:
    heading(f"Trasa {code}: {name}  {stars}", level=2, color=GREEN if stars.count("⭐") >= 5 else GREEN_LIGHT)
    add_table(
        ["Parametr", "Detail"],
        [
            ["Délka / čas", km],
            ["Mýtné", mytne],
            ["Trasa", trasa],
            ["Poznámka", pozn],
        ],
        col_widths=[3.5, 12.5]
    )

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════════
#  4. SOLNÉ DOLY A PODZEMÍ
# ═══════════════════════════════════════════════════════════════════════════════

heading("⛏️  SOLNÉ DOLY A ZAJÍMAVÉ PODZEMÍ")

heading("🧂  Salzwelten – Solné doly (Bad Dürrnberg nebo Hallstatt)", level=2, color=BROWN)
add_table(
    ["Parametr", "Detail"],
    [
        ["Bad Dürrnberg (bližší)", "55 km | 50 min jízdy"],
        ["Hallstatt (slavnější)",  "75 km | 70 min jízdy"],
        ["Vstupné", "~€22 dospělý | ~€12 dítě 4–15 let | do 4 let zdarma"],
        ["Vaše skupina", "~€122  |  se SalzburgerLand Card: ZDARMA"],
        ["Web", "salzwelten.at"],
    ],
    col_widths=[4.5, 11.5]
)
body("Podzemní solný důl s tunely, podzemním jezerem a skluzavkami! Hallstatt navíc UNESCO vesnička u jezera.")
tip_box("Přesně ta věc co doma nezažijete. Skvělé na deštivý den – pod zemí je vždy sucho.")

section_divider()

heading("⛏️  Bergwerk Schwaz – Středověké stříbrné doly (Tyrolsko)", level=2, color=BROWN)
add_table(
    ["Parametr", "Detail"],
    [
        ["Vzdálenost", "~130 km | 100 min jízdy"],
        ["Vstupné", "~€20 dospělý"],
        ["Web", "silberbergwerk.at"],
    ],
    col_widths=[4.5, 11.5]
)
body("Největší středověký stříbrný důl světa, průjezd vláčkem. Kombinovat s Gerlospassem nebo Zillertálem.")

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════════
#  5. FARMY, SÝRY, LOKÁLNÍ PRODUKTY
# ═══════════════════════════════════════════════════════════════════════════════

heading("🐄  FARMY, SÝRY A LOKÁLNÍ PRODUKTY")
tip_box("Tohle je ten zážitek – dřevěné prkno se sýry, uzené maso, čerstvé mléko přímo od krávy. Ve Filzmoos to máte doslova za rohem.")

heading("🛒  Ab Hof prodejny přímo ve Filzmoos", level=2, color=BROWN)
add_table(
    ["Místo", "Adresa", "Otevírací doba", "Co nabízí"],
    [
        ["⭐ Nestlerhof  (7 min od vašeho ubytování!)", "Nestlerweg 10", "do 19:00", "Vejce, mléko, sýr"],
        ["Selbstbedienungs-Hofladen Kirchgasshof", "Kirchgassweg 1", "do 22:00 (!)", "Samoobslužný hofladen"],
        ["Wexlerhof", "Wexlerweg 2", "zavolat: +43 664 3828200", "Ab hof prodej"],
        ["Bauernmarkt Kirchgasshof", "Kirchgassweg 1", "sezónně – ověřit", "Farmářský trh"],
    ],
    col_widths=[5, 4, 3.5, 3.5]
)

heading("🧀  Almy s Brettljause a domácími sýry", level=2, color=BROWN)
add_table(
    ["Alm", "Výška", "Přístup", "Co dostat"],
    [
        ["Kirchgasshütte / Aualm", "1366 m", "Shuttle nebo pěšky 1h", "Brettljause, Kaspressknödel, Kaiserschmarrn"],
        ["Oberhofalm", "1268 m", "Pěšky 45 min", "Brettljause, čerstvé mléko, rodinná farma"],
        ["Unterhofalm", "1280 m", "Pěšky 30 min", "Almjause, domácí produkty"],
        ["Schwaigalm", "1358 m", "Pěšky 1h", "Käsesnitte, Jause, alpský sýr"],
        ["Hofpürglhütte", "1705 m", "Pěšky 2,5h nebo 45 min z Aualm", "Domácí jídla – jen HOTOVOST! Otev. do 27.9."],
    ],
    col_widths=[4, 2, 4, 6]
)

heading("🏺  Sennerei a výrobna sýra", level=2, color=BROWN)
body("• Sennerei Ramsau am Dachstein (~25 km) – tradiční alpský sýr. Zeptat se v Filzmoos Tourismus (Dorfstraße 12).")
body("• Dachstein-Destillerie Mandlberggut – 10 km od Filzmoos (Mandlbergweg 11, Mandling). Lokální palírna alpských destilátů. Otevřeno do 18h.")
tip_box("AlmCard Filzmoos = slevy na jídlo na 20+ almách. Ověřit cenu u Filzmoos Tourismus, tel. +43 6453 8235.")

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════════
#  6. RESTAURACE
# ═══════════════════════════════════════════════════════════════════════════════

heading("🍽️  RESTAURACE A ALMY")

heading("Ve Filzmoos – přehled", level=2, color=GREEN)
add_table(
    ["Restaurace", "Cena/os.", "Specialita", "Poznámka"],
    [
        ["☕ Café Hirschenau",            "€5–15",  "Dorty, snídaně, káva",         "Nejlepší kavárna ve Filzmoos"],
        ["🍺 De Hiatabuam",               "€12–22", "Lokální kuchyně, Schnitzel",   "Oblíbené u místních"],
        ["🏘️ Gasthof Fiakerwirt",         "€12–20", "Guláš, Schnitzel",             "Klasický gasthof"],
        ["🍔 MeiZeit Lodge",              "€15–25", "Burgery, steaky",              "Otevřeno do 22h"],
        ["🍝 Mamma Norica",               "€15–25", "Italsko-alpská kuchyně",       "Rodinná atmosféra"],
        ["🥩 Panorama Steakrest. Reithof","€25–40", "Steaky, výhled na Bischofsmütze","Nejhezčí výhled"],
    ],
    col_widths=[4.5, 2.5, 5, 4]
)

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════════
#  7. DÉŠŤ
# ═══════════════════════════════════════════════════════════════════════════════

heading("☔  PROGRAM NA DEŠTIVÝ DEN")
body("Deštivý den? Žádný problém – v okolí máte skvělé kryté atrakce.")
add_table(
    ["Atrakce", "Vzdálenost", "Cena skupiny", "Proč je OK za deště"],
    [
        ["🧊 Eisriesenwelt", "30 km", "~€194 / zdarma s kartou", "Pod zemí, mráz, nezapomenutelné"],
        ["🧂 Salzwelten Bad Dürrnberg", "55 km", "~€122 / zdarma s kartou", "Tunely, skluzavky, podzemní jezero"],
        ["🏔️ Dachstein Eispalast", "15 km", "~€67 (jen atrakce)", "Gondola + ledový palác"],
        ["🏙️ Haus der Natur Salzburg", "80 km", "~€60", "Příroda, vesmír, reptilie – ne jen dějepis!"],
        ["🏰 Burg Mauterndorf", "35 km", "~€45", "Středověký hrad, interiéry"],
        ["🏺 Dachstein-Destillerie", "10 km", "levné", "Palírna alpských destilátů"],
    ],
    col_widths=[4.5, 2.5, 4, 5]
)

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════════
#  8. SLEVOVÉ KARTY
# ═══════════════════════════════════════════════════════════════════════════════

heading("💳  SLEVOVÉ KARTY")

heading("SalzburgerLand Card (6 dní) – vyplatí se?", level=2, color=GREEN)
add_table(
    ["", "Cena"],
    [
        ["Dospělý (6 dní)",   "€103"],
        ["Dítě (6 dní)",      "€52"],
        ["Pro skupinu 5+1",   "5×€103 + 1×€52 = €567"],
    ],
    col_widths=[6, 10]
)
body("Zahrnuje zdarma: Eisriesenwelt, Hohenwerfen, Liechtensteinklamm, Krimmler Wasserfälle, Salzwelten, Zoo Salzburg, Festung Hohensalzburg, Hellbrunn a ~180 dalších atrakcí.")
tip_box("Karta se vyplatí pokud navštívíte alespoň 5 atrakcí ze seznamu. Zejm. pokud jedete do Salzburgu – tam jsou 3–4 atrakce v ceně.")
body("Web: salzburgerland.com/de/salzburgerland-card", color=BLUE)

heading("Filzmoos Sommer Card – pravděpodobně zdarma pro hosty", level=2, color=GREEN_LIGHT)
body("Slevy na místní lanovky, turistické autobusy, atrakce. Zeptat se na recepci Haus Dachstein ihned po příjezdu!")

heading("AlmCard Filzmoos", level=2, color=GREEN_LIGHT)
body("Slevy na jídlo na 20+ almách. Ověřit cenu u Filzmoos Tourismus (Dorfstraße 12, tel. +43 6453 8235).")

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════════
#  9. PLÁN DNÍ
# ═══════════════════════════════════════════════════════════════════════════════

heading("📅  NAVRHOVANÝ PROGRAM 5 DNÍ")

days = [
    ("Den 1 – Středa 1.7. – Příjezd + okolí", [
        ("Po příjezdu", "Nestlerhof (7 min pěšky!) – koupit sýr, mléko, vejce"),
        ("Odpoledne",   "Procházka Filzmoos, Bachlalm + svišti (Murmeltierexpress)"),
        ("Večeře",      "De Hiatabuam nebo Café Hirschenau"),
    ]),
    ("Den 2 – Čtvrtek 2.7. – Werfen kombo", [
        ("9:00",  "Burg Hohenwerfen – ukázka sokolnictví v 11h!"),
        ("Oběd",  "Gasthof Stegerwirt Werfen (~€15/os.)"),
        ("13:30", "Eisriesenwelt – slot REZERVOVAT ONLINE PŘEDEM!"),
        ("17:00", "Zpět do Filzmoos (~70 km okruh)"),
    ]),
    ("Den 3 – Pátek 3.7. – Grossglockner (motorkáři)", [
        ("9:00",      "Motorky vyjíždějí na Grossglockner (Fusch vstup, €36,50/moto)"),
        ("Paralelně", "Auto s dítětem: Liechtensteinklamm nebo volný den"),
        ("Oběd",      "Restaurace Kaiser-Franz-Josefs-Höhe nebo Edelweißhütte"),
        ("~17h",      "Návrat do Filzmoos"),
    ]),
    ("Den 4 – Sobota 4.7. – Soutěsky + farmy", [
        ("9:00",      "Liechtensteinklamm (přijet brzy!)"),
        ("Oběd",      "Gasthof v Gollingu"),
        ("Odpoledne", "Bluntautal + Gollinger Wasserfall (kočárek OK!)"),
        ("Večeře",    "Kirchgasshütte / Aualm – Brettljause na almě!"),
    ]),
    ("Den 5 – Neděle 5.7. – Salzburg nebo Krimmler", [
        ("Varianta A", "Salzburg – Zoo Hellbrunn + Mirabellgarten + Altstadt (kočárek OK!)"),
        ("Varianta B", "Krimmler Wasserfälle + Gerlospass (motorkáři)"),
        ("Večeře",     "Panorama Steakrestaurant Reithof – závěrečná večeře s výhledem"),
    ]),
    ("Den 6 – Pondělí 6.7. – Odjezd", [
        ("Ráno",     "Kirchgasshof hofladen – nákup sýrů na cestu (otevřeno do 22h)"),
        ("Dopoledne","Dachstein-Destillerie Mandlberggut (10 km) – schnapps na cestu 😄"),
        ("Poté",     "Odjezd domů"),
    ]),
]

for day_title, items in days:
    heading(day_title, level=2, color=GREEN)
    add_table(
        ["Čas", "Aktivita"],
        items,
        col_widths=[3, 13]
    )

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════════
#  10. DÍTĚ
# ═══════════════════════════════════════════════════════════════════════════════

heading("👶  TIPY PRO CESTOVÁNÍ S DÍTĚTEM 1,5 ROKU")

heading("Co vzít s sebou", level=2, color=GREEN_LIGHT)
items_to_bring = [
    "✅  Ergonomické nosítko / šátek – nutné pro soutěsky a jeskyně",
    "✅  Terénní kočárek – Bluntautal, Salzburg Zoo, Filzmoos vesnice",
    "✅  ZIMNÍ OBLEČENÍ pro dítě (kombinéza, čepice, rukavice) – Eisriesenwelt a Dachstein Eispalast jsou pod 0°C!",
    "✅  Pláštěnka pro dítě – Krimmler, Liechtensteinklamm, Salzachöfen",
    "✅  Opalovací krém SPF50+ – ve výšce 2 500 m je UV záření 2× silnější",
]
for item in items_to_bring:
    body(item, size=10)

heading("Kočárek vs nosítko – přehled", level=2, color=GREEN_LIGHT)
add_table(
    ["Místo", "Kočárek", "Nosítko", "Poznámka"],
    [
        ["Bluntautal",              "✅ celá trasa", "✅", "Nejlepší výlet s kočárkem"],
        ["Bachlalm / Filzmoos",     "✅",            "✅", ""],
        ["Salzburg Zoo Hellbrunn",  "✅",            "✅", "Ideální pro dítě"],
        ["Grossglockner parkoviště","✅",            "✅", "Zpevněné plochy OK"],
        ["Liechtensteinklamm",      "❌",            "✅ nutné", "900 schodů"],
        ["Eisriesenwelt",           "❌",            "✅ + zimní oblečení!", "Pod 0°C celou hodinu"],
        ["Dachstein Eispalast",     "❌",            "✅ + zimní oblečení!", "Pod 0°C"],
        ["Burg Hohenwerfen",        "❌",            "✅",        "Schodiště"],
        ["Salzachöfen",             "❌",            "✅",        "Schody"],
    ],
    col_widths=[4.5, 2.5, 3, 6]
)

heading("Kde bude dítě nejvíce nadšené", level=2, color=GREEN_LIGHT)
best_for_kid = [
    "1.  🦫  Bachlalm – svišti živě, jezírko, klid, kočárek OK",
    "2.  🦅  Burg Hohenwerfen – sokolníci s dravými ptáky",
    "3.  🦒  Zoo Hellbrunn Salzburg – zvířata, kočárek OK",
    "4.  💧  Krimmler Wasserfälle (spodní stupeň) – obrovský vodopád, nezapomenutelný zvuk",
]
for item in best_for_kid:
    body(item, size=10.5)

# ═══════════════════════════════════════════════════════════════════════════════
#  11. LINKY
# ═══════════════════════════════════════════════════════════════════════════════

doc.add_page_break()
heading("🔗  DŮLEŽITÉ LINKY")
add_table(
    ["Atrakce / Zdroj", "URL"],
    [
        ["Filzmoos Tourismus",              "www.filzmoos.at"],
        ["Eisriesenwelt (REZERVACE!)",       "www.eisriesenwelt.at/de/"],
        ["Dachstein (REZERVACE!)",           "tickets.derdachstein.at"],
        ["Grossglockner ceny + stav silnice","www.grossglockner.at/de/fuer-ihren-besuch/preise-oeffnungszeiten"],
        ["Liechtensteinklamm",               "www.liechtensteinklamm.at"],
        ["Burg Hohenwerfen",                 "www.salzburg-burgen.at/de/hohenwerfen/"],
        ["Krimmler Wasserfälle",             "www.krimmler-wasserwelten.at"],
        ["Salzwelten (solné doly)",          "www.salzwelten.at"],
        ["SalzburgerLand Card",              "www.salzburgerland.com/de/salzburgerland-card/"],
        ["AlmCard Filzmoos",                 "www.filzmoos.at/de/services/almcard-filzmoos.html"],
        ["Počasí Filzmoos",                  "www.filzmoos.at/de/wetter-webcams.html"],
        ["Dachstein-Destillerie Mandlberggut","Mandlbergweg 11, Mandling (10 km od Filzmoos)"],
    ],
    col_widths=[5.5, 10.5]
)

# Footer note
doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("⚠️  Ceny ověřeny červen 2026 z oficiálních webů. Vždy ověřit aktuálnost před návštěvou.")
r.font.size = Pt(9)
r.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
r.italic = True

p2 = doc.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
r2 = p2.add_run("Filzmoos, Salzburger Land, Pongau, Rakousko  |  červen 2026")
r2.font.size = Pt(9)
r2.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
r2.italic = True

# ── Save ──────────────────────────────────────────────────────────────────────
out = "/Users/I550836/Library/CloudStorage/OneDrive-SAPSE/Osobni/Dovolena/Rakousko/2026/Filzmoos/pruvodce.docx"
doc.save(out)
print(f"✅  Uloženo: {out}")

"""Gera apresentacao-resultados-2026.pptx (para importar no Google Slides).

Mesmos dados do bloco DATA em apresentacao-resultados-2026.html.
Uso: python3 scripts/build_pptx.py
"""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ---------- dados ----------
MESES = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov*']
FAT = [49005, 37889, 26351, 55342, 56606, 28239, 37538, 40693, 45143, 50456, 30741]
INV = [2467, 1978, 2729, 2846, 3435, 3918, 2597, 2951, 4925, 5793, 4065]
PED = [15, 13, 11, 13, 23, 13, 10, 18, 24, 21, 16]
BENCH = (5, 10)
KW = {
    'Prata': [('joias em prata', 1600, 2.37), ('joias de prata', 2400, 1.94),
              ('prata 925', 49500, 1.65), ('anel de prata', 18100, 1.14)],
    'Ouro': [('joias em ouro', 2900, 4.25), ('joias de ouro', 2400, 4.06),
             ('ouro 18k', 8100, 3.87), ('anel de ouro', 18100, 1.31)],
}

fat_tot, inv_tot = sum(FAT), sum(INV)
roas_ano = fat_tot / inv_tot
roas_mes = [f / i for f, i in zip(FAT, INV)]
R = [7, 8, 9, 10]
fat_r, inv_r, ped_r = (sum(x[i] for i in R) for x in (FAT, INV, PED))
vol = {g: sum(k[1] for k in v) for g, v in KW.items()}
cpc = {g: sum(k[1] * k[2] for k in v) / vol[g] for g, v in KW.items()}


def br(n, d=0):
    s = f'{n:,.{d}f}'
    return s.replace(',', 'X').replace('.', ',').replace('X', '.')


def k(v):
    return 'R$ ' + br(v / 1000, 1 if v < 100000 else 0) + 'k'


# ---------- estilo ----------
BG = RGBColor(0x0D, 0x0B, 0x0B)
SURF = RGBColor(0x16, 0x12, 0x12)
RED = RGBColor(0xE5, 0x09, 0x14)
H = RGBColor(0xF4, 0xEF, 0xEC)
TXT = RGBColor(0xB5, 0xAA, 0xA7)
MUT = RGBColor(0x6A, 0x62, 0x60)
LINE = RGBColor(0x2A, 0x24, 0x24)
OK = RGBColor(0x52, 0xB7, 0x88)
PEND = RGBColor(0x5E, 0xA4, 0xD4)
HEAD, BODY = 'Oswald', 'Inter'

prs = Presentation()
prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)  # 16:9
PX = 12192000 / 1920  # 1 px do HTML (1920x1080) em EMU


def px(v):
    return Emu(int(v * PX))


def slide():
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = BG
    return s


def rect(s, x, y, w, h, color, line=None):
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, px(x), px(y), px(w), px(h))
    r.fill.solid()
    r.fill.fore_color.rgb = color
    if line:
        r.line.color.rgb = line
        r.line.width = Pt(0.75)
    else:
        r.line.fill.background()
    r.shadow.inherit = False
    return r


def text(s, x, y, w, h, runs, size=20, font=BODY, color=TXT, bold=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=None, upper=False):
    """runs: str ou lista de parágrafos; cada parágrafo é str ou lista de (texto, {opções})."""
    tb = s.shapes.add_textbox(px(x), px(y), px(w), px(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    paras = runs if isinstance(runs, list) else [runs]
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if spacing:
            p.space_after = Pt(spacing)
        for t, o in ([(para, {})] if isinstance(para, str) else para):
            r = p.add_run()
            r.text = t.upper() if upper else t
            f = r.font
            f.name = o.get('font', font)
            f.size = Pt(o.get('size', size) * 0.5)  # 1920px → 960pt: px * 0.5 = pt
            f.bold = o.get('bold', bold)
            f.color.rgb = o.get('color', color)
    return tb


def header(s, kicker, title, n):
    text(s, 120, 88, 1600, 40, kicker, size=22, font=HEAD, color=RED, upper=True)
    text(s, 120, 130, 1700, 90, title, size=64, font=HEAD, color=H, bold=True, upper=True)
    text(s, 120, 1010, 900, 30, 'Official Time · Resultados 2026', size=16, color=MUT, upper=True)
    text(s, 1700, 1010, 100, 30, f'{n:02d}', size=16, color=H, bold=True, align=PP_ALIGN.RIGHT)


def side(s, label, value, sub, mini=None):
    rect(s, 1340, 250, 460, 700, SURF)
    rect(s, 1340, 250, 6, 700, RED)
    paras = [[(label.upper(), {'size': 20, 'color': MUT})],
             [(value, {'size': 118, 'font': HEAD, 'bold': True, 'color': RED})],
             [(sub, {'size': 22})]]
    if mini:
        paras += [[('', {'size': 12})]] + mini
    text(s, 1390, 250, 380, 700, paras, anchor=MSO_ANCHOR.MIDDLE, spacing=8)


def style_chart(ch):
    ch.font.name = BODY
    ch.font.size = Pt(10)
    ch.font.color.rgb = TXT
    ch.has_legend = False
    for ax in (ch.category_axis, ch.value_axis):
        ax.format.line.color.rgb = LINE
        ax.tick_labels.font.color.rgb = TXT
        ax.tick_labels.font.size = Pt(10)
    ch.value_axis.has_major_gridlines = True
    ch.value_axis.major_gridlines.format.line.color.rgb = LINE


def legend(s, items, y=270):
    x = 150
    for label, color, is_line in items:
        rect(s, x, y + (8 if is_line else 2), 18, 4 if is_line else 16, color)
        text(s, x + 28, y - 4, 700, 30, label, size=20)
        x += 60 + len(label) * 11


def point(s, x, y, w, title, body):
    text(s, x, y, 40, 40, '→', size=36, font=HEAD, color=RED, bold=True)
    text(s, x + 76, y, w - 76, 220,
         [[(title.upper(), {'font': HEAD, 'size': 32, 'color': H})], [(body, {'size': 22})]], spacing=4)


def status(s, x, y, label, color):
    r = rect(s, x, y, 30 + len(label) * 14, 44, SURF, line=color)
    text(s, x + 18, y + 9, 600, 30, '● ' + label, size=19, color=color, bold=True, upper=True)
    return r


# ---------- 01 capa ----------
s = slide()
rect(s, 0, 0, 1920, 6, RED)
text(s, 120, 330, 1600, 40, 'Official Time · Joias e Relógios', size=22, font=HEAD, color=RED, upper=True)
text(s, 120, 380, 1600, 300, ['RESULTADOS', '2026'], size=132, font=HEAD, color=H, bold=True)
rect(s, 120, 700, 120, 6, RED)
text(s, 120, 740, 1600, 40, 'Fechamento do ano  ·  Jan–Nov 2026  ·  V4 Carvalho', size=26)

# ---------- 02 resultado do ano ----------
s = slide()
header(s, 'Resultado do ano · Jan–Nov 2026', 'O ano em dois números', 2)
for i, (lab, val, sub) in enumerate([
        ('Faturamento total', k(fat_tot), 'E-commerce · somente pedidos aprovados · fonte: Tray'),
        ('ROAS total do projeto', br(roas_ano, 1) + '×',
         f'Faturamento ÷ {k(inv_tot)} investidos em Meta + Google')]):
    x = 120 + i * 864
    rect(s, x, 270, 816, 680, SURF)
    rect(s, x, 270, 816, 6, RED)
    text(s, x + 64, 270, 700, 680,
         [[(lab.upper(), {'size': 24, 'color': MUT})],
          [(val, {'size': 150, 'font': HEAD, 'bold': True, 'color': H})],
          [(sub, {'size': 26})]], anchor=MSO_ANCHOR.MIDDLE, spacing=6)

# ---------- 03 ROAS vs benchmark ----------
s = slide()
header(s, 'ROAS 2026 · mês a mês', 'Acima do benchmark do segmento', 3)
rect(s, 120, 250, 1180, 700, SURF)
legend(s, [('ROAS do mês (faturamento ÷ investimento) · benchmark do segmento: 5× a 10×', RED, False)])
cd = CategoryChartData()
cd.categories = MESES
cd.add_series('ROAS', [round(v, 1) for v in roas_mes])
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, px(140), px(320), px(1140), px(610), cd)
ch = gf.chart
style_chart(ch)
ch.has_title = False
ch.value_axis.maximum_scale = 22
ch.value_axis.minimum_scale = 0
ch.value_axis.major_unit = 5
ch.value_axis.tick_labels.number_format = '0"×"'
ch.value_axis.tick_labels.number_format_is_linked = False
pl = ch.plots[0]
pl.gap_width = 70
pl.has_data_labels = True
pl.data_labels.number_format = '0.0"×"'
pl.data_labels.number_format_is_linked = False
pl.data_labels.font.size = Pt(9)
pl.data_labels.font.color.rgb = H
ser = pl.series[0]
ser.format.fill.solid()
ser.format.fill.fore_color.rgb = RED
acima = sum(1 for v in roas_mes if v > BENCH[1])
side(s, 'Acima do teto do benchmark', f'+{round((roas_ano / BENCH[1] - 1) * 100)}%',
     f'ROAS de {br(roas_ano, 1)}× em 2026 contra 5–10× do segmento de joias e relógios.',
     [[(f'{acima} de {len(roas_mes)} meses acima de 10×', {'size': 22, 'color': H})],
      [('Benchmark: WordStream Industry Benchmarks 2024', {'size': 16, 'color': MUT})]])

# ---------- 04 recorte ago–nov ----------
s = slide()
header(s, 'Recorte recente · Ago–Nov 2026', 'Faturamento x investimento', 4)
rect(s, 120, 250, 1180, 700, SURF)
legend(s, [('Faturamento', RED, False), ('Investimento (Meta + Google)', H, False)])
cd = CategoryChartData()
cd.categories = ['Ago', 'Set', 'Out', 'Nov*']
cd.add_series('Faturamento', [FAT[i] for i in R])
cd.add_series('Investimento', [INV[i] for i in R])
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, px(140), px(320), px(1140), px(610), cd)
ch = gf.chart
style_chart(ch)
ch.value_axis.tick_labels.number_format = '"R$ "#,##0'
ch.value_axis.tick_labels.number_format_is_linked = False
pl = ch.plots[0]
pl.gap_width = 60
pl.overlap = -10
pl.has_data_labels = True
pl.data_labels.number_format = '"R$ "#,##0'
pl.data_labels.number_format_is_linked = False
pl.data_labels.font.size = Pt(9)
pl.data_labels.font.color.rgb = H
for sr, c in zip(pl.series, (RED, H)):
    sr.format.fill.solid()
    sr.format.fill.fore_color.rgb = c
side(s, 'ROAS médio do período', br(fat_r / inv_r, 1) + '×',
     'Agosto, setembro e outubro fechados + novembro parcial (até 24/11).',
     [[('Faturado: ', {'size': 20}), (k(fat_r), {'size': 26, 'font': HEAD, 'color': H})],
      [('Investido: ', {'size': 20}), (k(inv_r), {'size': 26, 'font': HEAD, 'color': H})],
      [('Pedidos: ', {'size': 20}), (str(ped_r), {'size': 26, 'font': HEAD, 'color': H}),
       ('  ·  Ticket: ', {'size': 20}), ('R$ ' + br(fat_r / ped_r), {'size': 26, 'font': HEAD, 'color': H})]])

# ---------- 05 API de conversão ----------
s = slide()
header(s, 'Entregue · Novembro 2026', 'API de Conversão Tray → Meta', 5)
point(s, 120, 280, 780, 'O problema', 'O Meta Ads era o canal com mais compras canceladas por tentativa de golpe. '
      'O pixel contava esses pedidos como venda, e o algoritmo aprendia com compras que não existiram.')
point(s, 120, 500, 780, 'O que fizemos', 'Configuramos a API de Conversão da Tray, que devolve ao Meta a informação '
      'direto do servidor da loja: só a compra real, aprovada.')
point(s, 120, 700, 780, 'O que muda', 'Dado mais limpo significa otimização mais precisa: o Meta passa a buscar '
      'perfis parecidos com quem de fato compra, não com quem tenta fraudar.')
rect(s, 1000, 270, 800, 680, SURF)
status(s, 1048, 318, 'Concluído · V4', OK)
text(s, 1048, 400, 700, 500, [
    [('A demanda estava com o ', {}), ('time de tecnologia da Official Time', {'bold': True, 'color': H}),
     (' e não avançou. A ', {}), ('V4 assumiu e concluiu', {'bold': True, 'color': H}),
     (': API configurada e ativa desde a semana passada.', {})],
    [('Próximo passo: ', {'bold': True, 'color': H}),
     ('acompanhar nas próximas semanas a diferença entre compras reportadas pelo Meta e pedidos aprovados na loja.', {})]],
    size=24, spacing=14)

# ---------- 06 próximos passos ----------
s = slide()
text(s, 120, 300, 1600, 40, '2027', size=22, font=HEAD, color=RED)
text(s, 120, 345, 1600, 90, 'PRÓXIMOS PASSOS', size=64, font=HEAD, color=H, bold=True)
for i, (t, sub) in enumerate([('Joias em prata no catálogo', 'abrir uma nova frente de demanda'),
                              ('WhatsApp de atendimento das lojas', 'levar o digital para o PDV')]):
    y = 500 + i * 130
    text(s, 120, y, 1300, 80, [[(f'0{i + 1}   ', {'size': 32, 'color': RED, 'bold': True}),
                                (t.upper(), {'size': 50, 'color': H})]], font=HEAD, anchor=MSO_ANCHOR.BOTTOM)
    text(s, 1300, y + 30, 500, 50, sub, size=24, color=MUT, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.BOTTOM)
    rect(s, 120, y + 100, 1680, 2, LINE)
text(s, 120, 1010, 900, 30, 'Official Time · Resultados 2026', size=16, color=MUT, upper=True)
text(s, 1700, 1010, 100, 30, '06', size=16, color=H, bold=True, align=PP_ALIGN.RIGHT)

# ---------- 07 joias em prata ----------
s = slide()
header(s, 'Próximo passo 01', 'Joias em prata no catálogo', 7)
rect(s, 120, 250, 900, 720, SURF)
rows = [('PALAVRA-CHAVE', 'BUSCAS/MÊS', 'CPC MÉDIO', 'th')]
for g, items in KW.items():
    rows.append((g.upper(), '', '', 'grp'))
    rows += [(a, br(b), 'R$ ' + br(c, 2), 'td') for a, b, c in items]
    rows.append((f'Total {g.lower()}', br(vol[g]), 'R$ ' + br(cpc[g], 2), 'tot'))
for r, (a, b, c, kind) in enumerate(rows):
    y = 280 + r * 44
    o = {'size': 18 if kind == 'th' else 21, 'font': HEAD if kind == 'grp' else BODY,
         'bold': kind == 'tot', 'color': {'th': MUT, 'grp': RED}.get(kind, H)}
    text(s, 160, y, 440, 40, [[(a, o)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, 600, y, 180, 40, [[(b, o)]], align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 800, y, 180, 40, [[(c, o)]], align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    if kind in ('th', 'td', 'tot'):
        rect(s, 160, y + 42, 820, 2 if kind == 'tot' else 1, H if kind == 'tot' else LINE)
text(s, 160, 900, 820, 50, 'Google Ads, Planejador de Palavras-chave · Brasil · média dos últimos 12 meses · '
     'total com CPC ponderado', size=15, color=MUT)
point(s, 1100, 270, 700, 'A demanda já existe',
      f'Prata é buscada {br(vol["Prata"] / vol["Ouro"], 1)}× mais que ouro no Google, com clique mais barato.')
point(s, 1100, 480, 700, 'Ticket de entrada', 'Peças em prata abrem a marca para um público que ainda não compra '
      'relógio ou ouro, e alimentam a base para recompra.')
point(s, 1100, 700, 700, 'Mídia mais barata', f'O clique em prata custa {round((1 - cpc["Prata"] / cpc["Ouro"]) * 100)}% '
      'menos que em ouro. Com a mesma verba, isso traz mais tráfego qualificado para o site.')

# ---------- 08 whatsapp ----------
s = slide()
header(s, 'Próximo passo 02', 'WhatsApp de atendimento das lojas', 8)
point(s, 120, 280, 780, 'O cliente já está lá', 'Hoje o primeiro contato é digital. Quem mora na região quer tirar '
      'dúvida, ver a peça e reservar pelo WhatsApp antes de ir à loja, com atendimento individual.')
point(s, 120, 500, 780, 'Ponte para o PDV', 'Um número por loja, divulgado nos anúncios e no Google Meu Negócio. '
      'O primeiro atendimento acontece no WhatsApp, com o objetivo de levar o cliente até a loja.')
point(s, 120, 720, 780, 'Venda que hoje não medimos', 'Com campanhas de clique para WhatsApp, passamos a enxergar e '
      'otimizar a venda feita na loja física, não só no e-commerce.')
rect(s, 1000, 270, 800, 680, SURF)
status(s, 1048, 318, 'Proposta · decisão Official Time', RED)
text(s, 1048, 400, 700, 520, [
    [('3 lojas', {'font': HEAD, 'size': 96, 'bold': True, 'color': H})],
    [('em Jundiaí, cada uma com seu WhatsApp', {})],
    [('2 lojas', {'bold': True, 'color': H}), (' no Maxi Shopping  ·  ', {}),
     ('1 loja', {'bold': True, 'color': H}), (' no Jundiaí Shopping', {})],
    [('O que precisamos: ', {'bold': True, 'color': H}),
     ('número dedicado por loja e uma pessoa responsável pelo atendimento em cada uma, com tempo de resposta combinado.', {})]],
    size=24, spacing=14)

out = 'apresentacao-resultados-2026.pptx'
prs.save(out)
print('ok', out, f'faturamento {fat_tot} investimento {inv_tot} roas {roas_ano:.2f} recorte {fat_r / inv_r:.2f}')

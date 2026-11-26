"""Gera relatorio-2026.pptx — versão em slides de relatorio-2026/index.html
(official-time-nov-26.vercel.app), um slide por bloco, para importar no Google Slides.

Uso: python3 scripts/build_relatorio_pptx.py
"""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ---------------- dados (mesmos do relatório) ----------------
MESES = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov*']
FAT = [49005, 37889, 26351, 55342, 56606, 28239, 37538, 40693, 45143, 50456, 30741]
INV = [2467, 1978, 2729, 2846, 3435, 3918, 2597, 2951, 4925, 5793, 4065]
PED = [15, 13, 11, 13, 23, 13, 10, 18, 24, 21, 16]
FT, IT, PT = sum(FAT), sum(INV), sum(PED)
ROAS = FT / IT
ROAS_MES = [f / i for f, i in zip(FAT, INV)]

DECAY = [(0.10, 0, 0.03), (0.20, 0.03, 0.06), (0.30, 0.06, 0.10), (0.50, 0.10, 0.15),
         (1.00, 0.15, 0.28), (2.00, 0.28, 0.45), (4.00, 0.45, 0.60), (9e9, 0.60, 0.75)]
SIM_BASE, SIM_ROAS = 4900, (7, 8.5, 10)


def sim(inv):
    pct = (inv - SIM_BASE) / SIM_BASE
    if abs(pct) <= 0.005:
        return SIM_ROAS
    lo, hi = next((a, b) for m, a, b in DECAY if pct <= m)
    return (SIM_ROAS[0] * (1 - hi), SIM_ROAS[1] * (1 - (lo + hi) / 2), SIM_ROAS[2] * (1 - lo))


DOSSIE = {
    'E-commerce': [
        ('Mídia Meta + Google', 'Escalar tráfego e vendas com mensuração contínua', 'Contínuo', 'done', 'Executada'),
        ('API de Conversão Tray → Meta', 'Envio server-side só das compras aprovadas', 'Nov 26', 'done', 'Executada'),
        ('Campanhas sazonais', 'Datas comemorativas com criativos integrados', 'Mar – Ago 26', 'done', 'Executada'),
        ('Régua de CRM', 'Boas-vindas, carrinho, base e aniversários', 'Jun – Ago 26', 'lim', 'Aprovada c/ limitação'),
        ('CRM + reativação da base', 'Disparo para base com oferta 5% + mote emocional', 'Jul – Ago 26', 'lim', 'Aprovada c/ limitação'),
        ('Ajustes estruturais no site', 'Hierarquia, banners, links internos e mobile', '—', 'lim', 'Aprovada c/ limitação'),
        ('Planejamento Antecipado de Black Friday', 'Estrutura integrada PDV + e-com com antecedência', 'Set – Nov 26', 'done', 'Realizado'),
        ('31 anos + Mês do Cliente', 'KV aprovado e campanha veiculada no e-com', 'Ago – Set 26', 'done', 'Executada'),
        ('Migração para Shopify', 'Plataforma com mais flexibilidade de CRO e integrações', 'Fev 26', 'no', 'Não aprovada'),
        ('CRO — otimização de conversão', 'Clareza da oferta, checkout e redução de fricções', 'Jan + Jun 26', 'no', 'Reprovada'),
        ('Expansão com influenciadores', 'UGC, microinfluenciadores e roteiros de performance', 'Mai – Jun 26', 'no', 'Reprovada'),
        ('Novas captações próprias', 'Rotina de vídeos e fotos para criativos de performance', '—', 'no', 'Reprovada'),
    ],
    'PDV': [
        ('Diagnóstico e planejamento', 'Posicionamento, canais e prioridades iniciais', 'Out – Nov 25', 'done', 'Executada'),
        ('31 anos + Mês do Cliente', 'KV aprovado e campanha veiculada', 'Ago – Set 26', 'done', 'Executada'),
        ('Planejamento Antecipado de Black Friday', 'Criativos físicos e ativação no PDV com antecedência', 'Set – Nov 26', 'done', 'Realizado'),
        ('Adesivação de vitrines', 'Aplicação do KV de 31 anos no ponto de venda', 'Ago 26', 'done', 'Executada'),
        ('Quadro Permanente de Vitrine', 'Painel físico substituível para o ponto de venda', '—', 'pend', 'Buscando fornecedor'),
        ('Ações para alianças — Mês das Noivas', 'Foco em PDV com KV aprovado e depois reprovado', 'Mai 26', 'no', 'Reprovada'),
        ('Temporada Dourada — Troca do Ouro', 'Ouro antigo como crédito para nova joia', 'Abr – Mai 26', 'no', 'Reprovada'),
        ('Brinde — Dia da Mulher', 'Vela artesanal acima de valor mínimo de compra', 'Mar 26', 'no', 'Reprovada'),
        ('Brinde — Dia das Mães', 'Pingente de prata (formato infinito)', 'Mai 26', 'no', 'Reprovada'),
        ('Mídia de Shopping', 'Display e ativações no shopping center', '—', 'no', 'Reprovada'),
        ('Mídia off', 'Ativações de marca em canais offline', '—', 'no', 'Reprovada'),
    ],
    'B2B': [
        ('Landing Page de Vendas Corporativas', 'Captação de clientes PJ e presentes corporativos', 'Dez 25 – Jan 26', 'no', 'Não aprovada'),
        ('Abertura de frente B2B em Jundiaí', 'Estrutura de vendas corporativas local', 'Jul 26', 'no', 'Reprovada'),
        ('Mala direta corporativa', 'Prospecção com material físico para empresas', 'Jul 26', 'no', 'Reprovada'),
        ('Lista ICP + prospecção fria', 'Lista de empresas de Jundiaí no perfil ideal', 'Jul 26', 'no', 'Reprovada'),
        ('Diagnóstico gratuito como porta de entrada', 'Análise de marketing para empresas ICP', 'Jul 26', 'no', 'Reprovada'),
    ],
}

CAL = [
    ('Jan', 'Outlet no Site', 'R$49k no 1º mês do ano.', 'done', 'Concluído'),
    ('Fev', 'Outlet no Site', 'Faturamento: R$37,9k.', 'done', 'Concluído'),
    ('Mar', 'Dia da Mulher', 'Campanhas sazonais para o Mês da Mulher.', 'done', 'Concluído'),
    ('Abr', 'Pico de Receita', 'Total do mês: R$55,3k.', 'key', 'Pico do período'),
    ('Mai', 'Dia das Mães', 'Melhor mês: R$56,6k e 23 pedidos.', 'key', 'Melhor mês'),
    ('Jun', 'Namorados + CRM', 'Início da régua de CRM: boas-vindas e carrinho.', 'done', 'Concluído'),
    ('Jul', 'Reativação de Base', 'CRM com oferta de 5% para base inativa.', 'done', 'Concluído'),
    ('Ago', 'Dia dos Pais', 'Faturamento: R$40,7k · ROAS 13,8×.', 'done', 'Concluído'),
    ('Set', 'Aniversário da Loja', '31 anos — PDV + e-com. Faturamento: R$45,1k.', 'key', 'Destaque'),
    ('Out', 'Pré-Black Friday', 'Criativos, ofertas e vitrine. Faturamento: R$50,5k.', 'done', 'Realizado'),
    ('Nov', 'Black Friday', 'Campanha integrada PDV + e-com. Parcial até 24/11: R$30,7k.', 'done', 'Realizado'),
    ('Dez', 'Natal', 'Campanha de Natal integrada PDV + e-com.', 'exec', 'Em andamento'),
]

KW = {
    'Prata': [('joias em prata', 1600, 2.37), ('joias de prata', 2400, 1.94),
              ('prata 925', 49500, 1.65), ('anel de prata', 18100, 1.14)],
    'Ouro': [('joias em ouro', 2900, 4.25), ('joias de ouro', 2400, 4.06),
             ('ouro 18k', 8100, 3.87), ('anel de ouro', 18100, 1.31)],
}
VOL = {g: sum(k[1] for k in v) for g, v in KW.items()}
CPC = {g: sum(k[1] * k[2] for k in v) / VOL[g] for g, v in KW.items()}


def br(n, d=0):
    return f'{n:,.{d}f}'.replace(',', 'X').replace('.', ',').replace('X', '.')


# ---------------- estilo (tema claro do relatório) ----------------
BG = RGBColor(0xF5, 0xF2, 0xEF)
SURF = RGBColor(0xFF, 0xFF, 0xFF)
SURF2 = RGBColor(0xED, 0xE9, 0xE5)
RED = RGBColor(0xB2, 0x07, 0x10)
REDDIM = RGBColor(0xF0, 0xD6, 0xD7)
H = RGBColor(0x17, 0x0F, 0x0E)
TXT = RGBColor(0x5A, 0x52, 0x50)
MUT = RGBColor(0x9A, 0x92, 0x90)
LINE = RGBColor(0xE2, 0xDC, 0xD8)
STATUS = {  # (fundo, texto)
    'done': (RGBColor(0xD4, 0xEF, 0xE2), RGBColor(0x1A, 0x5E, 0x38)),
    'lim': (RGBColor(0xFF, 0xF9, 0xC4), RGBColor(0x8B, 0x60, 0x00)),
    'no': (RGBColor(0xFC, 0xDC, 0xDC), RGBColor(0x8B, 0x2B, 0x2B)),
    'pend': (RGBColor(0xD0, 0xE4, 0xF8), RGBColor(0x1A, 0x48, 0x78)),
    'exec': (RGBColor(0xDC, 0xFF, 0xCF), RGBColor(0x1A, 0x60, 0x10)),
    'key': (RGBColor(0xFC, 0xDC, 0xDC), RED),
}
ICON = {'done': '✓', 'lim': '≈', 'no': '✗', 'pend': '?', 'exec': '▶', 'key': '★'}
HEAD, BODY = 'Oswald', 'Inter'

prs = Presentation()
prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
PXE = 12192000 / 1920


def px(v):
    return Emu(int(v * PXE))


def slide():
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = BG
    return s


def rect(s, x, y, w, h, color, line=None, shape=MSO_SHAPE.RECTANGLE):
    r = s.shapes.add_shape(shape, px(x), px(y), px(w), px(h))
    r.fill.solid()
    r.fill.fore_color.rgb = color
    if line:
        r.line.color.rgb = line
        r.line.width = Pt(0.75)
    else:
        r.line.fill.background()
    r.shadow.inherit = False
    return r


def text(s, x, y, w, h, paras, size=20, font=BODY, color=TXT, bold=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=None, upper=False):
    tb = s.shapes.add_textbox(px(x), px(y), px(w), px(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    paras = paras if isinstance(paras, list) else [paras]
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if spacing:
            p.space_after = Pt(spacing)
        for t, o in ([(para, {})] if isinstance(para, str) else para):
            r = p.add_run()
            r.text = t.upper() if (upper or o.get('upper')) else t
            f = r.font
            f.name = o.get('font', font)
            f.size = Pt(o.get('size', size) * 0.5)
            f.bold = o.get('bold', bold)
            f.color.rgb = o.get('color', color)
    return tb


N = [0]


def header(s, eyebrow, title, desc=None):
    N[0] += 1
    text(s, 100, 70, 1100, 34, eyebrow, size=20, color=RED, bold=True, upper=True)
    text(s, 100, 104, 1100, 90, title, size=60, font=HEAD, color=H, bold=True)
    if desc:
        text(s, 1240, 80, 580, 140, desc, size=21, anchor=MSO_ANCHOR.MIDDLE)
    rect(s, 100, 215, 1720, 2, LINE)
    text(s, 100, 1020, 900, 30, 'Official Time · Relatório Executivo · Jan–Nov 2026', size=15, color=MUT, upper=True)
    text(s, 1720, 1020, 100, 30, f'{N[0]:02d}', size=15, color=H, bold=True, align=PP_ALIGN.RIGHT)


def card(s, x, y, w, h, accent_left=False, accent_top=False):
    rect(s, x, y, w, h, SURF, line=LINE)
    if accent_left:
        rect(s, x, y, 6, h, RED)
    if accent_top:
        rect(s, x, y, w, 6, RED)


def pill(s, x, y, kind, label, size=16):
    bg, fg = STATUS[kind]
    w = 36 + len(label) * size * 0.55
    r = rect(s, x, y, w, size * 2, bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    r.adjustments[0] = 0.2
    text(s, x, y, w, size * 2, f'{ICON[kind]} {label}', size=size, color=fg, bold=True,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def kpi(s, x, y, w, h, eyebrow, val, note, hi=False):
    card(s, x, y, w, h, accent_left=hi)
    text(s, x + 40, y, w - 70, h, [
        [(eyebrow.upper(), {'size': 18, 'color': MUT, 'bold': True})],
        [(val, {'size': 84, 'font': HEAD, 'bold': True, 'color': RED if hi else H})],
        [(note, {'size': 20, 'color': RED if hi else MUT, 'bold': hi})]], anchor=MSO_ANCHOR.MIDDLE, spacing=4)


def style_chart(ch):
    ch.font.name = BODY
    ch.font.size = Pt(10)
    ch.font.color.rgb = TXT
    ch.has_legend = False
    ch.has_title = False
    for ax in (ch.category_axis, ch.value_axis):
        ax.format.line.color.rgb = LINE
        ax.tick_labels.font.color.rgb = TXT
        ax.tick_labels.font.size = Pt(10)
    ch.value_axis.has_major_gridlines = True
    ch.value_axis.major_gridlines.format.line.color.rgb = LINE


def bar_colors(series, values, hi_fn):
    for i, v in enumerate(values):
        pt = series.points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = RED if hi_fn(i, v) else REDDIM


# ================= 01 CAPA =================
s = slide()
rect(s, 1180, 0, 740, 1080, H)
text(s, 1180, 0, 740, 1080, 'OT', size=420, font=HEAD, color=RGBColor(0x2A, 0x22, 0x21), bold=True,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, 100, 150, 1000, 34, 'V4 Carvalho · Relatório Executivo', size=22, color=RED, bold=True, upper=True)
text(s, 100, 190, 1000, 380, ['Official', 'Time'], size=150, font=HEAD, color=H, bold=True)
text(s, 100, 600, 1000, 90, [[('Janeiro — Novembro 2026', {'size': 34, 'color': H})],
                             [('Performance & Plano de Ação', {'size': 24})]])
x = 100
for p_ in ['Estratégia', 'Mídia', 'CRM', 'PDV', 'B2B']:
    w = 40 + len(p_) * 13
    rect(s, x, 720, w, 40, SURF, line=LINE)
    text(s, x, 720, w, 40, p_, size=17, color=TXT, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, upper=True)
    x += w + 14
for i, (v, l) in enumerate([('R$ 458k', 'Faturamento 2026'), (br(ROAS, 1) + '×', 'ROAS real'), (str(PT), 'Pedidos 2026')]):
    xx = 100 + i * 330
    rect(s, xx, 820, 2, 130, RED if i == 0 else LINE)
    text(s, xx + 24, 820, 300, 130, [[(v, {'size': 64, 'font': HEAD, 'bold': True, 'color': H})],
                                     [(l.upper(), {'size': 16, 'color': MUT, 'bold': True})]], anchor=MSO_ANCHOR.MIDDLE)

# ================= 02 O QUE ENTREGAMOS — ROAS =================
s = slide()
header(s, 'V4 · Prova de entrega · Jan–Nov 2026', 'O que entregamos',
       f'De janeiro a novembro de 2026: R$ {br(FT / 1000)}k em pedidos aprovados no e-commerce, {PT} pedidos '
       f'e ROAS de {br(ROAS, 1)}× sobre o investimento em Meta + Google.')
card(s, 100, 260, 1720, 300, accent_top=True)
text(s, 150, 260, 520, 300, br(ROAS, 1) + '×', size=190, font=HEAD, color=RED, bold=True, anchor=MSO_ANCHOR.MIDDLE)
rect(s, 700, 320, 2, 180, LINE)
text(s, 750, 260, 1020, 300, [
    [('ROAS REAL COMBINADO · GOOGLE + META · JAN–NOV 2026', {'size': 18, 'color': MUT, 'bold': True})],
    [(f'R$ {br(IT)}', {'bold': True, 'color': H}), (' investidos → ', {}), (f'R$ {br(FT)}', {'bold': True, 'color': H}),
     (' faturados', {})],
    [('21% acima do teto do segmento', {'bold': True, 'color': RED}), (' (benchmark 5–10× · WordStream 2024)', {})]],
    size=28, anchor=MSO_ANCHOR.MIDDLE, spacing=8)
card(s, 100, 600, 1720, 380)
text(s, 150, 630, 1600, 40, [[('ROAS — OFFICIAL TIME VS BENCHMARK DO NICHO', {'bold': True, 'color': H}),
                              ('   Windsor AI + Tray · Jan–Nov 2026', {'color': MUT})]], size=18)
SCALE = 20  # 0–20×
TX, TW = 560, 1000
for i, (lbl, sub, a, b, val, hi) in enumerate([('Benchmark nicho', 'Joalheria / Relógios', 5, 10, '5 – 10×', False),
                                                ('★ Official Time', 'Real combinado', 0, ROAS, br(ROAS, 1) + '×', True)]):
    y = 710 + i * 110
    text(s, 150, y - 10, 400, 70, [[(lbl, {'size': 24, 'color': RED if hi else H})], [(sub, {'size': 18, 'color': MUT})]])
    rect(s, TX, y + 18, TW, 12, SURF2)
    rect(s, TX + TW * a / SCALE, y + 18, TW * (b - a) / SCALE, 12, RED if hi else REDDIM)
    text(s, TX + TW + 30, y, 200, 50, val, size=28, color=RED if hi else TXT, bold=hi, anchor=MSO_ANCHOR.MIDDLE)
text(s, 150, 910, 1620, 60, [[('A Official Time opera ', {}), ('21% acima do teto do segmento', {'bold': True, 'color': H}),
                              (' mesmo com o investimento escalado a partir de setembro — eficiência mantida enquanto a receita absoluta cresce.', {})]],
     size=20)

# ================= 03 KPIs =================
s = slide()
header(s, 'Resultado · Jan–Nov 2026', 'Faturamento, ticket e pedidos')
kpi(s, 100, 300, 556, 420, 'Faturamento 2026', f'R$ {br(FT / 1000)}k', 'Jan–Nov · pedidos aprovados', hi=True)
kpi(s, 682, 300, 556, 420, 'Ticket médio', f'R$ {br(FT / PT / 1000, 1)}k', f'R$ {br(FT / PT)} por pedido')
kpi(s, 1264, 300, 556, 420, 'Volume de pedidos', str(PT), 'Jan–Nov 2026')
text(s, 100, 780, 1720, 80, f'Investimento total em mídia (Meta + Google): R$ {br(IT)} · ROAS {br(ROAS, 1)}× · '
     'Fonte: Tray (pedidos aprovados) e Windsor AI', size=20, color=MUT)

# ================= 04 FATURAMENTO MÊS A MÊS =================
s = slide()
header(s, 'Resultado · Faturamento', 'Faturamento mês a mês — 2026',
       'Pedidos aprovados (Tray) · Jan–Nov · novembro parcial até 24/11. Melhor mês: maio, R$ 56,6k.')
card(s, 100, 250, 1720, 740)
cd = CategoryChartData()
cd.categories = MESES
cd.add_series('Faturamento', FAT)
ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, px(130), px(280), px(1660), px(690), cd).chart
style_chart(ch)
ch.value_axis.tick_labels.number_format = '"R$ "#,##0'
ch.value_axis.tick_labels.number_format_is_linked = False
pl = ch.plots[0]
pl.gap_width = 60
pl.has_data_labels = True
pl.data_labels.number_format = '"R$ "#,##0'
pl.data_labels.number_format_is_linked = False
pl.data_labels.font.size = Pt(10)
pl.data_labels.font.color.rgb = H
bar_colors(pl.series[0], FAT, lambda i, v: i == 4)

# ================= 05 ROAS MÊS A MÊS =================
s = slide()
header(s, 'Resultado · ROAS', 'ROAS mês a mês — 2026',
       'Faturamento ÷ investimento Meta + Google. Em vermelho, os meses acima do teto do benchmark (10×).')
card(s, 100, 250, 1180, 740)
cd = CategoryChartData()
cd.categories = MESES
cd.add_series('ROAS', [round(v, 1) for v in ROAS_MES])
ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, px(130), px(280), px(1120), px(690), cd).chart
style_chart(ch)
ch.value_axis.minimum_scale, ch.value_axis.maximum_scale, ch.value_axis.major_unit = 0, 22, 5
ch.value_axis.tick_labels.number_format = '0"×"'
ch.value_axis.tick_labels.number_format_is_linked = False
pl = ch.plots[0]
pl.gap_width = 60
pl.has_data_labels = True
pl.data_labels.number_format = '0.0"×"'
pl.data_labels.number_format_is_linked = False
pl.data_labels.font.size = Pt(10)
pl.data_labels.font.color.rgb = H
bar_colors(pl.series[0], ROAS_MES, lambda i, v: v > 10)
card(s, 1310, 250, 510, 740, accent_left=True)
acima = sum(v > 10 for v in ROAS_MES)
text(s, 1350, 250, 440, 740, [
    [('BENCHMARK DO SEGMENTO', {'size': 18, 'color': MUT, 'bold': True})],
    [('5 – 10×', {'size': 80, 'font': HEAD, 'bold': True, 'color': H})],
    [(f'{acima} de {len(ROAS_MES)} meses', {'size': 30, 'bold': True, 'color': RED}), (' acima do teto de 10× e nenhum abaixo do piso de 5×.', {'size': 22})],
    [('', {'size': 10})],
    [('A partir de setembro o investimento quase dobrou (média de R$ 4,9k/mês) e o ROAS se acomodou entre 7,6× e 9,2× — '
      'ainda dentro ou acima da faixa do segmento.', {'size': 20})]], anchor=MSO_ANCHOR.MIDDLE, spacing=10)

# ================= 06 PROPOSTA DE INVESTIMENTO (calculadora) =================
s = slide()
header(s, 'Investimento em mídia', 'Proposta de investimento em mídia',
       'Base: patamar atual de R$ 4.900/mês e ROAS 8,5× (média Set–Nov). Acima disso, a projeção aplica perda de '
       'eficiência proporcional ao aumento. Versão interativa no relatório online.')
cols = ['Investimento mensal', 'Pessimista', 'Conservador', 'Otimista']
CW = [460, 420, 420, 420]
y0 = 270
xs = [100]
for w in CW[:-1]:
    xs.append(xs[-1] + w)
rect(s, 100, y0, 1720, 70, H)
for c, x, w in zip(cols, xs, CW):
    text(s, x + 30, y0, w - 60, 70, c, size=20, color=SURF, bold=True, upper=True,
         align=PP_ALIGN.LEFT if c == cols[0] else PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
for r_, inv in enumerate([4900, 6000, 8000, 10000]):
    y = y0 + 70 + r_ * 140
    rect(s, 100, y, 1720, 140, SURF if r_ else RGBColor(0xFB, 0xEE, 0xEE), line=LINE)
    pct = (inv - SIM_BASE) / SIM_BASE
    text(s, xs[0] + 30, y, CW[0] - 40, 140, [
        [(f'R$ {br(inv)}', {'size': 40, 'font': HEAD, 'bold': True, 'color': H})],
        [('patamar atual' if r_ == 0 else f'+{round(pct * 100)}% vs. atual', {'size': 18, 'color': RED if r_ == 0 else MUT})]],
        anchor=MSO_ANCHOR.MIDDLE)
    for k, (ro, x, w) in enumerate(zip(sim(inv), xs[1:], CW[1:])):
        text(s, x, y, w, 140, [
            [(f'R$ {br(inv * ro)}/mês', {'size': 34, 'font': HEAD, 'bold': True, 'color': RED if k == 2 else H})],
            [(f'ROAS {br(ro, 1)}×  ·  ≈ R$ {br(inv * ro * 12 / 1000)}k/ano', {'size': 18, 'color': MUT})]],
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, 100, 960, 1720, 40, 'Calibrado com dados reais Windsor AI + Tray · Set–Nov 2026 · cenários: pessimista 7×, conservador 8,5×, otimista 10× no patamar atual',
     size=16, color=MUT)

# ================= 07–09 DOSSIÊ =================
for aba, rows in DOSSIE.items():
    s = slide()
    header(s, 'Dossiê de Ações · ' + aba, 'Dossiê de Ações — ' + aba,
           'Ações propostas, aprovadas e reprovadas desde o início do projeto. Status atualizado em 25/11/2026.')
    n = len(rows)
    rh = min(64, 700 // n)
    y0 = 250
    for c, x, w, al in [('Ação', 100, 900, PP_ALIGN.LEFT), ('Período', 1020, 300, PP_ALIGN.LEFT), ('Status', 1360, 460, PP_ALIGN.LEFT)]:
        text(s, x, y0, w, 36, c, size=16, color=MUT, bold=True, upper=True, align=al)
    rect(s, 100, y0 + 38, 1720, 2, H)
    for i, (nome, obj, per, kind, lab) in enumerate(rows):
        y = y0 + 48 + i * rh
        text(s, 100, y, 900, rh, [[(nome, {'size': 21, 'bold': True, 'color': H}), ('  ·  ' + obj, {'size': 17, 'color': MUT})]],
             anchor=MSO_ANCHOR.MIDDLE)
        text(s, 1020, y, 300, rh, per, size=19, anchor=MSO_ANCHOR.MIDDLE)
        pill(s, 1360, y + (rh - 32) / 2, kind, lab)
        rect(s, 100, y + rh - 1, 1720, 1, LINE)

# ================= 10 CALENDÁRIO =================
s = slide()
header(s, 'Conclusão mês a mês', 'Calendário de Campanhas — 2026',
       'Campanha principal, destaque estratégico e status de janeiro a dezembro.')
CWd, CHd = 280, 350
for i, (mo, camp, det, kind, lab) in enumerate(CAL):
    x = 100 + (i % 6) * (CWd + 8)
    y = 250 + (i // 6) * (CHd + 20)
    card(s, x, y, CWd, CHd)
    if kind in ('key', 'exec'):
        rect(s, x, y, CWd, 5, RED if kind == 'key' else STATUS['exec'][1])
    text(s, x + 22, y + 20, CWd - 44, 60, [[(mo, {'size': 44, 'font': HEAD, 'bold': True, 'color': H}),
                                            ('  2026', {'size': 15, 'color': MUT})]])
    text(s, x + 22, y + 95, CWd - 44, 40, camp, size=20, bold=True, color=H)
    text(s, x + 22, y + 140, CWd - 44, 140, det, size=17)
    pill(s, x + 22, y + CHd - 56, kind, lab, size=14)

# ================= 11 JOIAS — DESTAQUES =================
s = slide()
header(s, 'Próximo passo · Catálogo', 'Joias em Prata',
       'Incluir joias em prata no catálogo abre uma nova frente de demanda — com mais busca e clique mais barato do que ouro. '
       'Fonte: Planejador de Palavras-chave do Google Ads.')
ratio, diff = VOL['Prata'] / VOL['Ouro'], round((1 - CPC['Prata'] / CPC['Ouro']) * 100)
kpi(s, 100, 260, 556, 330, 'Buscas de prata / mês', f'{br(VOL["Prata"] / 1000, 1)}k',
    f'{br(ratio, 1)}× o volume de ouro ({br(VOL["Ouro"] / 1000, 1)}k)', hi=True)
kpi(s, 682, 260, 556, 330, 'CPC médio prata', f'R$ {br(CPC["Prata"], 2)}', f'{diff}% abaixo do ouro (R$ {br(CPC["Ouro"], 2)})')
kpi(s, 1264, 260, 556, 330, 'Termo âncora', 'prata 925', '49,5 mil buscas/mês sozinho')
for i, (tag, title, body, imp) in enumerate([
        ('Demanda', 'A demanda já existe', f'Prata é buscada {br(ratio, 1)}× mais que ouro no Google. "Prata 925" e "anel de prata" '
         'concentram quase 68 mil buscas por mês.', 'Nova frente de tráfego sem criar demanda do zero'),
        ('Público', 'Ticket de entrada', 'Peças em prata abrem a marca para um público que ainda não compra relógio ou ouro — '
         'e alimentam a base de CRM para recompra.', 'Aquisição de novos clientes com ticket menor'),
        ('Mídia', 'Clique mais barato', f'O clique em prata custa {diff}% menos que em ouro. Com a mesma verba, mais tráfego '
         'qualificado chega ao site.', 'Mais eficiência por real investido')]):
    x = 100 + i * 582
    card(s, x, 620, 556, 370)
    text(s, x + 36, 650, 484, 320, [
        [(tag.upper(), {'size': 15, 'bold': True, 'color': RED})],
        [(title, {'size': 34, 'font': HEAD, 'bold': True, 'color': H})],
        [(body, {'size': 19})],
        [('→ ' + imp, {'size': 16, 'color': MUT})]], spacing=10)

# ================= 12 JOIAS — TABELA =================
s = slide()
header(s, 'Joias em Prata · Palavras-chave', 'Prata x Ouro — busca e CPC',
       'Google Ads · Planejador de Palavras-chave · Brasil · média dos últimos 12 meses · total com CPC ponderado pelo volume.')
for gi, (g, items) in enumerate(KW.items()):
    x = 100 + gi * 880
    card(s, x, 260, 840, 700, accent_top=gi == 0)
    text(s, x + 40, 290, 760, 50, g.upper(), size=30, font=HEAD, bold=True, color=RED if gi == 0 else H)
    for c, xx, w, al in [('Palavra-chave', x + 40, 400, PP_ALIGN.LEFT), ('Buscas/mês', x + 440, 180, PP_ALIGN.RIGHT),
                         ('CPC médio', x + 620, 180, PP_ALIGN.RIGHT)]:
        text(s, xx, 360, w, 30, c, size=15, color=MUT, bold=True, upper=True, align=al)
    rect(s, x + 40, 395, 760, 1, LINE)
    rows = items + [(f'Total {g.lower()}', VOL[g], CPC[g])]
    for i, (k, v, c) in enumerate(rows):
        y = 405 + i * 100
        tot = i == len(rows) - 1
        o = {'size': 26 if tot else 24, 'bold': tot, 'color': H}
        text(s, x + 40, y, 400, 90, [[(k, o)]], anchor=MSO_ANCHOR.MIDDLE)
        text(s, x + 440, y, 180, 90, [[(br(v), o)]], align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        text(s, x + 620, y, 180, 90, [[(f'R$ {br(c, 2)}', o)]], align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        rect(s, x + 40, y + 94, 760, 2 if i == len(rows) - 2 else 1, H if i == len(rows) - 2 else LINE)

out = 'relatorio-2026.pptx'
prs.save(out)
print('ok', out, len(prs.slides), 'slides')

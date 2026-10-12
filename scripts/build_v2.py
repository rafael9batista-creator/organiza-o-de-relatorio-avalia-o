# -*- coding: utf-8 -*-
"""Gera a v2 da caracterização consolidada (Ultracargo Palmeirante) a partir da base de 07/10/2026
+ cotações OPE154. Uso: python3 -I build_v2.py <base.xlsx> <saida.xlsx>"""
import re, sys, copy
from collections import OrderedDict
import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter

sys.path.insert(0, __file__.rsplit('/', 1)[0])
from dados_v2 import COT_LINHA, COT_COMP, NOVO_LINHA, NOVO_COMP, COTACOES, ACHADOS

BASE, OUT = sys.argv[1], sys.argv[2]

# cores
PRETO, VERM, VERDE, AMAR, AZUL = '000000', 'C00000', '00703C', 'B38600', '1F4E9A'
AZESC, ROXO = '002060', '7030A0'
BASE_COLORS = {PRETO, VERM, VERDE, AMAR, AZUL}
LABEL_COLOR = {'Crítica: ': PRETO, 'NF: ': VERM, 'SAP: ': VERDE, 'Projeto: ': AMAR, 'Placa/campo: ': AZUL}
LABEL_RE = re.compile(r'(Crítica: |NF: |SAP: |Projeto: |Placa/campo: )')

thin = Side(style='thin', color='BFBFBF')
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
HDR_FILL_V2 = PatternFill('solid', fgColor='E4DFEC')   # lilás claro = novo na v2
HDR_FONT_V2 = Font(name='Arial', size=9, bold=True, color=ROXO)
ALIGN_TOP = Alignment(wrap_text=True, vertical='top')

# ---------------------------------------------------------------- helpers
def norm_color(c):
    if c is None or c.rgb is None or not isinstance(c.rgb, str):
        return PRETO
    return c.rgb[-6:].upper()

def cell_text(v):
    if v is None:
        return ''
    if isinstance(v, CellRichText):
        return ''.join(p if isinstance(p, str) else p.text for p in v)
    return str(v)

def seg(t, c=PRETO, b=False, s=False):
    return {'t': t, 'c': c, 'b': b, 's': s}

def split_labels(text):
    """'Crítica: ...NF: ...' -> [(label, content), ...]; texto sem rótulo vira ('', texto)."""
    parts = LABEL_RE.split(text)
    out = []
    if parts[0].strip():
        out.append(('', parts[0].strip()))
    for i in range(1, len(parts), 2):
        out.append((parts[i], parts[i + 1].strip()))
    return out

def rich(segs):
    blocks = []
    for sg in segs:
        if not sg['t']:
            continue
        f = InlineFont(rFont='Arial', sz=9, color=Color(rgb='FF' + sg['c']),
                       b=True if sg['b'] else None, strike=True if sg['s'] else None)
        blocks.append(TextBlock(f, sg['t']))
    return CellRichText(blocks)

def plain_font(color, bold=False):
    return Font(name='Arial', size=9, color=color, bold=bold)

# ---------------------------------------------------------------- redundância
PHRASES = [
 r'LI-3100-\d{4}(?:-\d{3})?', r'\(lista de linhas\)',
 r'(?:DE|FD|RM)-3100-\d{4}-\d{3}(?:/-\d{3})?(?:_\d{2}A\d{2})?',
 r'DE-3100-4300-014(?: rev\. 04)?',
 r'BOMBA CENTRIFUGA(?: DESCARREG)?',
 r'não consta(?:m)? na base SAP \(seleção\)',
 r'ABC [ABC]\b', r'ano 2024', r'Finame \d{8}',
 r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}',
 r'KSB Brasil Ltda', r'Metroval Controle de Flu[ií]dos Ltda\.?', r'Romão Tecnologias Industriais Ltda\.?',
 r'Volga Engenharia Ind\. e Com\. Ltda', r'STEMAC S\.A\. Grupos Geradores', r'PHD Master Consultoria Ltda',
 r'TSPRO Fab\. Dist\. e Rep\. de Equip\. Industriais',
 r'Sem dado teórico individual\.', r'Só leitura de campo \(sem 2ª fonte\)\.', r'Vínculo a 14a/14b/14d não documentado\.',
 r'Modelo da placa \([^)]*\) = crítica \([^)]*\)\.', r'Fabricação 2024 compatível com NF de 2024\.',
 r'Confronto campo×ANP: [^.]*\.', r'OP só na ANP', r'OP confere campo × ANP\.', r'ANP preenche OP nao lida/parcial em campo',
 r'Placa completa\. Sem dado teórico individual; vínculo à linha 14a pela posição \(top loading\)\.',
 r'Placa fotografada \(terceiros, 23/09/2026\)\.', r'SEM placa fotografada\.',
 r'Produto do projeto: [^.]*\.', r'Projeto: bomba de drenagem \(recalque 3"\)\. Sem cadastro SAP\.',
 r'Dados KSB só na ANP; motor diesel sem placa lida\.', r'Projeto e SAP: bomba principal do SCI\.',
 r'Placa KSB com data 07/02/2025 posterior à NF\.',
 r'NF 34(?:370|511) \(decomposição aritmética única\)', r'decomposição aritmética única',
 r'KSB (?:MEGAFLOW V 80-250 K(?: \(vertical\))?|MCPK 125-?0?80-315 EE|MCPK 125-100-315(?: EE)?|MCPK 100-65-315 EE|MCPK 125-100-250 EE|MCPK 50-32-200\.1 GG|MCPK 250-200-500 GG)',
 r'(?:MEGAFLOW V 80-250 K(?: \(vertical\))?|megaflow V 80-250 K|MCPK 125-?0?80-315 EE|MCPK 125-100-315(?: EE)?|MCPK 100-65-315 EE|MCPK 125-100-250 EE|MCPK 50-32-200\.1 GG|MCPK 250-200-500 GG|MCPK \(serie 315\)(?: EE)?)',
 r'IPW55 \d+ kg \| 220/380/440 V \| [\d,/]+ A \| [\d,]+ kW / \d+ CV \| FS 1,00 \| 1780 rpm \| FP 0,8\d \| ISOL F \| rend [\d,]+% \| Ex db IIB T4 Gb \| DNV 18\.0055 X',
 r'Ex db IIB T4 Gb', r'DNV 18\.0055 X', r'WEG W22Xdb? ?\|? ?(?:200M/L|180M/L)?',
 r'Medidor de vazao tipo turbina – Metroval, MTF-?80', r'DIMEL(?: no)? 365/2020', r'Classe exatidao 0,3', r'QMM 100',
 r'60/-20 C', r'Temp max 60 min -20', r'[Vv]azao 150/15',
 r'Valvula de controle de vazao – Metroval, VCDM-(?:100|80|50)', r'DN (?:100|80|50) mm; T\.max 70 C; PN 10 bar',
 r'Combustiveis', r'Valvula ON-OFF – Metroval, VOOM-100',
 r'Transmissor de temperatura – ALUTAL, CEX-A-00015', r'Ex ia IIC T6\.\.T1 Ga', r'Ex ic IIC(?: T6\.\.T4 Ga)?',
 r'IP66W/IP66X', r'INMETRO', r'Instrumentacao',
 r'Medidor massico – Metroval, OI (?:600 SIMT/F2|50 SMPF/F2)', r'Vazao 30-300 L/min; 2" 150# RF; 20 bar; 60 C',
 r'Vazao min 300 / max 3000 \| Conexao 4" 150#RF \| Pressao max 19,5 \| Temp max 60 C', r'Biodiesel B100',
 r'Protecao de transbordamento e aterramento – KFW \(Qingdao Aiptec\), SLA-S-Y', r'Cert\. DNV 18\.0113 X',
 r'85-250 Vca; Ex db ia \[ia Ga\] IIB T5 Gb(?:; -20 a \+55 C)?', r'Carregamento rodoviario', r'Carregamento ferroviario',
 r'braço de carregamento superior Redlands CC3201408310, Ø 4"', r'P\.C\. 11715(?: / 11716)?', r'CC3201408310',
 r'Braco de carregamento superior \(ferroviario\) – TSPRO, BTC-42CA\.0466-4', r'ASME B31\.3 ?/ ?ASME IX',
 r'Cliente ULTRACARGO PALMEIRANTE', r'4" CARB(?:ONO)?\.? AL(?:UM\.)? / VITON',
 r'Computador de vazao \(presset [AB]\) – TOPTECH, MultiLoad II Ex', r'MultiLoad II Ex',
 r'VOLGA Cliente ULTRACARGO (?:SOLUÇÕES LOGÍSTICA SA|TERMINAIS LTDA\.)', r'doc\. DE-3100-4300-014_00A04',
 r'0,6/1 kV', r'isolação 1,0 kV', r'NBR IEC 61439-1 & 2',
 r'ANP Fig\. \d+(?:[/,] ?\d+)*', r'ver grupo 15-17', r'NF 2739 \(Romão, 29/05/2024\)', r'não lido',
 r'\(TAG pendente\) – Medidor de vazao tipo turbina',
 r'\b(?:B|BR|FT|FCV|SDV|TE|XS|TQ|GD|PN|PNA|QGMT|QGBT|TR|TL|BC|NB|QL|QD-UPS|BB\.NB|TR\.NB|H|BOTTOM)-3\d{3,4}(?: ?[AB])?\b',
 r'1\.000 L',
]
PHRASES_RE = [re.compile(p) for p in PHRASES]

def strike_redundancies(segs):
    """Marca com tachado a 2ª ocorrência em diante de cada frase, só em texto de cor-base e não-negrito."""
    text = ''.join(s['t'] for s in segs)
    n = len(text)
    color = [None] * n; bold = [False] * n; strike = [False] * n
    i = 0
    for s in segs:
        for k in range(len(s['t'])):
            color[i + k] = s['c']; bold[i + k] = s['b']; strike[i + k] = s['s']
        i += len(s['t'])
    for rx in PHRASES_RE:
        seen = set()
        for m in rx.finditer(text):
            a, b = m.span()
            if b <= a:
                continue
            if any(color[k] not in BASE_COLORS for k in range(a, b)):
                continue          # ocorrência em texto novo (roxo/azul escuro): ignora
            key = re.sub(r'[^0-9a-zA-Z]', '', m.group(0)).lower()
            if key not in seen:
                seen.add(key)     # 1ª ocorrência deste texto fica normal
                continue
            for k in range(a, b):
                if not bold[k]:
                    strike[k] = True
    out = []
    for k in range(n):
        key = (color[k], bold[k], strike[k])
        if out and out[-1]['key'] == key:
            out[-1]['t'] += text[k]
        else:
            out.append({'key': key, 't': text[k]})
    return [seg(o['t'], o['key'][0], o['key'][1], o['key'][2]) for o in out]

# ---------------------------------------------------------------- leitura da base
wb = openpyxl.load_workbook(BASE, rich_text=True)
ws = wb['CARACTERIZACAO']
FIRST, LAST = 4, 161
rows = OrderedDict()
for r in range(FIRST, LAST + 1):
    d = {}
    for col in 'ABCDEFGHIJKLMNOPQRSTU':
        c = ws[f'{col}{r}']
        d[col] = cell_text(c.value)
        d[col + '_color'] = norm_color(c.font.color) if not isinstance(c.value, CellRichText) else PRETO
    rows[r] = d

# vínculo corrigido TAG -> linha (cotação KSB)
MOVE = {'B-3111': '4', 'B-3112': '4', 'B-3116': '6', 'B-3117': '6', 'B-3121': '9', 'B-3123': '8'}
MOVE_FROM = {'B-3111': '6', 'B-3112': '6', 'B-3116': '4', 'B-3117': '4', 'B-3121': '8', 'B-3123': '9'}

def linha_efetiva(d):
    return MOVE.get(d['C'], d['A'])

conj_rows = [r for r, d in rows.items() if d['B'] in ('CONJUNTO', 'GRUPO')]
comp_by_linha = OrderedDict()
for r, d in rows.items():
    if d['B'] == 'COMPONENTE':
        comp_by_linha.setdefault(linha_efetiva(d), []).append(r)

# ---------------------------------------------------------------- célula única
def d_segments(text, bullet_prefix=''):
    out = []
    for label, content in split_labels(text):
        if not content:
            continue
        col = LABEL_COLOR.get(label, PRETO)
        if label:
            out.append(seg(label, col, b=True))
        out.append(seg(content + ' ', col))
    return out

def ficha_segments(d):
    """E..J do conjunto com a cor própria de cada célula."""
    out = []
    for col, rot in (('E', 'Marca: '), ('F', 'Modelo: '), ('G', 'Capacidade: '), ('H', 'Série/OP: '), ('I', 'TAG/SAP: '), ('J', 'Fabricação: ')):
        if d[col]:
            out.append(seg(rot, d[col + '_color'], b=True))
            out.append(seg(d[col] + '. ', d[col + '_color']))
    return out

def comp_segments(r, linha):
    d = rows[r]
    out = [seg('\n• ' + d['C'] + ': ', PRETO, b=True)]
    if d['C'] in MOVE:
        out.append(seg(f'[CORREÇÃO v2: componente da linha {MOVE[d["C"]]}; na base de 07/10 estava na linha {MOVE_FROM[d["C"]]}] ', ROXO, b=True))
    out += d_segments(d['D'])
    if d['T']:
        out.append(seg('Obs.: ', PRETO, b=True)); out.append(seg(d['T'] + ' ', PRETO))
    if d['U']:
        out.append(seg('Fotos: ', PRETO, b=True)); out.append(seg(d['U'] + '. ', PRETO))
    if d['C'] in COT_COMP:
        out.append(seg('Cotação: ', AZESC, b=True)); out.append(seg(COT_COMP[d['C']] + ' ', AZESC))
    if d['C'] in NOVO_COMP:
        out.append(seg('Novo v2: ', ROXO, b=True)); out.append(seg(NOVO_COMP[d['C']] + ' ', ROXO))
    return out

def condensed_for(r):
    d = rows[r]
    L = d['A']
    segs = [seg(f'LINHA {L} – {d["C"]} [{d["B"]}]\n', PRETO, b=True)]
    segs += d_segments(d['D'])
    # NF: colunas K/L/M/Q/S
    nf_bits = []
    if d['K'] and not d['K'].startswith('↑'):
        nf_bits.append(f'NF nº {d["K"]}')
    if d['L']: nf_bits.append(f'de {d["L"]}')
    if d['M']: nf_bits.append(f'– {d["M"]}')
    if d['Q']: nf_bits.append(f'– valor comprovado R$ {float(d["Q"]):,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.'))
    if nf_bits:
        segs.append(seg('\nNF (colunas K-Q): ', VERM, b=True)); segs.append(seg(' '.join(nf_bits) + '. ', VERM))
    if d['S']:
        segs.append(seg('Atribuição: ', PRETO, b=True)); segs.append(seg(d['S'] + '. ', PRETO))
    # crítica: N/O/P
    if d['N'] or d['P']:
        vals = []
        if d['N']: vals.append(f'qtd {d["N"]}')
        if d['O']: vals.append('unitário R$ ' + f'{float(d["O"]):,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.'))
        if d['P']: vals.append('total R$ ' + f'{float(d["P"]):,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.'))
        segs.append(seg('Crítica (colunas N-P): ', PRETO, b=True)); segs.append(seg(', '.join(vals) + '. ', PRETO))
    segs.append(seg('\nFicha (colunas E-J): ', PRETO, b=True)); segs += ficha_segments(d)
    if d['T']:
        segs.append(seg('\nObs. (coluna T): ', PRETO, b=True)); segs.append(seg(d['T'] + ' ', PRETO))
    if d['U']:
        segs.append(seg('Fotos (coluna U): ', PRETO, b=True)); segs.append(seg(d['U'] + '. ', PRETO))
    if L in COT_LINHA:
        segs.append(seg('\nCotação OPE154: ', AZESC, b=True)); segs.append(seg(COT_LINHA[L] + ' ', AZESC))
    if L in NOVO_LINHA:
        segs.append(seg('\nNovo v2 (achados / correções / proposições): ', ROXO, b=True)); segs.append(seg(NOVO_LINHA[L] + ' ', ROXO))
    comps = comp_by_linha.get(L, [])
    if d['B'] == 'GRUPO':
        sub = '14a, 14b, 14c, 14d, 14e e "14 (comp.)"' if L == '14' else '15, 16 e 17'
        segs.append(seg(f'\nSubconjuntos: {sub} – ver as células únicas próprias. ', ROXO))
    if comps:
        segs.append(seg(f'\nCOMPONENTES ({len(comps)}):', PRETO, b=True))
        for cr in comps:
            segs += comp_segments(cr, L)
    return strike_redundancies(segs)

def condensed_comp_group():
    comps = comp_by_linha.get('14 (comp.)', [])
    segs = [seg('LINHA 14 (comp.) – Componentes de campo da linha 14 sem vínculo documentado a 14a/14b/14d [COMPONENTES SEM VÍNCULO]\n', PRETO, b=True)]
    segs.append(seg('Novo v2: ', ROXO, b=True))
    segs.append(seg(f'{len(comps)} instrumentos e válvulas lidos em campo (medidores de vazão turbina MTF-80 e mássicos OI 50, válvulas de controle VCDM, válvulas on-off VOOM, transmissores de temperatura Alutal, monitores de transbordamento/aterramento KFW) – todos Metroval, Alutal ou KFW, compatíveis com o fornecimento Metroval (faturamento direto Ecoserv para os KFW: 32 monitores, NF 110). Sem lista de equipamentos Metroval não é possível alocar cada item a 14a, 14b ou 14d; para o laudo propõe-se tratá-los como componentes do conjunto de skids (linha 14) e usar as fotos de placa como comprovação. ', ROXO))
    segs.append(seg(f'\nCOMPONENTES ({len(comps)}):', PRETO, b=True))
    for cr in comps:
        segs += comp_segments(cr, '14 (comp.)')
    return strike_redundancies(segs)

# ---------------------------------------------------------------- escrita em CARACTERIZACAO
hdr = {'V': 'Cotação OPE154 (v2 – azul escuro)', 'W': 'Achados / correções / proposições (v2 – roxo)',
       'X': 'CÉLULA ÚNICA v2 – conjunto + componentes (cores de origem; redundâncias taxadas; azul escuro = cotação; roxo = novo)'}
for col, t in hdr.items():
    c = ws[f'{col}3']; c.value = t; c.font = HDR_FONT_V2; c.fill = HDR_FILL_V2; c.alignment = Alignment(wrap_text=True, vertical='center'); c.border = BORDER
ws.column_dimensions['V'].width = 60; ws.column_dimensions['W'].width = 60; ws.column_dimensions['X'].width = 120
ws['X1'] = 'v2 – 12/10/2026: colunas V, W e X e abas CONDENSADO, COTACOES_OPE154 e ACHADOS_v2 (ver LEIA-ME). Colunas A-U = base de 07/10/2026 sem alteração.'
ws['X1'].font = Font(name='Arial', size=9, bold=True, color=ROXO)
ws.freeze_panes = 'D4'
ws.auto_filter.ref = f'A3:X{LAST}'

condensed_cells = OrderedDict()   # chave -> (rótulo, rich, fotos, r)
for r, d in rows.items():
    L = d['A']; C = d['C']
    v = w = None
    if d['B'] in ('CONJUNTO', 'GRUPO'):
        v = COT_LINHA.get(L); w = NOVO_LINHA.get(L)
    else:
        v = COT_COMP.get(C); w = NOVO_COMP.get(C)
        if C in MOVE:
            w = f'[CORREÇÃO v2: componente da linha {MOVE[C]}; na base de 07/10 estava na linha {MOVE_FROM[C]}] ' + (w or '')
    for col, val, color in (('V', v, AZESC), ('W', w, ROXO)):
        c = ws[f'{col}{r}']
        c.alignment = ALIGN_TOP; c.border = BORDER
        if val:
            c.value = val; c.font = plain_font(color)
    cx = ws[f'X{r}']; cx.alignment = ALIGN_TOP; cx.border = BORDER
    if d['B'] in ('CONJUNTO', 'GRUPO'):
        rt = condensed_for(r)
        cx.value = rich(rt)
        condensed_cells[L] = (f'{L} – {C}', rt, d['B'], r)
    elif L == '14 (comp.)' and r == comp_by_linha['14 (comp.)'][0]:
        rt = condensed_comp_group()
        cx.value = rich(rt)
        condensed_cells['14 (comp.)'] = ('14 (comp.) – componentes sem vínculo', rt, 'COMPONENTES', r)
        for rr in comp_by_linha['14 (comp.)'][1:]:
            ws[f'X{rr}'].value = '↑ ver célula única em X' + str(r); ws[f'X{rr}'].font = plain_font(ROXO)

# ordenar as células únicas na ordem das linhas (14 (comp.) depois de 14e)
ordem = ['1','2','3','4','5','6','7','8','9','10','11','12','13','14','14a','14b','14c','14d','14e','14 (comp.)','15-17','15','16','17','18']

# ---------------------------------------------------------------- aba CONDENSADO
wsc = wb.create_sheet('CONDENSADO', 2)
wsc['A1'] = 'CÉLULA ÚNICA por conjunto (v2 – 12/10/2026): todas as informações da base de 07/10 (conjunto + componentes) nas cores de origem, redundâncias taxadas, cotações OPE154 em azul escuro e achados/correções/proposições em roxo. Base para as legendas do relatório fotográfico.'
wsc['A1'].font = Font(name='Arial', size=10, bold=True, color=ROXO)
heads = ['Linha', 'Nível', 'Conjunto', 'Célula única', 'Fotos / fontes de campo (conjunto + componentes)', 'Valor crítica (R$)', 'Valor NF (R$)', 'Valor cotação (R$)', 'Nº de componentes', 'Linha na aba CARACTERIZACAO']
for i, h in enumerate(heads, 1):
    c = wsc.cell(row=3, column=i, value=h); c.font = Font(name='Arial', size=9, bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F3864'); c.alignment = Alignment(wrap_text=True, vertical='center'); c.border = BORDER
widths = [9, 12, 34, 230, 40, 16, 16, 16, 10, 14]
for i, wdt in enumerate(widths, 1):
    wsc.column_dimensions[get_column_letter(i)].width = wdt
wsc.row_dimensions[3].height = 30
COT_VAL = {'1': 393500.00, '2': 815201.04, '3': 34798.96, '4': 1223780.04, '5': 71978.04, '6': 174716.70, '7': 161876.22, '8': 308561.97, '9': 102010.24, '10': 188172.40, '11': 162486.68, '12': 76802.33, '13': 160106.86, '14': 8992832.34, '15-17': 546941.16, '15': 249823.52, '16': 177058.82, '17': 120058.82, '18': 3163087.46}
rr = 4
for key in ordem:
    if key not in condensed_cells:
        continue
    rot, rt, nivel, r = condensed_cells[key]
    d = rows[r]
    fotos = []
    if nivel in ('CONJUNTO', 'GRUPO') and d['U']:
        fotos.append(f'{d["C"]}: {d["U"]}')
    for cr in comp_by_linha.get(key, []):
        if rows[cr]['U']:
            fotos.append(f'{rows[cr]["C"]}: {rows[cr]["U"]}')
    wsc.cell(row=rr, column=1, value=key)
    wsc.cell(row=rr, column=2, value=nivel)
    wsc.cell(row=rr, column=3, value=d['C'] if nivel != 'COMPONENTES' else 'Componentes de campo sem vínculo documentado (linha 14)')
    wsc.cell(row=rr, column=4, value=rich(rt))
    wsc.cell(row=rr, column=5, value='\n'.join(fotos))
    if nivel in ('CONJUNTO', 'GRUPO'):
        if d['P']: wsc.cell(row=rr, column=6, value=float(d['P']))
        if d['Q']: wsc.cell(row=rr, column=7, value=float(d['Q']))
        if key in COT_VAL:
            c = wsc.cell(row=rr, column=8, value=COT_VAL[key]); c.font = plain_font(AZESC)
    wsc.cell(row=rr, column=9, value=len(comp_by_linha.get(key, [])))
    wsc.cell(row=rr, column=10, value=f'CARACTERIZACAO!X{r}')
    for ci in range(1, 11):
        c = wsc.cell(row=rr, column=ci); c.alignment = ALIGN_TOP; c.border = BORDER
        if c.font.name != 'Arial' or c.font.color is None:
            c.font = plain_font(AZESC if ci == 8 else PRETO)
        if ci in (6, 7, 8):
            c.number_format = '#,##0.00'
    wsc.row_dimensions[rr].height = 409
    rr += 1
wsc.cell(row=rr + 1, column=3, value='Totais (CONJUNTO)').font = Font(name='Arial', size=9, bold=True)
wsc.cell(row=rr + 1, column=6, value=f'=SUMIFS(F4:F{rr-1},B4:B{rr-1},"CONJUNTO")').number_format = '#,##0.00'
wsc.cell(row=rr + 1, column=7, value=f'=SUMIFS(G4:G{rr-1},B4:B{rr-1},"CONJUNTO")+SUMIFS(G4:G{rr-1},B4:B{rr-1},"GRUPO")').number_format = '#,##0.00'
wsc.cell(row=rr + 1, column=8, value=f'=SUMIFS(H4:H{rr-1},B4:B{rr-1},"CONJUNTO")').number_format = '#,##0.00'
wsc.cell(row=rr + 2, column=3, value='Nota: o total de cotação soma só as linhas CONJUNTO (as linhas 14 e 15-17 estão nos subitens); a linha 14 usa o preço único Metroval.').font = plain_font(ROXO)
wsc.freeze_panes = 'D4'

# ---------------------------------------------------------------- aba COTACOES_OPE154
wso = wb.create_sheet('COTACOES_OPE154')
wso['A1'] = 'Propostas comerciais do arquivo OPE154_48.zip (49 PDFs; OPE154_16 = cópia de OPE154_05) – posição 12/10/2026. Azul escuro = dado da proposta; roxo = enquadramento/observação do avaliador.'
wso['A1'].font = Font(name='Arial', size=10, bold=True, color=AZESC)
heads = ['Arquivo', 'Fornecedor (CNPJ)', 'Proposta nº / revisão', 'Data', 'Objeto', 'Valor total (R$)', 'Linha da crítica / NF relacionada', 'Enquadramento proposto', 'Observação (v2)']
for i, h in enumerate(heads, 1):
    c = wso.cell(row=3, column=i, value=h); c.font = Font(name='Arial', size=9, bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F3864'); c.alignment = Alignment(wrap_text=True, vertical='center'); c.border = BORDER
for i, wdt in enumerate([14, 40, 30, 14, 70, 16, 38, 28, 60], 1):
    wso.column_dimensions[get_column_letter(i)].width = wdt
for j, rowv in enumerate(COTACOES, 4):
    for i, val in enumerate(rowv, 1):
        c = wso.cell(row=j, column=i, value=val)
        c.font = plain_font(ROXO if i in (8, 9) else AZESC); c.alignment = ALIGN_TOP; c.border = BORDER
        if i == 6: c.number_format = '#,##0.00'
n = 4 + len(COTACOES)
wso.cell(row=n + 1, column=5, value='Soma das propostas listadas (inclui serviços, duplicata e itens fora da crítica – apenas conferência)').font = plain_font(ROXO)
wso.cell(row=n + 1, column=6, value=f'=SUM(F4:F{n-1})').number_format = '#,##0.00'
wso.cell(row=n + 2, column=5, value='Soma das propostas que cobrem a crítica de M&E (STEMAC + KSB ×2 + Metroval + Romão selos + VOLGA)').font = plain_font(ROXO)
wso.cell(row=n + 2, column=6, value='=393500+850000+2630491.47+8992832.34+546941.16+3163087.46').number_format = '#,##0.00'
wso.cell(row=n + 3, column=5, value='Total da crítica de M&E (RT GEREN/COSAN 2024/189)').font = plain_font(ROXO)
wso.cell(row=n + 3, column=6, value=16576852.44).number_format = '#,##0.00'
wso.cell(row=n + 4, column=5, value='Diferença cotações − crítica (1 centavo = arredondamento da KSB)').font = plain_font(ROXO)
wso.cell(row=n + 4, column=6, value=f'=F{n+2}-F{n+3}').number_format = '#,##0.00'
wso.freeze_panes = 'A4'; wso.auto_filter.ref = f'A3:I{n-1}'

# ---------------------------------------------------------------- aba ACHADOS_v2
wsa = wb.create_sheet('ACHADOS_v2')
wsa['A1'] = 'Achados, correções e proposições da v2 (12/10/2026) – tudo em roxo = contextualização/decisão do avaliador a partir das cotações OPE154 e da base de 07/10.'
wsa['A1'].font = Font(name='Arial', size=10, bold=True, color=ROXO)
heads = ['Nº', 'Linha(s)', 'Tipo', 'Descrição', 'Fonte', 'Efeito / ação']
for i, h in enumerate(heads, 1):
    c = wsa.cell(row=3, column=i, value=h); c.font = Font(name='Arial', size=9, bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='5B2C8A'); c.alignment = Alignment(wrap_text=True, vertical='center'); c.border = BORDER
for i, wdt in enumerate([6, 12, 18, 90, 40, 50], 1):
    wsa.column_dimensions[get_column_letter(i)].width = wdt
for j, rowv in enumerate(ACHADOS, 4):
    for i, val in enumerate(rowv, 1):
        c = wsa.cell(row=j, column=i, value=val); c.font = plain_font(ROXO); c.alignment = ALIGN_TOP; c.border = BORDER
wsa.freeze_panes = 'A4'

# ---------------------------------------------------------------- aba NF_ME: colunas I e J
wsn = wb['NF_ME']
NF_COT = {
 '149420': ('STEMAC 13107/2023 REV01 (29/08/2023): R$ 393.500,00 – exata.', 'CNPJ da proposta (92.753.268/0052-62) difere do da NF.'),
 '526572': ('KSB RAS0045265 Rev. 7 (19/09/2023): 2 × 407.600,52 + 34.798,96 = R$ 850.000,00 – exata.', 'Motor diesel MWM 6.12TCE; tanques 1.000 L inclusos.'),
 '34370': ('KSB RAS0051138 Rev. 10 (11/12/2023): preço por TAG (28 bombas, R$ 2.630.491,47).', 'As 3 NFs KSB somam R$ 2.701.551,48 = cotação + R$ 71.060,00. Vínculo TAG × linha corrigido (ver ACHADOS A01/A02).'),
 '34511': ('KSB RAS0051138 Rev. 10: preço por TAG.', 'Idem.'),
 '543048': ('KSB RAS0051138 Rev. 10, item 23: B-3132/33 = R$ 162.486,68.', 'Excesso de R$ 71.060,00 sem contrapartida na cotação – pedir NF detalhada.'),
 '41385': ('Metroval 230200112STM-SC-RevE (30/08/2023): pacote R$ 8.992.832,34, sem preço por subitem.', 'A proposta prevê faturamento direto por Redlands, MKG, PHD Master e Ecoserv.'),
 '1551': ('Metroval: item (d) skid de descarga com duas bombas; PHD Master na lista de faturamento direto.', 'Explica o emitente PHD; composição ainda pendente (bombas B-3124/25 cotadas pela KSB).'),
 '42239': ('Metroval: item (e) carregamento ferroviário (sem preço).', 'Candidata às 10 SDV (campo: VOOM-100).'),
 '7032': ('Sem cotação TSPRO no OPE154.', 'TSPRO não consta na lista de faturamento direto da Metroval – compra adicional ou substituição.'),
 '7488': ('Sem cotação TSPRO no OPE154.', 'Idem.'),
 '11642': ('Redlands 0760/24 Rev.01 (10/01/2024): 21 escadas pantográficas PF-1604 × 11.495,00 = R$ 241.395,00 – exata.', 'NF identificada: escadas pantográficas das plataformas.'),
 '11726': ('Sem cotação específica (Redlands na lista de faturamento direto da Metroval).', 'Candidata aos braços CC32 ou peças.'),
 '1621': ('Sem cotação específica (PHD Master na lista de faturamento direto).', '—'),
 '91': ('Sem cotação específica (Ecoserv na lista de faturamento direto).', '—'),
 '110': ('Ecoserv 036/24 (09/02/2024): 32 monitores de aterramento × 12.090,00 = R$ 386.880,00 – exata.', 'NF identificada: monitores KFW SLA-S-Y (XS-xxxx em campo).'),
 '142': ('Ecoserv 036/24 Rev. 00 (20/02/2024): 4 amostradores KFW × 25.000,00 = R$ 100.000,00 – exata.', 'NF identificada: amostradores do skid de descarga (14d).'),
 '161': ('Sem cotação específica.', '—'),
 '26460': ('Sem cotação MKG; Smart Control cotou 24 filtros T 6"/8"/10" por R$ 645.624,00 (OPE154_22).', 'MKG na lista de faturamento direto da Metroval; filtros MKG em campo.'),
 '4965': ('Metroval: supervisão de instalação/comissionamento "inclusa no preço" (condições de fornecimento).', 'Serviço faturado à parte apesar da proposta – conferir.'),
 '2739': ('Romão 00638 Rev 04 (17/01/2024): selos R$ 546.941,16 + supervisão R$ 109.200,00 + frete R$ 101.183,92 = R$ 757.325,08.', 'Excesso da NF (R$ 54.399,96) não bate com frete nem supervisão – pedir NF detalhada.'),
 '22': ('VOLGA 20230138 REV 6A (06/10/2023): R$ 3.163.087,46 – exata na soma das 3 NFs.', 'Proposta da matriz (0001-12); NF da filial (0006-27).'),
 '33': ('VOLGA 20230138 REV 6A.', 'Idem.'),
 '34': ('VOLGA 20230138 REV 6A.', 'Idem.'),
}
for col, t in (('I', 'Cotação OPE154 correspondente (v2 – azul escuro)'), ('J', 'Observação v2 (roxo)')):
    c = wsn[f'{col}2']; c.value = t; c.font = HDR_FONT_V2; c.fill = HDR_FILL_V2; c.alignment = Alignment(wrap_text=True, vertical='center'); c.border = BORDER
wsn.column_dimensions['I'].width = 55; wsn.column_dimensions['J'].width = 55
for r in range(3, 26):
    nf = str(wsn[f'A{r}'].value)
    if nf in NF_COT:
        a, b = NF_COT[nf]
        wsn[f'I{r}'].value = a; wsn[f'I{r}'].font = plain_font(AZESC); wsn[f'I{r}'].alignment = ALIGN_TOP; wsn[f'I{r}'].border = BORDER
        wsn[f'J{r}'].value = b; wsn[f'J{r}'].font = plain_font(ROXO); wsn[f'J{r}'].alignment = ALIGN_TOP; wsn[f'J{r}'].border = BORDER

# ---------------------------------------------------------------- LEIA-ME
wsl = wb['LEIA-ME']
wsl['A1'].value = 'Caracterização consolidada de M&E – Ultracargo Palmeirante-TO (base 07/10/2026; v2 de 12/10/2026 com as cotações OPE154)'
leia = [
 ('', PRETO, False),
 ('v2 – 12/10/2026 – o que mudou (base de 07/10 preservada nas colunas A-U; tudo que é novo está em colunas e abas à parte):', ROXO, True),
 ('Azul escuro (002060) – dados das propostas comerciais do OPE154_48.zip (49 PDFs: STEMAC, KSB ×2, Metroval, Romão, VOLGA, Chronus, Ecoserv, Redlands, Kidde, Endress, Solo, RI, Protego, Tecnofluid etc.). Coluna V da CARACTERIZACAO, coluna I da NF_ME e aba COTACOES_OPE154.', AZESC, False),
 ('Roxo (7030A0) – achados, correções e proposições do avaliador: combinação das fontes, contexto técnico ou de mercado e decisões propostas. Coluna W da CARACTERIZACAO, coluna J da NF_ME, aba ACHADOS_v2 e trechos "Novo v2" das células únicas. "CORREÇÃO v2:" marca texto que substitui o que está escrito na base (o texto da base fica, sem tachado, para rastreabilidade).', ROXO, False),
 ('Coluna X (CARACTERIZACAO) e aba CONDENSADO – CÉLULA ÚNICA por conjunto: reúne crítica, NF, SAP, projeto, placa/campo, observação e fotos do conjunto e de todos os seus componentes, cada trecho na cor da fonte de origem, seguidos da cotação (azul escuro) e do novo (roxo). Texto tachado = redundância (2ª ocorrência em diante de modelo, TAG, documento de projeto, frase-padrão etc. dentro da mesma célula); a 1ª ocorrência fica normal. Texto novo (roxo/azul escuro) nunca é tachado. Para os componentes não se repetem as colunas E-J (já contidas na descrição D); para o conjunto elas entram na "Ficha".', PRETO, False),
 ('Correção de formato na célula única: na base a coluna D estava toda na cor do 1º rótulo (NF e SAP apareciam em preto dentro de "Crítica"); na célula única cada rótulo tem a cor da legenda (NF vermelho, SAP verde, Projeto amarelo, Placa/campo azul). A coluna D original não foi alterada.', ROXO, False),
 ('Correções de vínculo TAG × linha pela planilha de preços KSB RAS0051138 Rev. 10 (ver ACHADOS_v2 A01/A02): linha 4 = B-3103, 3104, 3105, 3106, 3107, 3108, 3110, 3111, 3112, 3113, 3114, 3115; linha 6 = B-3116, B-3117; linha 8 = B-3120, B-3122, B-3123; linha 9 = B-3121. As linhas da base não foram movidas; a célula única já agrupa os componentes pelo vínculo corrigido e marca "[CORREÇÃO v2]" nos seis TAGs afetados.', ROXO, False),
 ('Limite de exibição: a célula única tem até ~25 mil caracteres; no Excel a linha mostra no máximo 409 pt de altura (aba CONDENSADO já está nessa altura, coluna de 230 de largura). Para ler tudo, use a barra de fórmulas ou copie a célula para o Word.', PRETO, False),
 ('Painel de congelamento da CARACTERIZACAO ajustado para D4 (na base estava em D122) e autofiltro estendido até a coluna X.', ROXO, False),
 ('Anexo não usado: 2026-10-11_PRE_ANALISE_UNICA_v3.zip trata da documentação das Fazendas Paloma (Paulo Roberto de Paula) e não contém nada da Ultracargo – nenhum dado dele entrou nesta planilha (ACHADOS_v2 A19).', ROXO, False),
]
r0 = 20
for i, (t, col, b) in enumerate(leia):
    c = wsl.cell(row=r0 + i, column=1, value=t); c.font = Font(name='Arial', size=10, color=col, bold=b); c.alignment = Alignment(wrap_text=True, vertical='top')

wb.save(OUT)
print('salvo', OUT)
# QA: tamanho das células únicas
for k, (rot, rt, nivel, r) in condensed_cells.items():
    txt = ''.join(s['t'] for s in rt)
    struck = sum(len(s['t']) for s in rt if s['s'])
    print(f'{k:>10} | {len(txt):6d} chars | tachado {struck:5d} | runs {len(rt)}')

# -*- coding: utf-8 -*-
"""v4 (13/10/2026): assimila o 00_LEIA_ME correto (briefing unificado + pré-análise única de 05/10) e o
LOG_REPOSITORIO / LOG_FOTOS / Inventário de fotos (árvore de pastas ANEXO_IV_ME_fotos organizada por Rafael em 13/10).
Tudo o que vem do LEIA-ME entra em LILÁS; tudo o que vem da árvore de pastas / log entra em VERDE CLARO.
Uso: python3 -I build_v4.py <v3.xlsx> <saida_v4.xlsx> <ARVORE.csv> <PLANO.csv>"""
import sys, csv, re
import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter

SRC, OUT, ARV_CSV, PLANO_CSV = sys.argv[1:5]
K, B, R, G, Y, N, P = '000000', '1F4E9A', 'C00000', '00703C', 'B38600', '002060', '7030A0'
L = 'A97BD6'   # lilás – alterações a partir do 00_LEIA_ME (briefing unificado / pré-análise única)
V = '00B050'   # verde claro – alterações a partir do LOG_REPOSITORIO / LOG_FOTOS / árvore de pastas de Rafael
thin = Side(style='thin', color='BFBFBF'); BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
TOP = Alignment(wrap_text=True, vertical='top')

def blk(col, txt, bold=False):
    return TextBlock(InlineFont(rFont='Arial', sz=9, b=bold or None, color=Color(rgb='FF' + col)), txt)

def rich(parts):
    out = []
    for item in parts:
        if item[0] == 'H': out.append(blk(K, item[1], True))
        elif item[0] == 'HL': out.append(blk(L, item[1], True))
        elif item[0] == 'HV': out.append(blk(V, item[1], True))
        else: out.append(blk(item[0], item[1]))
    return CellRichText(out)

def append_rich(cell, parts):
    """Acrescenta blocos ao fim da célula, preservando o que já existe (texto simples ou rich text)."""
    old = cell.value
    if isinstance(old, CellRichText): base = list(old)
    elif old is None: base = []
    else: base = [blk(K, str(old))]
    cell.value = CellRichText(base + list(rich(parts)))
    cell.alignment = TOP

wb = openpyxl.load_workbook(SRC, rich_text=True)

# =====================================================================================================
# 1. SINTESE – acréscimos por conjunto (lilás = LEIA-ME; verde claro = árvore de pastas / log)
# =====================================================================================================
ADD = {}  # linha -> dict(sint=[...], obs=[...], nome=[...])
def add(linha, sint=(), obs=(), nome=()):
    ADD[linha] = dict(sint=list(sint), obs=list(obs), nome=list(nome))

add("1",
 sint=[('HV', "\nÁrvore de fotos (Rafael, 13/10): "), (V, "pasta 'Item_01_Grupo_gerador_STEMAC ok' – 00_vistas_gerais (1) e 'Grupo gerador (motor)' (9 fotos, 3 delas – 102219/102259/102325 – copiadas de F:\\ em 13/10); placa Baudouin 6M21G460/6, série 5123S001673, ordem DHM21D0052*01, lida completa (foto 0055).")],
 obs=[(V, " Foto 0061 (acessórios do alternador – termostatos) ficou ilegível e não foi copiada (NAO_ALOCADAS).")])

add("2",
 sint=[('HV', "\nÁrvore de fotos (13/10): "), (V, "'Sintema combate a Incendio Casa de Máquinas/Item_02_Bombas_SCI_megacpk_250-200-500' – 00_vistas_gerais (8), 'BOMBA 3101-A' (6), 'BOMBA 3101-B' (9), 'Tanques diesel SCI (2)' (11); na casa de máquinas: 00_vista_geral (24, com 3 fotos 170822/170827/170847 copiadas de F:\\ em 13/10), 'Tanque de agua compbate incendio' (3) e 'sistema de geração de espuma' (23).")],
 obs=[('HL', " LEIA-ME: "), (L, "A42/Q03 – as fotos de terceiros de 23/10/2025 mostram motobombas KSB Ômega 150-460A com motor MTU 447A 320 CV, placa 'MB-502', patrimônio 005543: NÃO são as bombas SCI financiadas (MCPK 250-200-500) – local a confirmar, não usar; para as motobombas vermelhas usar as fotos da vistoria de 15/09. A31/A44/L3-08 – os 2 tanques de diesel KSB de 1.000 L estão confirmados pela ANP Fig. 199 e por foto de terceiros de 05/09/2025 (complementar); capacidade sem leitura de placa. L3-04 – o proporcionador FireDos FD4000/3-PP-S da casa de espuma (SAP H-3150, série 11600455) foi fabricado em 11/2016 e faturado em 2024 (NF SECUR 45694): é Outros Fixos, fora desta linha, mas a divergência de idade precisa constar no laudo de Instalações. "),
      ('HV', "Árvore: "), (V, "o inventário aponta conflito – as páginas ANP 45 (tanques), 75, 76 e 77 (rodoviário) estão na pasta do SCI; decisão de Rafael pendente (PLANO/CONFLITO).")])

add("3", sint=[('HV', "\nÁrvore de fotos (13/10): "), (V, "'Item_03_Bomba_jockey_Hydrobloc' – 00_vistas_gerais (1) e 'BOMBA JOCKEY' (2).")])

add("4",
 sint=[('HV', "\nÁrvore de fotos (13/10): "), (V, "'Centran 1 ok/Item_04_Bombas_megacpk_125-080-315_12un' com subpastas B-3103 (8), 3104 (8), 3105 (5), 3106 (6), 3107 (9), 3108 (7), 3110 (8), 3113 (7), 3114 (6), 3115 (6), 3116 (8), 3117 (6); 00_vistas_gerais do CETRAN 01 (10). A árvore de Rafael confirma a realocação da v2 (A01): B-3109 em Item_05 e B-3111/3112 em Item_06.")])
add("5", sint=[('HV', "\nÁrvore de fotos (13/10): "), (V, "'Item_05_Bomba_megacpk_100-065-315_1un' – B-3109 (4) e 00_vistas_gerais (1).")])
add("6", sint=[('HV', "\nÁrvore de fotos (13/10): "), (V, "'Item_06_Bombas_megacpk_125-080-315_2un' – B-3111 (5), B-3112 (8).")])

add("7",
 sint=[('HV', "\nÁrvore de fotos (13/10): "), (V, "'Centran 2 ok/Item_07_Bombas_megacpk_125-100-315_2un' – B-3118 (5), B-3119 (14), 00_vistas_gerais (1); vistas gerais do CETRAN 02 (12).")],
 obs=[('HL', " LEIA-ME: "), (L, "L3-03 – com a placa da B-3119 lida no lote 3 (OP M482301627, 180 m³/h, 32 mca, 1776 rpm, 2024) fecham 6/6 bombas da descarga ferroviária com placa; a B-3119 sai da lista de ressalvas. "),
      ('HV', "Árvore: "), (V, "foto 152835 lida como B-3120 dentro da pasta B-3119 (REVISAR – decisão pendente).")])
add("8", sint=[('HV', "\nÁrvore de fotos (13/10): "), (V, "'Item_08_Bombas_megacpk_125-100-315_3un' – B-3120 (6), B-3121 (6), B-3122 (7), 00_vistas_gerais (4) – confirma A02.")])
add("9", sint=[('HV', "\nÁrvore de fotos (13/10): "), (V, "'Item_09_Bomba_megacpk_125-100-315_1un' – B-3123 (14), 00_vistas_gerais (1).")])

add("10",
 sint=[('HV', "\nÁrvore de fotos (13/10): "), (V, "'sistema ETE/Item_10_Bombas_megaflow_V80-250K_2un ok' – B-3130 (10), B-3131 (8), 00_vistas_gerais (4); 'tanque da ete' (28).")],
 obs=[('HL', " LEIA-ME: "), (L, "L3-01/L3-10 – com as placas KSB lidas (B-3130 OP M502300011, B-3131 OP M502300012; 35 m³/h, 25 mca, 445 kg) as megaflow saem do vermelho da decisão D06 e passam a 'verificadas com placa'; a atribuição par × linha (10 ou 11) continua sem documento – ressalva mantida. M12 – a hipótese anterior 'linha 12 = B-3130/3131' foi abandonada. "),
      ('HV', "Árvore: "), (V, "fotos 152123/152145 lidas como B-3131 dentro da pasta B-3130 (REVISAR).")])
add("11",
 sint=[('HV', "\nÁrvore de fotos (13/10): "), (V, "'sistema ETE/Item_11_Bombas_megaflow_V80-250K_2un ok' – B-3132 (2), B-3133 (2), 00_vistas_gerais (2).")],
 obs=[('HL', " LEIA-ME: "), (L, "L3-01/L3-06 – B-3132 OP M502300013 e B-3133 OP M502300014 (50 m³/h, 10 mca, 570 kg) lidas; B-3132 é MEGAFLOW (a leitura 'MSBPDKV' era de miniatura); saem do vermelho. M11 – a NF 543048 excede a linha em R$ 71.060,00: usar o valor do plano até o proponente explicar.")])

add("12",
 sint=[('HV', "\nÁrvore de fotos (13/10): "), (V, "'Item_12_Bomba_megacpk_050-032-200 ok' – B-31100 (6), 00_vistas_gerais (2). Motor lido: WEG W22 Premium, gaiola, 3~ 60 Hz, carcaça 132M, IPW55, 68 kg, 220/380 V (código 17837444/5, dígito final incerto).")],
 obs=[('HL', " LEIA-ME: "), (L, "L3-02 – linha 12 = B-31100 verificada com placa (fecha também a S19: a OP M482301632 existe); motor WEG 7,5 kW; o serviço bombeado não foi identificado em campo (bomba ao lado do eletrocentro) – o vínculo com os chuveiros de emergência vem só do projeto; a divergência B-31100 (placa) × 'BOMBA 3160' (botoeira) vai para a lista ao proponente, sem resolver. "),
      ('HV', "Árvore: "), (V, "2 fotos da pasta 'Sitema de água potável artesiano ok' (171246 vista geral, 171302 painel) leem B-31100 – conflito ITEM12 × ÁGUA no inventário, decisão de Rafael pendente; a pasta onde ele as guardou sugere serviço de água potável/artesiana – registrar, não resolver.")])

add("13",
 sint=[('HV', "\nÁrvore de fotos (13/10): "), (V, "Rafael colocou 'Item_13_Bombas_megacpk_125-100-250_2un' DENTRO de 'Item_14d_Skid_descarga_caminhoes' – B-3124 (6), B-3125 (9), 00_vistas_gerais (4): as bombas são componentes do skid de descarga (confirma a v3). Placas: B-3124 OP M482301633 e B-3125 OP M482301634, MCPK 125-100-250 EE, 82 m³/h, 33,19 mca, 1721 rpm, 2024; linhas Gasolina A / Etanol anidro / Etanol hidratado na B-3124 e Biodiesel B100 / Diesel S-500 / S-10 na B-3125; motores WEG com placa por reler (reflexo).")],
 obs=[('HL', " LEIA-ME: "), (L, "M13 – a ANP corrige a vazão (140,82 m³/h): citar como fonte de terceiro.")])

add("14",
 sint=[('HV', "\nÁrvore de fotos (Rafael, 13/10): "), (V, "as válvulas, filtros, medidores e sensores que estavam soltos foram postos por Rafael nas pastas dos skids a que pertencem – Item_14a_Top_loading e Item_14b_Bottom_loading (dentro de 'Estação de carga e descarga Caminhão ok'), Item_14c_Skid_vagoes e Item_14e_Carregamento_ferroviario (dentro de 'ferrovia, incendio e estaçoes de carga e descarga') e Item_14d_Skid_descarga_caminhoes; o organizar_fotos.py criou as subpastas por TAG (LOG_FOTOS 13/10: 487 fotos organizadas, 105 duplicatas para _REPOSITORIO, 6 copiadas de F:\\, pasta 'Item_14x_Componentes_Metroval_sem_vinculo' extinta). A alocação por skid está nas sínteses 14a, 14b, 14c, 14d e na seção '14 (comp.)'.")],
 obs=[('HL', " LEIA-ME: "), (L, "Q04 ('seguir sem o OPE154') ficou superada – o OPE154 foi lido na v2 (aba COTACOES_OPE154). M14A – NF→subitem não conclusivo (só 14d fecha); data provisória 02/07/2024 para o INPC dos skids 14a–c (fator 1,0936960 prov.), 14d 1,1076498. T04 – vida útil do skid a dividir (estrutura / mecânica / instrumentação), fonte ASA onde o IBAPE não tiver similar. Q05 – software/preset (Toptech) entra na composição de custo do conjunto (máquina + componente + software), quando identificado.")])

add("14a",
 sint=[('HV', "\nComponentes alocados ao TOP LOADING pela árvore de pastas (Rafael, 13/10) – 32 subpastas, 295 fotos\n"),
  (V, "• 7 braços Redlands: BR-3101 a BR-3106 e a pasta 'BR - 3107' (5 fotos) – o 7º braço, ponto de etanol hidratado que faltava; as duas placas lidas nessa pasta trazem séries 3634 e 3635 (= BR-3103/3104 pelo padrão 3631+n) – conferir em resolução cheia.\n"
      "• Válvulas de controle Metroval VCDM (12): FCV-3140 (s/n 0281, DN 100), 3141 (0302, DN 50), 3142 (0289), 3144 (0308), 3145 (0270, DN 80), 3146 (0310), 3147 (0274, DN 80), 3148 (0312), 3149 (0290), 3150 (0317, DN 50), 3151 (0291), 3152 (0318, DN 50).\n"
      "• Medidores turbina MTF-80 (8): FT-3141 (01311123, gasolina), FT-3142 (01?3?1123, S-500), FT-3148 (01421123, etanol hidratado), FT-3151 (01341123, S-500) e 4 sem TAG lida – séries 01441123 (S-500), 01471123 (gasolina), 01491123 (S-10), 01501123 (etanol anidro).\n"
      "• Medidor mássico OI 50 SMPF/F2: FT-3143 (12527, B100).\n"
      "• Transmissores de temperatura Alutal CEX-A-00015: TE-3101A, 3141A, 3147A, 3149A, 3152A (+3 em _a_identificar com TAG 'TE-31?1A').\n"
      "• Monitores de transbordamento/aterramento KFW SLA-S-Y: XS-3103A, XS-3104 (s/n 2011011), XS-3104B.\n"
      "• Filtros MKG (6): F-0368 (FP8S, fabr. 44895-3-6, S500), F-0415 (FP8S, 44895-3-8, S10), F-0545 (44895-3-7, gasolina), F-0635 (FP4S, 44895-1-3, etanol), F-0724 (FP6S, 44895-2-2, etanol), F-0835 (FP4S, 44895-1-4, B100).\n"
      "• Pastas auxiliares: 00_vistas_gerais (78), _a_identificar_instrumentos (45: Alutal ilegíveis, turbinas MTF-P, presets Toptech), _sem_TAG_caixas_juncao (2).")],
 obs=[('HL', " LEIA-ME: "), (L, "A14/A15 – os 'TAGs' dos filtros são números de linha da LI-3100-6105; destinos pela lista: F-0545/0635/0724/0835 → U-3103; F-0415 → B-3120; F-0368 não consta na rev. 00A04; F-0545 tem placa 8\" numa linha 6\" (registrar, não corrigir). "),
      ('HV', "Árvore: "), (V, "a alocação de Rafael bate com o projeto – os 4 filtros de U-3103 estão todos em 14a (U-3103 = top loading); F-0415 (linha da B-3120, CETRAN 02) e F-0368 (sem linha) também ficaram em 14a: manter como Rafael pôs e registrar. FT-3151 aparece em 14a (4 fotos, série 01341123, S-500) e em 14b (3 fotos): a foto 0315 leu 'FT-3151' com série 01401123 e etanol hidratado – são dois medidores, e o de etanol hidratado é a série 'órfã' das fotos de terceiros (A39). FT-3141: a foto 0667 lê série 01311123 e a 0670 lê 01371123 – reler. Com o 'BR - 3107' a divergência '7 pontos × 6 braços' se resolve, se a placa confirmar o 7º braço.")])

add("14b",
 sint=[('HV', "\nComponentes alocados ao BOTTOM LOADING pela árvore de pastas (Rafael, 13/10) – 36 subpastas, 208 fotos\n"),
  (V, "• Presets Toptech MultiLoad II Ex: BOTTOM-3107 (preset A), BOTTOM-3108 (preset B).\n"
      "• Válvulas on-off Metroval VOOM-100 (4): SDV-3103 (s/n 0331), SDV-3111 (0336), SDV-3112 (0338; uma foto lê 0339), SDV-3113 (0339) – as 4 SDV lidas em campo são do bottom loading, não do carregamento ferroviário.\n"
      "• Válvulas de controle VCDM (9): FCV-3155 (0293), 3156 (0321, DN 50), 3157 (0294), 3158 (0271, DN 80), 3159 (0295), 3160 (0272, DN 80), 3161 (0307), 3164 (0311), 3165 (0323, DN 50).\n"
      "• Medidores turbina MTF-80 (8): FT-3151 (2º medidor – ver obs.), FT-3159 (01451123, gasolina) e 6 sem TAG lida – séries 01291123 (etanol anidro), 01381123 (S-10), 01391123 (S-10), 01411123 (etanol anidro), 01431123 (S-500), 01481123 (S-500).\n"
      "• Medidores mássicos OI 50 SMPF/F2 (3): FT-3154 (12529), FT-3156 (12531), FT-3163 (12528) – B100.\n"
      "• Transmissor de temperatura Alutal: TE-3159A (etiqueta de linha).\n"
      "• Monitores KFW SLA-S-Y: XS-3107A, XS-3107B, XS-3108A (s/n 21022?42).\n"
      "• Filtros MKG (6): F-0351 (FP8S, 44895-3-3, S500), F-0455 (FP8S, 44895-3-5, S10), F-0548 (FP8S, 44895-3-4, gasolina), F-0638 (FP4S, 44895-1-1, etanol), F-0725 (FP6S, 44895-2-1, etanol), F-0841 (FP4S, 44895-1-2, B100).\n"
      "• Pastas auxiliares: 00_vistas_gerais (66), _a_identificar_instrumentos (53: Alutal ilegíveis, turbinas MTF-P, preset Toptech EXL, 2 FCV com TAG parcial).")],
 obs=[('HL', " LEIA-ME: "), (L, "A14 – os filtros de U-3104 pela LI-3100-6105 são F-0351/0455/0548/0638/0725/0841, exatamente os que Rafael pôs em 14b (U-3104 = bottom loading); A15 – F-0351/0455/0548 têm placa 8\" em linha 6\" e F-0725 placa 6\" em linha 4\". M14B – os 7 braços bottom continuam sem foto: não subir o grau de vistoria. A39 – SDV-3114 (provável), FCV-3153 e FCV-3162, vistas só em fotos de terceiros, estão na faixa 3140–3165 deste skid ou do top loading. "),
      ('HV', "Árvore: "), (V, "as SDV deixam de ser atribuídas ao ferroviário (ver 14e). FT-3151: 3 fotos aqui e 4 em 14a – dois medidores com a mesma TAG lida (um deles é o de série 01401123, etanol hidratado).")])

add("14c",
 sint=[('HV', "\nÁrvore de fotos (13/10): "), (V, "'ferrovia, incendio e estaçoes de carga e descarga/Item_14c_Skid_vagoes' – FCV-3180 (5), 3181 (4), 3182 (5), 3183 (7), FT-3180 (3), 3181 (5), 3182 (2), TE-3181A (2), TE-3183A (2), 00_vistas_gerais (7) com subpastas 'JB-3106P' (4: caixas de junção Ex TELBRA TMX-80GR, série 00027, JB-3106D e JB-3106P) e 'PLCVT-01_U-3105' (3: placa do skid). Rafael também pôs os filtros MKG F-0829 (FP8S, 44895-3-1, B100; 3 fotos) e F-0629 (FP8S, 44895-3-2, etanol; 3 fotos) dentro de 14c (pasta 'filtro f-0829/filtro f-0629'). Inventário: a FCV-3183 estava com aviso 'EQUIPAMENTO EM MANUTENÇÃO' na vistoria de 17/09.")],
 obs=[('HL', " LEIA-ME: "), (L, "A14 – F-0629 e F-0829 são as linhas de 8\" de etanol anidro e biodiesel da LI-3100-6105: coerente com os produtos do skid de vagões. G02 – equipamento parado = em manutenção (planta em operação, equipe de manutenção presente): registrar no item 08 do laudo. "),
      ('HV', "Árvore: "), (V, "fotos 154429/154604 lidas como JB-3106D dentro da pasta JB-3106P (REVISAR).")])

add("14d",
 sint=[('HV', "\nÁrvore de fotos (13/10): "), (V, "'Item_14d_Skid_descarga_caminhoes' – 00_vistas_gerais (11: caminhões-tanque em posição, braços, medidores, caminhão ONU 33/1170 – etanol), 6 fotos na raiz e 'Item_13_Bombas_megacpk_125-100-250_2un' (B-3124, B-3125) dentro do skid. As 6 fotos da raiz são os monitores KFW SLA-S-Y XS-31209 (s/n 911?03?), XS-31211 e XS-31212 (s/n 2011002?): o inventário marcou CONFLITO (TAG da série 312xx é do ferroviário; pasta é rodoviária) e Rafael não decidiu; como foram fotografados na descarga de caminhões (15/09, 14:14–14:17), ficam alocados a 14d até decisão.")],
 obs=[('HL', " LEIA-ME: "), (L, "M14D – NF 1551 fecha exato, mas o emitente PHD Master ≠ Metroval (N29): registrar; 14d é o único subitem da linha 14 com NF conclusiva.")])

add("14e",
 sint=[('HV', "\nÁrvore de fotos (13/10): "), (V, "'Item_14e_Carregamento_ferroviario' tem só 00_vistas_gerais (6: linhas Gasolina A, Diesel S-10 e S-500 da área de descarga ferroviária, sem placa); as plataformas ficaram como pastas irmãs 'plataforma 08' a 'plataforma 14' (14, 10, 11, 16, 6, 9 e 6 fotos – placas TSPRO BTC-42CA.0466-4; rastreamento individual -03 legível na plataforma 10) e '00_vista_geral' (41) dentro de 'ferrovia, incendio e estaçoes de carga e descarga'. Nenhuma SDV, sensor ou alvo ficou nesta pasta: as SDV-3103/3111/3112/3113 foram alocadas por Rafael ao bottom loading (14b). Inventário: fotos 160127/160134 lidas como PLATAFORMA 12 dentro da pasta 'plataforma 11' (REVISAR).")],
 obs=[('HL', " LEIA-ME: "), (L, "M14E/A43/A33 – os 14 braços TSPRO (BR-3130..3143 pela lista de linhas) são inferência, LADO A não fotografado: maior pendência de campo do M&E; existe placa 'PLATAFORMA 6 LADO A' (e 'LADO B') em foto de terceiros de 02/10/2025 (complementar), fora da faixa 8–14; o 'Lado A/B' do ANP se refere às bombas, não às posições da plataforma. "),
      ('HV', "Árvore: "), (V, "as 10 SDV da crítica continuam sem foto – nenhuma das SDV lidas pertence a esta linha pela árvore de Rafael. Para o Anexo IV: braço TSPRO da plataforma 10 (placa legível) + placa de posição.")])

add("14 (comp.)",
 nome=[(V, " → v4: alocados a 14a / 14b / 14c / 14d pela árvore de pastas de Rafael (13/10)")],
 sint=[('HV', "\nAlocação pela árvore de pastas (Rafael, 13/10)\n"),
  (V, "Top loading (14a): FCV-3140/3141/3142/3144/3145/3146/3147/3148/3149/3150/3151/3152; FT-3141/3142/3143/3148/3151; turbinas sem TAG 01441123, 01471123, 01491123, 01501123; TE-3101A/3141A/3147A/3149A/3152A; XS-3103A/3104/3104B; filtros F-0368/0415/0545/0635/0724/0835. "
      "Bottom loading (14b): FCV-3155/3156/3157/3158/3159/3160/3161/3164/3165; FT-3151 (2º)/3154/3156/3159/3163; turbinas sem TAG 01291123, 01381123, 01391123, 01411123, 01431123, 01481123; SDV-3103/3111/3112/3113; TE-3159A; XS-3107A/3107B/3108A; filtros F-0351/0455/0548/0638/0725/0841. "
      "Descarga de caminhões (14d): XS-31209/31211/31212 (conflito de região no inventário). Vagões (14c): filtros F-0629/F-0829.\n"),
  ('HL', "Componentes vistos só em fotos de terceiros (14/11/2025 – foto complementar; LEIA-ME A39/L3-09)\n"),
  (L, "FCV-3153 (provável) VCDM-50 s/n 0322; FCV-3162 VCDM-100 s/n 031?; 2 VCDM-100 com TAG ilegível, s/n 0292 e 0309 (terceiros leram 'FCV-3152' na de 0309, mas a FCV-3152 de campo é 0318); SDV-3114 (provável) VOOM-100; FT-3165 mássico OI 50 SMPF/F2 nº 12526 (B100); turbinas MTF-80 séries 01461123 (gasolina) e 01401123 (etanol hidratado – a série 'órfã' existe). Com estes, a faixa FCV-3140…3165 soma 26 válvulas de controle vistas (21 em campo + 5 de terceiros).")],
 obs=[('HL', " LEIA-ME: "), (L, "L3-09 – nenhum documento (crítica, RT0538, plano, NFs Metroval, SAP-seleção, LI/LM) lista válvulas e medidores Metroval por TAG: pedir à Metroval/proponente o data-book (de-para TAG × Nr.S) ou a base SAP completa; no laudo, caracterizar como componentes dos skids comprovados por foto de placa. A40/A41 – as divergências de Nr.S 0311 (FCV-3154 × FCV-3164) e 0272 (FCV-3130 × FCV-316?) já foram resolvidas pela releitura (FCV-3164 e FCV-3160); modelo dos mássicos = OI 50 SMPF/F2. "),
      ('HV', "Árvore: "), (V, "a seção deixa de ser 'sem vínculo' – o vínculo é o de campo (skid onde Rafael fotografou), não documental. Pendências em _a_identificar_instrumentos: 45 fotos em 14a e 53 em 14b (Alutal ilegíveis, 2 FCV com TAG parcial, turbinas MTF-P, presets) – subir em resolução cheia.")])

add("15-17",
 sint=[('HV', "\nÁrvore de fotos (13/10): "), (V, "pastas 'tanques/Item_15_Selos_Romao_14,4x13,33m', 'Item_16_Selos_Romao_9,6x11,45m' e 'Item_17_Selos_Romao_9,6x7,63m' criadas, todas com 0 fotos; as fotos externas dos tanques estão em 'tanques/TQ-3101' a 'TQ-3112' (11 a 15 cada) e 00_vistas_gerais (7).")],
 obs=[('HL', " LEIA-ME: "), (L, "D06/M15–M17/L3-10 – os selos continuam os únicos itens em vermelho ('declarados encontrados pelo vistoriador', R$ 546.941,16), sem foto nem certificado; A21 – o valor depende só da declaração no laudo; T05 – vida útil sem fonte: pedir à Romão a vida de projeto (estrutura flutuante × vedação).")])
for ln in ("15", "16", "17"):
    add(ln, sint=[('HV', "\nÁrvore (13/10): "), (V, "pasta Item_%s vazia (0 fotos); fotos externas dos tanques em tanques/TQ-xxxx." % ln)])

add("18",
 sint=[('HV', "\nÁrvore de fotos (13/10): "), (V, "'Item_18_Eletrocentro_VOLGA' com subpastas para os 26 componentes do as-built; com fotos só 00_vistas_gerais (16: painéis internos VOLGA/Schneider, vista externa dos contêineres), CTS-3103 (3), PN-3105 (3), PNA-3102 (1), PNA-3103 (1), PNA-3106 (3) e _sem_TAG_paineis (5: painel aspersores rodoviário, placas VOLGA ilegíveis, painel 'BOMBA DE DRENAGEM / ELETROCENTRO PN-311?'). As pastas QGMT-3100-01, SMF, TR-3100, PN-3100 lados A/B, BC-3100, TL-3100, QGBT-3100, TR.NB-3100, NB-3100, BB.NB-3100, QL-3100, QD-UPS-3100, HVAC-WM, BW.01_BW.02, EXT_GD.TR-3100, PN-AUT e 'Eletrocentro (container)' estão VAZIAS: as placas lidas são fotos de terceiros (WhatsApp 21 e 23/09/2026), que não estão na árvore ANEXO_IV_ME_fotos.")],
 obs=[('HL', " LEIA-ME: "), (L, "A34/A35/A36/A37/L3-07/L3-11 – 11 de 13 equipamentos com placa conferida com o as-built DE-3100-4300-014 rev. 04; as TAGs são '-3100' (a crítica usa '-3101'); nome do cliente diverge entre placas; as fotos de terceiros valem como evidência documental (Anexo VI) e entram no Anexo IV como 'foto complementar'; as 15 fotos de Pictures\\volga são idênticas (SHA-1) às já lidas e não trazem TR-3100, TR.NB-3100 nem placa geral – continuam no pedido ao proponente. T06 – vida útil pelo IBAPE (transformador de força 40 a, distribuição 25 a). "),
      ('HV', "Árvore: "), (V, "CTS-3103, PNA-3102, PNA-3103 e PNA-3106 estão na pasta do eletrocentro, mas são painéis Chronus (OE-638-02, OE-620-04/05/01) fora do as-built VOLGA – Outros Fixos, como o PNA 3105. Para o Anexo IV: copiar as 15 fotos de terceiros para as subpastas vazias com a legenda 'foto complementar'.")])

ws = wb['SINTESE']
ws['A1'].value = 'SÍNTESE por conjunto (v3 – 13/10/2026; v4 – 13/10/2026 com LEIA-ME e árvore de pastas): descrição do conjunto (crítica + placa), componentes (função, marca, modelo, capacidade, série), notas fiscais e cotação; divergências na coluna Observações; valores da crítica, das NFs e da cotação para comparação.'
ws['A2'].value = 'Cores: preto = crítica de orçamento | azul = placa/campo/ANP | vermelho = NF (RT0538) | verde = SAP | amarelo = projeto | azul escuro = cotação OPE154 | roxo = avaliador (divergências, correções, proposições) | LILÁS = o que mudou pelo 00_LEIA_ME de 05/10 (briefing unificado + pré-análise única) | VERDE CLARO = o que mudou pela árvore de pastas ANEXO_IV_ME_fotos organizada por Rafael em 13/10 (LOG_REPOSITORIO / LOG_FOTOS / inventário). Sem texto tachado. Valores em branco = não atribuível ao subitem.'
ws['A2'].font = Font(name='Arial', size=9, color=P)
rows_by_linha = {str(ws.cell(row=r, column=1).value): r for r in range(5, ws.max_row + 1) if ws.cell(row=r, column=1).value}
for linha, d in ADD.items():
    r = rows_by_linha[linha]
    if d['nome']: append_rich(ws.cell(row=r, column=3), d['nome'])
    if d['sint']: append_rich(ws.cell(row=r, column=4), d['sint'])
    if d['obs']: append_rich(ws.cell(row=r, column=5), d['obs'])
    ws.row_dimensions[r].height = 409

# =====================================================================================================
# 2. CARACTERIZACAO – colunas Y (lilás) e Z (verde claro), realocação dos componentes "14 (comp.)",
#    linhas novas (BR - 3107, filtros MKG, caixas de junção, itens só de terceiros)
# =====================================================================================================
wc = wb['CARACTERIZACAO']
HEAD_FILL = PatternFill('solid', fgColor='1F3864')
for col, title, color in ((25, 'LEIA-ME 05/10 – briefing unificado e pré-análise única (v4 – lilás)', L), (26, 'Árvore de pastas ANEXO_IV_ME_fotos / LOG_REPOSITORIO 13/10 (v4 – verde claro)', V)):
    c = wc.cell(row=3, column=col, value=title)
    c.font = Font(name='Arial', size=9, bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor=color); c.alignment = Alignment(wrap_text=True, vertical='center'); c.border = BORDER
    wc.column_dimensions[get_column_letter(col)].width = 70
wc['A1'].value = (wc['A1'].value or '') + ' | v4 – 13/10/2026: colunas Y (LEIA-ME, lilás) e Z (árvore de pastas/log, verde claro); coluna A dos componentes da linha 14 realocada pela árvore (verde claro); linhas novas no fim da aba.'

def plain_from(parts, color):
    """Texto simples (uma cor) para as colunas Y/Z, a partir das partes da síntese."""
    return ''.join(t for c, t in parts if c in (color, 'H' + ('L' if color == L else 'V')))

carac_rows = {}
for r in range(4, wc.max_row + 1):
    a, b = wc.cell(row=r, column=1).value, wc.cell(row=r, column=2).value
    if a is not None and b in ('CONJUNTO', 'GRUPO', 'COMPONENTES'):
        carac_rows.setdefault(str(a), r)
for linha, d in ADD.items():
    r = carac_rows.get(linha)
    if r is None: continue
    ytxt = plain_from(d['sint'] + d['obs'], L).strip(); ztxt = plain_from(d['sint'] + d['obs'], V).strip()
    if ytxt: wc.cell(row=r, column=25, value=ytxt).font = Font(name='Arial', size=9, color=L)
    if ztxt: wc.cell(row=r, column=26, value=ztxt).font = Font(name='Arial', size=9, color=V)
    for col in (25, 26):
        wc.cell(row=r, column=col).alignment = TOP; wc.cell(row=r, column=col).border = BORDER

# Realocação dos componentes "14 (comp.)" pela árvore (TAG ou nº de série -> skid, pasta, nº de fotos)
ALOC = {  # TAG/série -> (linha nova, pasta, fotos)
 'FCV-3140': ('14a', 'Item_14a_Top_loading/FCV-3140', 6), 'FCV-3141': ('14a', 'Item_14a_Top_loading/FCV-3141', 6), 'FCV-3142': ('14a', 'Item_14a_Top_loading/FCV-3142', 2),
 'FCV-3144': ('14a', 'Item_14a_Top_loading/FCV-3144', 3), 'FCV-3145': ('14a', 'Item_14a_Top_loading/FCV-3145', 2), 'FCV-3146': ('14a', 'Item_14a_Top_loading/FCV-3146', 7),
 'FCV-3147': ('14a', 'Item_14a_Top_loading/FCV-3147', 5), 'FCV-3148': ('14a', 'Item_14a_Top_loading/FCV-3148', 3), 'FCV-3149': ('14a', 'Item_14a_Top_loading/FCV-3149', 4),
 'FCV-3150': ('14a', 'Item_14a_Top_loading/FCV-3150', 4), 'FCV-3151': ('14a', 'Item_14a_Top_loading/FCV-3151', 2), 'FCV-3152': ('14a', 'Item_14a_Top_loading/FCV-3152', 3),
 'FT-3141': ('14a', 'Item_14a_Top_loading/FT-3141', 9), 'FT-3142': ('14a', 'Item_14a_Top_loading/FT-3142', 2), 'FT-3143': ('14a', 'Item_14a_Top_loading/FT-3143', 2),
 'FT-3148': ('14a', 'Item_14a_Top_loading/FT-3148', 3), 'FT-3151': ('14a', 'Item_14a_Top_loading/FT-3151 (4) e Item_14b_Bottom_loading/FT-3151 (3)', 7),
 '01441123': ('14a', 'Item_14a_Top_loading/FT_sem_TAG_serie_01441123', 2), '01471123': ('14a', 'Item_14a_Top_loading/FT_sem_TAG_serie_01471123', 6),
 '01491123': ('14a', 'Item_14a_Top_loading/FT_sem_TAG_serie_01491123', 3), '01501123': ('14a', 'Item_14a_Top_loading/FT_sem_TAG_serie_01501123', 3),
 'TE-3101A': ('14a', 'Item_14a_Top_loading/TE-3101A', 1), 'TE-3141A': ('14a', 'Item_14a_Top_loading/TE-3141A', 2), 'TE-3147A': ('14a', 'Item_14a_Top_loading/TE-3147A', 1),
 'TE-3149A': ('14a', 'Item_14a_Top_loading/TE-3149A', 1), 'TE-3152A': ('14a', 'Item_14a_Top_loading/TE-3152A', 2),
 'XS-3103A': ('14a', 'Item_14a_Top_loading/XS-3103A', 2), 'XS-3104': ('14a', 'Item_14a_Top_loading/XS-3104', 3), 'XS-3104B': ('14a', 'Item_14a_Top_loading/XS-3104B', 4),
 'FCV-3155': ('14b', 'Item_14b_Bottom_loading/FCV-3155', 1), 'FCV-3156': ('14b', 'Item_14b_Bottom_loading/FCV-3156', 7), 'FCV-3157': ('14b', 'Item_14b_Bottom_loading/FCV-3157', 2),
 'FCV-3158': ('14b', 'Item_14b_Bottom_loading/FCV-3158', 3), 'FCV-3159': ('14b', 'Item_14b_Bottom_loading/FCV-3159', 4), 'FCV-3160': ('14b', 'Item_14b_Bottom_loading/FCV-3160', 6),
 'FCV-3161': ('14b', 'Item_14b_Bottom_loading/FCV-3161', 3), 'FCV-3164': ('14b', 'Item_14b_Bottom_loading/FCV-3164', 2), 'FCV-3165': ('14b', 'Item_14b_Bottom_loading/FCV-3165', 4),
 'FT-3154': ('14b', 'Item_14b_Bottom_loading/FT-3154', 2), 'FT-3156': ('14b', 'Item_14b_Bottom_loading/FT-3156', 2), 'FT-3159': ('14b', 'Item_14b_Bottom_loading/FT-3159', 5), 'FT-3163': ('14b', 'Item_14b_Bottom_loading/FT-3163', 2),
 '01291123': ('14b', 'Item_14b_Bottom_loading/FT_sem_TAG_serie_01291123', 2), '01381123': ('14b', 'Item_14b_Bottom_loading/FT_sem_TAG_serie_01381123', 4),
 '01391123': ('14b', 'Item_14b_Bottom_loading/FT_sem_TAG_serie_01391123', 3), '01411123': ('14b', 'Item_14b_Bottom_loading/FT_sem_TAG_serie_01411123', 4),
 '01431123': ('14b', 'Item_14b_Bottom_loading/FT_sem_TAG_serie_01431123', 3), '01481123': ('14b', 'Item_14b_Bottom_loading/FT_sem_TAG_serie_01481123', 3),
 'SDV-3103': ('14b', 'Item_14b_Bottom_loading/SDV-3103', 2), 'SDV-3111': ('14b', 'Item_14b_Bottom_loading/SDV-3111', 3), 'SDV-3112': ('14b', 'Item_14b_Bottom_loading/SDV-3112', 4), 'SDV-3113': ('14b', 'Item_14b_Bottom_loading/SDV-3113', 3),
 'XS-3107A': ('14b', 'Item_14b_Bottom_loading/XS-3107A', 2), 'XS-3107B': ('14b', 'Item_14b_Bottom_loading/XS-3107B', 2), 'XS-3108A': ('14b', 'Item_14b_Bottom_loading/XS-3108A', 2),
 'XS-31209': ('14d', 'Item_14d_Skid_descarga_caminhoes (raiz)', 2), 'XS-31211': ('14d', 'Item_14d_Skid_descarga_caminhoes (raiz)', 2), 'XS-31212': ('14d', 'Item_14d_Skid_descarga_caminhoes (raiz)', 2),
}
SKID_NOME = {'14a': 'skid + plataforma TOP LOADING', '14b': 'unidade BOTTOM LOADING', '14d': 'skid de DESCARGA de caminhões', '14c': 'skid de carregamento de vagões'}
n_realoc = 0
for r in range(4, wc.max_row + 1):
    if str(wc.cell(row=r, column=1).value) != '14 (comp.)': continue
    tag = str(wc.cell(row=r, column=3).value); serie = str(wc.cell(row=r, column=8).value)
    key = tag if tag in ALOC else (serie if serie in ALOC else None)
    if key is None: continue
    novo, pasta, nf = ALOC[key]
    wc.cell(row=r, column=1).value = f'{novo} (comp.)'; wc.cell(row=r, column=1).font = Font(name='Arial', size=9, color=V, bold=True)
    txt = f"Era '14 (comp.)' (sem vínculo documentado). Árvore de Rafael (13/10): pasta 'Estação de carga e descarga Caminhão ok/{pasta}' ({nf} fotos) → componente do {SKID_NOME[novo]} ({novo})."
    if novo == '14d': txt += " Inventário: CONFLITO – a TAG da série 312xx é da região ferroviária (14e) e a pasta é rodoviária; decisão de Rafael pendente."
    if key == 'FT-3151': txt += " Duas pastas com a mesma TAG: 14a (série 01341123, diesel S-500) e 14b (foto 0315 lê série 01401123, etanol hidratado) – dois medidores; o 2º é a série 'órfã' das fotos de terceiros."
    if key == 'FT-3141': txt += " Inventário: foto 0667 lê série 01311123 e 0670 lê 01371123 – reler."
    if key == 'FCV-3151': txt += " A placa na pasta lê 'TAG FCV-3161 / Nr.S 0291' (leitura antiga); a plaqueta de TAG lê FCV-3151."
    if key == 'SDV-3112': txt += " Uma foto da pasta lê Nr.S 0339 (conferir com 0496 = 0338)."
    c = wc.cell(row=r, column=26, value=txt); c.font = Font(name='Arial', size=9, color=V); c.alignment = TOP; c.border = BORDER
    n_realoc += 1
# componentes já vinculados (14a braços, 14b presets, 14c, 14e, 18) – pasta na coluna Z
PASTA = {
 'BR-3101': ('Item_14a_Top_loading/BR-3101', 7), 'BR-3102': ('Item_14a_Top_loading/BR-3102', 4), 'BR-3103': ('Item_14a_Top_loading/BR-3103', 4), 'BR-3104': ('Item_14a_Top_loading/BR-3104', 2),
 'BR-3105': ('Item_14a_Top_loading/BR-3105', 6), 'BR-3106': ('Item_14a_Top_loading/BR-3106', 5), 'BOTTOM-3107': ('Item_14b_Bottom_loading/BOTTOM-3107', 1), 'BOTTOM-3108': ('Item_14b_Bottom_loading/BOTTOM-3108', 1),
 '(sem TAG) – Preset de carregamento': ('Item_14b_Bottom_loading/_a_identificar_instrumentos (preset Toptech MultiLoad II EXL)', 2),
 'FT-3180': ('Item_14c_Skid_vagoes/FT-3180', 3), 'FT-3181': ('Item_14c_Skid_vagoes/FT-3181', 5), 'FT-3182': ('Item_14c_Skid_vagoes/FT-3182', 2),
 'FCV-3180': ('Item_14c_Skid_vagoes/FCV-3180', 5), 'FCV-3181': ('Item_14c_Skid_vagoes/FCV-3181', 4), 'FCV-3182': ('Item_14c_Skid_vagoes/FCV-3182', 5), 'FCV-3183': ('Item_14c_Skid_vagoes/FCV-3183 (aviso EQUIPAMENTO EM MANUTENÇÃO na vistoria)', 7),
 'TE-3181A': ('Item_14c_Skid_vagoes/TE-3181A', 2), 'TE-3183A': ('Item_14c_Skid_vagoes/TE-3183A', 2),
 'BR ferrov. (lote 1)': ('ferrovia, incendio e estaçoes de carga e descarga/plataforma 08 a 14', 72), 'BR ferrov. (lote 2)': ('ferrovia, incendio e estaçoes de carga e descarga/plataforma 08 a 14', 72),
 'Plataformas 8 a 14 – lado B': ('ferrovia, incendio e estaçoes de carga e descarga/plataforma 08 (14), 09 (10), 10 (11), 11 (16), 12 (6), 13 (9), 14 (6)', 72),
 'CTS-3103': ('Item_18_Eletrocentro_VOLGA/CTS-3103 – painel Chronus, fora do as-built VOLGA', 3), 'PN-3105': ('Item_18_Eletrocentro_VOLGA/PN-3105', 3),
 'PNA 3105': ('Item_18_Eletrocentro_VOLGA/PNA 3105 (vazia – foto é de terceiros, WhatsApp)', 0), 'Painel VOLGA sem TAG': ('Item_18_Eletrocentro_VOLGA/_sem_TAG_paineis', 5),
 'Motor diesel do GMG': ('Item_01_Grupo_gerador_STEMAC ok/Grupo gerador (motor)', 9), 'Alternador (gerador) do GMG': ('Item_01_Grupo_gerador_STEMAC ok/Grupo gerador (motor)', 9),
 'B-3101A': ('Sintema combate a Incendio Casa de Máquinas/Item_02_Bombas_SCI_megacpk_250-200-500/BOMBA 3101-A', 6), 'B-3101B': ('.../Item_02_Bombas_SCI_megacpk_250-200-500/BOMBA 3101-B', 9),
 'TQ-3122 / TQ-3123 (tanques diesel das motobombas)': ('.../Item_02_Bombas_SCI_megacpk_250-200-500/Tanques diesel SCI (2)', 11),
 'B-3109': ('Centran 1 ok/Item_05_Bomba_megacpk_100-065-315_1un/B-3109', 4), 'B-3111': ('Centran 1 ok/Item_06_Bombas_megacpk_125-080-315_2un/B-3111', 5), 'B-3112': ('Centran 1 ok/Item_06_Bombas_megacpk_125-080-315_2un/B-3112', 8),
 'B-3118': ('Centran 2 ok/Item_07_Bombas_megacpk_125-100-315_2un/B-3118', 5), 'B-3119': ('Centran 2 ok/Item_07_Bombas_megacpk_125-100-315_2un/B-3119 (1 foto lida como B-3120 – REVISAR)', 14),
 'B-3120': ('Centran 2 ok/Item_08_Bombas_megacpk_125-100-315_3un/B-3120', 6), 'B-3121': ('Centran 2 ok/Item_08_Bombas_megacpk_125-100-315_3un/B-3121', 6), 'B-3122': ('Centran 2 ok/Item_08_Bombas_megacpk_125-100-315_3un/B-3122', 7),
 'B-3123': ('Centran 2 ok/Item_09_Bomba_megacpk_125-100-315_1un/B-3123', 14),
 'B-3130': ('sistema ETE/Item_10_Bombas_megaflow_V80-250K_2un ok/B-3130 (2 fotos lidas como B-3131 – REVISAR)', 10), 'B-3131': ('sistema ETE/Item_10_Bombas_megaflow_V80-250K_2un ok/B-3131', 8),
 'B-3132': ('sistema ETE/Item_11_Bombas_megaflow_V80-250K_2un ok/B-3132', 2), 'B-3133': ('sistema ETE/Item_11_Bombas_megaflow_V80-250K_2un ok/B-3133', 2),
 'B-31100': ('Item_12_Bomba_megacpk_050-032-200 ok/B-31100 (+2 fotos na pasta "Sitema de água potável artesiano ok" – CONFLITO)', 6),
 'B-3124': ('Estação de carga e descarga Caminhão ok/Item_14d_Skid_descarga_caminhoes/Item_13_Bombas_megacpk_125-100-250_2un/B-3124', 6),
 'B-3125': ('Estação de carga e descarga Caminhão ok/Item_14d_Skid_descarga_caminhoes/Item_13_Bombas_megacpk_125-100-250_2un/B-3125', 9),
}
for i in range(3103, 3118):
    t = f'B-{i}'
    if t not in PASTA and i not in (3109, 3111, 3112): PASTA[t] = (f'Centran 1 ok/Item_04_Bombas_megacpk_125-080-315_12un/{t}', None)
ARV_N = {'B-3103': 8, 'B-3104': 8, 'B-3105': 5, 'B-3106': 6, 'B-3107': 9, 'B-3108': 7, 'B-3110': 8, 'B-3113': 7, 'B-3114': 6, 'B-3115': 6, 'B-3116': 8, 'B-3117': 6}
for r in range(4, wc.max_row + 1):
    tag = str(wc.cell(row=r, column=3).value)
    if wc.cell(row=r, column=2).value == 'COMPONENTE' and tag in PASTA and wc.cell(row=r, column=26).value is None:
        pasta, nf = PASTA[tag]; nf = ARV_N.get(tag, nf)
        txt = f"Árvore de Rafael (13/10): pasta '{pasta}'" + (f" ({nf} fotos)." if nf is not None else '.')
        c = wc.cell(row=r, column=26, value=txt); c.font = Font(name='Arial', size=9, color=V); c.alignment = TOP; c.border = BORDER
    if wc.cell(row=r, column=2).value == 'COMPONENTE' and str(wc.cell(row=r, column=1).value) == '18' and wc.cell(row=r, column=26).value is None:
        c = wc.cell(row=r, column=26, value=f"Árvore de Rafael (13/10): pasta 'Item_18_Eletrocentro_VOLGA/{tag}' criada, VAZIA – a placa está em foto de terceiros (WhatsApp 21/23-09-2026), fora da árvore ANEXO_IV; copiar como 'foto complementar'.")
        c.font = Font(name='Arial', size=9, color=V); c.alignment = TOP; c.border = BORDER

# Linhas novas no fim da aba (componentes acrescentados na v4)
NEW = [  # (cor, linha, TAG, descrição, marca, modelo, capacidade, série, fabricação, observação, fotos)
 (V, '14a', 'BR - 3107', "Árvore (13/10): pasta 'Item_14a_Top_loading/BR - 3107' criada por Rafael – 7º braço de carregamento superior (ponto de etanol hidratado que faltava na contagem da crítica: 7 pontos). Placas Redlands lidas nessa pasta: modelo CC3201408310, Ø 4\", ano 24, séries 3634 e 3635.", 'Redlands', 'CC3201408310', 'Ø 4"', '3634 / 3635 lidas (= BR-3103/3104 pelo padrão 3631+n – DIVERGÊNCIA, conferir em resolução cheia)', '2024', 'Inventário: DIVERGENCIA de série – subir resolução cheia. Até confirmar, o 7º braço fica como "visto, placa a conferir".', "pasta 'BR - 3107' (5 fotos)"),
 (V, '14a', 'F-0368', "Filtro MKG (árvore: Item_14a_Top_loading/F-0368). LEIA-ME A14: F-0368 não consta na LI-3100-6105 rev. 00A04.", 'MKG Equipamentos', 'FP8S', 'Conexões 8" 150#; PMTA 20,1 kgf/cm²; PTH 30,3; vazão -/120/240; NR-13; diesel S500', '44895-3-6', '2024', 'Filtro tipo cesto; TAG = nº de linha (LEIA-ME A14).', '0589'),
 (V, '14a', 'F-0415', "Filtro MKG (árvore: Item_14a_Top_loading/F-0415). LEIA-ME A14: linha F-0415 da LI-3100-6105 tem destino B-3120 (CETRAN 02) – Rafael o pôs no top loading; registrar.", 'MKG Equipamentos', 'FP8S', 'Conexões 8" 150#; PMTA 20,1; PTH 30,3; vazão -/120/240; diesel S10', '44895-3-8', '2024', 'TAG = nº de linha; destino de projeto B-3120 × pasta 14a.', '0575, 0576'),
 (V, '14a', 'F-0545', "Filtro MKG (árvore: Item_14a_Top_loading/F-0545). LEIA-ME A14: linha → U-3103 (top loading); A15: placa 8\" em linha 6\".", 'MKG Equipamentos', 'FP8S (uma foto lê FP6S / 44895-2-1?)', 'Conexões 8" 150#; PMTA 20,1; PTH 30,3; vazão -/120/240; gasolina', '44895-3-7 (foto 0591)', '2024', 'Inventário: DIVERGENCIA entre fotos 0591 e 0727/0728 (FP8S 44895-3-7 × FP6S 44895-2-1?) – reler.', '0591, 0727, 0728'),
 (V, '14a', 'F-0635', "Filtro MKG (árvore: Item_14a_Top_loading/F-0635). LEIA-ME A14: linha → U-3103.", 'MKG Equipamentos', 'FP4S', 'Conexões 4" 150#; PMTA 20,1; PTH 30,3; vazão -/33/66; etanol', '44895-1-3', '2024', '', '0586, 0587'),
 (V, '14a', 'F-0724', "Filtro MKG (árvore: Item_14a_Top_loading/F-0724). LEIA-ME A14: linha → U-3103; A15: placa 6\" em linha 4\".", 'MKG Equipamentos', 'FP6S', 'Conexões 6" 150#; PMTA 20,1; PTH 30,3; vazão -/120/-; etanol', '44895-2-2', '2024', '', '0581, 0582'),
 (V, '14a', 'F-0835', "Filtro MKG (árvore: Item_14a_Top_loading/F-0835). LEIA-ME A14: linha → U-3103.", 'MKG Equipamentos', 'FP4S', 'PMTA 20,1; PTH 30,3; vazão 12/18/72; NR-13 GR.5/CL.A/CAT.III; biodiesel B100', '44895-1-4', '2024', '', '0578, 0579'),
 (V, '14b', 'F-0351', "Filtro MKG (árvore: Item_14b_Bottom_loading/F-0351 – placa completa, CNPJ 05.115.921/0001-13). LEIA-ME A14: linha → U-3104 (bottom loading); A15: placa 8\" em linha 6\".", 'MKG Equipamentos', 'FP8S (placa também lida como FP6S)', 'Conexões 8" 150#; PMTA 20,1; PTH 30,3; oper. 3 kgf/cm²; diesel S500', '44895-3-3', '2024', '', '0424-0426'),
 (V, '14b', 'F-0455', "Filtro MKG (árvore: Item_14b_Bottom_loading/F-0455). LEIA-ME A14: linha → U-3104; A15: placa 8\" em linha 6\".", 'MKG Equipamentos', 'FP8S', 'Conexões 8" 150#; PMTA 20,1; PTH 30,3; vazão -/120/240; diesel S10', '44895-3-5', '2024', '', '0433'),
 (V, '14b', 'F-0548', "Filtro MKG (árvore: Item_14b_Bottom_loading/F-0548). LEIA-ME A14: linha → U-3104; A15: placa 8\" em linha 6\".", 'MKG Equipamentos', 'FP8S', 'Conexões 8" 150#; PMTA 20,1; PTH 30,3; vazão -/120/240; gasolina', '44895-3-4', '2024', '', '0301, 0302'),
 (V, '14b', 'F-0638', "Filtro MKG (árvore: Item_14b_Bottom_loading/F-0638 – placa completa). LEIA-ME A14: linha → U-3104.", 'MKG Equipamentos', 'FP4S', 'Conexões 4" 150#; PMTA 20,1; PTH 30,3; vazão -/33/66; etanol', '44895-1-1', '2024', '', '0363, 0364'),
 (V, '14b', 'F-0725', "Filtro MKG (árvore: Item_14b_Bottom_loading/F-0725 – placa completa). LEIA-ME A14: linha → U-3104; A15: placa 6\" em linha 4\".", 'MKG Equipamentos', 'FP6S', 'Conexões 6" 150#; PMTA 20,1; PTH 30,3; vazão -/120/-; etanol', '44895-2-1', '2024', '', '0360-0362'),
 (V, '14b', 'F-0841', "Filtro MKG (árvore: Item_14b_Bottom_loading/F-0841). LEIA-ME A14: linha → U-3104.", 'MKG Equipamentos', 'FP4S', 'Conexões 4" 150#; PMTA 20,1; PTH 30,3; vazão 12/18/22; biodiesel B100', '44895-1-2', '2024', '', '0435, 0436'),
 (V, '14c', 'F-0829', "Filtro MKG (árvore: Item_14c_Skid_vagoes/filtro f-0829). LEIA-ME A14: linha de 8\" de biodiesel da LI-3100-6105. Inventário: fecha a D13 – último filtro da sequência 44895-3-x.", 'MKG Equipamentos', 'FP8S', 'ASME B16.34 ed. 2017; oper. 3 kgf/cm² / ambiente; PMTA 20,1; biodiesel B100', '44895-3-1', '2024', '', '0883, 0884'),
 (V, '14c', 'F-0629', "Filtro MKG (árvore: Item_14c_Skid_vagoes/filtro f-0829/filtro f-0629). LEIA-ME A14: linha de 8\" de etanol anidro.", 'MKG Equipamentos', 'FP8S', 'ASME B16.34 ed. 2017; oper. 3 kgf/cm² / ambiente; PMTA 20,1; etanol', '44895-3-2', '2024', '', '0839'),
 (V, '14c', 'JB-3106D / JB-3106P', "Caixas de junção Ex TELBRA do skid de vagões (árvore: Item_14c_Skid_vagoes/00_vistas_gerais/JB-3106P; 2 fotos lidas como JB-3106D dentro dela – REVISAR).", 'TELBRA', 'TMX-80GR', 'Ex d IIB Gb / Ex tb IIIC Db', '00027 (JB-3106P)', 'não lido', '3º modelo TELBRA do projeto; instrumentação (Outros Fixos) – componente do skid só como evidência.', '0860, 0864, 0865'),
 (V, '14c', 'PLCVT-01 / U-3105 (placa do skid)', "Placa Metroval do skid (árvore: Item_14c_Skid_vagoes/00_vistas_gerais/PLCVT-01_U-3105, 3 fotos): PLCVT-01 TOP LOADING U-3105, BPA/Palmeirante-TO, plataforma de carregamento de vagão-tanque – placa completa (confirma lote 2, foto 0836).", 'Metroval', 'PLCVT-01', '4 braços; 120 m³/h por braço; 10 kgf/cm² projeto', 'sem nº de série', '2024', '', '0836, 1095'),
 (L, '14 (comp.)', 'FCV-3153 (provável)', "LEIA-ME A39/INSTRUMENTOS: válvula vista só em foto de terceiros de 14/11/2025 (WA0310) – foto complementar; TAG provável.", 'Metroval', 'VCDM-50', 'DN 50; combustíveis', '0322', 'não lido', 'Não fotografada em campo; skid não identificado (faixa 3140–3165 = 14a/14b).', 'IMG-20251114-WA0310 (terceiros)'),
 (L, '14 (comp.)', 'FCV-3162', "LEIA-ME A39: válvula vista só em foto de terceiros (WA0306) – foto complementar.", 'Metroval', 'VCDM-100', 'DN 100; combustíveis', '031? (ilegível)', 'não lido', 'Não fotografada em campo.', 'IMG-20251114-WA0306 (terceiros)'),
 (L, '14 (comp.)', '(TAG ilegível) – VCDM-100 s/n 0292', "LEIA-ME A39: válvula vista só em foto de terceiros (WA0307).", 'Metroval', 'VCDM-100', 'DN 100', '0292', 'não lido', 'Não fotografada em campo.', 'IMG-20251114-WA0307 (terceiros)'),
 (L, '14 (comp.)', '(TAG ilegível) – VCDM-100 s/n 0309', "LEIA-ME A39/A40: válvula vista só em foto de terceiros (WA0315); terceiros leram 'FCV-3152 (provável)', mas a FCV-3152 de campo tem s/n 0318 – diverge.", 'Metroval', 'VCDM-100', 'DN 100', '0309', 'não lido', 'Não fotografada em campo.', 'IMG-20251114-WA0315 (terceiros)'),
 (L, '14 (comp.)', 'SDV-3114 (provável)', "LEIA-ME A39: válvula on-off vista só em foto de terceiros (WA0314) – foto complementar; TAG provável.", 'Metroval', 'VOOM-100', 'DN 100', 'ilegível', 'não lido', 'Não fotografada em campo; junto com SDV-3103/3111/3112/3113 (bottom loading) somam 5 das 10 SDV da crítica.', 'IMG-20251114-WA0314 (terceiros)'),
 (L, '14 (comp.)', 'FT-3165', "LEIA-ME A39: medidor mássico visto só em foto de terceiros (WA0330) – foto complementar.", 'Metroval', 'OI 50 SMPF/F2', 'Biodiesel B100', '12526', 'não lido', 'Não fotografado em campo; completa a sequência 12526–12531 dos mássicos do rodoviário.', 'IMG-20251114-WA0330 (terceiros)'),
 (L, '14 (comp.)', '(TAG ilegível) – MTF-80 s/n 01461123', "LEIA-ME A39: medidor turbina visto só em foto de terceiros (WA0352).", 'Metroval', 'MTF-80', 'Gasolina', '01461123', 'não lido', 'Não fotografado em campo.', 'IMG-20251114-WA0352 (terceiros)'),
 (L, '14 (comp.)', '(TAG ilegível) – MTF-80 s/n 01401123', "LEIA-ME A39: a série 'órfã' 01401123 existe (etanol hidratado) – foto de terceiros WA0354; a foto de campo 0315 (pasta FT-3151 em 14a) leu essa série com TAG FT-3151.", 'Metroval', 'MTF-80', 'Etanol hidratado', '01401123', 'não lido', 'Provável 2º medidor lido como FT-3151 (ver 14a/14b).', 'IMG-20251114-WA0354 (terceiros); 0315'),
]
r0 = wc.max_row + 2
c = wc.cell(row=r0, column=1, value='COMPONENTES ACRESCENTADOS NA v4 (verde claro = árvore de pastas de Rafael; lilás = LEIA-ME / fotos de terceiros). Sem valor próprio na crítica – entram nos conjuntos indicados na coluna A.')
c.font = Font(name='Arial', size=9, bold=True, color=P)
for i, (col, linha, tag, desc, marca, modelo, cap, serie, fab, obs, fotos) in enumerate(NEW, r0 + 1):
    vals = {1: linha, 2: 'COMPONENTE', 3: tag, 4: desc, 5: marca, 6: modelo, 7: cap, 8: serie, 9: 'não consta', 10: fab, 11: '↑ NF do conjunto', 20: obs, 21: fotos, 26 if col == V else 25: 'linha nova na v4'}
    for cc, vv in vals.items():
        cell = wc.cell(row=i, column=cc, value=vv); cell.font = Font(name='Arial', size=9, color=col, bold=(cc == 3)); cell.alignment = TOP; cell.border = BORDER
    wc.row_dimensions[i].height = 60
last = r0 + len(NEW)
wc.auto_filter.ref = f'A3:Z{last}'

# =====================================================================================================
# 3. Aba ARVORE_FOTOS_v4 (verde claro): árvore de pastas de Rafael × linha da crítica + conflitos do inventário
# =====================================================================================================
def linha_da_pasta(p):
    m = re.search(r'Item_(\d{2})([a-e]?)', p)
    if m: return str(int(m.group(1))) + m.group(2)
    if p.startswith('tanques'): return 'OF (tanques) / 15-17'
    if p.startswith('Sintema'): return '2 / 3 (SCI) – demais itens: Outros Fixos'
    if p.startswith('Sitema de água'): return 'fora do M&E (água potável) – 2 fotos leem B-31100 (12)'
    if 'plataforma' in p: return '14e'
    if p.startswith('ferrovia'): return '14c / 14e (vistas)'
    if p.startswith('Centran 1'): return '4 / 5 / 6 (vistas)'
    if p.startswith('Centran 2'): return '7 / 8 / 9 (vistas)'
    if p.startswith('Estação'): return '14a / 14b / 14d (vistas)'
    if p.startswith('sistema ETE'): return '10 / 11 (tanque da ETE: Outros Fixos)'
    if p.startswith('pipe hack'): return 'Outros Fixos (pipe rack)'
    return ''
wa = wb.create_sheet('ARVORE_FOTOS_v4')
wa['A1'] = 'Árvore de pastas ANEXO_IV_ME_fotos organizada por Rafael (13/10/2026) – fonte: 2026-10-13_Inventario_Fotos_AnexoIV.xlsx (aba ARVORE, após o ORGANIZAR), LOG_FOTOS_20261013_003759.csv e LOG_REPOSITORIO.csv. Verde claro = alteração feita com base no log/árvore.'
wa['A1'].font = Font(name='Arial', size=10, bold=True, color=V)
wa['A2'] = 'Válvulas, filtros, medidores e sensores que estavam soltos foram postos por Rafael dentro das pastas dos skids/grupos a que pertencem; o organizar_fotos.py criou as subpastas por TAG. A coluna "Linha da crítica" é a atribuição da v4 a partir da pasta.'
wa['A2'].font = Font(name='Arial', size=9, color=V)
heads = ['Pasta (árvore de Rafael)', 'Fotos + páginas', 'Linha da crítica atribuída', 'TAG / subpasta', 'Observação']
for i, h in enumerate(heads, 1):
    c = wa.cell(row=4, column=i, value=h); c.font = Font(name='Arial', size=9, bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor=V); c.border = BORDER; c.alignment = Alignment(wrap_text=True, vertical='center')
for i, w in enumerate([95, 12, 34, 34, 70], 1): wa.column_dimensions[get_column_letter(i)].width = w
OBS_PASTA = {
 'BR - 3107': '7º braço top loading (pasta criada por Rafael); placas com séries 3634/3635 – conferir.',
 'FT-3151': 'TAG em duas pastas (14a e 14b) – dois medidores.',
 'Item_14d_Skid_descarga_caminhoes': 'Raiz: XS-31209/31211/31212 – CONFLITO de região (TAG ferroviária em pasta rodoviária), decisão pendente.',
 '_a_identificar_instrumentos': 'Pendência: subir em resolução cheia (Alutal, FCV com TAG parcial, MTF-P, presets).',
 'filtro f-0829': 'Filtro MKG F-0829 (44895-3-1) alocado ao skid de vagões.', 'filtro f-0629': 'Filtro MKG F-0629 (44895-3-2) alocado ao skid de vagões (subpasta dentro de f-0829).',
 'B-31100': 'Mais 2 fotos de B-31100 na pasta de água potável (CONFLITO).',
 'Item_13_Bombas_megacpk_125-100-250_2un': 'Rafael colocou o Item 13 dentro do skid de descarga 14d.',
}
r = 5
with open(ARV_CSV, encoding='utf-8') as f:
    for row in csv.reader(f, delimiter=';'):
        pasta, n = row[0], row[1]
        sub = pasta.rsplit('/', 1)[-1]
        obs = OBS_PASTA.get(sub, '')
        if sub.startswith('Item_1') and 'Eletrocentro' not in pasta and n == '0' and 'Selos' in sub: obs = 'Pasta vazia – selos sem foto (vermelho, D06).'
        if 'Eletrocentro' in pasta and n == '0' and sub not in ('Item_18_Eletrocentro_VOLGA', 'sem_TAG'): obs = 'Vazia – placa em foto de terceiros (WhatsApp 21/23-09-2026), fora da árvore.'
        for i, v in enumerate([pasta, int(n), linha_da_pasta(pasta), sub if sub != pasta else '', obs], 1):
            c = wa.cell(row=r, column=i, value=v); c.font = Font(name='Arial', size=9, color=V); c.border = BORDER; c.alignment = TOP
        r += 1
r += 1
wa.cell(row=r, column=1, value='PLANO do inventário – itens sem decisão (CONFLITO / REVISAR / ANP_FALTA) e não alocadas – decisão de Rafael pendente').font = Font(name='Arial', size=9, bold=True, color=V); r += 1
for i, h in enumerate(['Ação', 'Decisão', 'Origem', 'Destino sugerido', 'Motivo'], 1):
    c = wa.cell(row=r, column=i, value=h); c.font = Font(name='Arial', size=9, bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor=V); c.border = BORDER
r += 1
with open(PLANO_CSV, encoding='utf-8') as f:
    for row in csv.reader(f, delimiter=';'):
        for i, v in enumerate(row[:5], 1):
            c = wa.cell(row=r, column=i, value=v); c.font = Font(name='Arial', size=9, color=V); c.border = BORDER; c.alignment = TOP
        r += 1
wa.freeze_panes = 'A5'

# =====================================================================================================
# 4. Aba ACHADOS_v4 (lilás = LEIA-ME; verde claro = log / árvore)
# =====================================================================================================
ACH = [
 (V, 'B01', '14, 14a, 14b, 14c, 14d', 'ALOCAÇÃO', "Rafael colocou as válvulas, filtros, medidores e sensores soltos dentro das pastas dos skids a que pertencem (Item_14a_Top_loading, Item_14b_Bottom_loading, Item_14c_Skid_vagoes, Item_14d_Skid_descarga_caminhoes); o organizar_fotos.py criou 32 + 36 + 13 subpastas por TAG (LOG_FOTOS 13/10: 487 fotos organizadas, 105 duplicatas). A seção '14 (comp.)' (58 instrumentos sem vínculo) fica resolvida pelo vínculo de campo.", 'LOG_FOTOS_20261013 / ARVORE', "Coluna A dos componentes realocada ('14a (comp.)', '14b (comp.)', '14d (comp.)'); listas de componentes nas sínteses 14a/14b/14c/14d."),
 (V, 'B02', '14a', 'ITEM NOVO', "Pasta 'BR - 3107' criada por Rafael (5 fotos): 7º braço de carregamento superior Redlands – fecha a divergência '7 pontos × 6 braços' da crítica, se a placa confirmar; as placas lidas nessa pasta trazem séries 3634 e 3635 (= BR-3103/3104 pelo padrão 3631+n) – conferir em resolução cheia.", 'ARVORE / INVENTARIO (DIVERGENCIA)', 'Linha nova BR - 3107 na CARACTERIZACAO (verde claro).'),
 (V, 'B03', '14b, 14e', 'CORREÇÃO', "As 4 válvulas on-off SDV-3103/3111/3112/3113 lidas em campo estão no BOTTOM LOADING (pasta 14b), não no carregamento ferroviário: a síntese 14e da v3 as atribuía provisoriamente às 10 SDV da crítica. As 10 SDV do ferroviário continuam sem foto.", 'ARVORE', 'Síntese 14b (componentes) e 14e (observação) atualizadas.'),
 (V, 'B04', '14d, 14e', 'CONFLITO', "XS-31209, XS-31211 e XS-31212 (KFW SLA-S-Y) foram fotografados na descarga de caminhões (pasta 14d, 15/09 14:14–14:17), mas o inventário diz que a TAG da série 312xx é do ferroviário; Rafael não decidiu (PLANO/CONFLITO em branco).", 'INVENTARIO/PLANO', "Realocados a '14d (comp.)' até decisão; registrado como conflito."),
 (V, 'B05', '14a, 14b', 'DIVERGÊNCIA', "FT-3151 aparece em duas pastas: 14a (4 fotos, série 01341123, diesel S-500) e 14b (3 fotos; a foto 0315 lê série 01401123 e etanol hidratado). São dois medidores; o de etanol hidratado é a série 'órfã' confirmada nas fotos de terceiros. FT-3141: fotos 0667 e 0670 leem séries 01311123 e 01371123.", 'INVENTARIO (DIVERGENCIA)', 'Reler as duas placas em resolução cheia; linha nova para a série 01401123.'),
 (V, 'B06', '14a, 14b, 14c', 'ALOCAÇÃO', "Filtros MKG (14): 6 em 14a (F-0368, 0415, 0545, 0635, 0724, 0835), 6 em 14b (F-0351, 0455, 0548, 0638, 0725, 0841), 2 em 14c (F-0629, F-0829). Coincide com a LI-3100-6105 (LEIA-ME A14: U-3103 = top loading, U-3104 = bottom loading), exceto F-0415 (linha da B-3120) e F-0368 (sem linha), que Rafael pôs em 14a.", 'ARVORE + LEIA-ME A14', '14 linhas novas de filtros na CARACTERIZACAO (verde claro), com a leitura das placas (INSTRUMENTOS da Caracterizacao v4 de 05/10).'),
 (V, 'B07', '12', 'CONFLITO', "Duas fotos da pasta 'Sitema de água potável artesiano ok' (vista geral e painel) leem B-31100; a pasta sugere que a bomba serve ao sistema de água potável/artesiana, e não aos chuveiros de emergência do projeto. Registrar, não resolver.", 'INVENTARIO/PLANO', 'Observação na síntese 12.'),
 (V, 'B08', '13, 14d', 'CONFIRMAÇÃO', "Rafael colocou 'Item_13_Bombas_megacpk_125-100-250_2un' dentro de 'Item_14d_Skid_descarga_caminhoes': confirma a leitura da v3 (as 2 bombas da linha 13 são as do skid de descarga).", 'ARVORE', 'Sem mudança de valor; nota nas sínteses 13 e 14d.'),
 (V, 'B09', '15-17, 18', 'LACUNA', "Pastas Item_15/16/17 (selos) criadas e vazias; no eletrocentro, 20 das 26 subpastas por componente estão vazias porque as placas lidas são fotos de terceiros (WhatsApp 21/23-09-2026) que não estão na árvore ANEXO_IV_ME_fotos.", 'ARVORE', "Para o Anexo IV: copiar as 15 fotos de terceiros para as subpastas vazias, com legenda 'foto complementar'; selos continuam sem foto."),
 (V, 'B10', '2, 18, OF', 'CONFLITO', "Páginas do relatório ANP em pasta de outra região: 45 (tanques) e 75/76/77 (rodoviário) na pasta do SCI; 58 (tanques) em 'pipe hack de cabos'; 75 na ETE; página 121 (detectores) é também de ELETRO e não está em pasta dessa região.", 'INVENTARIO/PLANO', 'Decisão de Rafael; não altera a caracterização.'),
 (V, 'B11', '7, 10, 14c, 14e', 'REVISAR', "Fotos de uma TAG dentro da pasta de outra: 152835 (B-3120 em B-3119); 152123/152145 (B-3131 em B-3130); 154429/154604 (JB-3106D em JB-3106P); 160127/160134 (PLATAFORMA 12 em 'plataforma 11'). Podem ser vistas de dois equipamentos.", 'INVENTARIO/PLANO', 'Decisão de Rafael.'),
 (V, 'B12', '1, 2', 'FOTOS', "6 fotos da vistoria que não estavam em nenhuma pasta foram copiadas de F:\\ (gerador: 102219/102259/102325; casa de máquinas SCI: 170822/170827/170847); 3 ficaram fora (0040 e 0247 descarte; 0061 acessórios do alternador, ilegível). A pasta 'Item_14x_Componentes_Metroval_sem_vinculo' (só .txt) foi para _REPOSITORIO.", 'LOG_REPOSITORIO 13/10', 'Fotos citadas nas sínteses 1 e 2.'),
 (V, 'B13', '18', 'CLASSIFICAÇÃO', "CTS-3103, PNA-3102, PNA-3103 e PNA-3106 estão na pasta Item_18_Eletrocentro_VOLGA, mas são painéis Chronus (OE-638-02, OE-620-04/05/01) fora do as-built VOLGA – Outros Fixos, como o PNA 3105.", 'ARVORE + INSTRUMENTOS', 'Observação na síntese 18; não entram no valor da linha 18.'),
 (V, 'B14', '14 (comp.)', 'CONFIRMAÇÃO', "O LOG_REPOSITORIO de 05-06/10 registra as releituras já incorporadas na v3: FCV-3130→3160, FCV-3154→3164 (+0387 = FCV-3155), FT-3145→3143, FCV-3151 (0291), FCV-3156 (0321), SDV-3103 (0331) e 11 válvulas 'TAG pendente' resolvidas pelo Nr.S; pastas vazias FCV-3168/SDV-3109/_TAG_pendente foram para _REPOSITORIO.", 'LOG_REPOSITORIO 05-06/10', 'Sem mudança.'),
 (L, 'B15', 'geral', 'FONTE', "O zip 2026-10-11_PRE_ANALISE_UNICA_v3 anexado na v2 era de outro trabalho (fazendas); o 00_LEIA_ME correto (05/10/2026) = BRIEFING_UNIFICADO_Ultracargo.md + PRE_ANALISE_UNICA_Ultracargo.xlsx (blocos 1–7, L3-01 a L3-11, Q01–Q05) + COMO_ORGANIZAR.md + PROMPT_nova_sessao.md + MAPA_PROJETO.csv.", '00_LEIA_ME', 'Nota da v2 sobre o anexo não usado fica superada (ver LEIA-ME).'),
 (L, 'B16', '14a, 14b, 14c', 'CADEIA DOCUMENTAL', "Os 'TAGs' dos filtros MKG F-0351…F-0841 são números de LINHA da LI-3100-6105; produto da placa = produto da linha em 13/13; destinos U-3103 (0545, 0635, 0724, 0835), U-3104 (0351, 0455, 0548, 0638, 0725, 0841), B-3120 (0415), linhas 8\" de etanol anidro e biodiesel (0629, 0829); F-0368 não consta na rev. 00A04; em 6 filtros o diâmetro da placa é maior que o da linha (A15).", 'Briefing §6 / A14 / A15', 'Vínculo documental filtro × skid; coincide com a árvore de Rafael (B06).'),
 (L, 'B17', '14 (comp.)', 'ITENS NOVOS', "Placas Metroval em fotos de terceiros de 14/11/2025 (31 placas lidas): novos FCV-3153 (provável, 0322), FCV-3162 (031?), 2 VCDM-100 com TAG ilegível (0292, 0309), SDV-3114 (provável), FT-3165 (OI 50 SMPF/F2 nº 12526), turbinas MTF-80 01461123 (gasolina) e 01401123 (etanol hidratado – série 'órfã' existe).", 'Briefing §5 / A39 / INSTRUMENTOS', "8 linhas novas na CARACTERIZACAO (lilás), marcadas 'foto complementar'."),
 (L, 'B18', '2', 'NÃO USAR', "Fotos de terceiros de 23/10/2025: motobombas KSB Ômega 150-460A + MTU 447A, placa MB-502 – não batem com as bombas SCI financiadas (MCPK 250-200-500); local a confirmar (A42). Regra Q03: fotos de 14 a 18/09/2026 = vistoria; qualquer outra data = 'foto complementar' na legenda.", 'A42 / Q03', 'Observação na síntese 2; regra aplicada a todas as fotos de terceiros.'),
 (L, 'B19', '10, 11, 15-17', 'STATUS', "Decisão D06: bens não encontrados/sem foto entram em vermelho como 'declarados encontrados'. Com as placas do lote 3 (L3-01/L3-10) as 4 megaflow saem do vermelho; ficam em vermelho só os 6 selos (R$ 546.941,16). M11: NF 543048 excede a linha 11 em R$ 71.060,00 – usar o plano.", 'D06 / L3-01 / L3-10 / M11', 'Observações nas sínteses 10, 11 e 15-17.'),
 (L, 'B20', '12, 7', 'STATUS', "L3-02: linha 12 = B-31100 verificada com placa (OP M482301632 fecha a S19); motor WEG 7,5 kW; serviço não identificado em campo; divergência B-31100 × 'BOMBA 3160' para o proponente. L3-03: B-3119 com placa (OP M482301627) – 6/6 bombas da descarga ferroviária com placa.", 'L3-02 / L3-03', 'Observações nas sínteses 12 e 7.'),
 (L, 'B21', '18', 'STATUS', "A34–A37, L3-07, L3-11: 11 de 13 equipamentos do eletrocentro com placa conferida com o as-built; TAGs '-3100' (crítica '-3101'); nomes de cliente divergentes nas placas; fotos de terceiros = evidência documental (Anexo VI), 'foto complementar' no Anexo IV; faltam TR-3100 (750 kVA), TR.NB-3100 e placa geral – pedido ao proponente.", 'Briefing §4', 'Observação na síntese 18.'),
 (L, 'B22', '14e', 'STATUS', "M14E: 14 braços TSPRO (BR-3130..3143 pela lista de linhas) são inferência – LADO A não fotografado, maior pendência de campo do M&E. A43: placa 'PLATAFORMA 6 LADO A' em foto de terceiros de 02/10/2025. A33: 'Lado A/B' do ANP refere-se às bombas.", 'M14E / A43 / A33', 'Observação na síntese 14e.'),
 (L, 'B23', '2 (OF)', 'DIVERGÊNCIA', "L3-04: proporcionador FireDos FD4000/3-PP-S (SAP H-3150, série 11600455) fabricado em 11/2016, NF SECUR 45694 de 2024 – equipamento 8 anos mais velho que a NF; é Outros Fixos (fora da linha 2), mas a divergência precisa constar.", 'L3-04', 'Observação na síntese 2.'),
 (L, 'B24', 'geral', 'PARÂMETROS', "Q04 ('seguir sem o OPE154') ficou superada – o OPE154 foi lido na v2. Q05: software/TI entra na composição de custo (máquina + componente + software), item a item. G02: planta em operação, equipe de manutenção presente – equipamento parado = em manutenção. P06 τ = 7, P08 Fc = 0,90, D02 instalado em operação / liquidez baixa, D04 idade desde a partida (30/09/2025) – parâmetros do laudo, não alteram a caracterização. L3-09: pedir data-book Metroval (de-para TAG × Nr.S).", 'PRE_ANALISE blocos 1, 5 e 6', 'Registro no LEIA-ME.'),
 (L, 'B25', 'geral', 'PENDÊNCIAS', "Lista única ao proponente (M16) no que toca ao M&E: de-para NF × subitem Metroval; excessos de NF (KSB 543048 +71.060,00; Romão 2739 +54.399,96; Metroval +80.487,06); placas TR-3100/TR.NB-3100/geral do eletrocentro; numeração completa das plataformas e LADO A; data de partida documentada; horas/ano das bombas; certificado Romão; cotações KSB/Metroval; base SAP completa.", 'Briefing – Depende do proponente', 'Sem mudança na planilha; lista para o pedido.'),
]
wk = wb.create_sheet('ACHADOS_v4')
wk['A1'] = 'Achados e alterações da v4 (13/10/2026) – LILÁS = a partir do 00_LEIA_ME correto (briefing unificado + pré-análise única de 05/10); VERDE CLARO = a partir do LOG_REPOSITORIO / LOG_FOTOS / árvore de pastas organizada por Rafael em 13/10.'
wk['A1'].font = Font(name='Arial', size=10, bold=True, color=P)
for i, h in enumerate(['Nº', 'Linha(s)', 'Tipo', 'Descrição', 'Fonte', 'Efeito / ação na planilha'], 1):
    c = wk.cell(row=3, column=i, value=h); c.font = Font(name='Arial', size=9, bold=True, color='FFFFFF'); c.fill = HEAD_FILL; c.border = BORDER
for i, w in enumerate([6, 16, 18, 110, 28, 50], 1): wk.column_dimensions[get_column_letter(i)].width = w
for r, (col, *vals) in enumerate(ACH, 4):
    for i, v in enumerate(vals, 1):
        c = wk.cell(row=r, column=i, value=v); c.font = Font(name='Arial', size=9, color=col, bold=(i == 1)); c.alignment = TOP; c.border = BORDER
    wk.row_dimensions[r].height = 75
wk.freeze_panes = 'A4'

# =====================================================================================================
# 5. LEIA-ME
# =====================================================================================================
wl = wb['LEIA-ME']
wl['A1'].value = 'Caracterização consolidada de M&E – Ultracargo Palmeirante-TO (base 07/10/2026; v2 12/10 com cotações OPE154; v3 13/10 com aba SINTESE; v4 13/10 com LEIA-ME e árvore de pastas)'
for r in range(1, wl.max_row + 1):
    v = wl.cell(row=r, column=1).value
    if isinstance(v, str) and v.startswith('Anexo não usado'):
        wl.cell(row=r, column=1).value = v + ' [v4: o 00_LEIA_ME correto foi recebido em 13/10 e assimilado – ver abaixo]'
n = wl.max_row + 2
notas = [
 ('v4 – 13/10/2026 – o que mudou (tudo em colunas, abas e cores à parte; nada da v3 foi apagado):', P, True),
 ('LILÁS (A97BD6) – alterações a partir do 00_LEIA_ME correto de 05/10/2026 (BRIEFING_UNIFICADO_Ultracargo.md + PRE_ANALISE_UNICA_Ultracargo.xlsx + COMO_ORGANIZAR.md + PROMPT_nova_sessao.md + MAPA_PROJETO.csv). O zip de pré-análise anexado na v2 era de outro trabalho. Fatos assimilados: decisões D01–D13 e respostas Q01–Q05 (CCB 183-25/0003-6; crítica vigente 05/11/2024; regra de "foto complementar" para fotos fora de 14–18/09/2026; software entra na composição de custo), bloco 6 do lote 3 (L3-01 a L3-11: megaflow com placa, B-31100 = linha 12, B-3119 com placa, FireDos 11/2016, releituras), achados A01–A46 (filtros MKG = linhas da LI-3100-6105; placas Metroval de terceiros com 8 itens novos; eletrocentro 11/13 placas; MB-502 não usar; PLATAFORMA 6 LADO A) e a lista de pendências ao proponente (M16).', L, False),
 ('VERDE CLARO (00B050) – alterações a partir do LOG_REPOSITORIO.csv (152 movimentos, 05 a 13/10), do LOG_FOTOS_20261013_003759.csv (599 movimentos) e da árvore ANEXO_IV_ME_fotos registrada em 2026-10-13_Inventario_Fotos_AnexoIV.xlsx (abas ARVORE, INVENTARIO, PLANO, NAO_ALOCADAS): Rafael colocou manualmente as válvulas, filtros, medidores e sensores que estavam soltos dentro das pastas dos skids/grupos a que pertencem (Item_14a_Top_loading, Item_14b_Bottom_loading, Item_14c_Skid_vagoes, Item_14d_Skid_descarga_caminhoes, Item_13 dentro de 14d etc.). A v4 usa essa árvore como vínculo de campo instrumento × skid: a seção "14 (comp.)" foi realocada (coluna A das linhas de componente passa a "14a (comp.)", "14b (comp.)" ou "14d (comp.)"), as sínteses 14a/14b/14c/14d/14e listam os componentes alocados, e cada componente recebeu a pasta e o nº de fotos na coluna Z da CARACTERIZACAO.', V, False),
 ('Onde está cada coisa na v4: aba SINTESE – acréscimos em lilás e verde claro no fim de cada célula de síntese e de observação (o texto da v3 ficou intacto); aba CARACTERIZACAO – colunas Y (lilás) e Z (verde claro), coluna A realocada em verde claro e 25 linhas novas no fim (BR - 3107, 14 filtros MKG, caixas de junção TELBRA, placa do skid PLCVT-01, 8 itens Metroval vistos só em fotos de terceiros); aba ARVORE_FOTOS_v4 – árvore completa de Rafael × linha da crítica + itens do PLANO sem decisão (CONFLITO/REVISAR/ANP); aba ACHADOS_v4 – B01 a B25 com fonte e efeito.', P, False),
 ('Regras mantidas: nenhuma leitura de campo foi ajustada para bater com documento (divergências registradas: BR - 3107 séries 3634/3635; FT-3151 em duas pastas; XS-312xx região; B-31100 na pasta de água potável; F-0415/F-0368 fora da linha de projeto); ilegível fica "não lido"; os itens do PLANO sem decisão (CONFLITO, REVISAR) ficam como Rafael deixou. As pastas de origem (Downloads, F:\\Ultracargo, Desktop\\geotins) não foram tocadas.', P, False),
 ('Pronto: assimilação do LEIA-ME e da árvore de pastas; falta: decisão de Rafael nos CONFLITO/REVISAR do inventário, releitura em resolução cheia (BR - 3107, FT-3141/3151, _a_identificar_instrumentos 45 + 53 fotos) e cópia das fotos de terceiros do eletrocentro para as subpastas vazias; depende do proponente: data-book Metroval (TAG × Nr.S), placas TR-3100/TR.NB-3100, LADO A das plataformas, certificado Romão.', P, False),
]
for i, (t, col, bold) in enumerate(notas):
    c = wl.cell(row=n + i, column=1, value=t); c.font = Font(name='Arial', size=10 if bold else 9, color=col, bold=bold); c.alignment = Alignment(wrap_text=True, vertical='top')

wb.save(OUT)
print('salvo', OUT, '| componentes realocados:', n_realoc, '| linhas novas:', len(NEW), '| achados v4:', len(ACH))

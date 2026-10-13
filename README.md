# Organização do relatório de avaliação – Ultracargo Palmeirante/TO (M&E, Banco da Amazônia)

- `2026-10-13_Caracterizacao_Consolidada_ME_Ultracargo_v3.xlsx` – versão viva. Abas: LEIA-ME, SINTESE (uma célula sintética por conjunto: descrição, componentes, NF, cotação; observações; valores crítica/NF/cotação), CARACTERIZACAO (colunas A-U = base; V cotação; W achados; X célula única), CONDENSADO, NF_ME, COTACOES_OPE154, ACHADOS_v2.
- `versoes_anteriores/` – v2 de 12/10/2026 (assimilada na v3).
- `base/2026-10-07_Caracterizacao_Consolidada_ME_Ultracargo_v1.xlsx` – base usada (sem alteração).
- `scripts/build_v2.py` + `scripts/dados_v2.py` – gerador da v2 a partir da base; `scripts/build_v3.py` + `scripts/sintese_v3.py` – acrescenta a aba SINTESE (`python3 -I scripts/build_v3.py <v2>.xlsx <v3>.xlsx`).

Legenda de cores: preto = crítica de orçamento; vermelho = NFs do RT0538; verde = SAP; amarelo = listas de projeto; azul = placa/campo/ANP; azul escuro = cotações OPE154; roxo = achados, correções e proposições do avaliador; tachado = redundância dentro da célula única.

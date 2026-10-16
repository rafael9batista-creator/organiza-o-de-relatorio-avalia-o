# Relatório fotográfico M&E Ultracargo – reorganização v4 → v5

## Contexto

Sou Rafael, engenheiro avaliador (laudo de M&E, NBR 14653-5, Ultracargo Palmeirante/TO, Banco da Amazônia). Refiz o relatório fotográfico das máquinas e equipamentos: retirei fotos repetidas, padronizei sequências e comecei a padronizar as legendas, sobretudo quando a legenda se refere a um subgrupo (ex.: "Bombas CETRAN 01 (Linha 4)"). Agora quero que você termine a padronização, sem alterar nenhum dado técnico.

**Arquivo de entrada (não alterar, não mover):**
`C:\projetos claude code\Ultracargo_Avaliacao\09_Laudo\2026-10-14_Relatorio_Fotografico_ME_Ultracargo_v4  limpo 2.docx`
(localize por `2026-10-14_Relatorio_Fotografico_ME_Ultracargo_v4*limpo*2*` dentro de `09_Laudo`; se a extensão for outra, me diga antes de continuar).

**Fotos originais:** árvore `ANEXO_IV_ME_fotos` (dentro de `06_Fotos_Laudo` ou pasta irmã do projeto; procure por esse nome). As subpastas seguem `Item_NN_...`, com subpastas por TAG (ex.: `Centran 1 ok/Item_04_Bombas_megacpk_125-080-315_12un/B-3117`), `tanques/TQ-3101` … `TQ-3112`, e páginas do `2025_04_11_RELATÓRIO FOTOGRÁFICO ANP_para envio-NN.pdf`. As fotos da vistoria têm nome `IMG_AAAAMMDD_HHMMSS.jpg` (vistoria em 15, 17 e 18/09/2026).

**Arquivo de saída:** `09_Laudo\2026-10-16_Relatorio_Fotografico_ME_Ultracargo_v5.docx` (nova versão; a v4 fica onde está, eu arquivo depois).

## Objetivo (5 tarefas, nesta ordem)

1. **Ordem das fotos:** corrigir a sequência dos grupos e subgrupos; dentro de cada equipamento, a sequência fixa é: vista geral → placa do equipamento → placa do motor/acionamento → detalhes. Subgrupos em ordem crescente de TAG (B-3103, B-3104, …; TQ-3101, TQ-3102, …; plataforma 08, 09, …).
2. **Numeração:** renumerar todas as fotos em sequência (FOTO 001 … FOTO NNN) depois do reordenamento, sem lacunas nem repetições. Gerar tabela de-para (número antigo → novo).
3. **Legendas:** aplicar o padrão único abaixo, sem perder nenhum dado que já está na legenda atual.
4. **Observações à parte:** tirar das legendas tudo que é observação ("leitura parcial", "ilegível", "reflexo", "conferir", "?", "provável", "REVISAR" etc.) e levar para o Quadro de Observações. No corpo do relatório, destacar **somente** onde provavelmente há erro ou inconsistência entre o que a legenda descreve e o que a foto mostra.
5. **Tanques / páginas ANP:** as fotos de tanque saíram de contexto (ex.: fotos do TQ-3102 estão na sequência do TQ-3101, e depois vem o TQ-3103). Reagrupar por tanque, cada TAG contígua e em ordem crescente, e conferir as páginas ANP que o inventário já apontou como fora de lugar (ver "Pontos conhecidos").

## Regras invioláveis

- **Nenhum dado técnico muda.** Modelo, OP, vazão, altura, rpm, ano, peso, carcaça, tensão, corrente, potência, data de fabricação: copiar exatamente como está na legenda atual. Se a foto mostrar outro valor, **não corrija a legenda**: registre a divergência no Quadro de Observações e destaque no corpo. Divergência se registra, não se resolve.
- **Nenhuma foto é excluída ou acrescentada** sem minha decisão. Duplicatas que ainda restarem entram no Quadro de Observações como "DUPLICATA provável" com os dois números, e eu decido.
- **Nenhuma observação é descartada.** Tudo que sair de uma legenda tem de reaparecer no Quadro, ligado ao número novo da foto.
- **Layout do Word:** manter uma foto por página, tamanho e proporção das imagens, estilos de parágrafo e fonte da v4. Não usar `pageBreakBefore` nem propriedades que engessem o layout; eu reposiciono quadros à mão.
- **Não reescrever texto do laudo** fora das legendas e do Quadro de Observações.
- Tarefa mecânica (ordenar, renumerar, reformatar) se faz sem pedir OK; só decisão de conteúdo vem para mim, e só no relato final, não no meio do trabalho.

## Padrão de legenda

Formato geral (campos separados por " – "; omitir campo vazio, nunca inventar):

```
FOTO NNN – [Grupo] – [Subgrupo/TAG] – [Equipamento: fabricante, modelo] – [placa: dados] – [acionamento: dados]. [Fonte, data].
```

Regras:
- **Grupo** = nome fixo do conjunto, igual em todas as fotos do conjunto (ex.: `CETRAN 01 (Linha 4)`, `CETRAN 02`, `SCI – casa de máquinas`, `Skid de descarga de caminhões (Item 14d)`, `Top loading (Item 14a)`, `Tanques – TQ-3101`, `Eletrocentro VOLGA`). Faça primeiro a lista dos grupos e subgrupos existentes na v4, unifique as variações de grafia (Centran/CETRAN/Cetran 1/01) e use um nome só por grupo.
- **TAG** aparece uma vez só, no campo Subgrupo/TAG. Não repetir "TAG B-3117" dentro de "placa:".
- **placa:** só os dados de placa do equipamento, na ordem fixa: modelo completo, nº de série/OP, grandezas de processo (Q, H, rpm), peso, ano.
- **acionamento:** "motor WEG W22Xdb, carc. 180M/L, IPW55, 22 kW/30 cv, 220/380/440 V, 75,2/43,5/37,6 A, 1775 rpm, 248 kg, fab. 04/06/2024" (ordem fixa: fabricante e linha, carcaça, grau de proteção, potência, tensão, corrente, rpm, peso, fabricação).
- **Fonte e data**, sempre no fim, em uma destas três formas, e só nestas:
  - `Vistoria 15/09/2026.` (data = data do nome do arquivo IMG_AAAAMMDD; se a legenda atual diz outra data, é inconsistência: destacar)
  - `Relatório Fotográfico ANP, 11/04/2025, p. NN.`
  - `Foto complementar (terceiros), DD/MM/AAAA.`
- Vista geral de subgrupo: `FOTO NNN – CETRAN 01 (Linha 4) – vista geral das bombas B-3103 a B-3117. Vistoria 15/09/2026.`
- Foto de detalhe sem placa: `FOTO NNN – CETRAN 01 (Linha 4) – B-3117 – detalhe: flange de recalque e manômetro. Vistoria 15/09/2026.`
- Unidades e grafia: `m³/h`, `mca`, `rpm`, `kW/cv`, `V`, `A`, `kg`; vírgula decimal; "carc." para carcaça; "fab." para data de fabricação; "s/n" para número de série; "OP" para ordem de produção KSB.

Exemplo – legenda atual:

> FOTO 644: Bombas CETRAN 01 (Linha 4) – B-3117 – Bomba KSB – placa: KSB MCPK 125-80-315 EE, OP M482301625, Q 120 m³/h, H 50 mca, 1771 rpm, ano 2024, TAG B-3117, peso 158 kg, com motor elétrico WEG W22Xdb: carc. 180M/L, IPW55, 248 kg, 220/380/440 V, 75,2/43,5/37,6 A, 22 kW / 30 cv, 1775 rpm, fab. 04/06/2024. Fotos da vistoria, 15/09/2026.

Legenda padronizada (mesmos dados, nova ordem, sem redundância):

> FOTO 644 – CETRAN 01 (Linha 4) – B-3117 – Bomba KSB Megacpk 125-80-315 EE – placa: OP M482301625, Q 120 m³/h, H 50 mca, 1771 rpm, 158 kg, ano 2024 – motor WEG W22Xdb, carc. 180M/L, IPW55, 22 kW/30 cv, 220/380/440 V, 75,2/43,5/37,6 A, 1775 rpm, 248 kg, fab. 04/06/2024. Vistoria 15/09/2026.

Em um subgrupo com várias fotos do mesmo equipamento (vista geral, placa, placa do motor, detalhe), os dados de placa entram **uma vez**, na foto da placa; as demais fotos do mesmo TAG trazem só Grupo – TAG – o que a foto mostra – fonte.

## Observações e inconsistências

- **Quadro de Observações** (tabela no fim do relatório, antes dos anexos, e aba `OBSERVACOES` do mapa): colunas `FOTO (nova)`, `FOTO (antiga)`, `TAG`, `Tipo`, `Observação`, `Decisão do Rafael` (vazia). Tipos: `LEITURA PARCIAL`, `ILEGÍVEL`, `INCONSISTÊNCIA`, `DUPLICATA`, `ORDEM` (foto realocada), `DATA` (data da legenda ≠ data do arquivo), `FONTE` (página ANP em pasta de outra região).
- **No corpo do relatório**, abaixo da legenda, uma linha só para o tipo `INCONSISTÊNCIA`, em destaque (texto vermelho ou realce amarelo, uma vez só, sem repetir a observação inteira): `Verificar: legenda diz 1771 rpm; placa na foto lê 1750 rpm.` Nenhum outro tipo aparece no corpo.
- **O que é inconsistência:** modelo, série, TAG, grandeza, ano ou data na legenda diferente do que se lê na foto; foto cuja placa é de outro TAG que não o da legenda; foto alocada a um subgrupo mas que mostra outro equipamento; legenda de bomba com motor de outra bomba.
- **Leitura em dois níveis.** Nível 1: abrir a imagem extraída do próprio docx e comparar campo a campo com a legenda; se tudo se lê e bate, registrar `conferida_docx = SIM` no mapa e seguir. Nível 2: só quando a imagem do docx deixa dúvida em algum campo (resolução, reflexo, recorte), abrir o arquivo original da árvore `ANEXO_IV_ME_fotos` (já localizado pelo hash no mapa) e reler; registrar `conferida_original = SIM`. Se nem a original resolve, registrar `ILEGÍVEL` no campo em dúvida, nunca `INCONSISTÊNCIA` por dúvida.
- Alvo da conferência: todas as fotos de placa (equipamento e motor); nas vistas gerais e detalhes, conferir só TAG/subgrupo e data.

## Pontos conhecidos do inventário de 13/10 (conferir com prioridade)

- `B-3119`: foto `IMG_20260917_152835` foi lida como B-3120 dentro da pasta B-3119.
- `B-3130`: fotos `152123` e `152145` lidas como B-3131.
- `JB-3106P`: fotos `154429` e `154604` lidas como JB-3106D.
- `plataforma 11`: fotos `160127` e `160134` lidas como PLATAFORMA 12.
- `B-31100` (Item 12): fotos `171246` e `171302` estavam na pasta de água potável.
- Monitores `XS-31209/31211/31212` fotografados na descarga de caminhões (TAG de série do ferroviário).
- Páginas ANP fora de região: p. 45 e p. 58 são de TANQUES (estavam em SCI e pipe rack); p. 75, 76 e 77 são de RODOVIÁRIO (estavam em SCI e ETE).
- `FT-3151`: dois medidores com a mesma TAG (séries 01341123 e 01401123); `FT-3141`: fotos leem séries 01311123 e 01371123.
- `SDV-3112`: uma foto lê s/n 0339 e outra 0338.

Se alguma dessas fotos estiver no relatório, a legenda tem de bater com o que a foto mostra; se não bater, é `INCONSISTÊNCIA`.

## Método (executar nesta ordem, sem perguntar entre passos)

**Passo 1 – Mapa da v4 (antes de mudar qualquer coisa).** Com `python-docx`, percorrer o documento na ordem e extrair, para cada foto, a imagem (salvar em `07_Saidas_do_Claude\fotos_v4_extraidas\NNN.jpg`), o parágrafo de legenda e o número atual. Gravar `07_Saidas_do_Claude\MAPA_FOTOS_v4_v5.xlsx`, aba `MAPA`, uma linha por foto, com as colunas ANTES já preenchidas:

| ANTES (passo 1) | DEPOIS (passos 3 a 6) |
|---|---|
| `num_antigo`, `pos_antiga` (ordem no docx), `legenda_antiga`, `grupo_antigo`, `subgrupo_antigo`, `tag_lida_na_legenda`, `tipo_foto` (vista geral / placa / placa motor / detalhe), `fonte`, `data_legenda`, `sha1_docx`, `arquivo_original`, `pasta_original`, `data_arquivo`, `duplicata_de` | `num_novo`, `pos_nova`, `grupo_novo`, `subgrupo_novo`, `legenda_nova`, `mudou_posicao` (S/N), `conferida_docx`, `conferida_original`, `tag_lida_na_foto`, `tipo_obs`, `observacao`, `tokens_numericos_perdidos` |

Ainda no passo 1 (tudo por script, sem olhar imagem): calcular SHA-1 (e pHash, se o docx recomprimiu) de cada imagem e casar com os arquivos de `ANEXO_IV_ME_fotos` para preencher `arquivo_original`, `pasta_original`, `data_arquivo` e `duplicata_de`. Conferir que cada imagem tem exatamente uma legenda (imagem sem legenda ou legenda sem imagem entram como `ORDEM`). Dizer em uma linha quantas fotos há, quantas casaram com a árvore e quantas duplicatas restam, e seguir.

**Passo 2 – Grupos.** Listar os grupos e subgrupos que existem no mapa, unificar grafias (Centran/CETRAN/Cetran 1/01) e fixar um nome por grupo (aba `GRUPOS`: `nome_na_v4` → `nome_padrao`, ordem no laudo).

**Passo 3 – Ordem.** Montar a sequência Item 01 → Item 18 (skids 14a–14e, tanques TQ-3101 → TQ-3112, plataformas 08 → 14), usando `pasta_original` como critério de desempate quando a legenda não diz o TAG. Preencher `pos_nova`, `num_novo`, `mudou_posicao`; toda foto com `mudou_posicao = S` ganha linha `ORDEM` nas observações ("era FOTO 312, entre TQ-3101; agora FOTO 330, em TQ-3102").

**Passo 4 – Conferência visual (dois níveis, conforme a seção acima).** Preencher `conferida_docx`, `conferida_original`, `tag_lida_na_foto`, `tipo_obs`, `observacao`. Começar pelos pontos conhecidos, depois todas as placas, depois amostra das demais.

**Passo 5 – Legendas novas.** Preencher `legenda_nova` aplicando o padrão e validar por script: começa com `FOTO NNN – `, termina com uma das três fontes, não contém "(", ")", "?", "parcial", "ilegível", "conferir", "provável", "REVISAR"; numeração contínua; `tokens_numericos_perdidos` vazio para toda foto (comparar o conjunto de números da legenda antiga × nova; diferença só é aceita quando o número era a própria TAG repetida ou o número antigo da foto).

**Passo 6 – Gerar a v5.** Copiar a v4 e, na cópia, reordenar os blocos foto+legenda pela `pos_nova`, substituir o texto das legendas pela `legenda_nova`, inserir a linha "Verificar:" nas `INCONSISTÊNCIA` e o Quadro de Observações no fim. Salvar `2026-10-16_Relatorio_Fotografico_ME_Ultracargo_v5.docx`. Reabrir com `python-docx` e conferir contra o mapa: mesmo número de imagens da v4, legendas na ordem de `pos_nova`, nenhuma imagem perdida ou duplicada pelo script; registrar o resultado na aba `VALIDACAO`.

**Passo 7 – Entregáveis em `07_Saidas_do_Claude\`:** `MAPA_FOTOS_v4_v5.xlsx` (abas `MAPA`, `GRUPOS`, `OBSERVACOES`, `DE_PARA` com só `num_antigo` → `num_novo`, `VALIDACAO`), e o script usado (`scripts\relatorio_fotografico_v5.py`), para que eu rode de novo se mudar algo.

## Relato final (uma vez só, no fim)

- **Pronto:** total de fotos, total de legendas padronizadas, quantas mudaram de posição, lista dos grupos na ordem final.
- **Inconsistências encontradas:** lista curta, foto por foto (nº novo, TAG, legenda diz × foto mostra). Esta é a parte que eu mais vou ler.
- **Duplicatas prováveis:** pares de números.
- **Falta / depende de mim:** decisões de conteúdo (duplicatas a excluir, inconsistências a corrigir na origem, páginas ANP a realocar), em lista numerada para eu responder OK / DIVERGE / DÚVIDA em cada uma.
- Sem explicação do método, sem repetir as regras.

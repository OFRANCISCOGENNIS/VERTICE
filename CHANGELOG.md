# Changelog — TRIPÉ

## v0.1.5 — CKCP v17: menos falso DEVOLVER (2026-11-12)

- `vba/AnaliseCKCP.bas` atualizado para a v17. Num teste com 2 CJI3 (1.356 PEPs), o DEVOLVER caiu de 1.332 para 253.
- Instalação/MO: o catálogo de serviços é consultado antes da classe de custo. As classes 8210040000 e 8210760000 passaram de RISCO para MÃO DE OBRA.
- ATV DRT: contado uma vez só (ATIVAÇÃO DIRETA); teste de 25% (8% em EME/EMM) ± `ATV_TOL`; MO zero com ATV vira QUESTIONAR.
- Linhas sem código de material (ICMS, viagem) entram na nova categoria OUTROS CUSTOS, fora de material, UC/COM e QTD x VALOR.
- MOP com base única na CHECK CAP e na ANÁLISE DE CA: chaves `MOP_BASE` (TOTAL_SEM_MOP), `PERC_MOP` (5,483) e `MOP_TOL` (1). ODD fica N/A e SEM MOP vira QUESTIONAR.
- Colunas: os 10 cabeçalhos padrão da CJI3 são reconhecidos, com correção de acento quebrado (mojibake). Unidades normalizadas (PEÇ/PEC → PC, un → UN).
- Layout: ANÁLISE DE CA ganha as colunas 40 a 43; na CHECK CAP o cabeçalho "% MOP/SERVICO" passa a "% MOP/BASE".

## v0.1.4 — CKCP v16: robustez e avisos (2026-11-11)

- `vba/AnaliseCKCP.bas` atualizado para a v16. Regras, vereditos, abas, colunas e chaves da `CONFIG` iguais à v15.
- Catálogo já aberto no Excel é reaproveitado e não é mais fechado sem salvar; erro de leitura de catálogo vira aviso.
- `CHECK CAPITALIZACAO` avisa data de corte inválida, coluna de data ausente e aba `CN43N` ausente.
- Mensagem indica a coluna obrigatória que falta e a aba lida; etapa na barra de status; avisos reunidos no fim.

## v0.1.3 — CKCP v15: CHECK CAPITALIZACAO (2026-10-31)

- `vba/AnaliseCKCP.bas` atualizado para a v15 (v14 + `Gerar_CheckCapitalizacao`), com os testes da planilha
  de capitalização RD de GO (TST GERAL, TST UC-COM, TST MOP, TST MAT, CABO em KILO e motivos da aba Devolver).
- Nova aba `CHECK CAPITALIZACAO` (1 linha por PEP): MOP/SERVIÇO (11% ± 1), UC/UAR, COM, serviço de instalação,
  quantidade zerada, valor negativo, abaixo do preço médio, serviço na classe de material, cabo UC em KG < 7 m,
  lançamentos após a data de corte, `MOTIVOS DEVOLUCAO` e `STATUS CAPITALIZACAO`. Motivos `QUESTIONAR:` não reprovam.
- Com uma aba `CN43N` no arquivo, também testa CC solicitante/responsável, TI, ODI, perfil de investimento,
  tranche e data de energização.
- Nova aba `CAP DETALHE` com os itens apontados. Parâmetros novos `CAP_*` na `CONFIG`.
- `Tripe.bas`: as duas abas novas entram na ordenação final logo depois de `ALERTAS CRITICOS`.
- Regras conferidas em Python sobre uma CJI3 real (431 mil linhas). Ainda falta compilar e rodar no Excel.

## v0.1.2 — CKCP v14 (2026-10-30)

- `vba/AnaliseCKCP.bas` atualizado para a versão 14 (AnaliseCKCP_OTIMIZADO de 23/10/2026). Novas regras da
  aderência MATERIAL vs SERVICO:
  - grupos de **LIGACAO SERV X MAT** (planilha "LIGACAO DE MATERIAL E SERVICO"): serviço e material do
    mesmo grupo lançados no mesmo PEP4 ficam aderentes; nova coluna `LIGACAO SERV X MAT`;
  - **famílias críticas**: NAO ADERENTE em qualquer PEP4 (ODD/ODI/ODM/ODS) reprova o PEP3;
  - CH FUS dentro do esperado pelos trafos (1F = 1, 3F = 3); serviço 5025300148 paga 3 CH FUS + 3 PARA RAIO MT;
  - 122060001 (aço cobreado) abaixo de 10 kg e CABO ISOLADO abaixo de 15 não reprovam;
  - material com "Classificação" vazia no SAP (ex.: sucata em ODD) entra como material pelo catálogo;
  - ranking de tempo por etapa ao final (desligável com `MOSTRAR_TEMPOS = 0` na aba CONFIG).
- A coluna de veredito da aba MATERIAL vs SERVICO passou de `STATUS_PEP4` para `STATUS_PEP3NIVEL`;
  o `Tripe.bas` lê os dois nomes.
- O CKCP v14 não gera mais a aba `RACIONALIZACAO COM` (e a apaga no início). O TRIPÉ já preservava a aba
  do inventário com nome temporário, então ela continua intacta.
- `AnaliseCKCP.bas` agora com quebras de linha CRLF, como os outros módulos (a versão anterior estava só com LF).

## v0.1.1 (2026-10-20)

- TRIPE DETALHE: nova coluna `DESCRICAO DO MATERIAL` ao lado de `FAMILIA`, com os materiais do
  inventário de cada PEP4 + família (`código - descrição`, até 15 por linha e depois `+n outros`).

## v0.1.0 — MVP (2026-10-20)

- Macro única `GerarTripe` que roda a análise Projeto × Baixa SAP (`inventario.bas`) e a análise
  Material × Serviço (`AnaliseCKCP.bas`) sobre a mesma planilha de upload.
- Planilha padrão de upload com as abas `INVENTARIO` e `RAZAO`, validação de cabeçalhos e busca
  automática do arquivo.
- Novas abas `TRIPE CONCLUSAO` (veredito final por obra) e `TRIPE DETALHE` (PEP4 + família com as
  três quantidades e a perna que quebrou).
- Cruzamento por PEP4 completo (`.I` / `.D` / `.M`), com fallback para PEP3 quando os PEP4 das duas
  abas não batem.
- Vereditos `SEM RAZAO` / `SEM INVENTARIO` para obras presentes em só uma das abas.
- Ao final ficam visíveis só as abas do relatório, em ordem fixa, com `REGRAS` oculta e as demais
  abas ocultas (nada é apagado).
- Conflito da aba `RACIONALIZACAO COM` resolvido: a do CKCP passa a se chamar `RACIONALIZACAO COM CKCP`.
- As macros são procuradas em todos os arquivos abertos, com mensagem clara quando falta um módulo.

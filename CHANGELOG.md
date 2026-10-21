# Changelog — TRIPÉ

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

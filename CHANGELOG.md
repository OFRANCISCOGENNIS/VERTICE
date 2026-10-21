# Changelog — TRIPÉ

## v0.2.0 (2026-10-21)

Refatoração interna sem mudar resultados, mais a versão nova do `inventario.bas` (essa sim com
regras novas — ver abaixo).

### Tabelas embutidas divididas (`AnaliseCKCP.bas`)

Cada procedimento gigante virou um procedimento principal que chama blocos de até 75 linhas,
com os mesmos dados na mesma ordem (conferido linha a linha contra a v0.1.1):

| Procedimento | Linhas de dados | Blocos |
|---|---|---|
| `CarregarTipoClassifAuto` | 824 | `CarregarTipoClassifAutoP1` … `P11` |
| `CarregarClassesCustoAuto` | 740 | `CarregarClassesCustoAutoP1` … `P10` |
| `CarregarFamiliaNivel` | 632 | `CarregarFamiliaNivelP1` … `P9` |
| `CarregarDescServico` | 605 | `CarregarDescServicoP1` … `P9` |

### Revisão de `On Error Resume Next`

Correções (o tratamento de erro não era restaurado corretamente):

- `Tripe.bas` › `GerarTripe`: depois de ativar a aba `TRIPE CONCLUSAO`, voltava para
  `On Error GoTo 0` e desligava o handler `Falha`. Agora restaura `On Error GoTo Falha`.
- `inventario.bas` › `GerarInventario`: mesmo caso ao ativar o `PAINEL DO GESTOR`. Agora
  restaura `On Error GoTo ErrHandler`.
- `inventario.bas` › `GravarLog` e `RegistrarKpiRun`: o `Resume Next` cobria o procedimento
  inteiro. Agora protege só a busca da aba (que pode não existir) e a ativação/ocultação final.
- `AnaliseCKCP.bas` › `AddTxt`, `MetricBlock` e a limpeza `SemCat:` de
  `CarregarCatalogoMateriais`, `CarregarCatalogoServicos`, `CarregarCatalogoClasse`,
  `CarregarConversoesCabo` e `CarregarComboServico`: o bloco ia até o `End Sub` sem restaurar.
  Agora termina com `On Error GoTo 0` (sem efeito no comportamento, só deixa explícito).

Mantidos de propósito (só formatação visual ou exclusão de abas que podem não existir),
agora com comentário explicando:

- `inventario.bas`: exclusão das abas antigas em `GerarInventario`; cabeçalho, cards e
  formatação condicional em `ProcessarSAPxPRJ`; botão em `ProcessarResumoPEP3`; `DefinirNome`.
- `AnaliseCKCP.bas`: `MostrarTelaFuturista`, `AplicarFreeze`, barras de dados em
  `FormatarVisualAba` e `GerarChecklistCKCP`, `PrepararImpressao`.

Os demais `On Error Resume Next` já restauravam o tratamento logo após a linha que pode falhar.

### `inventario.bas` atualizado

Nova versão do módulo de inventário (os ajustes de `On Error` acima foram reaplicados por cima):

- Discrepância de valor de material em PEP ODI: preço unitário de cada linha comparado com a
  média do mesmo código (tolerância `DISCREP_VALOR_PCT`, padrão 30%). Vira seção própria no
  RANKING DE RISCO e entra no SCORE (`PESO_DISCREP` / `CAP_DISCREP`) e no diagnóstico.
- "LACRE P/ MEDIDOR" passa a ser classificado como lacre, não medidor (`EhMedidorODM`).
- Desempenho: cache em `NormStr` e ordenação do ranking em O(n log n), mantendo a mesma ordem.
- `EQUIV_VERSAO_ATUAL` = 5.

### Outros

- `AnaliseCKCP.bas` passa a ter quebras de linha CRLF, como os outros `.bas` (estava só com LF).

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

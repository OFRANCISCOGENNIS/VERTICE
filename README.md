# VÉRTICE — TRIPÉ: Projeto × Baixa SAP × Serviço pago (v0.2.0)

> *A obra só fica de pé se as três pernas baterem.*

O TRIPÉ é uma automação em **VBA para Excel** que junta duas análises de obras que
antes eram feitas separadamente e entrega **uma conclusão única por obra (PEP3)**:

| Perna | Fonte | Pergunta |
|---|---|---|
| 1. **Projeto** | aba `INVENTARIO` — `TOMBADO PROJ` | O que foi projetado? |
| 2. **Baixa SAP** | aba `INVENTARIO` — `ATENDIDO SAP` | O que foi baixado no SAP? |
| 3. **Serviço pago** | aba `RAZAO` — lançamentos de serviço | O que foi pago de serviço? |

A obra é **APROVADA** só quando passa nas duas análises. Se a comparação Projeto × Baixa ou a comparação
Material × Serviço reprovar, a obra fica **REPROVADA**.

---

## Estrutura

```
VERTICE/
├── vba/
│   ├── Tripe.bas          # orquestrador + cruzamento das 3 pernas (macro GerarTripe)
│   ├── inventario.bas     # análise Projeto x Baixa SAP (macro GerarInventario)
│   └── AnaliseCKCP.bas    # análise Material x Serviço  (macro GerarRelatorio)
├── modelo/
│   └── TRIPE_UPLOAD_MODELO.xlsx   # planilha padrão de upload (só cabeçalhos)
├── CHANGELOG.md
└── README.md
```

## Instalação (uma vez)

1. Crie um arquivo novo no Excel e salve como **`TRIPE.xlsm`** (Pasta de Trabalho Habilitada para Macro).
2. `Alt+F11` → **Arquivo → Importar arquivo** e importe os **3** arquivos da pasta `vba/`.
3. Salve o `TRIPE.xlsm`.

> As macros devem ficar no `.xlsm`, **não** na planilha de upload: um `.xlsx` perde os módulos quando é salvo.

## Uso

1. Preencha uma cópia de `modelo/TRIPE_UPLOAD_MODELO.xlsx`:
   - **1ª aba `INVENTARIO`**: extração do inventário (projeto × SAP).
   - **2ª aba `RAZAO`**: razão de custo exportada do SAP (material + serviço).
2. Com o `TRIPE.xlsm` aberto, abra a planilha de upload.
3. `Alt+F8` → **`GerarTripe`** → Executar.

O TRIPÉ acha a planilha de upload sozinho. Ele usa o arquivo ativo; se esse não servir, procura nos
outros arquivos abertos; se ainda não achar, abre uma janela para você escolher. Antes de rodar, ele
confere se os cabeçalhos obrigatórios estão presentes.

## Planilha de upload

**`INVENTARIO`**: `PEP3NIVEL`, `PEP4NIVEL`, `NOTA`, `CLASSE_CUSTO`, `MATERIAL`, `VALOR`,
`DESC_MATERIAL`, `UNIDADE`, `ATENDIDO SAP`, `TOMBADO PROJ`, `TIPO`, `FAMILIA`, `SITUACAO_DO_MATERIAL`.
Também aceita o layout antigo: `COD MAT`, `MAT LIB SAP`, `MAT PRJ CAD`, `SIT MAT`.

**`RAZAO`**: export padrão do SAP, com pelo menos `Elemento PEP`, `Material`, `Qtd.total entrada`,
`Valor/moeda objeto` e `Classificação`.

O PEP4 deve estar no formato completo nas duas abas, por exemplo `GO-2200112LPT1.8.0027.I` (ou `.D`, `.M`).

## Abas geradas

Ao final da execução ficam visíveis, nesta ordem:

| Aba | Origem | Conteúdo |
|---|---|---|
| **TRIPE CONCLUSAO** | TRIPÉ | 1 linha por obra: veredito final, veredito de cada análise, nº de famílias divergentes, motivo |
| **TRIPE DETALHE** | TRIPÉ | 1 linha por PEP4 + família: família, descrição dos materiais (código - descrição), qtd projeto, baixada SAP, material CKCP, serviço pago, e qual perna quebrou |
| ALERTAS CRITICOS | CKCP | alertas da análise de custo |
| MATERIAL vs SERVICO | CKCP | aderência material × serviço por PEP4/família |
| MATERIAL | CKCP | lançamentos de material |
| SERVICO | CKCP | lançamentos de serviço |
| SERVICO SEM MATERIAL | CKCP | serviços pagos sem material correspondente |
| ANALISE DE CA | CKCP | custos administrativos por PEP |
| NAO CLASSIFICADOS | CKCP | itens sem classificação no catálogo |
| RAZAO | upload | base de custo original |
| REGRAS *(oculta)* | CKCP | regras aplicadas |
| ANALISE SAP x PRJ | inventário | aderência projeto × SAP por material |
| RESUMO SAP x PRJ | inventário | 1 linha por obra |
| ALERTA CRITICO | inventário | alertas do inventário |

As demais abas (INVENTARIO, CONFIG, painéis, históricos) ficam **ocultas, não apagadas**. Para ver
alguma delas, clique com o botão direito numa guia e escolha **Reexibir**.

## Situações do TRIPE DETALHE

| Situação | Perna que quebrou |
|---|---|
| ADERENTE (PRJ = SAP = SRV) | — |
| PROJETADO NAO BAIXADO | Baixa SAP |
| DIVERGENTE PROJETO x BAIXA | Projeto |
| MATERIAL BAIXADO SEM SERVICO PAGO | Serviço |
| SERVICO PAGO SEM MATERIAL / ACIMA / ABAIXO DO MATERIAL | Serviço |
| DIVERGENTE NAS TRES PERNAS | Projeto + Serviço |
| SERVICO PURO, SO NO CKCP, OK (COM - SERVICO NAO EXIGIDO) | — (informativo) |

O TRIPÉ **não reimplementa** as regras de aderência. Ele usa o resultado de cada macro, com todas as
exceções que elas já têm (equivalências, cabo, ramal, religador, trafo etc.), e só cruza os dois
resultados para dizer qual perna quebrou.

## Veredito por obra

| Inventário | CKCP | Final |
|---|---|---|
| REPROVADO | qualquer | **REPROVADO** |
| qualquer | REPROVADO | **REPROVADO** |
| APROVADO | APROVADO | **APROVADO** |
| tem dados | sem dados | **SEM RAZAO** |
| sem dados | tem dados | **SEM INVENTARIO** |
| outros (ex.: SEM UC) | | **VERIFICAR** |

## Configuração (aba `CONFIG`, oculta)

| Chave | Padrão | Uso |
|---|---|---|
| `TRIPE_ALIAS_FAMILIA` | *(vazio)* | De-para extra de família entre inventário e CKCP: `DE=PARA;DE2=PARA2` |
| `TRIPE_MIN_MATCH_PEP4` | `20` | % mínimo de PEP4 em comum (medido sobre a base menor) para cruzar por PEP4. Abaixo disso, o cruzamento é feito por PEP3 |

As duas macros de origem continuam usando as próprias chaves da `CONFIG`, como tolerâncias e caminhos
dos catálogos.

## Dependências externas

O `AnaliseCKCP.bas` lê catálogos de materiais, serviços, classes de custo, conversão de cabo e SRV
combo. Os caminhos ficam na `CONFIG` e, por padrão, apontam para a pasta `%USERPROFILE%\Downloads`.
Esses catálogos **não** estão neste repositório.

## Limitações conhecidas (MVP)

- As mensagens finais de cada macro de origem ainda aparecem durante a execução; basta clicar OK.
- As duas macros criam uma aba `RACIONALIZACAO COM`. O TRIPÉ mantém a do inventário e renomeia a do CKCP
  para `RACIONALIZACAO COM CKCP`; as duas ficam ocultas.
- Os nomes de família vêm de fontes diferentes: a coluna `FAMILIA` no inventário e o `CLS2` do catálogo
  no CKCP. Divergências de nome se resolvem com `TRIPE_ALIAS_FAMILIA`.
- Não há testes automatizados do VBA. A validação é feita rodando com uma planilha real.

## Próximos passos

- [ ] Modo silencioso nas duas macros de origem (sem MsgBox intermediário).
- [ ] Filtro opcional para concluir só as obras presentes nas duas abas.
- [ ] Versão web do TRIPÉ dentro da plataforma DexterIA (upload do `.xlsx` pelo navegador).

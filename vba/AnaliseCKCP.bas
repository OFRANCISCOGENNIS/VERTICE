Attribute VB_Name = "AnaliseCKCP"
'==============================================================================
'  CKCP RS2 - Anlise de Custos (VBA)
'  Transforma a base crua do SAP (EXPORT) na estrutura de relatrio.
'
'  COMO USAR:
'   1. Abra o EXPORT.XLSX (ou cole os dados crus na 1 aba).
'   2. Alt+F11 > Inserir > Mdulo > cole este cdigo (ou importe este .bas).
'   3. Feche o editor, Alt+F8 > GerarRelatorio > Executar.
'
'  A base crua deve ter os cabealhos do SAP na linha 1, incluindo:
'   "Elemento PEP", "Classe de custo", "Material", "Qtd.total entrada",
'   "Valor/moeda objeto", "Classificao", "Descrio SA", etc.
'==============================================================================
Option Explicit

' Rastreio de etapa para diagnostico de erros (exibido no handler Falha)
Private gEtapa As String

' --- ndices das colunas da base crua (preenchidos por MapearColunas) ---------
Private cPEP As Long, cClasse As Long, cDescClasse As Long, cMaterial As Long
Private cTexto As Long, cQtd As Long, cUML As Long, cValor As Long
Private cClassif As Long, cDescSA As Long, cCentro As Long, cEmpresa As Long
Private cObj As Long, cDenObj As Long, cDenClasse As Long, cDocCompra As Long
Private cNumDoc As Long, cDenominacao As Long, cUsuario As Long, cNumDocRef As Long
Private cDataLanc As Long, cHora As Long, cDataEntrada As Long, cTipoDoc As Long
Private cAno As Long, cDivisao As Long, cDataDoc As Long, cLinhaLanc As Long
Private cODI As Long, cSA As Long, cDocEstorno As Long, cOrgEstorno As Long
Private cEstorno As Long, cRefEstorno As Long, cOperRef As Long
Private cCLS1Raw As Long, cCLS2Raw As Long, cCLS3Raw As Long, cTipoAplicRaw As Long

' --- Estilo -------------------------------------------------------------------
Private Const COR_HDR As Long = &H7D491F          ' azul escuro 1F497D (BGR)
Private Const COR_OK As Long = &HCEEFC6           ' verde claro
Private Const COR_BAD As Long = &HCEC7FF          ' vermelho claro

' --- Design tokens da camada visual (fonte unica de verdade) ------------------
' Observacao: Const nao aceita RGB(); os numeros sao r + g*256 + b*65536.
Private Const FONTE_UI As String = "Segoe UI"
' Superficies e estrutura
Private Const COR_UI_HDR As Long = 4286720            ' RGB(0,105,65)   verde institucional
Private Const COR_UI_HDR_ACENTO As Long = 2505728     ' RGB(0,60,38)    verde escuro
Private Const COR_UI_ZEBRA As Long = 16251381         ' RGB(245,249,247)
Private Const COR_UI_BORDA As Long = 15132390         ' RGB(230,230,230)
Private Const COR_UI_GRADE As Long = 15658734         ' RGB(238,238,238)
Private Const COR_UI_TINTA As Long = 4739144          ' RGB(72,80,72)   grafite
Private Const COR_UI_CHAVE As Long = 3033344          ' RGB(0,73,46)
Private Const COR_UI_MUDO As Long = 8620931           ' RGB(131,139,131)
' Barras de dados
Private Const COR_UI_BARRA As Long = 9879160          ' RGB(120,190,150)
Private Const COR_UI_BARRA_NEG As Long = 6053078      ' RGB(214,92,92)
Private Const COR_UI_BARRA_PCT As Long = 13413253     ' RGB(133,171,204)
' Semantica: fundo / texto
Private Const COR_UI_OK_BG As Long = 15531448         ' RGB(232,245,236)
Private Const COR_UI_OK_FG As Long = 4681237          ' RGB(21,110,71)
Private Const COR_UI_BAD_BG As Long = 15526909        ' RGB(253,235,236)
Private Const COR_UI_BAD_FG As Long = 2629800         ' RGB(168,32,40)
Private Const COR_UI_WARN_BG As Long = 14087423       ' RGB(255,244,214)
Private Const COR_UI_WARN_FG As Long = 26258          ' RGB(146,102,0)
Private Const COR_UI_NEU_BG As Long = 15922161        ' RGB(241,243,242)
Private Const COR_UI_NEU_FG As Long = 8159352         ' RGB(120,128,124)
Private Const COR_UI_INFO_BG As Long = 16314600       ' RGB(232,240,248)
Private Const COR_UI_INFO_FG As Long = 9721621        ' RGB(21,87,148)

Private wsRaw As Worksheet
Private dados As Variant
Private nLin As Long
Private dCatMat As Object   ' COD_MATERIAL(str) -> "FAMILIA|CLS1|CLS2|CLS3"
Private dCatSrv As Object   ' COD_SERVICO(str) -> "CLS1|CLS2|CLS3|TIPO_APLIC|SEGMENTO"
Private dCatCC As Object    ' CLASSE_CUSTO(str) -> "CLS1|CLS2|CLS3|TIPO_APLIC"
Private dCabo As Object     ' COD_MATERIAL(str) -> fator KG->metros (Double)
Private dCombo As Object    ' COD_SERVICO(str)  -> fator multiplicador (Double)
Private dTipoCls As Object  ' CLASSIFICACAO(CLS2 normalizada) -> TIPO (COM/UC/UAR)
Private dFamEquiv As Object ' CLS2 servico (norm) -> familia material equivalente
Private dSrvPuro As Object  ' CLS2 servico (norm) -> 1 (servico sem material esperado)
Private dFamNivel As Object ' FAMILIA (norm) -> nivel (PRIMARIO/SECUNDARIO/NA)
Private rawHeaders As Variant
Private rawColCount As Long
Private gSplashOK As Boolean

' FASE 1.2: vereditos ODI da aba MATERIAL vs SERVICO compartilhados em memoria
' (fonte unica de verdade entre Gerar_MaterialVsServico e Gerar_PainelExecutivo)
Private dMvSVerd As Object    ' PEP3NIVEL -> "APROVADO"/"REPROVADO" (linhas ODI)
Private dMvSFamNC As Object   ' PEP3NIVEL -> qtd de familias NAO ADERENTES (ODI)
Private dMvSDif As Object     ' PEP3NIVEL -> soma das diferencas (ODI NAO ADERENTE)

' FASE 4.1: cache da aba CONFIG (chave -> valor) e das classes de viagem
Private dCfg As Object          ' CHAVE(str) -> valor(str)
Private dClsViagem As Object    ' CLASSE_CUSTO(str) -> 1


'==============================================================================
'  ROTINA PRINCIPAL
'==============================================================================
' ============================================================
'  [AT] MAT vs SERV AT  -  estruturas do modulo de avaliacao
' ============================================================
Private Const COR_HEADER      As Long = &H660301
Private Const COR_INCONF_BG   As Long = &HD6E4FC
Private Const COR_INCONF_FG   As Long = &H115AC5
Private Const COR_ADER_OK     As Long = &HCEEFCE
Private Const COR_ADER_DIV    As Long = &H9CEBFF
Private Const COR_ADER_ERR    As Long = &HCEC7FF
Private Const COR_GRUPO_A     As Long = &HFFF4F0
Private Const COR_GRUPO_B     As Long = &HFFFFFF
Private Const COR_SEM_GRUPO_A As Long = &HFFF9F9
Private Const COR_TIPO_D_BG   As Long = &HF7EBDD
Private Const COR_TIPO_D_FG   As Long = &H794E1F
Private Const COR_TIPO_C_BG   As Long = &HDAEFE2
Private Const COR_TIPO_C_FG   As Long = &H235637

Private Type tItem
    Empresa As String
    Segmento As String
    TipoObraAneel As String
    Pep3Nivel As String
    Pep As String
    Tipo As String
    Material As String
    TextoMaterial As String
    Uml As String
    ValorMoeda As Double
    QtdEntrada As Double
    Cls1 As String
    Cls2 As String
    Cls2Orig As String
    TipoCusto As String
    Mat As Double
    Srv As Double
    Aderencia As String
    Inconformidade As String
    PctMop As Double
    GrupoKey As String
End Type

Private aItens() As tItem
Private nItens As Long
Private aMatCorr() As String
Private aSrvCorr() As String
Private aTipoCorr() As String
Private nCorr As Long
Private dNormCache As Object   ' cache memoizado de NormClassif
Private dSemAcCache As Object  ' cache memoizado de SemAcento
Private dDescSrv As Object     ' COD_SERVICO -> descricao (Denominacao)

Sub GerarRelatorio()
    Dim t As Double: t = Timer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    On Error GoTo Falha

    ' 0a) Garante e carrega a aba CONFIG (parametros centralizados) - FASE 4
    gEtapa = "GarantirConfig": GarantirConfig
    gEtapa = "CarregarConfig": CarregarConfig

    ' 0) Remove abas obsoletas de execucoes anteriores (se existirem)
    Dim wsLixo As Worksheet
    Dim abasObsoletas As Variant
    abasObsoletas = Array("RISCO", "ANOMALIAS", "ANALISE REGIONAL", _
                          "PEND CLASSIFICACAO", "RESUMO POR PEP", "CLASSIFICACAO", _
                          "MAT vs SERV AT", "PREMISSAS", "CLASSE DE CUSTO", _
                          "PAINEL EXECUTIVO", "PORTFOLIO OBRA")
    Dim ix As Long
    For ix = 0 To UBound(abasObsoletas)
        On Error Resume Next
        Set wsLixo = ActiveWorkbook.Worksheets(abasObsoletas(ix))
        If Not wsLixo Is Nothing Then
            wsLixo.Delete
            Set wsLixo = Nothing
        End If
        On Error GoTo Falha
    Next ix

    ' 1) Localiza a aba com a base crua
    Set wsRaw = LocalizarBase()
    If wsRaw Is Nothing Then
        MsgBox "Nao encontrei a base crua. Verifique se ha uma aba com a coluna " & _
               "'Elemento PEP'.", vbExclamation: GoTo Fim
    End If

    ' 2) Mapeia colunas pelos cabecalhos
    If Not MapearColunas(wsRaw) Then
        MsgBox "Nao encontrei as colunas obrigatorias (Elemento PEP, Classificacao, " & _
               "Valor/moeda objeto, Qtd.total entrada e Material).", vbExclamation: GoTo Fim
    End If

    ' 3) Carrega os dados para memoria (rapido)
    CarregarDados wsRaw
    If nLin = 0 Then
        MsgBox "A base foi localizada, mas nao ha linhas de dados para processar.", vbExclamation
        GoTo Fim
    End If

    ' 3b) Carrega o catalogo de MATERIAIS (FAMILIA, CLS1/2/3)
    gEtapa = "CarregarCatalogoMateriais": CarregarCatalogoMateriais

    ' 3c) Descricoes embutidas de servico (prioridade) ANTES do catalogo externo,
    '     que preenche as descricoes faltantes a partir de TEXTO BREVE do arquivo.
    gEtapa = "CarregarDescServico": CarregarDescServico

    ' 3c') Carrega o catalogo de SERVICOS (CLS1/2/3, TIPO_APLIC, SEGMENTO + descricoes)
    gEtapa = "CarregarCatalogoServicos": CarregarCatalogoServicos

    ' 3c'') Reclassificacao curada de servicos (sobrepoe o catalogo)
    gEtapa = "CarregarReclassificacaoServicos": CarregarReclassificacaoServicos

    ' 3d) Carrega o catalogo de CLASSE DE CUSTO (CLS1/2/3 + TIPO_APLIC; marca RISCO)
    gEtapa = "CarregarCatalogoClasse": CarregarCatalogoClasse

    ' 3e) Carrega conversao de cabo (KG->m) e SRV COMBO (fator multiplicador)
    '     - usados para refinar a aderencia MATERIAL vs SERVICO (ajuste fino)
    gEtapa = "CarregarConversoesCabo": CarregarConversoesCabo
    gEtapa = "CarregarComboServico": CarregarComboServico

    ' 3f) De-para CLASSIFICACAO (familia/CLS2) -> TIPO (COM/UC/UAR)
    gEtapa = "CarregarTipoClassif": CarregarTipoClassif

    ' 3g) Equivalencias servico->material e familias de servico puro
    gEtapa = "CarregarEquivSrvMat": CarregarEquivSrvMat
    gEtapa = "CarregarSrvPuro": CarregarSrvPuro
    gEtapa = "CarregarFamiliaNivel": CarregarFamiliaNivel

    ' 4) Gera cada aba
    gEtapa = "Gerar_RazaoCJ": Gerar_RazaoCJ
    gEtapa = "Gerar_AlertasCriticos": Gerar_AlertasCriticos
    gEtapa = "Gerar_MaterialVsServico": Gerar_MaterialVsServico
    gEtapa = "Gerar_Material": Gerar_Material
    gEtapa = "Gerar_Servico": Gerar_Servico
    gEtapa = "Gerar_AnaliseCA": Gerar_AnaliseCA
    gEtapa = "Gerar_ServicoSemMaterial": Gerar_ServicoSemMaterial
    gEtapa = "Gerar_NaoClassificados": Gerar_NaoClassificados
    gEtapa = "Gerar_RacionalizacaoCOM": Gerar_RacionalizacaoCOM
    gEtapa = "Gerar_Regras": Gerar_Regras

    ' Checklist do processo: criado automaticamente (preserva marcacoes se ja existir)
    gEtapa = "GerarChecklistCKCP": GerarChecklistCKCP True

    ' FASE 5: ordena as guias por fluxo de leitura e ativa o painel
    gEtapa = "OrganizarAbas": OrganizarAbas

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    gEtapa = "MostrarTelaFuturista": MostrarTelaFuturista nLin, (Timer - t)
    GoTo Fim

Falha:
    MsgBox "Erro " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Etapa: " & gEtapa, vbCritical
Fim:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub


'==============================================================================
'  TELA DE RESULTADO FUTURISTA  (painel HUD desenhado com Shapes)
'==============================================================================
Private Sub MostrarTelaFuturista(ByVal nLin As Long, ByVal seg As Double)
    ' Best-effort (so formatacao visual): falha aqui nao altera a analise.
    On Error Resume Next
    LimparSplash

    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim win As Window: Set win = ActiveWindow
    Dim vr As Range: Set vr = win.VisibleRange

    Dim PW As Double, PH As Double, L As Double, T As Double
    PW = 380: PH = 200
    L = vr.Left + (vr.Width - PW) / 2
    T = vr.Top + (vr.Height - PH) / 2
    If L < vr.Left Then L = vr.Left + 20
    If T < vr.Top Then T = vr.Top + 20

    Dim acc As Long, txt As Long, mut As Long, bord As Long, painel As Long
    acc = RGB(0, 150, 80)        ' verde discreto
    txt = RGB(33, 37, 41)        ' quase preto
    mut = RGB(140, 146, 153)     ' cinza rotulo
    bord = RGB(222, 226, 230)    ' borda cinza clara
    painel = RGB(255, 255, 255)  ' branco

    ' painel principal - flat, branco, borda fina
    Dim pn As Shape
    Set pn = ws.Shapes.AddShape(msoShapeRoundedRectangle, L, T, PW, PH)
    pn.Name = "FX_PANEL"
    pn.Adjustments(1) = 0.04
    pn.Fill.ForeColor.RGB = painel
    pn.Line.ForeColor.RGB = bord
    pn.Line.Weight = 0.75
    pn.Shadow.Type = msoShadow25
    pn.Shadow.Transparency = 0.85
    pn.Shadow.Blur = 8

    ' titulo + subtitulo
    AddTxt ws, "FX_TITLE", L + 28, T + 30, PW - 56, 30, _
           ChrW(10003) & "  Relatorio gerado", "Segoe UI", 16, True, acc, msoAlignLeft
    AddTxt ws, "FX_SUB", L + 28, T + 60, PW - 56, 16, _
           "Analise de custo  /  CKCP", "Segoe UI", 9.5, False, mut, msoAlignLeft

    ' linha separadora
    Dim sep As Shape
    Set sep = ws.Shapes.AddLine(L + 28, T + 90, L + PW - 28, T + 90)
    sep.Name = "FX_SEP"
    sep.Line.ForeColor.RGB = bord
    sep.Line.Weight = 0.75

    ' metricas (texto simples, sem caixas)
    MetricBlock ws, "FX_M1", L + 28, T + 104, (PW - 56) / 2, _
                "Linhas processadas", Format(nLin, "#,##0"), txt, mut
    MetricBlock ws, "FX_M2", L + 28 + (PW - 56) / 2, T + 104, (PW - 56) / 2, _
                "Tempo", Format(seg, "0.0") & "s", txt, mut

    ' botao OK - outline simples
    Dim ok As Shape
    Set ok = ws.Shapes.AddShape(msoShapeRoundedRectangle, L + PW - 116, T + PH - 44, 88, 28)
    ok.Name = "FX_OK"
    ok.Adjustments(1) = 0.18
    ok.Fill.ForeColor.RGB = acc
    ok.Line.Visible = msoFalse
    With ok.TextFrame2.TextRange
        .Text = "OK"
        .Font.Name = "Segoe UI"
        .Font.Size = 10.5
        .Font.Bold = msoTrue
        .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .ParagraphFormat.Alignment = msoAlignCenter
    End With
    ok.TextFrame2.VerticalAnchor = msoAnchorMiddle
    ok.OnAction = "FecharSplashCKCP"

    ' O painel permanece na tela; o botao OK (FecharSplashCKCP) apaga as formas.
    On Error GoTo 0
End Sub

Private Sub AddTxt(ws As Worksheet, nm As String, L As Double, T As Double, _
                   W As Double, H As Double, s As String, fname As String, _
                   sz As Single, bold As Boolean, cor As Long, al As Long)
    ' Best-effort (so formatacao visual): falha aqui nao altera a analise.
    On Error Resume Next
    Dim tb As Shape
    Set tb = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, L, T, W, H)
    tb.Name = nm
    tb.Line.Visible = msoFalse
    tb.Fill.Visible = msoFalse
    With tb.TextFrame2.TextRange
        .Text = s
        .Font.Name = fname
        .Font.Size = sz
        .Font.Bold = IIf(bold, msoTrue, msoFalse)
        .Font.Fill.ForeColor.RGB = cor
        .ParagraphFormat.Alignment = al
    End With
    tb.TextFrame2.VerticalAnchor = msoAnchorMiddle
    tb.TextFrame2.WordWrap = msoFalse
    On Error GoTo 0
End Sub

Private Sub MetricBlock(ws As Worksheet, nm As String, L As Double, T As Double, _
                        W As Double, rotulo As String, valor As String, _
                        cor As Long, mut As Long)
    On Error Resume Next
    AddTxt ws, nm & "_L", L, T, W - 8, 14, rotulo, "Segoe UI", 8.5, False, mut, msoAlignLeft
    AddTxt ws, nm & "_V", L, T + 16, W - 8, 30, valor, "Segoe UI", 22, True, cor, msoAlignLeft
    On Error GoTo 0
End Sub

Private Sub LimparSplash()
    On Error Resume Next
    Dim wsx As Worksheet, sh As Shape
    For Each wsx In ActiveWorkbook.Worksheets
        For Each sh In wsx.Shapes
            If Left$(sh.Name, 3) = "FX_" Then sh.Delete
        Next sh
    Next wsx
    On Error GoTo 0
End Sub

Public Sub FecharSplashCKCP(Optional ByVal ignorar As Variant)
    gSplashOK = True
    LimparSplash
End Sub



'==============================================================================
'  PREPARAO
'==============================================================================
Private Function LocalizarBase() As Worksheet
    Dim ws As Worksheet, c As Range

    ' Prioriza a base crua do SAP. Abas geradas pelo proprio relatorio tambem
    ' podem ter "PEP" na linha 1, entao o fallback abaixo e mais restritivo.
    For Each ws In ActiveWorkbook.Worksheets
        Set c = ws.Rows(1).Find("Elemento PEP", LookAt:=xlWhole)
        If Not c Is Nothing Then Set LocalizarBase = ws: Exit Function
    Next ws

    Dim melhor As Worksheet, melhorScore As Long, score As Long
    For Each ws In ActiveWorkbook.Worksheets
        ' Aceita tambm base j normalizada com cabealho "PEP"
        Set c = ws.Rows(1).Find("PEP", LookAt:=xlWhole)
        If Not c Is Nothing Then
            If TemCabecalhosMinimos(ws) Then
                score = PontuarBase(ws)
                If score > melhorScore Then
                    Set melhor = ws
                    melhorScore = score
                End If
            End If
        End If
    Next ws
    If Not melhor Is Nothing Then Set LocalizarBase = melhor
End Function

Private Function MapearColunas(ws As Worksheet) As Boolean
    ' Busca por fragmentos SEM acento (robusto a problemas de codificacao).
    cPEP = ColLike(ws, Array("ELEMENTO PEP", "PEP"))
    cClasse = ColLike(ws, Array("CLASSE DE CUSTO", "CLASSE_CUSTO", "CLASSE CUSTO"))
    cDescClasse = ColLike(ws, Array("DESCR.CLASSE", "DENOM.CLASSE", "DESC_CLASSE"))
    cMaterial = ColLike(ws, Array("MATERIAL"))
    cTexto = ColLike(ws, Array("TEXTO BREVE", "TEXTO_MATERIAL"))
    cQtd = ColLike(ws, Array("QTD.TOTAL", "QTD_ENTRADA", "QTD ENTRADA"))
    cUML = ColLike(ws, Array("UNID.MEDIDA", "UML"))
    cValor = ColLike(ws, Array("VALOR/MOEDA", "VALOR_MOEDA", "VALOR MOEDA"))
    cClassif = ColLike(ws, Array("CLASSIFICA"))        ' Classificacao
    cDescSA = ColLike(ws, Array("DESCRICAO SA", "DESCRICAO_SA", "DESCR SA")) ' Descricao SA

    ' FASE 1.3: EMPRESA e DIVISAO priorizam o proprio cabecalho;
    ' fallback cruzado so quando o proprio nao existe na base.
    cEmpresa = ColLike(ws, Array("EMPRESA"))
    If cEmpresa = 0 Then cEmpresa = ColLike(ws, Array("DIVISAO"))
    cDivisao = ColLike(ws, Array("DIVISAO"))
    If cDivisao = 0 Then cDivisao = ColLike(ws, Array("EMPRESA"))

    cObj = ColLike(ws, Array("OBJETO"))
    cDenObj = ColLike(ws, Array("DENOMINACAO_OBJETO", "DENOMINACAO OBJETO"))
    cDenClasse = ColLike(ws, Array("DENOM_CLASSE_CUSTO", "DENOM CLASSE CUSTO"))
    cDocCompra = ColLike(ws, Array("DOC_COMPRA", "DOC COMPRA"))
    cNumDoc = ColLike(ws, Array("NUM_DOC", "NUM DOC"))
    cDenominacao = ColLike(ws, Array("DENOMINACAO"))
    cUsuario = ColLike(ws, Array("USUARIO"))
    cNumDocRef = ColLike(ws, Array("NUM_DOC_REF", "NUM DOC REF"))
    cDataLanc = ColLike(ws, Array("DATA_LANCAMENTO", "DATA LANCAMENTO"))
    cHora = ColLike(ws, Array("HORA"))
    cDataEntrada = ColLike(ws, Array("DATA_ENTRADA", "DATA ENTRADA"))
    cTipoDoc = ColLike(ws, Array("TIPO_DOC", "TIPO DOC"))
    cAno = ColLike(ws, Array("ANO"))
    cDataDoc = ColLike(ws, Array("DATA_DOCUMENTO", "DATA DOCUMENTO"))
    cLinhaLanc = ColLike(ws, Array("LINHA LANCAMENTO", "LINHA_LANCAMENTO"))
    cODI = ColLike(ws, Array("ODI_ANEEL", "ODI ANEEL"))

    ' FASE 1.3: "SA" so por igualdade exata (a busca por substring casava
    ' "DESCRICAO SA" quando a coluna SA nao existia).
    cSA = ColExata(ws, Array("SA"))

    cDocEstorno = ColLike(ws, Array("DOC_ESTORNO", "DOC ESTORNO"))
    cOrgEstorno = ColLike(ws, Array("ORG_ESTORNO", "ORG ESTORNO"))
    cEstorno = ColLike(ws, Array("ESTORNO"))
    cRefEstorno = ColLike(ws, Array("REF_ESTORNO", "REF ESTORNO"))
    cOperRef = ColLike(ws, Array("OPERACAO_REFERENCIA", "OPERACAO REFERENCIA"))
    cCLS1Raw = ColLike(ws, Array("CLS1"))
    cCLS2Raw = ColLike(ws, Array("CLS2"))
    cCLS3Raw = ColLike(ws, Array("CLS3"))
    cTipoAplicRaw = ColLike(ws, Array("TIPO_APLICACAO", "TIPO APLICACAO", "TIPO APLIC"))
    MapearColunas = TemCabecalhosMinimos(ws)
End Function


Private Function TemCabecalhosMinimos(ws As Worksheet) As Boolean
    TemCabecalhosMinimos = _
        (ColLike(ws, Array("ELEMENTO PEP", "PEP")) > 0) And _
        (ColLike(ws, Array("CLASSIFICA")) > 0) And _
        (ColLike(ws, Array("VALOR/MOEDA", "VALOR_MOEDA", "VALOR MOEDA")) > 0) And _
        (ColLike(ws, Array("QTD.TOTAL", "QTD_ENTRADA", "QTD ENTRADA")) > 0) And _
        (ColLike(ws, Array("MATERIAL")) > 0)
End Function

Private Function PontuarBase(ws As Worksheet) As Long
    PontuarBase = ws.Cells(ws.Rows.Count, ColLike(ws, Array("ELEMENTO PEP", "PEP"))).End(xlUp).Row
    If ColLike(ws, Array("NUM_DOC", "NUM DOC")) > 0 Then PontuarBase = PontuarBase + 1000000
    If ColLike(ws, Array("DATA_LANCAMENTO", "DATA LANCAMENTO")) > 0 Then PontuarBase = PontuarBase + 1000000
    If ColLike(ws, Array("OPERACAO_REFERENCIA", "OPERACAO REFERENCIA")) > 0 Then PontuarBase = PontuarBase + 1000000
    If ColLike(ws, Array("OBJETO")) > 0 Then PontuarBase = PontuarBase + 1000000
End Function

' Procura a 1 coluna cujo cabealho CONTM qualquer fragmento da lista
' (case-insensitive, ignora acentos). Prioriza correspondncia exata.
Private Function ColLike(ws As Worksheet, frags As Variant) As Long
    Dim ult As Long, j As Long, i As Long, hdr As String, fr As String
    ult = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    ' 1 passada: igualdade exata (sem acento)
    For i = LBound(frags) To UBound(frags)
        fr = SemAcento(UCase$(CStr(frags(i))))
        For j = 1 To ult
            hdr = SemAcento(UCase$(Trim$(CStr(ws.Cells(1, j).Value))))
            If hdr = fr Then ColLike = j: Exit Function
        Next j
    Next i
    ' 2 passada: contm o fragmento
    For i = LBound(frags) To UBound(frags)
        fr = SemAcento(UCase$(CStr(frags(i))))
        For j = 1 To ult
            hdr = SemAcento(UCase$(Trim$(CStr(ws.Cells(1, j).Value))))
            If InStr(hdr, fr) > 0 Then ColLike = j: Exit Function
        Next j
    Next i
    ColLike = 0
End Function

' Remove acentos comuns para comparao robusta (usa cdigos Unicode - imune
' a problemas de codificao do arquivo).
Private Function SemAcento(ByVal s As String) As String
    If dSemAcCache Is Nothing Then Set dSemAcCache = CreateObject("Scripting.Dictionary")
    If dSemAcCache.Exists(s) Then SemAcento = dSemAcCache(s): Exit Function
    Dim sOrig As String: sOrig = s
    Dim cods As Variant, subs As Variant, k As Long
    s = UCase$(s)
    ' maisculas:
    cods = Array(192, 193, 194, 195, 196, 201, 202, 200, 205, 206, 211, 212, 213, 214, 218, 219, 199)
    subs = Array("A", "A", "A", "A", "A", "E", "E", "E", "I", "I", "O", "O", "O", "O", "U", "U", "C")
    For k = LBound(cods) To UBound(cods)
        s = Replace(s, ChrW(CLng(cods(k))), subs(k))
    Next k
    dSemAcCache(sOrig) = s
    SemAcento = s
End Function

Private Sub CarregarDados(ws As Worksheet)
    Dim ult As Long, ultCol As Long
    ult = ws.Cells(ws.Rows.Count, cPEP).End(xlUp).Row
    If ult < 2 Then
        ReDim dados(1 To 1, 1 To 1)
        nLin = 0
        Exit Sub
    End If

    ultCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1 > ultCol Then
        ultCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    End If
    rawColCount = ultCol
    rawHeaders = ws.Range(ws.Cells(1, 1), ws.Cells(1, ultCol)).Value
    dados = ws.Range(ws.Cells(2, 1), ws.Cells(ult, ultCol)).Value
    nLin = 0
    Dim i As Long
    For i = 1 To UBound(dados, 1)
        If Trim$(CStr(dados(i, cPEP))) <> "" Then nLin = nLin + 1
    Next i
End Sub

Private Function ValorCampo(ByVal lin As Long, ByVal col As Long, Optional ByVal padrao As Variant = "") As Variant
    If col <= 0 Then
        ValorCampo = padrao
    ElseIf col > UBound(dados, 2) Then
        ValorCampo = padrao
    Else
        ValorCampo = dados(lin, col)
    End If
End Function

Private Function TextoCampo(ByVal lin As Long, ByVal col As Long, Optional ByVal padrao As String = "") As String
    TextoCampo = Trim$(CStr(ValorCampo(lin, col, padrao)))
End Function

Private Function ValorMatriz(m As Variant, ByVal lin As Long, ByVal col As Long, Optional ByVal padrao As Variant = "") As Variant
    If col <= 0 Then
        ValorMatriz = padrao
    ElseIf col > UBound(m, 2) Then
        ValorMatriz = padrao
    Else
        ValorMatriz = m(lin, col)
    End If
End Function

Private Function TextoMatriz(m As Variant, ByVal lin As Long, ByVal col As Long, Optional ByVal padrao As String = "") As String
    TextoMatriz = Trim$(CStr(ValorMatriz(m, lin, col, padrao)))
End Function

Private Function LinhaCLS1(ByVal lin As Long) As String
    LinhaCLS1 = TextoCampo(lin, cCLS1Raw)
End Function

Private Function LinhaCLS2(ByVal lin As Long) As String
    LinhaCLS2 = TextoCampo(lin, cCLS2Raw)
End Function

Private Function LinhaCLS3(ByVal lin As Long) As String
    LinhaCLS3 = TextoCampo(lin, cCLS3Raw)
End Function

Private Function LinhaTipoAplic(ByVal lin As Long) As String
    LinhaTipoAplic = TextoCampo(lin, cTipoAplicRaw)
End Function

Private Function MatInfoLinha(ByVal lin As Long, ByVal idx As Long) As String
    MatInfoLinha = CatInfo(dados(lin, cMaterial), idx)
    If MatInfoLinha <> "" Then Exit Function

    ' Fallback: classificacao por CLASSE DE CUSTO (idx 1=CLS1, 2=CLS2, 3=CLS3
    ' mapeia para CCInfo 0=CLS1, 1=CLS2, 2=CLS3)
    If idx >= 1 And idx <= 3 Then
        MatInfoLinha = CCInfo(ValorCampo(lin, cClasse), idx - 1)
        If MatInfoLinha <> "" Then Exit Function
    End If

    Select Case idx
        Case 0: MatInfoLinha = "(SEM FAMILIA)"
        Case 1: MatInfoLinha = LinhaCLS1(lin)
        Case 2: MatInfoLinha = LinhaCLS2(lin)
        Case 3: MatInfoLinha = LinhaCLS3(lin)
    End Select
End Function

' De-para fixo: codigo de servico -> CLS2 (familia) forcada.
' Estes servicos instalam cabo (equivalem a 3 cabos) -> familia COND PROT,
' para casar com o material cabo na aba MATERIAL vs SERVICO.
Private Function Cls2SrvOverride(codSrv As Variant) As String
    Cls2SrvOverride = ""
    Dim c As String: c = NormCod(codSrv)
    Select Case c
        Case "5500000582", "5500000575"
            Cls2SrvOverride = "COND PROT"
    End Select
End Function

Private Function SrvInfoLinha(ByVal lin As Long, ByVal idx As Long) As String
    ' Override de CLS2 para servicos especificos (forca a familia do cabo)
    If idx = 1 Then
        Dim ov As String: ov = Cls2SrvOverride(dados(lin, cMaterial))
        If ov <> "" Then SrvInfoLinha = ov: Exit Function
    End If
    SrvInfoLinha = SrvInfo(dados(lin, cMaterial), idx)
    If SrvInfoLinha <> "" Then Exit Function

    ' Fallback: classificacao por CLASSE DE CUSTO (idx 0=CLS1, 1=CLS2, 2=CLS3
    ' mapeia direto para CCInfo)
    If idx >= 0 And idx <= 2 Then
        SrvInfoLinha = CCInfo(ValorCampo(lin, cClasse), idx)
        If SrvInfoLinha <> "" Then Exit Function
    End If

    Select Case idx
        Case 0: SrvInfoLinha = LinhaCLS1(lin)
        Case 1: SrvInfoLinha = LinhaCLS2(lin)
        Case 2: SrvInfoLinha = LinhaCLS3(lin)
        Case 3: SrvInfoLinha = LinhaTipoAplic(lin)
    End Select
End Function

Private Function TipoPEPCodigo(ByVal pep As String) As String
    Select Case UCase$(Right$(Trim$(pep), 2))
        Case ".I": TipoPEPCodigo = "I"
        Case ".D": TipoPEPCodigo = "D"
        Case ".M": TipoPEPCodigo = "M"
        Case Else: TipoPEPCodigo = "S"
    End Select
End Function

Private Function TipoPEPANEEL(ByVal pep As String) As String
    Select Case TipoPEPCodigo(pep)
        Case "I": TipoPEPANEEL = "ODI"
        Case "D": TipoPEPANEEL = "ODD"
        Case "M": TipoPEPANEEL = "ODM"
        Case Else: TipoPEPANEEL = "OUTRO"
    End Select
End Function
Private Function ClassificacaoPendente(ByVal cls1 As String, ByVal cls2 As String, ByVal cls3 As String) As Boolean
    Dim s1 As String, s2 As String, s3 As String
    s1 = UCase$(SemAcento(Trim$(cls1)))
    s2 = UCase$(SemAcento(Trim$(cls2)))
    s3 = UCase$(SemAcento(Trim$(cls3)))
    ClassificacaoPendente = (s3 = "" Or s3 = "CLASSIFICAR" Or s2 = "CLASSIFICAR" Or s1 = "CLASSIFICAR")
End Function


'==============================================================================
'  CATLOGO DE MATERIAIS (FAMILIA, CLS1, CLS2, CLS3)
'==============================================================================
Private Sub CarregarCatalogoMateriais()
    Set dCatMat = CreateObject("Scripting.Dictionary")

    ' 1) Tenta o caminho padro em Downloads
    Dim caminho As String
    caminho = CaminhoCatalogo("CAT_MATERIAIS", _
        "%USERPROFILE%\Downloads\MATERIAS_ATUAIS_4.xlsx;%USERPROFILE%\Downloads\MATERIAS_ATUAIS (4).xlsx;" & _
        "%USERPROFILE%\Downloads\MATERIAS_ATUAIS (2).xlsx;%USERPROFILE%\Downloads\MATERIAS_ATUAIS.xlsx")

    ' 2) Se no achar, pede para o usurio selecionar
    If caminho = "" Then
        Dim f As Variant
        f = Application.GetOpenFilename( _
            "Excel (*.xls*),*.xls*", , _
            "Selecione o catalogo de MATERIAIS (MATERIAS_ATUAIS). Cancele para pular.")
        If f = False Then Exit Sub   ' usurio cancelou -> segue sem catlogo
        caminho = CStr(f)
    End If

    On Error GoTo SemCat
    Dim wb As Workbook, ws As Worksheet, arr As Variant
    Set wb = Workbooks.Open(caminho, ReadOnly:=True, UpdateLinks:=0)
    Set ws = wb.Worksheets(1)

    Dim cCod As Long, cFam As Long, c1 As Long, c2 As Long, c3 As Long
    cCod = ColLike(ws, Array("COD MATERIAL", "COD_MATERIAL", "MATERIAL"))
    cFam = ColLike(ws, Array("FAMILIA"))
    c1 = ColLike(ws, Array("CLS1"))
    c2 = ColLike(ws, Array("CLS2"))
    c3 = ColLike(ws, Array("CLS3"))
    If cCod = 0 Then wb.Close SaveChanges:=False: Exit Sub

    Dim ult As Long
    ult = ws.Cells(ws.Rows.Count, cCod).End(xlUp).Row
    arr = ws.Range(ws.Cells(2, 1), ws.Cells(ult, ws.UsedRange.Columns.Count)).Value

    Dim i As Long, cod As String, fam As String, v1 As String, v2 As String, v3 As String
    For i = 1 To UBound(arr, 1)
        cod = NormCod(arr(i, cCod))
        If cod <> "" And Not dCatMat.Exists(cod) Then
            fam = TextoMatriz(arr, i, cFam)
            v1 = TextoMatriz(arr, i, c1)
            v2 = TextoMatriz(arr, i, c2)
            v3 = TextoMatriz(arr, i, c3)
            dCatMat(cod) = fam & "|" & v1 & "|" & v2 & "|" & v3
        End If
    Next i
    wb.Close SaveChanges:=False
    Exit Sub
SemCat:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

' Normaliza cdigo de material para casar base x catlogo
Private Sub CarregarDescServico()
    ' De-para COD_SERVICO -> Denominacao (catalogo embutido).
    Set dDescSrv = CreateObject("Scripting.Dictionary")
    ' Tabela embutida dividida em 9 blocos (limite de tamanho de procedimento).
    CarregarDescServicoP1
    CarregarDescServicoP2
    CarregarDescServicoP3
    CarregarDescServicoP4
    CarregarDescServicoP5
    CarregarDescServicoP6
    CarregarDescServicoP7
    CarregarDescServicoP8
    CarregarDescServicoP9
End Sub

Private Sub CarregarDescServicoP1()
    dDescSrv("5013300045") = "ATEND EMERG INST TRAFO 1F ATE 15KVA"
    dDescSrv("5013300046") = "ATEND EMERG RET TRAFO 1F ATE 15KVA"
    dDescSrv("5015000019") = "RET COND NU ALUMINIO <=1/0 SRD"
    dDescSrv("5015000022") = "INST CON 4PLEX BT 3X10(10)-3X120(70) SRD"
    dDescSrv("5020000092") = "INST POSTE CONCRETO ATE 14M LM"
    dDescSrv("5020000093") = "RET POSTE CONCRETO ATE 14M LM"
    dDescSrv("5020000098") = "INST POSTE DE FIBRA 7 A 19M LM"
    dDescSrv("5020000099") = "RET POSTE DE FIBRA 7 A 19M LM"
    dDescSrv("5020000107") = "RESTAURAR PASSEIO SIMPLES CIMENTADO LM"
    dDescSrv("5020000118") = "RET POSTE DE MADEIRA LM"
    dDescSrv("5020100056") = "BASE CONCRETO 1CAMADA PARA POSTE LM"
    dDescSrv("5020100057") = "BASE CONCRETO 2CAMADAS PARA POSTE LM"
    dDescSrv("5020100058") = "BASE CONCRETO C/ MANILHA PARA POSTE LM"
    dDescSrv("5020200034") = "INST ESTAI ANCORA LM"
    dDescSrv("5020200035") = "RET ESTAI ANCORA LM"
    dDescSrv("5020200036") = "INST ESTAI SUBSOLO BASE REFORCADA LM"
    dDescSrv("5021000131") = "INST ESTR PRI 2/3F SUS/PAS SP CRUZ SP LM"
    dDescSrv("5021000132") = "RET ESTR PRI 2/3F SUS/PAS SP CRUZ SP LM"
    dDescSrv("5021000133") = "INST ESTR PRI 2/3F SUS/PAS DP CRUZ DP LM"
    dDescSrv("5021000134") = "RET ESTR PRI 2/3F SUS/PAS DP CRUZ DP LM"
    dDescSrv("5021000135") = "INST ESTR PRI 2/3F ANC/AMA SP CRUZ DP LM"
    dDescSrv("5021000136") = "RET ESTR PRI 2/3F ANC/AMA SP CRUZ DP LM"
    dDescSrv("5021000137") = "INST ESTR PRI 2/3F ANC/AMA DP CRUZ DP LM"
    dDescSrv("5021000138") = "RET ESTR PRI 2/3F ANC/AMA DP CRUZ DP LM"
    dDescSrv("5021000139") = "INST ESTR PRI 1F SUSP/PAS SP S/CRUZ LM"
    dDescSrv("5021000140") = "RET ESTR PRI 1F SUSP/PAS SP S/CRUZ LM"
    dDescSrv("5021000142") = "RET ESTR PRI 1F SUSP/PAS DP S/CRUZ LM"
    dDescSrv("5021000143") = "INST ESTR PRI 1F ANC/AMA SP S/CRUZ LM"
    dDescSrv("5021000144") = "RET ESTR PRI 1F ANC/AMA SP S/CRUZ LM"
    dDescSrv("5021000145") = "INST ESTR PRI 1F ANC/AMA DP S/CRUZ LM"
    dDescSrv("5021000146") = "RET ESTR PRI 1F ANC/AMA DP S/CRUZ LM"
    dDescSrv("5021000147") = "INST ESTR PRI 2/3F SUS/PAS SP S/CRUZ LM"
    dDescSrv("5021000153") = "INST ESTR PRI 3F ANC/AMA CRUZ DP HT LM"
    dDescSrv("5021000154") = "RET ESTR PRI 3F ANC/AMA CRUZ DP HT LM"
    dDescSrv("5021000155") = "INST ESTR PRI 3F ANC/AMA S/CRUZ HTE LM"
    dDescSrv("5021000158") = "INST ESTR PRI RDC PASSANTE SIMPLES LM"
    dDescSrv("5021000159") = "RET ESTR PRI RDC PASSANTE SIMPLES LM"
    dDescSrv("5021000160") = "INST ESTR PRI RDC 3F ANCORAGEM SIMPLE LM"
    dDescSrv("5021000161") = "RET ESTR PRI RDC 3F ANCORAGEM SIMPLES LM"
    dDescSrv("5021000162") = "INST ESTR PRI RDC 3F ANCORAGEM DUPLA LM"
    dDescSrv("5021000163") = "RET ESTR PRI RDC 3F ANCORAGEM DUPLA LM"
    dDescSrv("5021000164") = "INST ESTR PRI RDC N3S-CE LM"
    dDescSrv("5021100014") = "INST ESTRUTURA SECUNDARIA LM"
    dDescSrv("5021100015") = "RET ESTRUTURA SECUNDARIA LM"
    dDescSrv("5021200044") = "INST ISOLADOR ROLDANA LM"
    dDescSrv("5022000234") = "INST COND NU ALUMINIO > 1/0 LM"
    dDescSrv("5022000235") = "RET COND NU ALUMINIO > 1/0 LM"
    dDescSrv("5022000236") = "INST COND NU ALUMINIO <= 1/0 LM"
    dDescSrv("5022000237") = "RET COND NU ALUMINIO <= 1/0 LM"
    dDescSrv("5022000238") = "INST EMENDA COND NU AL/CU LM"
    dDescSrv("5022000250") = "INST FLY TAP/CRUZAM AEREO LM"
    dDescSrv("5022000251") = "RET FLY TAP /CRUZAM AEREO LM"
    dDescSrv("5022000266") = "INST COND 4PLEX BT 3X10(10)-3X120(70) LM"
    dDescSrv("5022000267") = "RET COND 4PLEX BT 3X10(10)-3X120(70) LM"
    dDescSrv("5022000276") = "INST COND COBERTO RDC 35-70 MM2 LM"
    dDescSrv("5022000278") = "INST COND COBERTO RDC 95-185 MM2 LM"
    dDescSrv("5022000280") = "INST CABO MENSAGEIRO RDC LM"
    dDescSrv("5022000286") = "INST CONEX OU EMENDA DE CABO RDC LM"
    dDescSrv("5022200033") = "INST ATERRAMENTO COMPL 1HASTE LM"
    dDescSrv("5022200034") = "INST ATERRAMENTO COMPL 5-8HASTES LM"
    dDescSrv("5022200037") = "INST SECCION E ATERRAM CERCA LD LM PDE"
    dDescSrv("5023000178") = "INST BANCO DE BATERIAS EQUIPAMENTOS ARD"
    dDescSrv("5023000179") = "RET BANCO DE BATERIAS EQUIPAMENTOS ARD"
    dDescSrv("5024000070") = "AVISO ENTREGA CARTA DESLIG PROG S/ PROT"
    dDescSrv("5024000071") = "AVISO ENTREGA CARTA DESLIG PROG C/ PROT"
    dDescSrv("5024000072") = "INST CONEXAO RAMAL DE CONEXAO LM"
    dDescSrv("5024000073") = "RET CONEXAO RAMAL DE CONEXAO LM"
    dDescSrv("5025300103") = "INST ILUMINACAO PUBLICA LM"
    dDescSrv("5025300104") = "RET ILUMINACAO PUBLICA LM"
    dDescSrv("5025300105") = "INST CONEXAO ILUMINACAO PUBLICA LM"
    dDescSrv("5025300106") = "RET CONEXAO ILUMINACAO PUBLICA LM"
    dDescSrv("5025300111") = "OPERAR/MANOB CHAVE/GRAMPO LV LM"
    dDescSrv("5025300122") = "INST ESTR CP RELIGADOR/SEC 3F >400A LM"
    dDescSrv("5025300124") = "INST REGULADOR TENSAO POSTE/PLATAF LM"
    dDescSrv("5025300125") = "RET REGULADOR TENSAO POSTE/PLATAF LM"
End Sub

Private Sub CarregarDescServicoP2()
    dDescSrv("5025300139") = "INST TRAFO DT POSTE/PLATAF 1/2F LM"
    dDescSrv("5025300140") = "RET TRAFO DT POSTE/PLATAF 1/2F LM"
    dDescSrv("5025300141") = "INST TRAFO DT POSTE/PLATAF 3F LM"
    dDescSrv("5025300142") = "RET TRAFO DT POSTE/PLATAF 3F LM"
    dDescSrv("5025300143") = "INST ESTR CP TRAFO DT 1F FN LM"
    dDescSrv("5025300144") = "RET ESTR CP TRAFO DT 1F FN LM"
    dDescSrv("5025300147") = "INST ESTR CP TRAFO DT 3F LM"
    dDescSrv("5025300148") = "RET ESTR CP TRAFO DT 3F LM"
    dDescSrv("5025300151") = "INST TRAFO DT C/ TALHA MANUAL LM"
    dDescSrv("5025300152") = "RET TRAFO DT C/ TALHA MANUAL LM"
    dDescSrv("5025300162") = "INST PARA-RAIOS MT LM"
    dDescSrv("5025300163") = "RET PARA-RAIOS MT LM"
    dDescSrv("5025300167") = "RET CHAVE FUSIVEL RELIG LM"
    dDescSrv("5025300172") = "INST CHAVE FUSIVEL 15/25KV LM"
    dDescSrv("5025300173") = "RET CHAVE FUSIVEL 15/25KV LM"
    dDescSrv("5025300174") = "INST CHAVE FUSIVEL 34,5KV LM"
    dDescSrv("5025300175") = "RET CHAVE FUSIVEL 34,6KV LM"
    dDescSrv("5025300176") = "INST CHAVE FACA LM"
    dDescSrv("5025300177") = "RET CHAVE FACA LM"
    dDescSrv("5025300234") = "TURMA LM 5H CAM EM ATEND EMERGENCIA"
    dDescSrv("5025300254") = "TURMA LM 5H CAM ATEND EMERGENCIAL HE25%"
    dDescSrv("5025300323") = "FISCALIZACAO DE OBRA LM"
    dDescSrv("5025300324") = "ENCERRAMENTO TECNICO DA OBRA LM"
    dDescSrv("5025300325") = "COMISSIONAMENTO DE OBRA LM"
    dDescSrv("5025300326") = "PROJETO ASBUILT LM"
    dDescSrv("5026000019") = "PODA ARV MED P >6<12M C REB (2MR) 3P PO"
    dDescSrv("5030000073") = "TURMA LV INST POSTE TRAFO LV"
    dDescSrv("5030000074") = "TURMA LV RET POSTE TRAFO LV"
    dDescSrv("5031200310") = "INST ESTR PRI 2/3F SUS/PAS SP CRUZ SP LV"
    dDescSrv("5031200311") = "RET ESTR PRI 2/3F SUS/PAS SP CRUZ SP LV"
    dDescSrv("5031200312") = "INST ESTR PRI 2/3F SUS/PAS DP CRUZ DP LV"
    dDescSrv("5031200313") = "RET ESTR PRI 2/3F SUS/PAS DP CRUZ DP LV"
    dDescSrv("5031200314") = "INST ESTR PRI 2/3F ANC/AMA SP CRUZ DP LV"
    dDescSrv("5031200315") = "RET ESTR PRI 2/3F ANC/AMA SP CRUZ DP LV"
    dDescSrv("5031200316") = "INST ESTR PRI 2/3F ANC/AMA DP CRUZ DP LV"
    dDescSrv("5031200317") = "RET ESTR PRI 2/3F ANC/AMA DP CRUZ DP LV"
    dDescSrv("5031200318") = "INSTALAR EST PRI 1F SUS/PAS SP S/CRUZ LV"
    dDescSrv("5031200319") = "RET ESTR PRI 1F SUSP/PAS SP S/CRUZ LV"
    dDescSrv("5031200320") = "INST ESTR PRI 1F SUSP/PAS DP S/CRUZ LV"
    dDescSrv("5031200321") = "RET ESTR PRI 1F SUSP/PAS DP S/CRUZ LV"
    dDescSrv("5031200322") = "INST ESTR PRI 1F ANC/AMA SP S/CRUZ LV"
    dDescSrv("5031200323") = "RET ESTR PRI 1F ANC/AMA SP S/CRUZ LV"
    dDescSrv("5031200324") = "INST ESTR PRI 1F ANC/AMA DP S/CRUZ LV"
    dDescSrv("5031200325") = "RET ESTR PRI 1F ANC/AMA DP S/CRUZ LV"
    dDescSrv("5031200338") = "TURMA LV INST POSTE ESTRUT PRIMARIA LV"
    dDescSrv("5031200339") = "TURMA LV RET POSTE ESTRUT PRIMARIA LV"
    dDescSrv("5031200347") = "INST ESTR PRI RDC PASSANTE SIMPLES LV"
    dDescSrv("5031200348") = "RET ESTR PRI RDC PASSANTE SIMPLES LV"
    dDescSrv("5031200349") = "INST ESTR PRI RDC 3F ANCORAGEM SIMPLE LV"
    dDescSrv("5031200350") = "RET ESTR PRI RDC 3F ANCORAGEM SIMPLES LV"
    dDescSrv("5031200351") = "INST ESTR PRI RDC 3F ANCORAGEM DUPLA LV"
    dDescSrv("5031200352") = "RET ESTR PRI RDC 3F ANCORAGEM DUPLA LV"
    dDescSrv("5031200353") = "INST ESTR PRI RDC N3S-CE LV"
    dDescSrv("5032000202") = "INST COND NU ALUMINIO > 1/0 LV"
    dDescSrv("5032000203") = "RET COND NU ALUMINIO > 1/0 LV"
    dDescSrv("5032000204") = "INST COND NU ALUMINIO <= 1/0 LV"
    dDescSrv("5032000205") = "RET COND NU ALUMINIO <= 1/0 LV"
    dDescSrv("5032000207") = "INST EMENDA COND NU AL/CU LV"
    dDescSrv("5032000219") = "INST ESTRIBO OU GRAMPO LINHA VIVA RD LV"
    dDescSrv("5032000224") = "INST COND COBERTO RDC 35-70 MM2 LV"
    dDescSrv("5032000225") = "RET COND COBERTO RDC 35-70 MM2 LV"
    dDescSrv("5032000226") = "INST COND COBERTO RDC 95-185 MM2 LV"
    dDescSrv("5032000227") = "RET COND COBERTO RDC 95-185 MM2 LV"
    dDescSrv("5032000228") = "INST CABO MENSAGEIRO RDC LV"
    dDescSrv("5032000229") = "RET CABO MENSAGEIRO RDC LV"
    dDescSrv("5033000292") = "INST CHAVE FUSIVEL 15/25KV LV"
    dDescSrv("5033000293") = "RET CHAVE FUSIVEL 15/25KV LV"
    dDescSrv("5033000294") = "INST CHAVE FUSIVEL 34,5KV LV"
    dDescSrv("5033000295") = "RET CHAVE FUSIVEL 34,5KV LV"
    dDescSrv("5033000296") = "INST CHAVE FACA LV"
    dDescSrv("5033000297") = "RET CHAVE FACA LV"
    dDescSrv("5033000308") = "TURMA LV INST RELIGADOR/SECCIONAD 3F LV"
    dDescSrv("5033000314") = "TURMA LV INST TRAFO DT LV"
    dDescSrv("5033000315") = "TURMA LV RET TRAFO DT LV"
    dDescSrv("5033000318") = "INST PARA-RAIOS MT LV"
End Sub

Private Sub CarregarDescServicoP3()
    dDescSrv("5044000003") = "APROVACOES E AUTORIZACOES"
    dDescSrv("5100100024") = "SERVICO DE TELECOM"
    dDescSrv("5150100002") = "SERVICOS LOGISTICOS"
    dDescSrv("5261000001") = "SERVICOS AMBIENTAIS"
    dDescSrv("5500000051") = "ABR119_A_INST COBERTA PROT VAO REDE-LM"
    dDescSrv("5500000054") = "ABR120_A_RET COBERTA PROT VAO REDE-LM"
    dDescSrv("5500000150") = "ABR229_A_INST EM SUB TERMINAL BIMETAL-LM"
    dDescSrv("5500000151") = "ABR229_B_INST EM SUB TERMINAL BIMETAL-LM"
    dDescSrv("5500000153") = "ABR230_A_RET EM SUB TERMINAL BIMETAL-LM"
    dDescSrv("5500000154") = "ABR230_B_RET EM SUB TERMINAL BIMETAL-LM"
    dDescSrv("5500000304") = "ABR501_A_INST SUB POS MAL EST PAS LV-LV"
    dDescSrv("5500000307") = "ABR502_A_RET SUB POS MAL EST PAS LV-LV"
    dDescSrv("5500000310") = "ABR503_A_INS SUB POS MAL EST C/ENCALV-LV"
    dDescSrv("5500000313") = "ABR504_A_RET SUB POS MAL EST C/ENCALV-LV"
    dDescSrv("5500000326") = "ABR508_B_INST 1ISO PINO/CADEIA C/LV-LV"
    dDescSrv("5500000397") = "ABR706_A_ATERR/SECCIONAMENTO CERCAS-LM"
    dDescSrv("5500000399") = "ABR706_B_ATERR/SECCIONAMENTO CERCAS-LM"
    dDescSrv("5500000400") = "ABR706_C_ATERR/SECCIONAMENTO CERCAS-LM"
    dDescSrv("5500000409") = "ABR801_A_INSTALACAO DE CHAVE FACA-LM"
    dDescSrv("5500000410") = "ABR801_B_INSTALACAO DE CHAVE FACA-LM"
    dDescSrv("5500000412") = "ABR802_A_RETIRADA DE CHAVE FACA-LM"
    dDescSrv("5500000413") = "ABR802_B_RETIRADA DE CHAVE FACA-LM"
    dDescSrv("5500000421") = "ABR805_A_INST BC CONDENS TP-TC C/LV-LV"
    dDescSrv("5500000476") = "AHO102_B_INST TIRAN POSTE MOCO MT/BT-LM"
    dDescSrv("5500000478") = "AHO103_A_RET TIRAN POSTE MOCO MT/BT-LM"
    dDescSrv("5500000479") = "AHO103_B_RET TIRAN POSTE MOCO MT/BT-LM"
    dDescSrv("5500000481") = "AHO108_A_CONCRETAGEM DE CIMENTO PARA POS"
    dDescSrv("5500000482") = "AHO108_B_CONCRETAGEM DE CIMENTO PARA POS"
    dDescSrv("5500000483") = "AHO108_C_CONCRETAGEM DE CIMENTO PARA POS"
    dDescSrv("5500000484") = "AHO110_A_INST DE POSTE MT E ESTRUTURA-LM"
    dDescSrv("5500000485") = "AHO110_B_INST DE POSTE MT E ESTRUTURA-LM"
    dDescSrv("5500000486") = "AHO110_C_INST DE POSTE MT E ESTRUTURA-LM"
    dDescSrv("5500000487") = "AHO111_A_INST DE POSTE BT E ESTRUTURA-LM"
    dDescSrv("5500000488") = "AHO111_B_INST DE POSTE BT E ESTRUTURA-LM"
    dDescSrv("5500000489") = "AHO111_C_INST DE POSTE BT E ESTRUTURA-LM"
    dDescSrv("5500000490") = "AHO112_A_INST POSTE DE MT CONCRETAGEM-LM"
    dDescSrv("5500000491") = "AHO112_B_INST POSTE DE MT CONCRETAGEM-LM"
    dDescSrv("5500000492") = "AHO112_C_INST POSTE DE MT CONCRETAGEM-LM"
    dDescSrv("5500000493") = "AHO113_A_INST POSTE DE BT CONCRETAGEM-LM"
    dDescSrv("5500000494") = "AHO113_B_INST POSTE DE BT CONCRETAGEM-LM"
    dDescSrv("5500000514") = "AHO121_A_INST ESTRUTURA MT REDE NUA-LM"
    dDescSrv("5500000515") = "AHO121_B_INST ESTRUTURA MT REDE NUA-LM"
    dDescSrv("5500000516") = "AHO121_C_INST ESTRUTURA MT REDE NUA-LM"
    dDescSrv("5500000517") = "AHO122_A_INST ESTRUT MT REDE ISOLADA-LM"
    dDescSrv("5500000520") = "AHO123_A_INST ARMACAO SECUNDARIA-LM"
    dDescSrv("5500000521") = "AHO123_B_INST ARMACAO SECUNDARIA-LM"
    dDescSrv("5500000522") = "AHO123_C_INST ARMACAO SECUNDARIA-LM"
    dDescSrv("5500000526") = "AHO125_A_RET ESTRUTURA ANCORAGEM MT-LM"
    dDescSrv("5500000529") = "AHO126_A_RETIRADA ARMACAO SECUNDARIA-LM"
    dDescSrv("5500000530") = "AHO126_B_RETIRADA ARMACAO SECUNDARIA-LM"
    dDescSrv("5500000532") = "AHO127_A_INST TIRANTE (VENTO, ANCORA)-LM"
    dDescSrv("5500000533") = "AHO127_B_INST TIRANTE (VENTO, ANCORA)-LM"
    dDescSrv("5500000534") = "AHO127_C_INST TIRANTE (VENTO, ANCORA)-LM"
    dDescSrv("5500000553") = "AHO135_A_INST PSTE MT CV CI PRE EXIST-LM"
    dDescSrv("5500000554") = "AHO135_B_INST PSTE MT CV CI PRE EXIST-LM"
    dDescSrv("5500000555") = "AHO135_C_INST PSTE MT CV CI PRE EXIST-LM"
    dDescSrv("5500000556") = "AHO136_A_INST PSTE BT CV CI PRE EXIST-LM"
    dDescSrv("5500000557") = "AHO136_B_INST PSTE BT CV CI PRE EXIST-LM"
    dDescSrv("5500000559") = "AHO137_A_RET POSTE MT-LM"
    dDescSrv("5500000560") = "AHO137_B_RET POSTE MT-LM"
    dDescSrv("5500000562") = "AHO138_A_RET POSTE BT-LM"
    dDescSrv("5500000563") = "AHO138_B_RET POSTE BT-LM"
    dDescSrv("5500000572") = "AHO201_A_INST CONDUTOR 16/50MM CB/AL-LM"
    dDescSrv("5500000573") = "AHO201_B_INST CONDUTOR 16/50MM CB/AL-LM"
    dDescSrv("5500000574") = "AHO201_C_INST CONDUTOR 16/50MM CB/AL-LM"
    dDescSrv("5500000575") = "AHO202_A_INST COND. 70/120 MM CB/AL-LM"
    dDescSrv("5500000581") = "AHO204_A_INST RD COMPACT(SPACE CAB)MT-LM"
    dDescSrv("5500000582") = "AHO204_B_INST RD COMPACT(SPACE CAB)MT-LM"
    dDescSrv("5500000590") = "AHO208_A_RET CON 16 50MM2 CU/EQUIV AL-LM"
    dDescSrv("5500000591") = "AHO208_B_RET CON 16 50MM2 CU/EQUIV AL-LM"
    dDescSrv("5500000593") = "AHO210_A_RET COND>120MM2 CU EQU AL-LM"
    dDescSrv("5500000596") = "AHO211_A_RETIRADA CONDUTOR AEREO BT-LM"
    dDescSrv("5500000597") = "AHO211_B_RETIRADA CONDUTOR AEREO BT-LM"
    dDescSrv("5500000617") = "AHO218_A_INST (FACA,FUSIVEIS,LAMINAS)-LM"
    dDescSrv("5500000618") = "AHO218_B_INST (FACA,FUSIVEIS,LAMINAS)-LM"
End Sub

Private Sub CarregarDescServicoP4()
    dDescSrv("5500000619") = "AHO218_C_INST (FACA,FUSIVEIS,LAMINAS)-LM"
    dDescSrv("5500000620") = "AHO219_A_INST ELEM DE CONT/TELECONT"
    dDescSrv("5500000621") = "AHO219_B_INST ELEM DE CONT/TELECONT"
    dDescSrv("5500000626") = "AHO221_A_RET (FACA, FUSIVEIS, LAMINA)-LM"
    dDescSrv("5500000629") = "AHO222_A_INSTALACAO DE PARA-RAIOS-LM"
    dDescSrv("5500000630") = "AHO222_B_INSTALACAO DE PARA-RAIOS-LM"
    dDescSrv("5500000632") = "AHO223_A_RETIRADA DE PARA-RAIOS-LM"
    dDescSrv("5500000633") = "AHO223_B_RETIRADA DE PARA-RAIOS-LM"
    dDescSrv("5500000686") = "AHO243_A_INST COND PRE REUN BT(PR BT)-LM"
    dDescSrv("5500000687") = "AHO243_B_INST COND PRE REUN BT(PR BT)-LM"
    dDescSrv("5500000688") = "AHO243_C_INST COND PRE REUN BT(PR BT)-LM"
    dDescSrv("5500000689") = "AHO244_A_RET RD COMPACTA SPACE CAB MT-LM"
    dDescSrv("5500000719") = "AHO337_A_INST PLACA EQUIP OU POSTE-LM"
    dDescSrv("5500000720") = "AHO337_B_INST PLACA EQUIP OU POSTE-LM"
    dDescSrv("5500000743") = "AHO721_A_AVISO DESLIGAMENTO S/ASSINATUR"
    dDescSrv("5500000744") = "AHO721_B_AVISO DESLIGAMENTO S/ASSINATUR"
    dDescSrv("5500000745") = "AHO721_C_AVISO DESLIGAMENTO S/ASSINATUR"
    dDescSrv("5500000746") = "AHO722_A_AVISO DESLIGAMENTO C/ASSINATUR"
    dDescSrv("5500000747") = "AHO722_B_AVISO DESLIGAMENTO C/ASSINATUR"
    dDescSrv("5500000749") = "AHO730_A_OPERACAO EQUIPAMENTO REDE MT-LM"
    dDescSrv("5500000750") = "AHO730_B_OPERACAO EQUIPAMENTO REDE MT-LM"
    dDescSrv("5500000751") = "AHO730_C_OPERACAO EQUIPAMENTO REDE MT-LM"
    dDescSrv("5500000760") = "AHO804_A_INST POST PASS MT 10-15MT LV-LM"
    dDescSrv("5500000761") = "AHO804_B_INST POST PASS MT 10-15MT LV-LM"
    dDescSrv("5500000763") = "AHO805_A_RET POST PASS MT 10-15MT LV-LM"
    dDescSrv("5500000764") = "AHO805_B_RET POST PASS MT 10-15MT LV-LM"
    dDescSrv("5500000766") = "AHO806_A_INST P MT 10 15MT/EST ENC LV-LM"
    dDescSrv("5500000767") = "AHO806_B_INST P MT 10 15MT/EST ENC LV-LM"
    dDescSrv("5500000769") = "AHO807_A_RET P MT 10A15MT/EST ENC LV-LV"
    dDescSrv("5500000770") = "AHO807_B_RET P MT 10A15MT/EST ENC LV-LV"
    dDescSrv("5500000784") = "AHO812_A_INST CRUZ SIMP/DUP DISP PASS-LM"
    dDescSrv("5500000787") = "AHO813_A_RET CRZ SIM/DUP DISP PASS LV-LM"
    dDescSrv("5500000788") = "AHO813_B_RET CRZ SIM/DUP DISP PASS LV-LM"
    dDescSrv("5500000790") = "AHO814_A_INST CRZ SIMP/DUP C ENCAB LV-LV"
    dDescSrv("5500000791") = "AHO814_B_INST CRZ SIMP/DUP C ENCAB LV-LV"
    dDescSrv("5500000796") = "AHO816_A_RET CRUZ SIMP/DUP ENCAB LV-LV"
    dDescSrv("5500000797") = "AHO816_B_RET CRUZ SIMP/DUP ENCAB LV-LV"
    dDescSrv("5500000817") = "AHO823_A_INST EQ PROT 1F RRAIOMT/BTLV-LV"
    dDescSrv("5500000818") = "AHO823_B_INST EQ PROT 1F RRAIOMT/BTLV-LV"
    dDescSrv("5500000820") = "AHO824_A_RET EQP PROT MON RRAIO C/LV-LV"
    dDescSrv("5500000821") = "AHO824_B_RET EQP PROT MON RRAIO C/LV-LV"
    dDescSrv("5500000832") = "AHO828_A_INST/SEC RELIG S CX CONT LV-LV"
    dDescSrv("5500000835") = "AHO829_A_RET SEC/RECO SF6/OLEO/VAZ LV-LV"
    dDescSrv("5500000863") = "AHO860_B_FURAR POCO ARTESIANO ATERRAM._I"
    dDescSrv("5500000872") = "AHO863_B_ESCAVAR E/OU MOVIMENTAR TERR_II"
    dDescSrv("5500000883") = "AHO867_B_LANCAMENTO MALHA DE ATERRAMENTO"
    dDescSrv("5500100607") = "CHO501_A_INST RAMA BT MONO/BI ATE 36MM2"
    dDescSrv("5500100628") = "CHO508_A_RET DE CABO AEREO BT"
    dDescSrv("5500200066") = "EBR227_B_INST (FACA,FUSIVEIS,LAMINAS)-LM"
    dDescSrv("5500200067") = "EBR227_C_INST (FACA,FUSIVEIS,LAMINAS)-LM"
    dDescSrv("5500200069") = "EBR228_B_RET (FACA, FUSIVEIS, LAMINA)-LM"
    dDescSrv("5500200070") = "EBR228_C_RET (FACA, FUSIVEIS, LAMINA)-LM"
    dDescSrv("5500200323") = "EHO301_A_OPERACAO EQUIPAMENTO REDE MT-E"
    dDescSrv("5500200324") = "EHO301_B_OPERACAO EQUIPAMENTO REDE MT-E"
    dDescSrv("5500200325") = "EHO301_C_OPERACAO EQUIPAMENTO REDE MT-E"
    dDescSrv("5500200632") = "FHO103_A_INST MELH ATERR SERVIC PROT-LM"
    dDescSrv("5500200633") = "FHO103_B_INST MELH ATERR SERVIC PROT-LM"
    dDescSrv("5500200635") = "FHO104_A_INST PLACA EQUIP OU POSTE-LM"
    dDescSrv("5500200636") = "FHO104_B_INST PLACA EQUIP OU POSTE-LM"
    dDescSrv("5500200637") = "FHO104_C_INST PLACA EQUIP OU POSTE-LM"
    dDescSrv("5500200644") = "FHO107_A_INST ESTRUT AP TRAFO 1 POSTE-LM"
    dDescSrv("5500200645") = "FHO107_B_INST ESTRUT AP TRAFO 1 POSTE-LM"
    dDescSrv("5500200646") = "FHO107_C_INST ESTRUT AP TRAFO 1 POSTE-LM"
    dDescSrv("5500200653") = "FHO110_A_INST TF SOBRE POSTE SO TRAFO-LM"
    dDescSrv("5500200654") = "FHO110_B_INST TF SOBRE POSTE SO TRAFO-LM"
    dDescSrv("5500200655") = "FHO110_C_INST TF SOBRE POSTE SO TRAFO-LM"
    dDescSrv("5500200656") = "FHO112_A_RET ESTRUT AP TRAFO 1 POSTE-LM"
    dDescSrv("5500200657") = "FHO112_B_RET ESTRUT AP TRAFO 1 POSTE-LM"
    dDescSrv("5500200665") = "FHO115_A_RET TF SOBRE POSTE SO TRAFO-LM"
    dDescSrv("5500200666") = "FHO115_B_RET TF SOBRE POSTE SO TRAFO-LM"
    dDescSrv("5500400467") = "TBR144_B_ABERTURA DE FAIXA DE SERVIDAO"
    dDescSrv("5500400468") = "TBR144_C_ABERTURA DE FAIXA DE SERVIDAO"
    dDescSrv("5500400502") = "THO108_A_PODA POR UNIDADE MT S RECOLH"
    dDescSrv("5500400503") = "THO108_B_PODA POR UNIDADE MT S RECOLH"
    dDescSrv("5500400512") = "THO112_A_PODA POR UNID. BT COM RECOLH"
End Sub

Private Sub CarregarDescServicoP5()
    dDescSrv("5500400513") = "THO112_B_PODA POR UNID. BT COM RECOLH"
    dDescSrv("5023000144") = "NIVELAR CRUZETA TIPO HT LD LV PDE"
    dDescSrv("5025300059") = "COMISSIONAMENTO DE OBRA"
    dDescSrv("5033000334") = "SERVICO LINHA VIVA SE/LT PICK UP HE 50%"
    dDescSrv("5034000036") = "SERVICO APOIO LINHA VIVA SE/LT PICK UP"
    dDescSrv("5044000016") = "EXECUTAR MAP AEROFOTOGRAMETR MDT"
    dDescSrv("5044300003") = "DESL/INST MAN EQUIP FUROS SOND LD LM PDE"
    dDescSrv("5044400001") = "REVISAO FINAL E PRE-COMISSIONAMENTO LD"
    dDescSrv("5045200001") = "INDENIZACAO PROPRIETARIOS"
    dDescSrv("5045200002") = "DESPESAS CARTORIAIS"
    dDescSrv("5045200003") = "SERVIC DE LIBERACAO DE FAIXA DE SERVIDAO"
    dDescSrv("5045500007") = "ADMINISTRACAO DE CANTEIRO"
    dDescSrv("5045800001") = "AUTORIZACOES/LIC/TAXAS AMBIENTAIS"
    dDescSrv("5045800061") = "EA-PCA/RCA/EIV LD 01 A 03KM"
    dDescSrv("5045800067") = "EA-INVENT N 01 A 03KM"
    dDescSrv("5050000004") = "CONSTRUIR CASA DE COMANDO EM SE"
    dDescSrv("5050000025") = "CONSTRUIR ESCADA EM CONCRETO"
    dDescSrv("5050000026") = "EXECUCAO DE PORTICO DE IDENTIFICACAO SE"
    dDescSrv("5050100010") = "BOTA FORA DE SOBRA DE MATERIAL"
    dDescSrv("5050100011") = "EXECUTAR ATERRO MECANIZADO (JAZIDA)"
    dDescSrv("5050100013") = "EXECUTAR CORTE MECANIZADO (1A CATEGORIA)"
    dDescSrv("5050100015") = "EXECUTAR DESTOCAMENTO DE ARVORE"
    dDescSrv("5050100016") = "EXECUTAR LIMPEZA MANUAL DE TERRENO"
    dDescSrv("5050100017") = "EXECUTAR LIMPEZA MECANIZADA DE TERRENO"
    dDescSrv("5050100018") = "EXECUTAR LOCACAO DE OBRA DE SUBESTACAO"
    dDescSrv("5050100019") = "EXECUTAR REVESTIMENTO EM PO DE BRITA"
    dDescSrv("5050100020") = "REVESTIMENTO DE TALUDE CONCRETO"
    dDescSrv("5050200005") = "CONSTRUCAO RAMPA DE ACESSO"
    dDescSrv("5050200019") = "APLICAR TINTA ACRILICA EXTERNA"
    dDescSrv("5050200022") = "CONSTRUIR CALCADA OU PASSEIO"
    dDescSrv("5050200024") = "CONSTRUIR MEIO-FIO"
    dDescSrv("5050200025") = "CONSTRUIR MURO"
    dDescSrv("5050200027") = "EXECUTAR PINTURA DE ESTRUTURA EM SE"
    dDescSrv("5050200029") = "EXECUTAR PISTA EM BLOCOS DE CONCRETO"
    dDescSrv("5050200031") = "INSTALAR E FORNECER CERCA EM MOUROES"
    dDescSrv("5050200033") = "INSTALAR E FORNECER CONCERTINA"
    dDescSrv("5050200036") = "INSTALAR E FORNECER PORTAO PARA VEICULOS"
    dDescSrv("5050200041") = "LIGACAO PROV DE AGUA E INST SANITARIA SE"
    dDescSrv("5050300051") = "EXECUTAR BASE DE TR DE FORCA ATE 40MVA"
    dDescSrv("5050300057") = "EXECUTAR BASE PARA DISJUNTOR 69 KV"
    dDescSrv("5050300058") = "EXECUTAR BASE PARA DISJUNTOR ATE 34,5 KV"
    dDescSrv("5050300059") = "EXECUTAR BASE PARA ESTRURURA METALICA"
    dDescSrv("5050300060") = "EXECUTAR BASE PARA IP PR TC TP 138 KV"
    dDescSrv("5050300067") = "EXECUTAR FUNDACAO DE POSTE ATE 9M"
    dDescSrv("5050300068") = "EXECUTAR FUNDACAO DE POSTE DE 9M A 14M"
    dDescSrv("5050300073") = "MONTAGEM ESTRUTURA PRE-MOLDADA TF SERV"
    dDescSrv("5050400022") = "CONSTRUIR CAIXA DE DRENAGEM"
    dDescSrv("5050400025") = "CONSTRUIR VALETA PARA ESCOAMENTO DE AGUA"
    dDescSrv("5050400026") = "EXECUTAR CAIXA SEPARADORA DE AGUA E OLEO"
    dDescSrv("5050400027") = "EXECUTAR VALETA E BACIA DISSIPACAO"
    dDescSrv("5050400028") = "NIVELAMENTO DE BRITA NO PATIO SE PDE"
    dDescSrv("5050400030") = "INSTALAR LINHA DE DRENO SECO EM SE"
    dDescSrv("5050400032") = "INSTALAR TUBO DE CONCRETO SIMPLES (D=30C"
    dDescSrv("5050400033") = "INSTALAR TUBO DE CONCRETO SIMPLES (D=80C"
    dDescSrv("5050400035") = "INSTALAR TUBO DE PVC DE 150MM"
    dDescSrv("5050400036") = "INSTALAR TUBO DE PVC DE 200MM"
    dDescSrv("5050400037") = "REALOCAR E ESPALHAR BRITA NO PATIO DE SE"
    dDescSrv("5050400048") = "EXECUTAR CX PASSAG/BOCA LOBO 1000X1000MM"
    dDescSrv("5051000041") = "INSTALAR EST SUP MET IP PR TC TP 69 KV"
    dDescSrv("5051000042") = "INSTALAR POSTE DE CONCRETO 11M ATE 18M"
    dDescSrv("5052000058") = "CONECTAR TR FORCA  69/34,5KV OU 13,8KV"
    dDescSrv("5052000098") = "SERV INST/CONFEC PLACA OPERAC EQUIP SE"
    dDescSrv("5052000124") = "INST CHAV SECC SECO 1F 13,8KV SE LV PDE"
    dDescSrv("5052000129") = "INST CONEXAO EM EQUIPAMENTO MT SE LV PDE"
    dDescSrv("5052100007") = "CONSTRUIR CAIXA DE INSPECAO DE MALHA SE"
    dDescSrv("5052100008") = "EXECT ATERRAM EQUIPAMENT LD LM PDE"
    dDescSrv("5052100010") = "EXECUTAR MALHA DE ATERRAMENTO DE SE"
    dDescSrv("5052100011") = "SERVICO EXECUTAR SOLD EXOTERM EM SE"
    dDescSrv("5052200014") = "INSTALAR BAR TRIF RIGIDO ATE 34,5KV"
    dDescSrv("5052300000") = "EXECUTAR CONEXOES PARA BARRAMENTO"
    dDescSrv("5052300002") = "EXECUTAR MUFLA ATE 34,5KV"
    dDescSrv("5052300003") = "INSTALAR BAR TRIF DUPLO FLEX 69 OU 138KV"
    dDescSrv("5052300004") = "INSTALAR BAR TRIF DUPLO FLEX ATE 34,5KV"
    dDescSrv("5052300005") = "INSTALAR BAR TRIF SIMP FLEX 69 OU 138KV"
    dDescSrv("5052400001") = "LANCAR CABOS PARA-RAIOS"
End Sub

Private Sub CarregarDescServicoP6()
    dDescSrv("5052500000") = "EXECUTAR BASE CHAMINE CABOS DE FORCA"
    dDescSrv("5052500001") = "EXECUTAR CAIXA DE PASSAGEM 400MM X 400MM"
    dDescSrv("5052500002") = "EXECUTAR CAIXA DE PASSAGEM 600MM X 600MM"
    dDescSrv("5052500003") = "EXECUTAR CAIXA DE PASSAGEM 1M X 1M"
    dDescSrv("5052500004") = "EXECUTAR CANALETA 400MM"
    dDescSrv("5052500005") = "EXECUTAR CANALETA 600MM"
    dDescSrv("5052500007") = "EXECUTAR DUTOS ENVELOPADOS 12X100MM"
    dDescSrv("5052500008") = "EXECUTAR DUTOS ENVELOPADOS 1X200MM"
    dDescSrv("5052500011") = "INSTALAR CAIXA DE DERIVACAO PARA TC/TP"
    dDescSrv("5052500015") = "INSTALAR ESTRUTURA DE SUPORTE PARA MUFLA"
    dDescSrv("5053000011") = "EXECUTAR BASE PROJETOR DE ILUMINACAO SE"
    dDescSrv("5053000012") = "EXECUTAR EST SUP PARA ILUMINACAO"
    dDescSrv("5053000013") = "INSTALAR DO CONJUNTO DE TOMADA EM SE"
    dDescSrv("5053000014") = "INSTALAR LUMINARIA ARANDELA EM SE"
    dDescSrv("5053000016") = "LANCAR CABOS PARA ILUMINACAO EM SE"
    dDescSrv("5053000017") = "LIGACAO PROVISORIA LUZ E FORCA SE"
    dDescSrv("5054000008") = "EXECUTAR INTERLIGACAO DE DISJUNTOR"
    dDescSrv("5054000009") = "EXECUTAR INTERLIGACAO DE TC"
    dDescSrv("5054000010") = "EXECUTAR INTERLIGACAO DE TP"
    dDescSrv("5054000011") = "EXECUTAR INTERLIGACAO DE TSA"
    dDescSrv("5054000015") = "EXECUTAR INTERLIGACAO TR C/ AUTOMACAO"
    dDescSrv("5054000045") = "EXEC RETROFIT DE TRANSFORMADOR POTENCIAL"
    dDescSrv("5054000050") = "INTERLIGAR RELE MONTADO EM PAINEL"
    dDescSrv("5054100005") = "LANCAR CABO DE FIBRA OPTICA"
    dDescSrv("5054200008") = "EXECUTAR INSTAL.E INTERL.PAINEL PROTECAO"
    dDescSrv("5055000004") = "EXECUTAR BASE PARA TSA ATE 34,5 KV"
    dDescSrv("5056000005") = "TAXAS E EMOLUMENTOS"
    dDescSrv("5056000038") = "CONTRUIR BARRACAO DE OBRAS"
    dDescSrv("5056000044") = "DESMOBILIZAR OBRA DE GRANDE PORTE"
    dDescSrv("5056000045") = "DESMOBILIZAR OBRA DE MEDIO PORTE"
    dDescSrv("5056000046") = "DESMOBILIZAR OBRA DE PEQUENO PORTE"
    dDescSrv("5056000048") = "EMITIR ART E DOCUMENTACAO DE AMPLI SES"
    dDescSrv("5056000053") = "INSTALAR PLACA DE IDENTIFICACAO DE OBRAS"
    dDescSrv("5056000054") = "SERVICO INSTALAR TAPUME P/ ISOLAM OBRAS"
    dDescSrv("5056000059") = "MOBILIZAR EQUIPE E MAQ NA CAPITAL"
    dDescSrv("5056000060") = "MOB/DES EQUIP OBRA GRAND PORTE LD LM PDE"
    dDescSrv("5056000061") = "MOB/DES EQUIP OBRA MEDIO PORTE LD LM PDE"
    dDescSrv("5056000062") = "MOB/DESM EQUIP OBRA MEDIO PEQ LD LM PDE"
    dDescSrv("5056000070") = "FRET EQ/MAT 201<DMT<400KM LD LM PDE"
    dDescSrv("5056000071") = "SERV FRETE EQ/MAT DMT<50KM LD LM PDE"
    dDescSrv("5056000075") = "MOBILIZACAO E DESMOBILIZACAO DE PESSOAL"
    dDescSrv("5056000076") = "LIMPEZA GERAL DA OBRA DE AUTOMACAO"
    dDescSrv("5056000078") = "FORN E INST DE DEPOSIT CONTAINER NA OBRA"
    dDescSrv("5056000085") = "DESMOBILIZAR 0 A 200KM EQP 5H TOPOG SE"
    dDescSrv("5056000090") = "DESMOBILIZAR ACIMA 600KM EQP 5H SOND SE"
    dDescSrv("5056000091") = "DESMOBILIZAR ACIMA 600KM EQP 5H TOPOG SE"
    dDescSrv("5056000096") = "MOBILIZAR 0 A 200KM EQP 5H TOPOG SE"
    dDescSrv("5056000101") = "MOBILIZAR ACIMA 600KM EQP 5H TOPOG SE"
    dDescSrv("5056000107") = "DESCARREGAR ARTEFATO DE CONCRETO NA OBRA"
    dDescSrv("5056000110") = "INST JUMPER EM CONDUTOR MT SE LV PDE"
    dDescSrv("5056000116") = "TURMA LV 5H PICK-UP EM DISP LD/SE PDE"
    dDescSrv("5056100001") = "ELABORACAO DE AS BUILT DO PROJETO DA SE"
    dDescSrv("5056100004") = "DISPONIBILIZAR TECNICO COMISSIONAMENTO"
    dDescSrv("5056100008") = "GUARDA FISICA DIURNA/NOTURNA DE OBRA SE"
    dDescSrv("5056300000") = "EXECUTAR ABRIGO PARA EXTINTOR DE SE"
    dDescSrv("5056300003") = "INSTALAR E FORNECER SINALIZACOES"
    dDescSrv("5056300004") = "INSTALAR EXTINTOR DE PO QUIMICO 12KG"
    dDescSrv("5056300008") = "INSTALAR EXTINTOR DE PO QUIMICO 50KG"
    dDescSrv("5056400006") = "DESATIVAR INTERLIGACAO DISJ 69KV 138KV"
    dDescSrv("5056400012") = "DESATIVAR INTERLIGACAO TP/TC 69KV 138KV"
    dDescSrv("5056400078") = "RETIRAR POSTE DE 11M ATE 18M"
    dDescSrv("5056500001") = "DEMOLICAO MEIO-FIO"
    dDescSrv("5056500002") = "DEMOLIR BASE DE CONCRETO ARMADO"
    dDescSrv("5056500003") = "DEMOLIR BASE DE CONCRETO SIMPLES"
    dDescSrv("5056500005") = "DEMOLIR CAIXA DE PASSAGEM"
    dDescSrv("5056500009") = "RETIRAR PISTA EM BLOCOS SEXTAVADO"
    dDescSrv("5057000003") = "ELABORAR RELATORIO DE OBRA DE AUTOMACAO"
    dDescSrv("5057100004") = "EXECUTAR PROJETO AS BUILT"
    dDescSrv("5057100008") = "ELABORACAO PROJ TECNICO DRENAGEM SE"
    dDescSrv("5057100013") = "EXECUTAR ENSAIO DE INFILTRACAO SE"
    dDescSrv("5057100015") = "EXECUTAR IMPLANT MARCO DE CONCR GEORR SE"
    dDescSrv("5057100016") = "EXECUTAR PROJ DE TERRAPLANAGEM SE"
    dDescSrv("5057100021") = "LEV SEMI-CADAST E PLANIALTIM SE"
    dDescSrv("5057100023") = "PROJETO ENTRADA DE LINHA ATE 34,5KV"
    dDescSrv("5057100024") = "PROJETO ENTRADA DE LINHA ACIMA DE 34,5KV"
End Sub

Private Sub CarregarDescServicoP7()
    dDescSrv("5057100025") = "ELABORACAO DE PROJETO DE AUTOM SISTEMAS"
    dDescSrv("5057100027") = "CONEXAO BANCO DE CAPACITOR ATE 34,5 KV"
    dDescSrv("5057100029") = "INTERLIGACAO DE BARRAS ATE 34,5KV"
    dDescSrv("5057100030") = "ELABORACAO DE PROJETO DE AUTOM DE ARQUIT"
    dDescSrv("5057100037") = "DIAGRAMA UNIFILAR DE PROTECAO E MEDICAO"
    dDescSrv("5057100038") = "DIAGRAMA INTERLIG PAINEL SERVIC AUX CC"
    dDescSrv("5057100040") = "ELABORACAO DE PROJETO DE AUTOM DIAGRAMA"
    dDescSrv("5057100041") = "PROJETO TRANSFORMADOR FORCA C/COMUTADOR"
    dDescSrv("5057100042") = "ELABORACAO DE PROJETO DE AUTOMACAO P/INT"
    dDescSrv("5057100043") = "PROJETO CONEXAO DE TRANSFOR ATE 34,5KV"
    dDescSrv("5071000061") = "PROJETO CIVIL, ELETROMEC. E AUTOM BIM SE"
    dDescSrv("5071000062") = "MAQUETE ELETRONICA VISTAS 3D SE"
    dDescSrv("5071000063") = "ANIMACAO TOUR 360 PELA SE VIDEO"
    dDescSrv("5100000005") = "SERVICO DE TRANSMISSAO DE DADOS"
    dDescSrv("5120000003") = "CONSTRUCAO DE IMOVEL"
    dDescSrv("5210100004") = "ANUNCIOS E PUBLICACOES LEGAIS"
    dDescSrv("5261000159") = "EA - PRO PLA GERDE RES SOL E EFLU LIQ"
    dDescSrv("5281000003") = "PERF LASER AEROTRANSP E MAPEAM AEROFOTOG"
    dDescSrv("5290000003") = "SERVICO DE ENGENHARIA DE PROJETOS"
    dDescSrv("5400099341") = "LICENCA 1 ANO FORTIGATE 60F"
    dDescSrv("5400099342") = "LICENCA 1 ANO FORTISWITCH 124E"
    dDescSrv("5051000040") = "INSTALAR ANEL DE CONCRETO"
    dDescSrv("5051000043") = "INSTALAR POSTE DE CONCRETO 5M ATE 10M"
    dDescSrv("5051000046") = "INSTALAR SUPORTE JABAQUARA"
    dDescSrv("5051000047") = "INSTALAR VIGA CONC ATE 6 METROS"
    dDescSrv("5052000066") = "INSTALAR CHAVE FACA 13,8KV E 34,5KV"
    dDescSrv("5052000067") = "INSTALAR CHAVE FUSIVEL ATE 34,5KV"
    dDescSrv("5052000068") = "INSTALAR CHAVE SEC 3F MH 13,8KV E 34,5KV"
    dDescSrv("5052000076") = "INSTALAR CHAVE TANDEM 13,8KV E 34,5KV"
    dDescSrv("5052000079") = "INSTALAR DISJUNTOR 13,8 KV OU 34,5 KV"
    dDescSrv("5052000082") = "INSTALAR PARA-RAIOS ATE 33 KV"
    dDescSrv("5052000088") = "INSTALAR TC 13,8KV OU 34,5KV"
    dDescSrv("5052000091") = "INSTALAR TP 13,8KV OU 34,5KV"
    dDescSrv("5052000094") = "INSTALAR TSA ATE 34,5KV"
    dDescSrv("5052200015") = "INSTALAR IP ATE 34,5KV"
    dDescSrv("5052300009") = "SERVICO LANCAR CABO ISOL ATE 34,5KV"
    dDescSrv("5052400000") = "INSTALAR HASTES PARA-RAIOS"
    dDescSrv("5052500012") = "INSTALAR ELETRODUTO ACIMA DE 2'"
    dDescSrv("5052500014") = "INSTALAR ELETRODUTO DE 1 1/2' ATE 2'"
    dDescSrv("5053000015") = "INSTALAR REFLETOR DE ILUMINACAO EM SE"
    dDescSrv("5054000023") = "LANCAR CABOS ATE 8 CONDUTORES"
    dDescSrv("5054000029") = "ADEQUACAO FIACAO RELE PAINEL DE PROTECAO"
    dDescSrv("5054000032") = "EXEC INSTAL ELETR CX CONCENTRACAO DE TC"
    dDescSrv("5054000033") = "EXEC INSTAL ELETR CX CONCENTRACAO DE TP"
    dDescSrv("5054000036") = "FIXACAO E INTERLIGACAO MEDIDOR EM PAINEL"
    dDescSrv("5054100006") = "LANCAMENTO DE CORDAO DE FIBRA OPTICA"
    ' Base de servicos (planilha base_servi_os.xlsx)
    dDescSrv("5028000005") = "FABRICACAO E INSTALACAO PLACA DA OBRA"
    dDescSrv("5040000002") = "SINALIZACAO DE ESTRUTURAS COM PINTURA"
    dDescSrv("5040000006") = "DISTRIBUICAO DE ESTRUTURA LD LM PDE"
    dDescSrv("5040000008") = "LOCACAO DE ESTRUTURAS LD LM PDE"
    dDescSrv("5040100002") = "FUNDACAO TIPO C1"
    dDescSrv("5040100004") = "FUNDACAO TIPO D2"
    dDescSrv("5040100005") = "FUNDACAO TIPO D3"
    dDescSrv("5040100046") = "FUNDACAO TIPO D4"
    dDescSrv("5040200011") = "IMPL ESTR CONC ACIMA DE 26M LD LM PDE"
    dDescSrv("5040300014") = "MONT ESTRUT CONC CIR SIMPLES SUSP URB LD"
    dDescSrv("5040300016") = "MONT ESTRUT CONC CIRC SIMPL ANCOR URB LD"
    dDescSrv("5042000005") = "LANC/NIV/GRAM CABOS PARA-RAIO LD LM PDE"
    dDescSrv("5042000020") = "LANC COND CIRC SIMP CAL 927,2 LD LM PDE"
    dDescSrv("5042100005") = "GRAMP COND CIRC SIMP CAL 927,2 LD LM PDE"
    dDescSrv("5042300008") = "NIV COND CIRC DUP CAL 927,2 LD LM PDE"
    dDescSrv("5042400003") = "EXEC DE ATERR RURAL - FASE I LD LM PDE"
    dDescSrv("5045400001") = "CONSERT DE CALCADA COM CIMENTO LD LM PDE"
    dDescSrv("5045800006") = "EXECUCAO DOS PROGRAMAS AMBIENTAIS"
    dDescSrv("5045900007") = "INSTALACAO DE DEFENSA LT"
    dDescSrv("5050200013") = "PINTURA DE BASE OU POSTE C/ CAL SE PDE"
    dDescSrv("5054100001") = "INSTALACAO E INTERLIG EQUIP COMUNICACAO"
    dDescSrv("5054200013") = "EXECUTAR INSTALACAO PAINEL MEDICAO OPER"
    ' Base de servicos (planilha classificar_servi_os.xlsx)
    dDescSrv("5025300082") = "EXEC SERVICO INTEGRACAO AUTOMACAO"
    dDescSrv("5025500008") = "ENCERRAMENTO TECNICO DA OBRA LM"
    dDescSrv("5041000108") = "INST ISOLAD SUSP SP 69KV LD LV PDE"
    dDescSrv("5041000109") = "RET ISOLAD SUSP SP 69KV LD LV PDE"
    dDescSrv("5050300014") = "CONSTRUCAO BASE PARA TRAFO 34,5/13,8 KV"
End Sub

Private Sub CarregarDescServicoP8()
    dDescSrv("5050300021") = "CONSTRUCAO BASE PARA DISJUNTOR 15KV"
    dDescSrv("5050300028") = "CONSTRUCAO BASE PARA DISJUNTOR 72,5KV"
    dDescSrv("5050300108") = "DEMOLIR BASE PARA DISJ. ATE 34,5KV"
    dDescSrv("5052000133") = "RET CONEX EM EQUIP 69KV SE LV PDE"
    dDescSrv("5054000022") = "INSTALAR RELE EM PAINEL NOVO"
    dDescSrv("5056500013") = "MONTAGEM CAIXA EQUIP. PATIO"
    dDescSrv("5091100037") = "SERVICOS DE DESENVOLVIMENTO DE SOFTWARE"
    dDescSrv("5120000002") = "INSTALACAO DE AR CONDICIONADO"
    dDescSrv("5300900040") = "CONST MURO COM DE BLOCO ESTRUT 14X19X39"
    dDescSrv("5400099371") = "MONTAGEM DE CRUZETA DUPLA"
    dDescSrv("5500000003") = "138 KV INST. CADEIA ISOLADORES SUS LDAT"
    dDescSrv("5500000945") = "BACIA COLETORA"
    dDescSrv("5500000965") = "BASE DE CONCRETO PARA DISJUNTOR 34,5KV -"
    dDescSrv("5500100769") = "COMISSIONAMENTO CH. SECC. TRIPOLAR (15"
    dDescSrv("5500100776") = "COMISSIONAMENTO DISJUNTOR DE 15KV (VACUO"
    dDescSrv("5500100777") = "COMISSIONAMENTO DISJUNTOR DE 345 69 E"
    dDescSrv("5500100783") = "COMISSIONAMENTO PR'S 15,34,5,69 E 138KV"
    dDescSrv("5500100788") = "COMISSIONAMENTO RELIGADORES 15/34,5KV"
    dDescSrv("5500100790") = "COMISSIONAMENTO TC'S (1534569 E 138 K"
    dDescSrv("5500100801") = "COMISSIONAMENTO/ATIVACAO DE EQUIPAMENT0"
    dDescSrv("5500100866") = "DEMOLICAO E RETIRO DE CONCRETO"
    dDescSrv("5500100914") = "DESMONTAGEM A MAO DE ESTRUTURAS METALICA"
    dDescSrv("5500100920") = "DESMONTAGEM DE BANCO REGULADOR DE TENSAO"
    dDescSrv("5500100929") = "DESMONTAGEM DE DISJUNTOR ENTRE 10 E 23 K"
    dDescSrv("5500100931") = "DESMONTAGEM DE DISJUNTOR ENTRE 345 E 69"
    dDescSrv("5500100977") = "DMONT. RELE COM MULTIFUNCAO"
    dDescSrv("5500200814") = "FORNECIMENTOS DE MATERIAIS MENORES ELET"
    dDescSrv("5500200957") = "INSTALACAO CABO DE FIBRA OPTICA"
    dDescSrv("5500200969") = "INSTALACAO DE CANTEIRO OBRAS SIMPLES"
    dDescSrv("5500200970") = "INSTALACAO DE CHAVE CONJUGADA PARA BCO."
    dDescSrv("5500200971") = "INSTALACAO DE CHAVE FUSIVEL DE 10 A 23KV"
    dDescSrv("5500200973") = "INSTALACAO DE CHAVE MONOPOLAR  10 A 23KV"
    dDescSrv("5500200980") = "INSTALACAO DE ELETRODUTOS E CONDUTORES A"
    dDescSrv("5500300005") = "INSTALACAO DE POSTES PARA SUPORTE DE EQU"
    dDescSrv("5500300006") = "INSTALACAO DE PROTECAO ANTI-FAUNA"
    dDescSrv("5500300053") = "INSTALACAO/ DESINSTALACAO DE SE MOVEL AT"
    dDescSrv("5500300079") = "LANCAMENTO DE CABO ISOLADO ATE 630 MM² D"
    dDescSrv("5500300081") = "LANCAMENTO DE CABO NU ATE 120 MM*2"
    dDescSrv("5500300088") = "LANCAMENTO E CONEXAO DE CABOS DE CONTROL"
    dDescSrv("5500300090") = "LANCAMENTO EM VALA DE CABO ISOLADO ATE 6"
    dDescSrv("5500300137") = "LOCACAO CAMINHAO MUNCK 20 TON. TARA 10 T"
    dDescSrv("5500300169") = "MANOBRA EM SE POR MANOBRA"
    dDescSrv("5500300171") = "MANOBRA EN SE_LV"
    dDescSrv("5500300358") = "MONTAGEM A MAO DE ESTRUTURAS METALICAS"
    dDescSrv("5500300365") = "MONTAGEM DE BANCO DE REGULADOR AEREO/SUB"
    dDescSrv("5500300375") = "MONTAGEM DE DISJUNTOR ENTRE 10 E 23 KV ("
    dDescSrv("5500300378") = "MONTAGEM DE DISJUNTOR ENTRE 345 E 69 KV"
    dDescSrv("5500300391") = "MONTAGEM DE PARA-RAIOS ENTRE 10 E 23 KV"
    dDescSrv("5500300399") = "MONTAGEM DE RELIGADOR ATE 34,5KV"
    dDescSrv("5500300402") = "MONTAGEM DE TRANSFORMADORES DE MEDIDA EN"
    dDescSrv("5500300415") = "MONTAGEM/CABEAMENTO DO RELE MULTIFUNCAO"
    dDescSrv("5500300419") = "MUFLA EM CABO ISOLADO ATE 630 MM² DE SEC"
    dDescSrv("5500300558") = "RETIRADA DE CADEIA DE SUSPENSAO"
    dDescSrv("5500300563") = "RETIRADA DE CHAVE MONOPOLAR 10 A 23KV"
    dDescSrv("5500300567") = "RETIRADA DE DETRITOS"
    dDescSrv("5500300584") = "RETIRADA MODEM - COMUNIC. TELECONTR."
    dDescSrv("5500300591") = "REVISAO E/OU SUBSTITUICAO DE CONEXOES DE"
    dDescSrv("5500300748") = "SERV ADM ENCERRAM ORDENS ODI - MANUT"
    dDescSrv("5500300801") = "SERV PROJ ELABORACAO PROJ ENGENHARIA"
    dDescSrv("5500300852") = "SERVICO INSTAL SIST COMUNIC VIA SATELIT"
    dDescSrv("5500400487") = "TESTES PARA CABOS ISOLADOS"
    dDescSrv("5500400583") = "TRANSP. DE EQUIPAM. E/OU MATERIAIS - SE"
    dDescSrv("5500400598") = "TRANSPORTE DE EQUIPAMENTOS MENORES"
    dDescSrv("5500400601") = "TRANSPORTE DE EQUIPAMENTOS/MATERIAIS COM"
    dDescSrv("5500500620") = "13KV DMONT. CHAVE SECCIONADORA MONOFASIC"
    dDescSrv("5500500624") = "13KV DMONT. RELIGADOR TRIPOLAR"
    dDescSrv("5500500626") = "13KV DMONT. TRANSFORMADOR DE CORRENTE"
    dDescSrv("5500500631") = "13KV LANC. CABO AL, DE 266,8 A 477MCM"
    dDescSrv("5500500632") = "13KV LANC. CABO ISOLADO, DE 101 A 300MM"
    dDescSrv("5500500633") = "13KV MONT. DISJUNTOR TRIPOLAR"
    dDescSrv("5500500634") = "13KV MONT. SUPORTE METALICO PARA TC"
    dDescSrv("5500500635") = "13KV MONT. TRANSFORMADOR DE CORRENTE"
    dDescSrv("5500500666") = "69KV LANC. CABO AL, DE 266,8 A 477MCM"
    dDescSrv("5500500691") = "DMONT. CABO DE ACO"
    dDescSrv("5500500918") = "US - INSTALACAO DE CADEIA DE ISOLADORES"
End Sub

Private Sub CarregarDescServicoP9()
    dDescSrv("5500500919") = "US - RETIRADA DE CADEIA DE ISOLADORES DE"
    dDescSrv("5500500921") = "LOCACAO MAT/EQ CACAMBA ESTAC ATE 10M3'"
    dDescSrv("5500500939") = "69 KV INST. CADEIA ISOLADORES SUS LDAT"
    dDescSrv("5500500940") = "69 KV INST. CADEIA ISOLADORES ANC LDAT"
    dDescSrv("5500900688") = "COMISSIONAMENTO/ATIVACAO DE EQUIP -AT"
End Sub

Private Function DescServico(ByVal cod As String) As String
    DescServico = ""
    If dDescSrv Is Nothing Then Exit Function
    Dim k As String: k = NormCod(cod)
    If dDescSrv.Exists(k) Then DescServico = dDescSrv(k)
End Function

Private Function NormCod(v As Variant) As String
    If IsNumeric(v) Then
        Dim d As Double: d = CDbl(v)
        If d = Int(d) Then NormCod = Format$(d, "0") Else NormCod = CStr(d)
    Else
        NormCod = Trim$(CStr(v))
    End If
End Function

' Devolve parte do catlogo: 0=FAMILIA, 1=CLS1, 2=CLS2, 3=CLS3
Private Function CatInfo(codMat As Variant, idx As Long) As String
    CatInfo = ""
    If dCatMat Is Nothing Then Exit Function
    Dim cod As String: cod = NormCod(codMat)
    If dCatMat.Exists(cod) Then
        Dim p() As String: p = Split(dCatMat(cod), "|")
        If idx <= UBound(p) Then CatInfo = Trim$(p(idx))
    End If
End Function
'==============================================================================
'  CATLOGO DE SERVIOS (CLS1, CLS2, CLS3, TIPO_APLIC, SEGMENTO)
'  Nas linhas de servio (CA), o cdigo do servio est na coluna "Material".
'==============================================================================
Private Sub CarregarCatalogoServicos()
    Set dCatSrv = CreateObject("Scripting.Dictionary")

    Dim caminho As String
    caminho = CaminhoCatalogo("CAT_SERVICOS", _
        "%USERPROFILE%\Downloads\SERVICOS_ATUAIS_3.xlsx;%USERPROFILE%\Downloads\SERVICOS_ATUAIS (3).xlsx;" & _
        "%USERPROFILE%\Downloads\SERVICOS_ATUAIS_2.xlsx;%USERPROFILE%\Downloads\SERVICOS_ATUAIS (2).xlsx;" & _
        "%USERPROFILE%\Downloads\SERVICOS_ATUAIS.xlsx")
    If caminho = "" Then
        Dim f As Variant
        f = Application.GetOpenFilename( _
            "Excel (*.xls*),*.xls*", , _
            "Selecione o catalogo de SERVICOS (SERVICOS_ATUAIS). Cancele para pular.")
        If f = False Then Exit Sub
        caminho = CStr(f)
    End If

    On Error GoTo SemCat
    Dim wb As Workbook, ws As Worksheet, arr As Variant
    Set wb = Workbooks.Open(caminho, ReadOnly:=True, UpdateLinks:=0)
    Set ws = wb.Worksheets(1)

    Dim cCod As Long, c1 As Long, c2 As Long, c3 As Long, cTA As Long, cSeg As Long, cTxt As Long, cUN As Long
    cCod = ColLike(ws, Array("COD SERVICO", "COD_SERVICO", "SERVICO"))
    c1 = ColLike(ws, Array("CLS1"))
    c2 = ColLike(ws, Array("CLS2"))
    c3 = ColLike(ws, Array("CLS3"))
    cTA = ColLike(ws, Array("TIPO APLICACAO", "TIPO_APLICACAO", "TIPO APLIC"))
    cSeg = ColLike(ws, Array("SEGMENTO"))
    cTxt = ColLike(ws, Array("TEXTO BREVE", "TEXTO_BREVE", "DENOMINACAO", "DESCRICAO"))
    cUN = ColLike(ws, Array("UN BASICA", "UNID.MEDIDA", "UNIDADE", "UN"))
    If cCod = 0 Then wb.Close SaveChanges:=False: Exit Sub
    If dDescSrv Is Nothing Then Set dDescSrv = CreateObject("Scripting.Dictionary")

    Dim ult As Long
    ult = ws.Cells(ws.Rows.Count, cCod).End(xlUp).Row
    arr = ws.Range(ws.Cells(2, 1), ws.Cells(ult, ws.UsedRange.Columns.Count)).Value

    Dim i As Long, cod As String, desc As String
    For i = 1 To UBound(arr, 1)
        cod = NormCod(arr(i, cCod))
        If cod <> "" And Not dCatSrv.Exists(cod) Then
            dCatSrv(cod) = _
                TextoMatriz(arr, i, c1) & "|" & _
                TextoMatriz(arr, i, c2) & "|" & _
                TextoMatriz(arr, i, c3) & "|" & _
                TextoMatriz(arr, i, cTA) & "|" & _
                TextoMatriz(arr, i, cSeg) & "|" & _
                TextoMatriz(arr, i, cUN)
        End If
        ' Descricao vinda do arquivo (embutido em dDescSrv tem prioridade)
        If cod <> "" And cTxt > 0 And Not dDescSrv.Exists(cod) Then
            desc = Trim$(TextoMatriz(arr, i, cTxt))
            If desc <> "" Then dDescSrv(cod) = desc
        End If
    Next i
    wb.Close SaveChanges:=False
    Exit Sub
SemCat:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

' Overrides curados de reclassificacao CLS1/CLS2/CLS3 por COD_SERVICO
' (RECLASSIFICA_SERVI_OS.xlsx). Roda DEPOIS do catalogo externo e SOBREPOE
' a classificacao, preservando TIPO_APLIC/SEGMENTO/UND ja carregados.
Private Sub CarregarReclassificacaoServicos()
    ReclasSrv "5023000185", "SERVICO", "TRAFO", "MAO DE OBRA"
    ReclasSrv "5023000184", "SERVICO", "TRAFO", "MAO DE OBRA"
    ReclasSrv "5023000186", "SERVICO", "TRAFO", "MAO DE OBRA"
    ReclasSrv "5024000086", "SERVICO", "CONEXAO", "MAO DE OBRA"
    ReclasSrv "5024000085", "SERVICO", "CONEXAO", "MAO DE OBRA"
    ReclasSrv "5023000187", "SERVICO", "REGULADOR", "MAO DE OBRA"
    ReclasSrv "5025300385", "SERVICO", "ESPACADOR", "MAO DE OBRA"
    ReclasSrv "5400099811", "SERVICO", "COND ISOLADO", "MAO DE OBRA"
    ReclasSrv "5400099848", "SERVICO", "CONEXAO", "MAO DE OBRA"
    ReclasSrv "5400099851", "SERVICO", "CONEXAO", "MAO DE OBRA"
    ReclasSrv "5301800152", "SERVICO", "ALUGUEL GERADOR", "RISCO"
End Sub

Private Sub ReclasSrv(ByVal cod As String, ByVal c1 As String, _
                      ByVal c2 As String, ByVal c3 As String)
    If dCatSrv Is Nothing Then Set dCatSrv = CreateObject("Scripting.Dictionary")
    Dim ta As String, seg As String, und As String, p() As String
    cod = NormCod(cod)
    If dCatSrv.Exists(cod) Then
        p = Split(dCatSrv(cod), "|")
        If UBound(p) >= 3 Then ta = p(3)
        If UBound(p) >= 4 Then seg = p(4)
        If UBound(p) >= 5 Then und = p(5)
    End If
    dCatSrv(cod) = c1 & "|" & c2 & "|" & c3 & "|" & ta & "|" & seg & "|" & und
End Sub

' Devolve parte do catlogo de servio: 0=CLS1,1=CLS2,2=CLS3,3=TIPO_APLIC,4=SEGMENTO,5=UND
Private Function SrvInfo(codSrv As Variant, idx As Long) As String
    SrvInfo = ""
    If dCatSrv Is Nothing Then Exit Function
    Dim cod As String: cod = NormCod(codSrv)
    If dCatSrv.Exists(cod) Then
        Dim p() As String: p = Split(dCatSrv(cod), "|")
        If idx <= UBound(p) Then SrvInfo = Trim$(p(idx))
    End If
End Function


'==============================================================================
'  CATLOGO DE CLASSE DE CUSTO (CLS1, CLS2, CLS3, TIPO_APLIC)
'  Usado para a aba RISCO: classes cujo CLS3 = "RISCO".
'==============================================================================
Private Sub CarregarCatalogoClasse()
    Set dCatCC = CreateObject("Scripting.Dictionary")
    CarregarClassificacaoClassesDados

    Dim caminho As String
    caminho = CaminhoCatalogo("CAT_CLASSE", _
        "%USERPROFILE%\Downloads\CLASSE_CUSTO_ATUAIS_2.xlsx;%USERPROFILE%\Downloads\CLASSE_CUSTO_ATUAIS (2).xlsx;" & _
        "%USERPROFILE%\Downloads\CLASSE_CUSTO_ATUAIS.xlsx;%USERPROFILE%\Downloads\CLASSE_CUSTO_ATUAIS (1).xlsx")
    If caminho = "" Then
        Dim f As Variant
        f = Application.GetOpenFilename( _
            "Excel (*.xls*),*.xls*", , _
            "Selecione o catalogo de CLASSE DE CUSTO (CLASSE_CUSTO_ATUAIS). Cancele para pular.")
        If f = False Then Exit Sub
        caminho = CStr(f)
    End If

    On Error GoTo SemCat
    Dim wb As Workbook, ws As Worksheet, arr As Variant
    Set wb = Workbooks.Open(caminho, ReadOnly:=True, UpdateLinks:=0)
    Set ws = wb.Worksheets(1)

    Dim cCod As Long, c1 As Long, c2 As Long, c3 As Long, cTA As Long
    cCod = ColLike(ws, Array("CLASSE CUSTO", "CLASSE_CUSTO", "CLASSE DE CUSTO"))
    c1 = ColLike(ws, Array("CLS1"))
    c2 = ColLike(ws, Array("CLS2"))
    c3 = ColLike(ws, Array("CLS3"))
    cTA = ColLike(ws, Array("TIPO APLICACAO", "TIPO_APLICACAO", "TIPO APLIC"))
    If cCod = 0 Then wb.Close SaveChanges:=False: Exit Sub

    Dim ult As Long
    ult = ws.Cells(ws.Rows.Count, cCod).End(xlUp).Row
    arr = ws.Range(ws.Cells(2, 1), ws.Cells(ult, ws.UsedRange.Columns.Count)).Value

    Dim i As Long, cod As String
    For i = 1 To UBound(arr, 1)
        cod = NormCod(arr(i, cCod))
        If cod <> "" And Not dCatCC.Exists(cod) Then
            dCatCC(cod) = _
                TextoMatriz(arr, i, c1) & "|" & _
                TextoMatriz(arr, i, c2) & "|" & _
                TextoMatriz(arr, i, c3) & "|" & _
                TextoMatriz(arr, i, cTA)
        End If
    Next i
    wb.Close SaveChanges:=False
    Exit Sub
SemCat:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Sub CarregarClassificacaoClassesDados()
    ' Tabela oficial (CLASSIFICACAO DAS CLASSES DE CUSTO.xlsx):
    ' CLASSE DE CUSTO -> CLS1 | CLS2 | CLS3
    ' Embutida para que ANALISE DE CA classifique MOP, tributos, suporte,
    ' frete, JOA e risco mesmo sem arquivo externo.

    ' Base completa CLASSE_CUSTO_ATUAIS (782 classes); os overrides curados
    ' abaixo rodam depois e tem prioridade sobre a base.
    CarregarClassesCustoAuto

    ' --- Alimentacao Viagem ----------------------------------------------
    AddClasseCusto "8111290000", "OUTROS", "ALIMENTACAO", "SUPORTE"

    ' --- Ativacao custos FISCALIZACAO ------------------------------------
    AddClasseCusto "8010280000", "OUTROS", "MOP", "MOP"              ' FIS01
    AddClasseCusto "8110280000", "OUTROS", "MOP", "MOP"              ' FIS11
    AddClasseCusto "8210280000", "OUTROS", "MOP", "MOP"              ' FIS21
    AddClasseCusto "6150282100", "OUTROS", "MOP_CUSTEIO", "RISCO"    ' FIS21 custeio
    AddClasseCusto "8990280000", "OUTROS", "MOP", "MOP"              ' FIS99

    ' --- Ativacao custos GERENCIAMENTO -----------------------------------
    AddClasseCusto "8010270000", "OUTROS", "MOP", "MOP"              ' GRN01
    AddClasseCusto "8110270000", "OUTROS", "MOP", "MOP"              ' GRN11
    AddClasseCusto "6150271100", "OUTROS", "MOP_CUSTEIO", "RISCO"    ' GRN11 custeio
    AddClasseCusto "8210270000", "OUTROS", "MOP", "MOP"              ' GRN21
    AddClasseCusto "8990270000", "OUTROS", "MOP", "MOP"              ' GRN99

    ' --- Ativacao custos PROJETO -----------------------------------------
    AddClasseCusto "8010260000", "OUTROS", "MOP", "MOP"              ' PRJ01
    AddClasseCusto "8110260000", "OUTROS", "MOP", "MOP"              ' PRJ11
    AddClasseCusto "8210260000", "OUTROS", "MOP", "MOP"              ' PRJ21
    AddClasseCusto "6150262100", "OUTROS", "MOP_CUSTEIO", "RISCO"    ' PRJ21 custeio
    AddClasseCusto "8990260000", "OUTROS", "MOP", "MOP"              ' PRJ99

    ' --- Outras ativacoes ------------------------------------------------
    AddClasseCusto "8999080000", "OUTROS", "OUTROS", "JOA"             ' Encargos financeiros EFT
    AddClasseCusto "8119020000", "OUTROS", "ODC_MATERIAL", "TRIBUTOS"  ' Ativacao ODC Material
    AddClasseCusto "8019020000", "OUTROS", "ODC", "MOP"                ' Ativacao ODC Pessoal
    AddClasseCusto "8999900000", "OUTROS", "RISCO", "RISCO"            ' Ativo em Curso - Estudos e Projetos

    ' --- Combustiveis / COM ----------------------------------------------
    AddClasseCusto "8119980000", "OUTROS", "CC_MATERIAL", "OUTROS"     ' COM - Combustiveis/Lubrificantes
    AddClasseCusto "8110990000", "OUTROS", "OUTROS", "RISCO"           ' COM Materiais de Investimento

    ' --- Direitos / Patentes ---------------------------------------------
    AddClasseCusto "8997260000", "OUTROS", "DIREITO_SERVIDAO", "MAT UC"  ' Direito Marca e Patentes (Servidao)

    ' --- Tributos --------------------------------------------------------
    AddClasseCusto "8110980000", "OUTROS", "OUTROS", "TRIBUTOS"        ' ICMS Maquinas e Equipamentos

    ' --- Materiais com classe especifica ---------------------------------
    AddClasseCusto "8113230000", "MATERIAL", "EMENDA", "MAT COM"       ' Emenda (RECLASSIFICAR.xlsx)

    ' --- Liquidacoes / Outros --------------------------------------------
    AddClasseCusto "8003800000", "OUTROS", "CC_LIQUIDACAO", "RISCO"    ' Liq. Externa - NG 38 Taxa Fiscalizacao
    AddClasseCusto "6150001101", "OUTROS", "MOP_CUSTEIO", "RISCO"      ' Materiais de Manutencao
    AddClasseCusto "6150009931", "OUTROS", "MOP_CUSTEIO", "RISCO"      ' Outras Despesas
    AddClasseCusto "6110009024", "OUTROS", "MOP_CUSTEIO", "RISCO"      ' Outras Receitas Parceiros
    AddClasseCusto "6750009001", "OUTROS", "MOP_CUSTEIO", "RISCO"      ' Falta Inventario de Investimento (NOVO)
    AddClasseCusto "8990010000", "OUTROS", "OUTROS", "OUTROS"          ' Outros
    AddClasseCusto "8210930000", "OUTROS", "CC_SERVICO", "RISCO"       ' Serv Abertura Fx Servidao
    AddClasseCusto "8210030000", "OUTROS", "OUTROS", "RISCO"           ' Servico em SE

    ' --- Servicos especificos --------------------------------------------
    AddClasseCusto "8210400000", "OUTROS", "MOBILIDADE", "SUPORTE"     ' Servicos de Conducao
    AddClasseCusto "8210040000", "OUTROS", "RISCO", "RISCO"            ' Servicos de Construcao
    AddClasseCusto "8210520000", "SERVICO", "FRETE/TRANSP", "FRETE/TRANSP"  ' Servicos de Fretes
    AddClasseCusto "8210550000", "OUTROS", "HOSPEDAGEM", "SUPORTE"     ' Servicos de Hospedagem
    AddClasseCusto "8210430000", "OUTROS", "MEIO-AMBIENTE", "MEIO-AMBIENTE"  ' Servicos de Meio Ambiente
    AddClasseCusto "6150002143", "OUTROS", "MOP_CUSTEIO", "RISCO"      ' Servicos de Meio Ambiente custeio
    AddClasseCusto "8210390000", "OUTROS", "PASSAGEM", "SUPORTE"       ' Servicos de Passagens
    AddClasseCusto "8210630000", "OUTROS", "OUTROS", "RISCO"           ' Servicos e Terceiros Demais
    AddClasseCusto "8210640000", "OUTROS", "OUTROS", "RISCO"           ' Servicos e Terceiros Projetos
End Sub

Private Sub AddClasseCusto(ByVal cod As String, ByVal cls1 As String, ByVal cls2 As String, ByVal cls3 As String)
    If dCatCC Is Nothing Then Set dCatCC = CreateObject("Scripting.Dictionary")
    dCatCC(NormCod(cod)) = cls1 & "|" & cls2 & "|" & cls3 & "|"
End Sub

Private Sub CarregarClassesCustoAuto()
    ' Base oficial CLASSE_CUSTO_ATUAIS (embutida). Chamada ANTES dos overrides
    ' curados de CarregarClassificacaoClassesDados, que tem prioridade.
    ' Tabela embutida dividida em 10 blocos (limite de tamanho de procedimento).
    CarregarClassesCustoAutoP1
    CarregarClassesCustoAutoP2
    CarregarClassesCustoAutoP3
    CarregarClassesCustoAutoP4
    CarregarClassesCustoAutoP5
    CarregarClassesCustoAutoP6
    CarregarClassesCustoAutoP7
    CarregarClassesCustoAutoP8
    CarregarClassesCustoAutoP9
    CarregarClassesCustoAutoP10
End Sub

Private Sub CarregarClassesCustoAutoP1()
    AddClasseCusto "6105107110", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6105207110", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "610530902", "OUTROS", "ALUGUEL", "RISCO"
    AddClasseCusto "6105407110", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6105407997", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6105408145", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6110000301", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6110000302", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6110000303", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6110001111", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6110001301", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6110002110", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6110002113", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6110002158", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6110002159", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6110002190", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6110002191", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6110002197", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6110002208", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6110002302", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6110009000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6110009008", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6110009018", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6110009020", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150000101", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150000112", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150000124", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150000126", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150000127", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150000128", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150000130", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150000140", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150000142", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150000153", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150000192", "OUTROS", "ODC_PESSOAL", "TRIBUTOS"
    AddClasseCusto "6150001110", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150001111", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150001112", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150001113", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150001114", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150001115", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150001117", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150001118", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150001120", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150001121", "OUTROS", "MAT_VENDAS", "RISCO"
    AddClasseCusto "6150001122", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150001125", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150001129", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150001130", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150001192", "OUTROS", "ODC_MATERIAL", "TRIBUTOS"
    AddClasseCusto "6150001195", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002102", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002103", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002104", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002105", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002106", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002107", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002109", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002110", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002112", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002113", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002114", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "6150002115", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002116", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002120", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002121", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002123", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002125", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002126", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002130", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002131", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002132", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002133", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002134", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002136", "OUTROS", "RISCO", "RISCO"
End Sub

Private Sub CarregarClassesCustoAutoP2()
    AddClasseCusto "6150002137", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002138", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002139", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002140", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002142", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002144", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002148", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002149", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002150", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002151", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002152", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002153", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002155", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002156", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002157", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002161", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002162", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002163", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002164", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002166", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002168", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002169", "OUTROS", "ODC_MATERIAL", "TRIBUTOS"
    AddClasseCusto "6150002171", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002175", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002176", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002178", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002179", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002180", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002182", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002183", "OUTROS", "CC_CUSTEIO", "RISCO"
    AddClasseCusto "6150002186", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002192", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002193", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002195", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002207", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002211", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002212", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002214", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002232", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002234", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002238", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002246", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150002293", "OUTROS", "CC_REC_DESPESAS", "RISCO"
    AddClasseCusto "6150004425", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150005301", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150009101", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150009102", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150009103", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "6150009104", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150009201", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150009301", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150009302", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150009303", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150009401", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150009403", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150009504", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150009804", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150009907", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150009908", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150009915", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150009917", "OUTROS", "CC_CUSTEIO", "RISCO"
    AddClasseCusto "6150009919", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150009922", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150009932", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150009942", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150009946", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150009988", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150009994", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150260100", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150260200", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150261100", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150269100", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150269200", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150269300", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150269500", "OUTROS", "MOP_CUSTEIO", "RISCO"
End Sub

Private Sub CarregarClassesCustoAutoP3()
    AddClasseCusto "6150269600", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150269800", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150269900", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150270100", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150270200", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150272100", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150279100", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150279200", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150279300", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150279500", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150279600", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150279800", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150279900", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6150280100", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150280200", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150281100", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150289100", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150289200", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150289300", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150289500", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150289600", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150289800", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150289900", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150290100", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150290200", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150291100", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150292100", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150299100", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150299200", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150299300", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150299500", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150299600", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150299800", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6150299900", "OUTROS", "MOP_CUSTEIO", "RISCO"
    AddClasseCusto "6310003003", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "6310009001", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6310009002", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6350001007", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6350001013", "OUTROS", "IOF", "TRIBUTOS"
    AddClasseCusto "6350003001", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6350003004", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6350003010", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6350009001", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6350009006", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6350009012", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6710003001", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6710009001", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6750001001", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6750002001", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6750009002", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "6999999998", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8000100000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8000200000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8001100000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8001200000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8002100000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8003600000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8003700000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8003900000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8004000000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8004100000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8004200000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8005300000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8005500000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8005600000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8008100000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8008200000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8008300000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8009100000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8009200000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8009300000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8009400000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8009500000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8009600000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8009700000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
End Sub

Private Sub CarregarClassesCustoAutoP4()
    AddClasseCusto "8009800000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8009900000", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8010000001", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8010000002", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8010000003", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8010000004", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8010000180", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8010000181", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8010000182", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8010000183", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8010000184", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8010240000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8010290000", "OUTROS", "MOP_MONTAGEM", "MOP"
    AddClasseCusto "8010420000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8010440000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8010490000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8010880000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8019010000", "OUTROS", "ENCERRAMENTO ODS", "TRIBUTOS"
    AddClasseCusto "8019040000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8019930000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8020000280", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8020000281", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8020260000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8020270000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8020280000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8020290000", "OUTROS", "MOP_MONTAGEM", "MOP"
    AddClasseCusto "8020420000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8020880000", "OUTROS", "ODR", "RISCO"
    AddClasseCusto "8020970000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8029040000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8110000000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8110001180", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8110001181", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8110010000", "OUTROS", "ENCERRAMENTO ODD", "BAIXA ODD"
    AddClasseCusto "8110010093", "OUTROS", "ENCERRAMENTO ODD", "BAIXA ODD"
    AddClasseCusto "8110010094", "OUTROS", "ENCERRAMENTO ODD", "BAIXA ODD"
    AddClasseCusto "8110010095", "OUTROS", "ENCERRAMENTO ODD", "BAIXA ODD"
    AddClasseCusto "8110010096", "OUTROS", "ENCERRAMENTO ODD", "BAIXA ODD"
    AddClasseCusto "8110010097", "OUTROS", "ENCERRAMENTO ODD", "BAIXA ODD"
    AddClasseCusto "8110010098", "OUTROS", "ENCERRAMENTO ODD", "BAIXA ODD"
    AddClasseCusto "8110010099", "OUTROS", "ENCERRAMENTO ODD", "BAIXA ODD"
    AddClasseCusto "8110020000", "OUTROS", "ENCERRAMENTO ODD", "BAIXA ODD"
    AddClasseCusto "8110020093", "OUTROS", "ENCERRAMENTO ODD", "BAIXA ODD"
    AddClasseCusto "8110020094", "OUTROS", "ENCERRAMENTO ODD", "BAIXA ODD"
    AddClasseCusto "8110020095", "OUTROS", "ENCERRAMENTO ODD", "BAIXA ODD"
    AddClasseCusto "8110020096", "OUTROS", "ENCERRAMENTO ODD", "BAIXA ODD"
    AddClasseCusto "8110020097", "OUTROS", "ENCERRAMENTO ODD", "BAIXA ODD"
    AddClasseCusto "8110020098", "OUTROS", "ENCERRAMENTO ODD", "BAIXA ODD"
    AddClasseCusto "8110020099", "OUTROS", "ENCERRAMENTO ODD", "BAIXA ODD"
    AddClasseCusto "8110030000", "OUTROS", "ODA_N/A_EM_OBRA", "RISCO"
    AddClasseCusto "8110040000", "OUTROS", "ODA_N/A_EM_OBRA", "RISCO"
    AddClasseCusto "8110050000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8110060000", "OUTROS", "ODC", "OUTROS"
    AddClasseCusto "8110070000", "OUTROS", "ODP", "RISCO"
    AddClasseCusto "8110080000", "OUTROS", "ODR", "RISCO"
    AddClasseCusto "8110100000", "OUTROS", "ODR", "RISCO"
    AddClasseCusto "8110110000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8110120000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8110130000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8110140000", "OUTROS", "BENFEITORIAS", "OUTROS"
    AddClasseCusto "8110150000", "OUTROS", "TRANSF. FAB. E REP M", "RISCO"
    AddClasseCusto "8110290000", "OUTROS", "MOP_MONTAGEM", "MOP"
    AddClasseCusto "8110940000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8110950000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8110970000", "OUTROS", "OUTROS", "TRIBUTOS"
    AddClasseCusto "8110999000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111010000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111060000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8111080000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111090000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111140000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8111160000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111190000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8111200000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111210000", "OUTROS", "CC_MATERIAL", "RISCO"
End Sub

Private Sub CarregarClassesCustoAutoP5()
    AddClasseCusto "8111220000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111230000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111250000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111260000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111270000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8111320000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111330000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8111340000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111350000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8111350001", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111350002", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111350003", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111350004", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111350005", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111350006", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111350007", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111350008", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111360000", "OUTROS", "MATERIAL", "RISCO"
    AddClasseCusto "8111370000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8111380000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8111390000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111410000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111480000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111490000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111493000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111500000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8111540000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111550000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111570000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111580000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111600000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111630000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111650000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111670000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8111680000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8111690000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111720000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111800000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8111810000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8111820000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111830000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111870000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111880000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111890000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8111900000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111910000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111920000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111930000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111940000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111950000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111960000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8111970000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8113070000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8113150000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8113200000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8113210000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8113220000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8113360000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8113450000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8113470000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8113480000", "MATERIAL", "RISCO", "RISCO"
    AddClasseCusto "8113490000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8113550000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8113560000", "OUTROS", "MATERIAL", "RISCO"
    AddClasseCusto "8113570000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8113600000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8113880000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8113900000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8113910000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8113920000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8113930000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8114680000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8114690000", "OUTROS", "SIST COMUNIC LOCAL", "RISCO"
    AddClasseCusto "8114770000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8114800000", "OUTROS", "CC_MATERIAL", "RISCO"
End Sub

Private Sub CarregarClassesCustoAutoP6()
    AddClasseCusto "8114810000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8114820000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8114830000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8114840000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8114850000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8114860000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8114870000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8114880000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8114890000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8114920000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8116290000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8116860000", "MATERIAL", "TERRENO", "MAT. UC"
    AddClasseCusto "8116960000", "OUTROS", "URBAN_BENFEITORIAS", "MAO DE OBRA CIVIL"
    AddClasseCusto "8116970000", "OUTROS", "URBAN_BENFEITORIAS", "MAO DE OBRA CIVIL"
    AddClasseCusto "8116990000", "OUTROS", "URBAN_BENFEITORIAS", "MAO DE OBRA CIVIL"
    AddClasseCusto "8117250000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8117260000", "OUTROS", "INDENIZACAO_SERVIDAO", "OUTROS"
    AddClasseCusto "8117270000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8117310000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8117320000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8117330000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8118980000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8118990000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8119000000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8119010000", "OUTROS", "ENCERRAMENTO ODS", "TRIBUTOS"
    AddClasseCusto "8119030000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8119040000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8119050000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8119060000", "OUTROS", "CC_MATERIAL", "TRIBUTOS"
    AddClasseCusto "8119870000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8119880000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8119900000", "OUTROS", "ODA_N/A_EM_OBRA", "RISCO"
    AddClasseCusto "8119920000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8119930000", "MATERIAL", "SALDO REMANESCENTE", "RISCO"
    AddClasseCusto "8119940000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8119960000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8119970000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8119990000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8210002290", "OUTROS", "CC_SERVICO", "RISCO"
    AddClasseCusto "8210002291", "OUTROS", "CC_SERVICO", "RISCO"
    AddClasseCusto "8210002292", "OUTROS", "CC_SERVICO", "RISCO"
    AddClasseCusto "8210010000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8210020000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8210050000", "OUTROS", "RECUP_EQP RT", "RISCO"
    AddClasseCusto "8210060000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8210070000", "OUTROS", "MANUT_VEICULO", "RISCO"
    AddClasseCusto "8210080000", "OUTROS", "CC_SERVICO", "RISCO"
    AddClasseCusto "8210090000", "OUTROS", "EQUIPAMENTOS G", "RISCO"
    AddClasseCusto "8210100000", "OUTROS", "CC_SERVICO", "RISCO"
    AddClasseCusto "8210110000", "OUTROS", "CC_SERVICO", "RISCO"
    AddClasseCusto "8210120000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8210130000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8210150000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8210160000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8210190000", "OUTROS", "CC_SERVICO", "RISCO"
    AddClasseCusto "8210200000", "OUTROS", "CC_SERVICO", "RISCO"
    AddClasseCusto "8210210000", "OUTROS", "CC_SERVICO", "RISCO"
    AddClasseCusto "8210230000", "OUTROS", "CONSULTORIA", "SUPORTE"
    AddClasseCusto "8210240000", "OUTROS", "OUTROS", "OUTROS"
    AddClasseCusto "8210250000", "OUTROS", "CC_SERVICO", "RISCO"
    AddClasseCusto "8210290000", "OUTROS", "MOP_MONTAGEM", "MOP"
    AddClasseCusto "8210300000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8210310000", "OUTROS", "CC_SERVICO", "RISCO"
    AddClasseCusto "8210320000", "OUTROS", "SOFTWARE", "OUTROS"
    AddClasseCusto "8210340000", "SERVICO", "FRETE/TRANSP", "FRETE/TRANSP"
    AddClasseCusto "8210350000", "OUTROS", "CC_SERVICO", "RISCO"
    AddClasseCusto "8210360000", "OUTROS", "CC_SERVICO", "RISCO"
    AddClasseCusto "8210380000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8210410000", "OUTROS", "CC_SERVICO", "RISCO"
    AddClasseCusto "8210420000", "OUTROS", "CC_SERVICO", "RISCO"
    AddClasseCusto "8210440000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8210460000", "OUTROS", "CC_SERVICO", "RISCO"
    AddClasseCusto "8210480000", "OUTROS", "CC_MATERIAL", "RISCO"
    AddClasseCusto "8210490000", "OUTROS", "PUBLICIDADE", "PUBLICIDADE"
    AddClasseCusto "8210500000", "OUTROS", "SERVICOS DE REPROGRAFIA", "RISCO"
End Sub

Private Sub CarregarClassesCustoAutoP7()
    AddClasseCusto "8210510000", "OUTROS", "AGUA ESGOTO", "RISCO"
    AddClasseCusto "8210530000", "SERVICO", "RISCO", "RISCO"
    AddClasseCusto "8210540000", "OUTROS", "CC_SERVICO", "RISCO"
    AddClasseCusto "8210560000", "OUTROS", "ALIMENTACAO", "SUPORTE"
    AddClasseCusto "8210570000", "OUTROS", "INDENIZACAO", "SUPORTE"
    AddClasseCusto "8210580000", "OUTROS", "CC_SERVICO", "RISCO"
    AddClasseCusto "8210620000", "OUTROS", "SERVICO DE IMPRESSAO", "RISCO"
    AddClasseCusto "8210660000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8210710000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8210750000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8210760000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8210770000", "OUTROS", "CC_SERVICO", "RISCO"
    AddClasseCusto "8210780000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8210790000", "OUTROS", "RESCISAO CONTRATUAL", "RISCO"
    AddClasseCusto "8210800000", "OUTROS", "OUTROS", "SUPORTE"
    AddClasseCusto "8210820000", "OUTROS", "LOGISTICA", "SUPORTE"
    AddClasseCusto "8210840000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8210880000", "OUTROS", "ODR", "RISCO"
    AddClasseCusto "8210920000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8210940000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8210950000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8210970000", "OUTROS", "OUTROS", "TRIBUTOS"
    AddClasseCusto "8219000000", "OUTROS", "ENCERRAMENTO ODS", "OUTROS"
    AddClasseCusto "8219010000", "OUTROS", "ENCERRAMENTO ODS", "TRIBUTOS"
    AddClasseCusto "8219020000", "OUTROS", "ODC", "MOP"
    AddClasseCusto "8219030000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8219040000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8219050000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8219060000", "OUTROS", "CC_MATERIAL", "TRIBUTOS"
    AddClasseCusto "8219070000", "OUTROS", "CC_SERVICO", "RISCO"
    AddClasseCusto "8219930000", "OUTROS", "SALDO REMANESCENTE", "RISCO"
    AddClasseCusto "8220070000", "OUTROS", "SOFTWARE", "OUTROS"
    AddClasseCusto "8220100000", "OUTROS", "CC_SERVICO", "RISCO"
    AddClasseCusto "8220310000", "OUTROS", "CC_SERVICO", "RISCO"
    AddClasseCusto "8220350000", "OUTROS", "CC_SERVICO", "RISCO"
    AddClasseCusto "8220360000", "OUTROS", "CC_SERVICO", "RISCO"
    AddClasseCusto "8220370000", "OUTROS", "CC_SERVICO", "RISCO"
    AddClasseCusto "8530010000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8530020000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8530030000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8530940000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8530970000", "OUTROS", "OUTROS", "TRIBUTOS"
    AddClasseCusto "8539000000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8539010000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8539020000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8539030000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8539040000", "OUTROS", "DEPRECIACAO", "RISCO"
    AddClasseCusto "8539050000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8539060000", "OUTROS", "CC_MATERIAL", "OUTROS"
    AddClasseCusto "8539930000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8550010000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8550020000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8550030000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8559010000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8559020000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8559030000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8559040000", "OUTROS", "AMORTIZACAO", "TRIBUTOS"
    AddClasseCusto "8559050000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8559060000", "OUTROS", "CC_MATERIAL", "OUTROS"
    AddClasseCusto "8611002101", "OUTROS", "PIS", "TRIBUTOS"
    AddClasseCusto "8611002102", "OUTROS", "COFINS", "TRIBUTOS"
    AddClasseCusto "8611002103", "OUTROS", "TRIBUTOS", "TRIBUTOS"
    AddClasseCusto "8611002301", "OUTROS", "TRIBUTOS", "TRIBUTOS"
    AddClasseCusto "8611009000", "OUTROS", "ENCERRAMENTO ODS", "TRIBUTOS"
    AddClasseCusto "8611009007", "OUTROS", "ODA_N/A_EM_OBRA", "RISCO"
    AddClasseCusto "8611009020", "OUTROS", "ENCERRAMENTO DE ODS", "OUTROS"
    AddClasseCusto "8611880000", "OUTROS", "ODR", "RISCO"
    AddClasseCusto "8810010000", "OUTROS", "OUTROS", "TRIBUTOS"
    AddClasseCusto "8810020000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8810030000", "OUTROS", "OUTROS", "TRIBUTOS"
    AddClasseCusto "8819930000", "OUTROS", "OUTROS", "TRIBUTOS"
    AddClasseCusto "8820010000", "OUTROS", "OUTROS", "OUTROS"
    AddClasseCusto "8820020000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8826860000", "MATERIAL", "TERRENO", "MAT. UC"
    AddClasseCusto "8829930000", "OUTROS", "OUTROS", "TRIBUTOS"
End Sub

Private Sub CarregarClassesCustoAutoP8()
    AddClasseCusto "8830010000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8830020000", "OUTROS", "INDENIZACAO", "RISCO"
    AddClasseCusto "8830030000", "OUTROS", "INDENIZACAO", "RISCO"
    AddClasseCusto "8839930000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8840010000", "OUTROS", "ENCERRAMENTO ODD", "BAIXA ODD"
    AddClasseCusto "8840020000", "OUTROS", "ODA_N/A_EM_OBRA", "RISCO"
    AddClasseCusto "8900000000", "OUTROS", "CC_SERVICO/MATERIAL", "RISCO"
    AddClasseCusto "8910009180", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8910009181", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8910009182", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8910010000", "OUTROS", "SUPORTE ALUGUEL", "RISCO"
    AddClasseCusto "8910020000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8910260000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8910270000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8910280000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8910290000", "OUTROS", "MOP_MONTAGEM", "MOP"
    AddClasseCusto "8910880000", "OUTROS", "ODR", "RISCO"
    AddClasseCusto "8910950000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8910970000", "OUTROS", "OUTROS", "TRIBUTOS"
    AddClasseCusto "8919000000", "OUTROS", "ENCERRAMENTO ODS", "OUTROS"
    AddClasseCusto "8919010000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8919020000", "OUTROS", "ODC", "MOP"
    AddClasseCusto "8919030000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8919040000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8919050000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8919080000", "OUTROS", "OUTROS", "JOA"
    AddClasseCusto "8919930000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8920009280", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8920009281", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8920010000", "OUTROS", "CC_SEGUROS", "RISCO"
    AddClasseCusto "8920020000", "OUTROS", "CC_SEGUROS", "RISCO"
    AddClasseCusto "8920260000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8920270000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8920280000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8920290000", "OUTROS", "MOP_MONTAGEM", "MOP"
    AddClasseCusto "8920950000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8920970000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8929000000", "OUTROS", "ENCERRAMENTO ODS", "OUTROS"
    AddClasseCusto "8929010000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8929020000", "OUTROS", "ODC", "MOP"
    AddClasseCusto "8929030000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8929040000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8929050000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8929930000", "OUTROS", "CC_SEGUROS", "RISCO"
    AddClasseCusto "8930009380", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8930009381", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8930010000", "OUTROS", "OUTROS", "TRIBUTOS"
    AddClasseCusto "8930020000", "OUTROS", "OUTROS", "TRIBUTOS"
    AddClasseCusto "8930260000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8930270000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8930280000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8930290000", "OUTROS", "MOP_MONTAGEM", "MOP"
    AddClasseCusto "8930950000", "OUTROS", "TRIBUTOS", "TRIBUTOS"
    AddClasseCusto "8930970000", "OUTROS", "OUTROS", "TRIBUTOS"
    AddClasseCusto "8939000000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8939010000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8939020000", "OUTROS", "TRIBUTOS", "TRIBUTOS"
    AddClasseCusto "8939030000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8939040000", "OUTROS", "OUTROS", "TRIBUTOS"
    AddClasseCusto "8939050000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8939930000", "OUTROS", "OUTROS", "TRIBUTOS"
    AddClasseCusto "8940009480", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8949020000", "OUTROS", "CC_TF_ODC", "TRIBUTOS"
    AddClasseCusto "8950009580", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8950009581", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8950040000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8950260000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8950270000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8950280000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8950290000", "OUTROS", "MOP_MONTAGEM", "MOP"
    AddClasseCusto "8950950000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8960009680", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8960009681", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8960260000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8960270000", "OUTROS", "MOP", "MOP"
End Sub

Private Sub CarregarClassesCustoAutoP9()
    AddClasseCusto "8960280000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8960290000", "OUTROS", "MOP_MONTAGEM", "MOP"
    AddClasseCusto "8960950000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8970009780", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8970950000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8970970000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8980009880", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8980009881", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8980009882", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8980010000", "OUTROS", "CC_REC_DESPESAS", "RISCO"
    AddClasseCusto "8980260000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8980270000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8980280000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8980290000", "OUTROS", "MOP_MONTAGEM", "MOP"
    AddClasseCusto "8989000000", "OUTROS", "ENCERRAMENTO ODS", "OUTROS"
    AddClasseCusto "8989010000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8989020000", "OUTROS", "ODC", "MOP"
    AddClasseCusto "8989040000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8989050000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8989060000", "OUTROS", "CC_TF_ODC", "OUTROS"
    AddClasseCusto "8989080000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8990009980", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8990009981", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8990010001", "OUTROS", "OUTROS", "TRIBUTOS"
    AddClasseCusto "8990020000", "OUTROS", "SALDO REMANESCENTE", "RISCO"
    AddClasseCusto "8990030000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8990040000", "OUTROS", "INDENIZACAO_SERVIDAO", "OUTROS"
    AddClasseCusto "8990050000", "OUTROS", "CC_POS_GRADUACAO", "RISCO"
    AddClasseCusto "8990080000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8990190000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8990290000", "OUTROS", "MOP_MONTAGEM", "MOP"
    AddClasseCusto "8990880000", "OUTROS", "ODR", "RISCO"
    AddClasseCusto "8990950000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8990970000", "OUTROS", "OUTROS", "TRIBUTOS"
    AddClasseCusto "8990990000", "OUTROS", "CC_PERDAS", "RISCO"
    AddClasseCusto "8997250000", "OUTROS", "DIR_PATENTES", "RISCO"
    AddClasseCusto "8997270000", "OUTROS", "OUTROS", "OUTROS"
    AddClasseCusto "8999000000", "OUTROS", "ENCERRAMENTO ODS", "OUTROS"
    AddClasseCusto "8999010000", "OUTROS", "ENCERRAMENTO ODS", "TRIBUTOS"
    AddClasseCusto "8999020000", "OUTROS", "ODC", "MOP"
    AddClasseCusto "8999030000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8999040000", "OUTROS", "MOP", "MOP"
    AddClasseCusto "8999050000", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8999910000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8999920000", "OUTROS", "TERRENO", "MAT. UC"
    AddClasseCusto "8999930000", "OUTROS", "SALDO REMANESCENTE", "RISCO"
    AddClasseCusto "8999940000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8999950000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8999960000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8999970000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8999980000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8999980001", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8999980002", "OUTROS", "CC_LIQUIDACAO", "RISCO"
    AddClasseCusto "8999990000", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "8999999997", "OUTROS", "OUTROS", "RISCO"
    AddClasseCusto "8999999998", "OUTROS", "OUTROS", "OUTROS"
    AddClasseCusto "8999999999", "OUTROS", "RISCO", "RISCO"
    AddClasseCusto "ADM01", "OUTROS", "MOP", "MOP"
    AddClasseCusto "ADM02", "OUTROS", "MOP", "MOP"
    AddClasseCusto "ADM11", "OUTROS", "MOP", "MOP"
    AddClasseCusto "ADM21", "OUTROS", "MOP", "MOP"
    AddClasseCusto "ADM53", "OUTROS", "DEPRECIACAO", "RISCO"
    AddClasseCusto "ADM55", "OUTROS", "OUTROS", "TRIBUTOS"
    AddClasseCusto "ADM91", "OUTROS", "MOP", "MOP"
    AddClasseCusto "ADM92", "OUTROS", "MOP", "MOP"
    AddClasseCusto "ADM93", "OUTROS", "OUTROS", "TRIBUTOS"
    AddClasseCusto "ADM98", "OUTROS", "MOP", "MOP"
    AddClasseCusto "ADM99", "OUTROS", "MOP", "MOP"
    AddClasseCusto "AHT01", "OUTROS", "MOP", "MOP"
    AddClasseCusto "AHT02", "OUTROS", "MOP", "MOP"
    AddClasseCusto "AHT11", "OUTROS", "MOP", "MOP"
    AddClasseCusto "AHT21", "OUTROS", "MOP", "MOP"
    AddClasseCusto "AHT53", "OUTROS", "OUTROS", "OUTROS"
    AddClasseCusto "AHT55", "OUTROS", "OUTROS", "OUTROS"
    AddClasseCusto "AHT91", "OUTROS", "MOP", "MOP"
End Sub

Private Sub CarregarClassesCustoAutoP10()
    AddClasseCusto "AHT92", "OUTROS", "MOP", "MOP"
    AddClasseCusto "AHT93", "OUTROS", "TRIBUTOS", "TRIBUTOS"
    AddClasseCusto "AHT95", "OUTROS", "MOP", "MOP"
    AddClasseCusto "AHT96", "OUTROS", "MOP", "MOP"
    AddClasseCusto "AHT97", "OUTROS", "MOP", "MOP"
    AddClasseCusto "AHT98", "OUTROS", "MOP", "MOP"
    AddClasseCusto "AHT99", "OUTROS", "MOP", "MOP"
    AddClasseCusto "EFT91", "OUTROS", "OUTROS", "JOA"
    AddClasseCusto "EFT93", "OUTROS", "OUTROS", "TRIBUTOS"
    AddClasseCusto "EFT99", "OUTROS", "JOA", "JOA"
    AddClasseCusto "FIS01", "OUTROS", "MOP", "MOP"
    AddClasseCusto "FIS02", "OUTROS", "MOP", "MOP"
    AddClasseCusto "FIS11", "OUTROS", "MOP", "MOP"
    AddClasseCusto "FIS21", "OUTROS", "MOP", "MOP"
    AddClasseCusto "FIS91", "OUTROS", "MOP", "MOP"
    AddClasseCusto "FIS92", "OUTROS", "MOP", "MOP"
    AddClasseCusto "FIS93", "OUTROS", "MOP", "MOP"
    AddClasseCusto "FIS95", "OUTROS", "MOP", "MOP"
    AddClasseCusto "FIS96", "OUTROS", "MOP", "MOP"
    AddClasseCusto "FIS98", "OUTROS", "MOP", "MOP"
    AddClasseCusto "FIS99", "OUTROS", "MOP", "MOP"
    AddClasseCusto "GRN01", "OUTROS", "MOP", "MOP"
    AddClasseCusto "GRN02", "OUTROS", "MOP", "MOP"
    AddClasseCusto "GRN11", "OUTROS", "MOP", "MOP"
    AddClasseCusto "GRN21", "OUTROS", "MOP", "MOP"
    AddClasseCusto "GRN91", "OUTROS", "MOP", "MOP"
    AddClasseCusto "GRN92", "OUTROS", "MOP", "MOP"
    AddClasseCusto "GRN93", "OUTROS", "MOP", "MOP"
    AddClasseCusto "GRN95", "OUTROS", "MOP", "MOP"
    AddClasseCusto "GRN96", "OUTROS", "MOP", "MOP"
    AddClasseCusto "GRN98", "OUTROS", "MOP", "MOP"
    AddClasseCusto "GRN99", "OUTROS", "MOP", "MOP"
    AddClasseCusto "MON01", "OUTROS", "MOP", "MOP"
    AddClasseCusto "MON02", "OUTROS", "MOP_MONTAGEM", "MOP"
    AddClasseCusto "MON11", "OUTROS", "MOP_MONTAGEM", "MOP"
    AddClasseCusto "MON21", "OUTROS", "MOP_MONTAGEM", "MOP"
    AddClasseCusto "MON91", "OUTROS", "MOP_MONTAGEM", "MOP"
    AddClasseCusto "MON92", "OUTROS", "MOP", "MOP"
    AddClasseCusto "MON93", "OUTROS", "MOP_MONTAGEM", "MOP"
    AddClasseCusto "MON95", "OUTROS", "MOP", "MOP"
    AddClasseCusto "MON96", "OUTROS", "MOP", "MOP"
    AddClasseCusto "MON98", "OUTROS", "MOP", "MOP"
    AddClasseCusto "MON99", "OUTROS", "MOP_MONTAGEM", "MOP"
    AddClasseCusto "ODC01", "OUTROS", "ODC", "TRIBUTOS"
    AddClasseCusto "ODC11", "OUTROS", "ODC", "TRIBUTOS"
    AddClasseCusto "ODC21", "OUTROS", "ODC", "TRIBUTOS"
    AddClasseCusto "ODC53", "OUTROS", "ODC", "OUTROS"
    AddClasseCusto "ODC55", "OUTROS", "ODC", "OUTROS"
    AddClasseCusto "ODC91", "OUTROS", "ODC", "TRIBUTOS"
    AddClasseCusto "ODC92", "OUTROS", "ODC", "TRIBUTOS"
    AddClasseCusto "ODC93", "OUTROS", "ODC", "TRIBUTOS"
    AddClasseCusto "ODC94", "OUTROS", "ODC", "TRIBUTOS"
    AddClasseCusto "ODC98", "OUTROS", "ODC", "TRIBUTOS"
    AddClasseCusto "ODC99", "OUTROS", "ODC", "TRIBUTOS"
    AddClasseCusto "PRJ01", "OUTROS", "MOP", "MOP"
    AddClasseCusto "PRJ02", "OUTROS", "MOP", "MOP"
    AddClasseCusto "PRJ11", "OUTROS", "MOP", "MOP"
    AddClasseCusto "PRJ21", "OUTROS", "MOP", "MOP"
    AddClasseCusto "PRJ91", "OUTROS", "MOP", "MOP"
    AddClasseCusto "PRJ92", "OUTROS", "MOP", "MOP"
    AddClasseCusto "PRJ93", "OUTROS", "MOP", "MOP"
    AddClasseCusto "PRJ95", "OUTROS", "MOP", "MOP"
    AddClasseCusto "PRJ96", "OUTROS", "MOP", "MOP"
    AddClasseCusto "PRJ98", "OUTROS", "MOP", "MOP"
    AddClasseCusto "PRJ99", "OUTROS", "MOP", "MOP"
End Sub


' Devolve parte do catlogo de classe: 0=CLS1,1=CLS2,2=CLS3,3=TIPO_APLIC
Private Function CCInfo(codCC As Variant, idx As Long) As String
    CCInfo = ""
    If dCatCC Is Nothing Then Exit Function
    Dim cod As String: cod = NormCod(codCC)
    If dCatCC.Exists(cod) Then
        Dim p() As String: p = Split(dCatCC(cod), "|")
        If idx <= UBound(p) Then CCInfo = Trim$(p(idx))
    End If
End Function


'==============================================================================
'  CONVERSAO DE CABO (KG -> METROS)   [ajuste fino da aderncia]
'  Cabos so lanados em KG; o servio mede em METROS. Converte:
'     metros = kg * FATOR   (FATOR = metros por kg, do catlogo)
'==============================================================================
Private Sub CarregarConversoesCabo()
    Set dCabo = CreateObject("Scripting.Dictionary")

    Dim caminho As String
    caminho = CaminhoCatalogo("CAT_CABO", "%USERPROFILE%\Downloads\CONVERSOES_CABO_ATUAIS.xlsx")
    If caminho = "" Then Exit Sub   ' opcional: sem arquivo, segue sem converter

    On Error GoTo SemCat
    Dim wb As Workbook, ws As Worksheet, arr As Variant
    Set wb = Workbooks.Open(caminho, ReadOnly:=True, UpdateLinks:=0)
    Set ws = wb.Worksheets(1)

    Dim cCod As Long, cFat As Long
    cCod = ColLike(ws, Array("COD MATERIAL", "COD_MATERIAL", "MATERIAL"))
    cFat = ColLike(ws, Array("FATOR"))
    If cCod = 0 Or cFat = 0 Then wb.Close SaveChanges:=False: Exit Sub

    Dim ult As Long
    ult = ws.Cells(ws.Rows.Count, cCod).End(xlUp).Row
    arr = ws.Range(ws.Cells(2, 1), ws.Cells(ult, ws.UsedRange.Columns.Count)).Value

    Dim i As Long, cod As String, f As Double
    For i = 1 To UBound(arr, 1)
        cod = NormCod(arr(i, cCod))
        f = ToNum(arr(i, cFat))
        If cod <> "" And f > 0 And Not dCabo.Exists(cod) Then dCabo(cod) = f
    Next i
    wb.Close SaveChanges:=False
    Exit Sub
SemCat:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

' Fator de converso do material (0 = sem converso)
Private Function CaboFator(codMat As Variant) As Double
    CaboFator = 0
    If dCabo Is Nothing Then Exit Function
    Dim cod As String: cod = NormCod(codMat)
    If dCabo.Exists(cod) Then CaboFator = dCabo(cod)
End Function


'==============================================================================
'  SRV COMBO (FATOR MULTIPLICADOR DE SERVIO)   [ajuste fino da aderncia]
'  Alguns servios instalam vrias unidades (ex.: "cruzeta dupla" = 2) ou so
'  medidos em KM (fator 1000). A QTD do servio  multiplicada pelo FATOR para
'  bater com a QTD de material instalado.
'==============================================================================
Private Sub CarregarComboServico()
    Set dCombo = CreateObject("Scripting.Dictionary")

    ' Combos fixos embutidos (garantidos mesmo sem o catalogo externo).
    ' O loop do arquivo usa Not dCombo.Exists(cod), entao estes prevalecem.
    dCombo("5500000582") = "3|"   ' equivale a 3 cabos -> multiplica QTD por 3
    dCombo("5500000575") = "3|"   ' equivale a 3 cabos -> multiplica QTD por 3

    Dim caminho As String
    caminho = CaminhoCatalogo("CAT_COMBO", "%USERPROFILE%\Downloads\SRV_COMBO_ATUAIS.xlsx")
    If caminho = "" Then
        Dim f0 As Variant
        f0 = Application.GetOpenFilename( _
            "Excel (*.xls*),*.xls*", , _
            "Selecione o catalogo SRV COMBO (SRV_COMBO_ATUAIS). Cancele para pular.")
        If f0 = False Then Exit Sub
        caminho = CStr(f0)
    End If

    On Error GoTo SemCat
    Dim wb As Workbook, ws As Worksheet, arr As Variant
    Set wb = Workbooks.Open(caminho, ReadOnly:=True, UpdateLinks:=0)
    Set ws = wb.Worksheets(1)

    Dim cCod As Long, cFat As Long, cFam As Long
    cCod = ColLike(ws, Array("SRV_PRINCIPAL", "SRV PRINCIPAL", "COD SERVICO", "SERVICO"))
    cFat = ColLike(ws, Array("FATOR"))
    cFam = ColLike(ws, Array("FAMILIA"))
    If cCod = 0 Or cFat = 0 Then wb.Close SaveChanges:=False: Exit Sub

    Dim ult As Long
    ult = ws.Cells(ws.Rows.Count, cCod).End(xlUp).Row
    arr = ws.Range(ws.Cells(2, 1), ws.Cells(ult, ws.UsedRange.Columns.Count)).Value

    Dim i As Long, cod As String, f As Double, fam As String
    For i = 1 To UBound(arr, 1)
        cod = NormCod(arr(i, cCod))
        f = ToNum(arr(i, cFat))
        fam = IIf(cFam > 0, Trim$(CStr(arr(i, cFam))), "")
        ' guarda "fator|familia"
        If cod <> "" And f > 0 And Not dCombo.Exists(cod) Then dCombo(cod) = CStr(f) & "|" & fam
    Next i
    wb.Close SaveChanges:=False
    Exit Sub
SemCat:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

' Fator multiplicador do servio. So aplica quando a FAMILIA do combo bate com o
' CLS2 do servico (evita inflar familia diferente, ex.: TRAFO x PARA RAIO MT).
Private Function ComboFator(codSrv As Variant) As Double
    ComboFator = 1
    If dCombo Is Nothing Then Exit Function
    Dim cod As String: cod = NormCod(codSrv)
    If Not dCombo.Exists(cod) Then Exit Function
    Dim p() As String: p = Split(dCombo(cod), "|")
    Dim fator As Double: fator = ToNum(p(0))
    Dim famCombo As String: famCombo = ""
    If UBound(p) >= 1 Then famCombo = p(1)
    ' CLS2 do servico (do catalogo de servicos)
    Dim cls2Srv As String: cls2Srv = SrvInfo(codSrv, 1)
    If famCombo = "" Or NormClassif(famCombo) = NormClassif(cls2Srv) Then
        ComboFator = fator        ' sem familia (assume ok) ou familia bate
    Else
        ComboFator = 1            ' familia nao bate -> nao multiplica
    End If
End Function
'==============================================================================
'  DE-PARA: CLASSIFICACAO (familia/CLS2) -> TIPO (COM / UC / UAR)
'  Mapeamento fixo (de Pasta5). Onde a mesma familia aparece com TIPO diferente,
'  vale o PRIMEIRO (comportamento de PROCV). Comparacao ignora acento/maiusc.
'==============================================================================
Private Sub CarregarTipoClassif()
    Set dTipoCls = CreateObject("Scripting.Dictionary")
    Dim s As String
    s = "CRUZETA=COM;COM_MAT_SERV=COM;PARA RAIO BT=COM;PARA RAIO MT=COM;POSTE RD=UC;" & _
        "SUPORTE MET=COM;CH FUS=COM;ISOLADOR PINO MT=COM;PINO=COM;TRAFO=UC;COND NU=UC;" & _
        "ISOLADOR BT=COM;ALCA=COM;PLACA ID=COM;CABO ISOLADO=UC;ELETRODUTO=COM;HASTE ATERR=COM;" & _
        "MAO FRANCESA=COM;COND COBRE=UC;SUPORTE TRAFO=COM;COND ISOLADO=UC;CINTA POSTE=COM;" & _
        "RAMAL=UC;COND PROT=UC;CH SU FACA=UC;CORDOALHA=COM;ISOLADOR ANC MT=COM;EMENDA=COM;" & _
        "ESPACADOR LOSAG=COM;PROTETOR RETICUL=UC;SUPORTE=COM;RELIGADOR=UC;POSTE PDR ACO=UC;" & _
        "TERMINAL ANEL/OLHAL=COM;CABO COBRE=COM;TC_COM=COM;MEDIDOR_FISCAL=UC;TERMINAL ILHOS=COM;" & _
        "CH AFER=COM;CAIXA PDR=COM;CAIXA_TC=COM;CABO CONTROLE=COM;PLACA ADVERT=COM;TERMINAL COND=COM;" & _
        "DUTO=COM;CX DISPLAY SMC=COM;CP_CS_MD=UAR;TER_LEITURA=UAR;MEDIDOR=UC;REGULADOR=UC;" & _
        "CH FUS RELIG=UC;HASTE ANCORA ESTAI=COM;TP_TC=UC;ISOLADOR ESPACADOR=COM;VIGA MET=COM;" & _
        "CABO NU=COM;BLOCO CONC ESTAI=COM;GRAXA (ANE)=COM;LUBRIFICANTE=COM;MUFLA=COM;" & _
        "PAINEL CONTR EXAUSTOR=UAR;TERMINAL BIMETALICO=COM;BOMBA SUBM=UAR;EXAUSTOR=COM;" & _
        "TAMPA DE FERRO=COM;ESPACADOR=COM;RELE=UAR;TERMINAL CABO=COM;CANALETA=COM;CH VAC 1F=UC;" & _
        "POSTE CAPITEL=UC;TORA EUCALIPTO=COM;CAIXA DE PASSAGEM=COM;POSTE LD=UC;CAPACITOR=UC;" & _
        "CONTROLADOR=UAR;MURO CONC=UC;GRAMPO=COM;DISJ BT=COM;PORTA FUSIVEL=COM;" & _
        "POSTE_TORRE=UC;TORRE MET=UC;TORRE CONC=UC;TRAFO DE FORCA=UC;DISJ SE=UC;" & _
        "CH SEC TRI=UC;BANCO CAPACITOR=UC;ISOLADOR AT=COM;CONECTOR=COM;" & _
        "CABO FIB OPT=COM;VIGA CONC=COM;SUPORTE CONC=COM;ANEL CONC=COM;" & _
        "CANTONEIRA=COM;PAINEL MET=COM;TUBO FOFO=COM"

    Dim parts() As String, kv() As String, i As Long, key As String
    parts = Split(s, ";")
    For i = 0 To UBound(parts)
        kv = Split(parts(i), "=")
        If UBound(kv) >= 1 Then
            key = NormClassif(kv(0))
            If key <> "" And Not dTipoCls.Exists(key) Then dTipoCls(key) = Trim$(kv(1))
        End If
    Next i

    ' Preenche familias fora da tabela curada a partir do CLS3 do catalogo
    ' (curadas acima tem prioridade; AddTipoCls nao sobrepoe).
    CarregarTipoClassifAuto

    ' Familias UAR adicionais declaradas na CONFIG (chave FAM_UAR, separadas por ;).
    ' UAR e equivalente a UC: PEP com essas familias nao gera alerta "SEM UC".
    Dim famUAR As String: famUAR = CfgTxt("FAM_UAR", "")
    If famUAR <> "" Then
        parts = Split(famUAR, ";")
        For i = 0 To UBound(parts)
            key = NormClassif(parts(i))
            If key <> "" Then dTipoCls(key) = "UAR"   ' sobrepoe (declaracao explicita vence)
        Next i
    End If
End Sub

Private Sub CarregarTipoClassifAuto()
    ' TIPO (UC/COM) por familia derivado do CLS3 dominante do catalogo de
    ' materiais (MAT. UC/MAT. COM). Preenche familias fora da tabela curada;
    ' AddTipoCls nao sobrepoe entradas ja definidas (curadas tem prioridade).
    ' Tabela embutida dividida em 11 blocos (limite de tamanho de procedimento).
    CarregarTipoClassifAutoP1
    CarregarTipoClassifAutoP2
    CarregarTipoClassifAutoP3
    CarregarTipoClassifAutoP4
    CarregarTipoClassifAutoP5
    CarregarTipoClassifAutoP6
    CarregarTipoClassifAutoP7
    CarregarTipoClassifAutoP8
    CarregarTipoClassifAutoP9
    CarregarTipoClassifAutoP10
    CarregarTipoClassifAutoP11
End Sub

Private Sub CarregarTipoClassifAutoP1()
    AddTipoCls "ABRACADEIRA TRIPARTIDA DN700", "UC"
    AddTipoCls "ABRACADEIRA_PRESSAO", "COM"
    AddTipoCls "ABRANDADOR", "UC"
    AddTipoCls "ACESS FLANGE", "COM"
    AddTipoCls "ACESSORIO PERFILADO", "COM"
    AddTipoCls "ACESS_SOLAR", "COM"
    AddTipoCls "ACESS_TERMOC_MUFLA", "COM"
    AddTipoCls "ACIONADOR CHAV", "UC"
    AddTipoCls "ACS COMUNIC", "UC"
    AddTipoCls "ADAPTADOR", "COM"
    AddTipoCls "ADAPTADOR_TUBULACAO", "COM"
    AddTipoCls "ADESIVO (ANE)", "COM"
    AddTipoCls "AGITADOR", "COM"
    AddTipoCls "ALARME DE COMUNIC", "COM"
    AddTipoCls "ALICATE", "UC"
    AddTipoCls "ALINHADOR EIXO", "COM"
    AddTipoCls "ALINHADOR POLIA", "UC"
    AddTipoCls "ALINHADOR TUBO", "COM"
    AddTipoCls "AMORTECEDOR", "COM"
    AddTipoCls "AMORTECEDOR GRAMPO", "COM"
    AddTipoCls "AMPERIMETRO", "UC"
    AddTipoCls "ANALISADOR", "UC"
    AddTipoCls "ANALISADOR QUAL ENERGIA", "UC"
    AddTipoCls "ANALISADOR REG", "UC"
    AddTipoCls "ANEL BORRAC", "COM"
    AddTipoCls "ANEL PLAST", "COM"
    AddTipoCls "ANEL VEDACAO", "COM"
    AddTipoCls "ANEL_CONC", "COM"
    AddTipoCls "ANTENA", "UC"
    AddTipoCls "ANTENA SATELITE", "UC"
    AddTipoCls "ANTENA_CELULAR", "COM"
    AddTipoCls "ANTENA_RADIO", "UC"
    AddTipoCls "ANTIFURTO_REGULADOR", "COM"
    AddTipoCls "AP_ANALISADOR", "UC"
    AddTipoCls "AP_ULTRASOM", "UC"
    AddTipoCls "AQUECEDOR", "UC"
    AddTipoCls "AR CONDICIONADO", "UC"
    AddTipoCls "ARAME", "COM"
    AddTipoCls "ARAME_LISO", "COM"
    AddTipoCls "AREIA", "COM"
    AddTipoCls "ARGAMASSA", "COM"
    AddTipoCls "ARRUELA", "COM"
    AddTipoCls "ASFALTO", "COM"
    AddTipoCls "ASFALTO FRIO", "COM"
    AddTipoCls "ATERRO", "COM"
    AddTipoCls "AUTOTRAC", "UC"
    AddTipoCls "AUTOTRAFO", "UC"
    AddTipoCls "AUTOTRAFO POT", "UC"
    AddTipoCls "BACIA_CONTENCAO_QUIMICA", "UC"
    AddTipoCls "BANCADA HIDROMETRACAO", "COM"
    AddTipoCls "BANCO BATERIA", "UC"
    AddTipoCls "BANCO_CAPACTOR", "UC"
    AddTipoCls "BANDEJA", "COM"
    AddTipoCls "BANDEJA_OPTIC", "COM"
    AddTipoCls "BARRA", "COM"
    AddTipoCls "BARRA ESTRUT", "COM"
    AddTipoCls "BARRAMENTO", "UC"
    AddTipoCls "BARRAMENTO_BIFASICO", "UC"
    AddTipoCls "BARRA_COBRE", "UC"
    AddTipoCls "BARRA_SUPORTE", "UC"
    AddTipoCls "BASE", "COM"
    AddTipoCls "BASE BANC CAPAC", "UC"
    AddTipoCls "BASE CONC CMB", "UC"
    AddTipoCls "BASE CONC DISJ", "UC"
    AddTipoCls "BASE CONC TAMBOR", "COM"
    AddTipoCls "BASE RELE ILUM", "COM"
    AddTipoCls "BASE_ANTENA", "COM"
    AddTipoCls "BASE_FUS", "COM"
    AddTipoCls "BASE_PARA", "UC"
    AddTipoCls "BASTAO(OPEX)", "COM"
    AddTipoCls "BATERIA", "COM"
    AddTipoCls "BATERIA ALCALIN (ANE)", "COM"
    AddTipoCls "BATERIA SOLAR", "COM"
    AddTipoCls "BATERIAS", "COM"
    AddTipoCls "BEBEDOURO", "UC"
End Sub

Private Sub CarregarTipoClassifAutoP2()
    AddTipoCls "BETONEIRA", "UC"
    AddTipoCls "BGAN KIT LEGACY", "UC"
    AddTipoCls "BLOCO", "COM"
    AddTipoCls "BLOCO_TESTE", "COM"
    AddTipoCls "BLOQUEADOR CILINDRICO", "UC"
    AddTipoCls "BOBINA", "COM"
    AddTipoCls "BOBINA ELETRICA", "COM"
    AddTipoCls "BOBINA TF", "COM"
    AddTipoCls "BOMBA", "COM"
    AddTipoCls "BOMBA CENTRIF", "UC"
    AddTipoCls "BOMBA D'AGUA", "COM"
    AddTipoCls "BOMBA DAGUA", "COM"
    AddTipoCls "BOMBA DE ENCHER PNEU", "COM"
    AddTipoCls "BOMBA DE FLUIDOS", "UC"
    AddTipoCls "BOMBA DOSADORA", "UC"
    AddTipoCls "BOMBA FLUTUANTE", "UC"
    AddTipoCls "BOMBA PERISTALICA", "UC"
    AddTipoCls "BOMBA SUBMERSA", "UC"
    AddTipoCls "BOMBA SUCCAO", "UC"
    AddTipoCls "BOMBA VAC", "COM"
    AddTipoCls "BORNE", "COM"
    AddTipoCls "BRACO", "COM"
    AddTipoCls "BRITA", "COM"
    AddTipoCls "BRUCHA TF FORCA", "UC"
    AddTipoCls "BUCHA", "COM"
    AddTipoCls "BUCHA CAP", "UC"
    AddTipoCls "BUCHA DE TRAFO", "UC"
    AddTipoCls "BUCHA TF", "UC"
    AddTipoCls "BUCHA TF 138KV", "UC"
    AddTipoCls "CABINE BCAP", "COM"
    AddTipoCls "CABO", "COM"
    AddTipoCls "CABO ACO", "COM"
    AddTipoCls "CABO COAXIAL", "COM"
    AddTipoCls "CABO CU", "UC"
    AddTipoCls "CABO FIBRA", "COM"
    AddTipoCls "CABO FLEX", "COM"
    AddTipoCls "CABO LAN UTP", "COM"
    AddTipoCls "CABO PATCH", "COM"
    AddTipoCls "CABO SOLAR", "COM"
    AddTipoCls "CABOS_PARA_RAIOS", "COM"
    AddTipoCls "CABO_ACO_COBREADO", "UC"
    AddTipoCls "CABO_ELET", "COM"
    AddTipoCls "CABO_FLEXIVEL", "COM"
    AddTipoCls "CABO_RELE", "COM"
    AddTipoCls "CADEIA_ISOLADORES", "UC"
    AddTipoCls "CAIXA", "COM"
    AddTipoCls "CAIXA 2MED", "COM"
    AddTipoCls "CAIXA 3MED", "COM"
    AddTipoCls "CAIXA 4MED", "COM"
    AddTipoCls "CAIXA 6MED", "COM"
    AddTipoCls "CAIXA 8MED", "COM"
    AddTipoCls "CAIXA BLIND PDR", "COM"
    AddTipoCls "CAIXA CONTROL TF", "UC"
    AddTipoCls "CAIXA D'AGUA", "COM"
    AddTipoCls "CAIXA DE PROT SM", "COM"
    AddTipoCls "CAIXA FIBRA OPT", "COM"
    AddTipoCls "CAIXA LIG", "COM"
    AddTipoCls "CAIXA LIG BCAP", "COM"
    AddTipoCls "CAIXA LINK", "COM"
    AddTipoCls "CAIXA MED", "COM"
    AddTipoCls "CAIXA MEDIDOR", "COM"
    AddTipoCls "CAIXA PAINEL ELET", "COM"
    AddTipoCls "CAIXA PASSAGEM", "COM"
    AddTipoCls "CAIXA PROTETOR", "COM"
    AddTipoCls "CAIXA SIGFI", "COM"
    AddTipoCls "CAIXA TERM", "UC"
    AddTipoCls "CAIXA_PASSAGEM", "COM"
    AddTipoCls "CAMARA EXTINC", "UC"
    AddTipoCls "CAMARA UV", "COM"
    AddTipoCls "CAMERA_FILM", "UC"
    AddTipoCls "CAMERA_FILM (ANE)", "UC"
    AddTipoCls "CAMINHAO", "UC"
    AddTipoCls "CANALETA CONC", "UC"
    AddTipoCls "CANALETA_PERFURADA", "COM"
    AddTipoCls "CANETA", "COM"
End Sub

Private Sub CarregarTipoClassifAutoP3()
    AddTipoCls "CANIVETE", "COM"
    AddTipoCls "CANTONEIRA SIGFI", "COM"
    AddTipoCls "CAP PVC", "COM"
    AddTipoCls "CAPITEL CONC", "UC"
    AddTipoCls "CARGA INDEVIDA X", "UC"
    AddTipoCls "CARREGADOR_FLU", "UC"
    AddTipoCls "CARRETEL DE ACO", "UC"
    AddTipoCls "CATRACA (ANE)", "UC"
    AddTipoCls "CCM", "UC"
    AddTipoCls "CC_MATERIAL", "UC"
    AddTipoCls "CENTELHADOR", "COM"
    AddTipoCls "CH ATERR", "UC"
    AddTipoCls "CH SACA FU", "UC"
    AddTipoCls "CH SEC 138KV", "UC"
    AddTipoCls "CH SEC 145KV", "UC"
    AddTipoCls "CH SEC 69KV", "UC"
    AddTipoCls "CH SEC 69KV_SUPORTE", "UC"
    AddTipoCls "CH SEC OLEO", "UC"
    AddTipoCls "CH SEC TANDEM", "UC"
    AddTipoCls "CH SEC_SUPORTE", "UC"
    AddTipoCls "CH TF", "UC"
    AddTipoCls "CH VAC 3F", "UC"
    AddTipoCls "CHAPA", "COM"
    AddTipoCls "CHAPA MET", "COM"
    AddTipoCls "CHAVE", "COM"
    AddTipoCls "CHAVE ATS", "COM"
    AddTipoCls "CHAVE TE (FERRAMENTA)", "COM"
    AddTipoCls "CHAVE_VACUO", "UC"
    AddTipoCls "CHUMBADOR", "COM"
    AddTipoCls "CHUMBADOR_CB", "COM"
    AddTipoCls "CILINDRO GAS", "COM"
    AddTipoCls "CIMENTO", "COM"
    AddTipoCls "CJ BOMBA", "UC"
    AddTipoCls "CJ MOTOBOMBA", "COM"
    AddTipoCls "CJ REAGENT", "COM"
    AddTipoCls "CJ_RADIADOR TF", "UC"
    AddTipoCls "COLARINHO", "COM"
    AddTipoCls "COLARINHO PEAD DN200", "UC"
    AddTipoCls "COLETOR REG", "UC"
    AddTipoCls "COLUN MET", "UC"
    AddTipoCls "COLUNA BARRA", "COM"
    AddTipoCls "COLUNA CONC", "UC"
    AddTipoCls "COMPORTA", "UC"
    AddTipoCls "COMPORTA 400MM", "UC"
    AddTipoCls "COMPORTA DN1000", "COM"
    AddTipoCls "COMPORTA DN500", "COM"
    AddTipoCls "COMPORTA DUPLA", "UC"
    AddTipoCls "COMPORTA VERT", "UC"
    AddTipoCls "COMPRESSOR AR", "UC"
    AddTipoCls "COMUTADOR", "UC"
    AddTipoCls "COMUTADOR TF", "UC"
    AddTipoCls "CONCENTRADOR_DIGITAL", "UC"
    AddTipoCls "CONCHA OLHAL", "COM"
    AddTipoCls "COND ACO", "COM"
    AddTipoCls "COND FIBRA", "UC"
    AddTipoCls "COND FIBRA OPT", "UC"
    AddTipoCls "COND SUBAQUAT", "UC"
    AddTipoCls "CONDENSADOR", "UC"
    AddTipoCls "CONDULET", "COM"
    AddTipoCls "CONDULETE", "COM"
    AddTipoCls "CONE ANTIPOUSO", "COM"
    AddTipoCls "CONEC EMEND", "UC"
    AddTipoCls "CONECTOR DB9", "COM"
    AddTipoCls "CONECTOR_MACHO", "COM"
    AddTipoCls "CONEXAO FOFO", "UC"
    AddTipoCls "CONEX_CRUZETA", "COM"
    AddTipoCls "CONJ CUBICULO", "UC"
    AddTipoCls "CONJ CX CPU", "COM"
    AddTipoCls "CONJ MED", "UC"
    AddTipoCls "CONJ PLACA", "UC"
    AddTipoCls "CONJ SIGFI", "UC"
    AddTipoCls "CONJ_CAD_ISOL", "UC"
    AddTipoCls "CONJ_POSTE ARMADO", "UC"
    AddTipoCls "CONTAINER", "COM"
    AddTipoCls "CONTATO", "UC"
End Sub

Private Sub CarregarTipoClassifAutoP4()
    AddTipoCls "CONTATO AUX", "COM"
    AddTipoCls "CONTATOR", "COM"
    AddTipoCls "CONTR ELETN", "UC"
    AddTipoCls "CONTRA VIGA", "COM"
    AddTipoCls "CONTRAVENTO LD", "COM"
    AddTipoCls "CONTROL VAZ FLUTUADOR", "UC"
    AddTipoCls "CONTROL. NV DE AGUA", "COM"
    AddTipoCls "CONTROLADOR_SOLAR", "UC"
    AddTipoCls "CONTROLE B_CAP", "COM"
    AddTipoCls "CONTROLE RELIG", "UC"
    AddTipoCls "CONVERSOR", "UC"
    AddTipoCls "CORDAO FIBRA", "COM"
    AddTipoCls "CORREIA", "COM"
    AddTipoCls "CORRENTE VALETADEIRA", "UC"
    AddTipoCls "CPU DA CP", "COM"
    AddTipoCls "CPU DA CS", "UC"
    AddTipoCls "CROMATOGRAFO", "UC"
    AddTipoCls "CRONOMETRO", "COM"
    AddTipoCls "CUBICULO", "UC"
    AddTipoCls "CURVA", "COM"
    AddTipoCls "CURVA 30º DN500", "COM"
    AddTipoCls "CURVA 45° DN800", "UC"
    AddTipoCls "CURVA 45º DN100", "COM"
    AddTipoCls "CURVA 45º DN200", "COM"
    AddTipoCls "CURVA 45º DN600", "UC"
    AddTipoCls "CURVA 90º DN1000", "UC"
    AddTipoCls "CURVA 90º DN150", "COM"
    AddTipoCls "CURVA 90º DN200", "COM"
    AddTipoCls "CURVA 90º DN300", "COM"
    AddTipoCls "CURVA 90º DN400", "COM"
    AddTipoCls "CURVA 90º DN600", "UC"
    AddTipoCls "CURVA 90º DN8000", "UC"
    AddTipoCls "CURVA 90º DN900", "UC"
    AddTipoCls "CURVA_STAND", "COM"
    AddTipoCls "CX BLIND", "COM"
    AddTipoCls "CX BLIND CMB", "COM"
    AddTipoCls "CX PAPELAO", "COM"
    AddTipoCls "CX PASS CONC_ARMADO", "COM"
    AddTipoCls "CX PDR HIDROM", "COM"
    AddTipoCls "CX PORTA DISJ", "COM"
    AddTipoCls "DATALOGGER", "COM"
    AddTipoCls "DEFENSA", "UC"
    AddTipoCls "DEFENSA MET", "UC"
    AddTipoCls "DESKTOP", "COM"
    AddTipoCls "DETECTOR_FALTA", "UC"
    AddTipoCls "DISJ", "COM"
    AddTipoCls "DISJ CUBICULO", "UC"
    AddTipoCls "DISJ MOTOR", "COM"
    AddTipoCls "DISJUNTOR", "UC"
    AddTipoCls "DISJUNTOR_AT", "UC"
    AddTipoCls "DISJUNTOR_BIP", "COM"
    AddTipoCls "DISJ_TERMOMAG", "COM"
    AddTipoCls "DISPLAY_LCD", "COM"
    AddTipoCls "DISPOSITIVO_OB", "UC"
    AddTipoCls "DISTRIBUIDOR DIO", "UC"
    AddTipoCls "DISTRIBUIDOR OPTICO", "UC"
    AddTipoCls "DOSADOR CLORADOR", "UC"
    AddTipoCls "DOSADORA HIDROG", "UC"
    AddTipoCls "DRIVER", "UC"
    AddTipoCls "DUTO PEAD - RAMAL", "UC"
    AddTipoCls "DUTO RETRATIL", "UC"
    AddTipoCls "ECAPORADOR", "UC"
    AddTipoCls "EIXO ACIONAMENTO", "UC"
    AddTipoCls "EIXO BOMBA", "UC"
    AddTipoCls "EIXO BOMBAMERSA", "COM"
    AddTipoCls "EIXO INTERMED", "UC"
    AddTipoCls "EIXO REGULAG", "UC"
    AddTipoCls "EIXO TRANSM", "UC"
    AddTipoCls "ELETROCALHA", "COM"
    AddTipoCls "ELETROCENTRO", "UC"
    AddTipoCls "ELO", "COM"
    AddTipoCls "ELO FUS", "COM"
    AddTipoCls "EMPILHADEIRA", "UC"
    AddTipoCls "ENCOD ABSOLUTO", "UC"
    AddTipoCls "ENGRAXADEIRA", "UC"
End Sub

Private Sub CarregarTipoClassifAutoP5()
    AddTipoCls "ENVELOPE MEDIDOR", "COM"
    AddTipoCls "ENVELOPE PLAST", "COM"
    AddTipoCls "EQP_LABORATORIO", "COM"
    AddTipoCls "ESCADA", "UC"
    AddTipoCls "ESCADA PEDAROLA", "COM"
    AddTipoCls "ESCORA ARM", "UC"
    AddTipoCls "ESFERA SINALIZ", "COM"
    AddTipoCls "ESMETILHADEIRA", "UC"
    AddTipoCls "ESPUMA POLIURET", "COM"
    AddTipoCls "ESTABILIZADOR COLUNAMERSA", "COM"
    AddTipoCls "ESTRU MET", "UC"
    AddTipoCls "ESTRUT ANC", "COM"
    AddTipoCls "ESTRUT SIGFI", "UC"
    AddTipoCls "ESTRUT SUSP", "COM"
    AddTipoCls "ESTRUTURA SUPORTE", "COM"
    AddTipoCls "ESTRUT_BCAP", "COM"
    AddTipoCls "ESTRUT_LD", "UC"
    AddTipoCls "ESTUFA", "UC"
    AddTipoCls "EST_GATEWAY", "UC"
    AddTipoCls "ETIQUETA MED", "COM"
    AddTipoCls "EXPAN ENTRAD", "UC"
    AddTipoCls "EXTENSAO OTICA", "COM"
    AddTipoCls "EXTENSAO_MONOFIBRA", "COM"
    AddTipoCls "EXTENSAO_TORRE", "COM"
    AddTipoCls "EXTIN INC", "UC"
    AddTipoCls "EXTRATOR", "UC"
    AddTipoCls "FERRO SOLDA (ANE)", "COM"
    AddTipoCls "FERROMAGNETICO", "UC"
    AddTipoCls "FILTRO DE AR", "COM"
    AddTipoCls "FIO COBRE", "COM"
    AddTipoCls "FIREWALL", "UC"
    AddTipoCls "FITA_IS", "COM"
    AddTipoCls "FIXACAO", "COM"
    AddTipoCls "FLANGE", "UC"
    AddTipoCls "FLANGE DN400", "COM"
    AddTipoCls "FONTE ALIM", "UC"
    AddTipoCls "FONTE NOTEBOOK", "UC"
    AddTipoCls "FORMULARIO", "COM"
    AddTipoCls "FORTISWITCH", "COM"
    AddTipoCls "FRIGOBAR", "UC"
    AddTipoCls "FURADEIRA", "UC"
    AddTipoCls "FUSIVEL", "COM"
    AddTipoCls "FUSIVEL CRT", "COM"
    AddTipoCls "GABARITO_TORRE", "COM"
    AddTipoCls "GARRAFA_RELI", "UC"
    AddTipoCls "GAS_INCOLOR", "COM"
    AddTipoCls "GATEWAY", "UC"
    AddTipoCls "GAXETA", "COM"
    AddTipoCls "GERAD PORTAT", "UC"
    AddTipoCls "GERADOR", "UC"
    AddTipoCls "GERADOR (FERRAMENTAL)", "UC"
    AddTipoCls "GERADOR DE CLORO", "UC"
    AddTipoCls "GPS", "UC"
    AddTipoCls "GRELHA", "COM"
    AddTipoCls "GRELHA_TORRE", "COM"
    AddTipoCls "GRUPO", "UC"
    AddTipoCls "GRUPO GERADOR SIGFI", "UC"
    AddTipoCls "GUARDA CORPO", "COM"
    AddTipoCls "GUARNICAO_BUCHA", "COM"
    AddTipoCls "HASTE", "COM"
    AddTipoCls "HIDROEJETOR", "UC"
    AddTipoCls "HIDROGEROX", "UC"
    AddTipoCls "HIDROMETRO", "UC"
    AddTipoCls "IED", "UC"
    AddTipoCls "IMPERMEABILIZANTE", "COM"
    AddTipoCls "IMPRESSORA ETIQUETA", "UC"
    AddTipoCls "IMPRESSORA_PLOTTER", "UC"
    AddTipoCls "IMPRESSORA_TERMIC", "COM"
    AddTipoCls "INDIC MICRO", "UC"
    AddTipoCls "INDICADOR OLEO", "UC"
    AddTipoCls "INDICADOR TF", "UC"
    AddTipoCls "INTERRUPTOR ODS", "COM"
    AddTipoCls "INTERRUPTOR TOM", "COM"
    AddTipoCls "INVERSOR", "UC"
    AddTipoCls "INVERSOR SIGFI", "UC"
End Sub

Private Sub CarregarTipoClassifAutoP6()
    AddTipoCls "ISOL PEDESTAL", "UC"
    AddTipoCls "ISOL PEDESTAL_SUPORTE", "UC"
    AddTipoCls "ISOLADOR", "COM"
    AddTipoCls "ISOLADOR (SUCATA)", "COM"
    AddTipoCls "ISOLADOR CAST", "COM"
    AddTipoCls "ISOLADOR MT", "COM"
    AddTipoCls "ISOLADOR PEDESTAL", "COM"
    AddTipoCls "ISOLADOR_AT", "UC"
    AddTipoCls "ISOLADOR_EPOXI", "COM"
    AddTipoCls "JAR TEST", "UC"
    AddTipoCls "JG JUNTAS KSB", "COM"
    AddTipoCls "JG RETENTOR", "UC"
    AddTipoCls "JOELHO", "COM"
    AddTipoCls "JUNCAO", "COM"
    AddTipoCls "JUNTA", "COM"
    AddTipoCls "JUNTA DESMONT. DN1000", "UC"
    AddTipoCls "JUNTA DESMONT. DN400", "UC"
    AddTipoCls "JUNTA DESMONT. DN500", "UC"
    AddTipoCls "JUNTA DESMONT. DN700", "UC"
    AddTipoCls "JUNTA DESMONT. DN800", "UC"
    AddTipoCls "JUNTA GIBAULT", "COM"
    AddTipoCls "JUNTA GIBAULT DN600", "COM"
    AddTipoCls "JUNTA KSB", "COM"
    AddTipoCls "KIT ATER", "COM"
    AddTipoCls "KIT BUSHLINK", "COM"
    AddTipoCls "KIT CALCO", "UC"
    AddTipoCls "KIT CAVALETE", "COM"
    AddTipoCls "KIT DISP_MED", "COM"
    AddTipoCls "KIT ELETR SIGFI", "COM"
    AddTipoCls "KIT HIDRANTE", "UC"
    AddTipoCls "KIT PLPT", "COM"
    AddTipoCls "KIT T TABLET", "UC"
    AddTipoCls "KIT_CONTROLE", "UC"
    AddTipoCls "KIT_PLACAS", "COM"
    AddTipoCls "LACRE", "COM"
    AddTipoCls "LAJE_PREFABRICADA", "UC"
    AddTipoCls "LAMPADA", "COM"
    AddTipoCls "LAMPADA_SINALIZ", "COM"
    AddTipoCls "LAVADORA", "UC"
    AddTipoCls "LIC S SOFTW", "UC"
    AddTipoCls "LIMITADOR_CORRENTE", "COM"
    AddTipoCls "LNB AUTO", "UC"
    AddTipoCls "LOTE TERRENO", "UC"
    AddTipoCls "LUMINARIA", "UC"
    AddTipoCls "LUVA", "COM"
    AddTipoCls "LUVA ELETROD", "COM"
    AddTipoCls "LUVA FERRO DN150", "COM"
    AddTipoCls "LUVA FERRO DN800", "UC"
    AddTipoCls "LUVA PVC", "COM"
    AddTipoCls "LUVA ULTRALINK", "UC"
    AddTipoCls "LUVA_ALUMINIO", "COM"
    AddTipoCls "MADEIRA", "COM"
    AddTipoCls "MALHA ATERRAMENTO", "UC"
    AddTipoCls "MALHA_DIFUS", "COM"
    AddTipoCls "MANGEUIRA", "COM"
    AddTipoCls "MANGUEIRA", "COM"
    AddTipoCls "MANGUEIRA (ANE)", "COM"
    AddTipoCls "MANGUEIRA INDUSTRIAL", "COM"
    AddTipoCls "MANILHA", "COM"
    AddTipoCls "MANOMETRO", "COM"
    AddTipoCls "MANOMETRO_BAR", "UC"
    AddTipoCls "MAQUINA", "UC"
    AddTipoCls "MAQUINA ELETROFUSAO", "UC"
    AddTipoCls "MAQ_SOLDA", "UC"
    AddTipoCls "MARRETA", "UC"
    AddTipoCls "MASSA ISOLANTE", "COM"
    AddTipoCls "MASTRO ANTENA", "COM"
    AddTipoCls "MAT.COM", "COM"
    AddTipoCls "MATRIZ_DIODO", "COM"
    AddTipoCls "MEDIDOR DE VAZAO", "UC"
    AddTipoCls "MEDIDOR ULTRASONICO", "UC"
    AddTipoCls "MEDIDOR VAZAO", "UC"
    AddTipoCls "MEDIDOR_FRONTEIRA", "UC"
    AddTipoCls "MED_VAZAO", "UC"
    AddTipoCls "MESA SE", "UC"
End Sub

Private Sub CarregarTipoClassifAutoP7()
    AddTipoCls "MICROCOMPUTADOR", "UC"
    AddTipoCls "MICROHMIMETRO", "UC"
    AddTipoCls "MICROSWITCH INVERS", "UC"
    AddTipoCls "MINI TECLADO", "COM"
    AddTipoCls "MINI_DISJ", "COM"
    AddTipoCls "MINI_DISJ BT", "COM"
    AddTipoCls "MISTURADOR RAP", "UC"
    AddTipoCls "MOD CONCENT", "UC"
    AddTipoCls "MODEM", "UC"
    AddTipoCls "MODULO", "COM"
    AddTipoCls "MODULO CONTROLE", "UC"
    AddTipoCls "MODULO DIODO", "COM"
    AddTipoCls "MODULO REMOT", "UC"
    AddTipoCls "MODULO REMOTO", "UC"
    AddTipoCls "MODULO SOLAR", "UC"
    AddTipoCls "MODULO TRANSCEPTOR", "UC"
    AddTipoCls "MODULO_TROC", "UC"
    AddTipoCls "MOD_RETIFICADOR", "UC"
    AddTipoCls "MOD_VENTILADOR", "UC"
    AddTipoCls "MONITOR", "UC"
    AddTipoCls "MONITOR GAS", "COM"
    AddTipoCls "MONITOR PC", "COM"
    AddTipoCls "MONITOR TEMP OLEO", "UC"
    AddTipoCls "MONITOR TEMP. TF", "UC"
    AddTipoCls "MONITOR_AUTOM", "UC"
    AddTipoCls "MOTO VENTILADOR", "UC"
    AddTipoCls "MOTOBOMBA", "COM"
    AddTipoCls "MOTOBOMBA CENTMERSA", "UC"
    AddTipoCls "MOTOBOMBA SUBMMERSA", "UC"
    AddTipoCls "MOTOPODA", "UC"
    AddTipoCls "MOTOR CC", "COM"
    AddTipoCls "MOTOSSERRA", "UC"
    AddTipoCls "MOTOVENTILADOR", "COM"
    AddTipoCls "MOTOVENTILADOR TF", "COM"
    AddTipoCls "MOUSE", "COM"
    AddTipoCls "NIPLE", "COM"
    AddTipoCls "NO-BREAK", "UC"
    AddTipoCls "NOBREAK", "COM"
    AddTipoCls "NOTEBOOK", "UC"
    AddTipoCls "OLEO ISOL", "COM"
    AddTipoCls "OLEO LUB", "COM"
    AddTipoCls "OLEO MINERAL", "COM"
    AddTipoCls "PAINEL", "UC"
    AddTipoCls "PAINEL CC", "UC"
    AddTipoCls "PAINEL COMANDO", "COM"
    AddTipoCls "PAINEL CSA", "UC"
    AddTipoCls "PAINEL ELET", "UC"
    AddTipoCls "PAINEL EXCLUSIVO", "UC"
    AddTipoCls "PAINEL FOTO BP", "COM"
    AddTipoCls "PAINEL FRONTAL", "COM"
    AddTipoCls "PAINEL MEDICAO", "UC"
    AddTipoCls "PAINEL PROT", "UC"
    AddTipoCls "PAINEL SERV AUX", "UC"
    AddTipoCls "PAINEL TELECOM", "UC"
    AddTipoCls "PAINEL_AUTO", "UC"
    AddTipoCls "PAINEL_AUTOMACAO", "UC"
    AddTipoCls "PAINEL_ELET", "UC"
    AddTipoCls "PAINEL_MED", "UC"
    AddTipoCls "PAINEL_PART", "UC"
    AddTipoCls "PAINEL_QGBT", "UC"
    AddTipoCls "PAINEL_SERVIDOR", "UC"
    AddTipoCls "PAINEL_TELEM", "UC"
    AddTipoCls "PAPEL QUIMICO", "COM"
    AddTipoCls "PAPEL TOALHA", "COM"
    AddTipoCls "PARA RAIO", "UC"
    AddTipoCls "PARA RAIO AT", "UC"
    AddTipoCls "PARA RAIO_SUPORTE", "UC"
    AddTipoCls "PARAF", "COM"
    AddTipoCls "PARAFUSO_SEXT", "COM"
    AddTipoCls "PARALELISMO SINCRONO TF", "UC"
    AddTipoCls "PARECE CORTA FOGO", "UC"
    AddTipoCls "PATCH PANEL", "COM"
    AddTipoCls "PATCH UTP", "COM"
    AddTipoCls "PDA_RELE", "UC"
    AddTipoCls "PE DE TORRE", "COM"
End Sub

Private Sub CarregarTipoClassifAutoP8()
    AddTipoCls "PECA_RET", "COM"
    AddTipoCls "PEDRA_BRITADA", "COM"
    AddTipoCls "PEDREGULHO/SEIXO", "COM"
    AddTipoCls "PERFIL", "COM"
    AddTipoCls "PERFIL_ESTRUT", "COM"
    AddTipoCls "PILAR_ESTACA", "UC"
    AddTipoCls "PILHA", "COM"
    AddTipoCls "PLACA", "COM"
    AddTipoCls "PLACA DE REDE", "UC"
    AddTipoCls "PLACA DE REDUCAO", "UC"
    AddTipoCls "PLACA ELETRONICA", "COM"
    AddTipoCls "PLACA MONTG SIGFI", "COM"
    AddTipoCls "PLACA SINALIZACAO", "COM"
    AddTipoCls "PLACA TERMOPLASTICA", "UC"
    AddTipoCls "POLO DISJ", "UC"
    AddTipoCls "POLYWALL", "UC"
    AddTipoCls "PONTALETE", "COM"
    AddTipoCls "PORCA", "COM"
    AddTipoCls "PORCA_SEXTAVADA", "COM"
    AddTipoCls "POSTE", "UC"
    AddTipoCls "POSTE MET", "UC"
    AddTipoCls "POSTE RD (SUCATA)", "UC"
    AddTipoCls "POSTE TELEFONICO", "UC"
    AddTipoCls "POSTE TELESCOPICO", "UC"
    AddTipoCls "POSTE/BARRA_SUPORTE", "UC"
    AddTipoCls "POSTE_MET_LUM", "UC"
    AddTipoCls "PREGO", "COM"
    AddTipoCls "PRENSA-FIO", "COM"
    AddTipoCls "PREPARADOR POLIMERO", "UC"
    AddTipoCls "PROCESSADOR", "UC"
    AddTipoCls "PROD_QUIMICO", "COM"
    AddTipoCls "PROLONGADOR", "COM"
    AddTipoCls "PROT", "COM"
    AddTipoCls "PROTETOR", "UC"
    AddTipoCls "PROTETOR DPS", "UC"
    AddTipoCls "PROTETOR ILUM", "UC"
    AddTipoCls "QGBT", "UC"
    AddTipoCls "QGS", "UC"
    AddTipoCls "QUADRO DE COMANDO", "UC"
    AddTipoCls "QUADRO DT", "UC"
    AddTipoCls "QUADRO ELET", "COM"
    AddTipoCls "QUADRO_COMANDO", "COM"
    AddTipoCls "QUADRO_DISJ", "COM"
    AddTipoCls "QUADRO_ELETR", "UC"
    AddTipoCls "QUADRO_MED_BLIND", "COM"
    AddTipoCls "QUADRO_MET", "COM"
    AddTipoCls "QUADRO_METALICO", "UC"
    AddTipoCls "RACK", "COM"
    AddTipoCls "RACK_CABO", "COM"
    AddTipoCls "RACK_SERVIDOR", "UC"
    AddTipoCls "RADIADOR_TRAFO", "COM"
    AddTipoCls "RADIO", "COM"
    AddTipoCls "REATOR", "UC"
    AddTipoCls "REATOR LAMP", "COM"
    AddTipoCls "REATOR POT", "UC"
    AddTipoCls "REBITE", "COM"
    AddTipoCls "RECEPTOR", "UC"
    AddTipoCls "REDUCAO", "COM"
    AddTipoCls "REDUCAO CONCENTRICA", "UC"
    AddTipoCls "REFIGERADO", "UC"
    AddTipoCls "REFLETOR", "UC"
    AddTipoCls "REGISTRO", "COM"
    AddTipoCls "REGISTRO DN500", "UC"
    AddTipoCls "REGISTRO GAVETA", "COM"
    AddTipoCls "REGUA BORNE", "COM"
    AddTipoCls "REGUA TOMADA", "COM"
    AddTipoCls "REGUA_TOMADA", "COM"
    AddTipoCls "REGULADOR SUCATA", "UC"
    AddTipoCls "REGULADOR_TENSAO", "UC"
    AddTipoCls "RELE AUX", "COM"
    AddTipoCls "RELE BCO CAP", "UC"
    AddTipoCls "RELE CONTATO", "COM"
    AddTipoCls "RELE DIGITAL", "UC"
    AddTipoCls "RELE PROT", "UC"
    AddTipoCls "RELE REFLETOR", "COM"
End Sub

Private Sub CarregarTipoClassifAutoP9()
    AddTipoCls "RELE REG", "UC"
    AddTipoCls "RELE REGULADOR", "UC"
    AddTipoCls "RELE SUPERV_TENSAO", "COM"
    AddTipoCls "RELE TF", "UC"
    AddTipoCls "RELE_RAPIDO", "COM"
    AddTipoCls "REPETIDORA", "UC"
    AddTipoCls "REPETIDORA RADIO", "UC"
    AddTipoCls "RESERVATORIO", "UC"
    AddTipoCls "RESERVATORIO 500L", "COM"
    AddTipoCls "RESERVATORIO TF", "COM"
    AddTipoCls "RESI ELETR", "UC"
    AddTipoCls "RESISTOR", "COM"
    AddTipoCls "RETENTOR", "COM"
    AddTipoCls "RETIFICADOR", "UC"
    AddTipoCls "RETIFICADOR + BC_BATERIA", "UC"
    AddTipoCls "RETROESCAVADEIRA", "UC"
    AddTipoCls "RISCO", "UC"
    AddTipoCls "ROLAMENTO", "UC"
    AddTipoCls "ROTAMETRO", "UC"
    AddTipoCls "ROTEADOR", "UC"
    AddTipoCls "ROTOR DE SUCCAO", "UC"
    AddTipoCls "ROTOR P/ BOMBAS", "UC"
    AddTipoCls "RTAC", "UC"
    AddTipoCls "SATURADOR", "UC"
    AddTipoCls "SE MOVEL", "UC"
    AddTipoCls "SECCIONADOR", "UC"
    AddTipoCls "SENSO OPT", "UC"
    AddTipoCls "SENSO TEMP", "UC"
    AddTipoCls "SENSOR", "UC"
    AddTipoCls "SENSOR ARCO ELETRICO", "UC"
    AddTipoCls "SENSOR PT100", "COM"
    AddTipoCls "SENSOR RESISTVO", "UC"
    AddTipoCls "SEPARADOR DE HIDROGENIO", "UC"
    AddTipoCls "SERVIDOR", "UC"
    AddTipoCls "SIGFI_PAINEL SOLAR", "UC"
    AddTipoCls "SILICAGEL", "COM"
    AddTipoCls "SINALIZADOR SOLAR", "COM"
    AddTipoCls "SINALIZ_AVIFAUNA", "COM"
    AddTipoCls "SINCRONIZADOR", "UC"
    AddTipoCls "SIS. DE VIGILANCIA", "UC"
    AddTipoCls "SIS. FOTOV BOMBEA_AGUA", "UC"
    AddTipoCls "SIST FILTR", "UC"
    AddTipoCls "SIST FLUT", "UC"
    AddTipoCls "SIST ILUM BASE RELE", "COM"
    AddTipoCls "SIST ILUM REFLETOR", "COM"
    AddTipoCls "SIST ILUM RELE", "COM"
    AddTipoCls "SIST ILUM SUPORTE", "COM"
    AddTipoCls "SIST ILUMINACAO", "UC"
    AddTipoCls "SIST MED", "UC"
    AddTipoCls "SIST RETIFICACAO", "UC"
    AddTipoCls "SIST TIA", "UC"
    AddTipoCls "SISTEMA DE ALIMENTACAO", "UC"
    AddTipoCls "SIST_RETIFICADOR", "UC"
    AddTipoCls "SOFTWARE", "UC"
    AddTipoCls "SOLDA", "UC"
    AddTipoCls "SONDA", "COM"
    AddTipoCls "SOPRADOR", "UC"
    AddTipoCls "SSD", "UC"
    AddTipoCls "STAND COMERCIAL", "UC"
    AddTipoCls "STUB MET", "UC"
    AddTipoCls "SUCATA MET", "COM"
    AddTipoCls "SUPORTE ATENA", "COM"
    AddTipoCls "SUPORTE CH SEC", "COM"
    AddTipoCls "SUPORTE GACHO", "COM"
    AddTipoCls "SUPORTE HIDROMETRO", "COM"
    AddTipoCls "SUPORTE ISOL PED", "UC"
    AddTipoCls "SUPORTE POSTE", "COM"
    AddTipoCls "SUPORTE TC", "COM"
    AddTipoCls "SUPORTE_ISOLADOR", "COM"
    AddTipoCls "SUPORTE_TC", "UC"
    AddTipoCls "SUSPENSAO_TORRE", "COM"
    AddTipoCls "SWITCH", "UC"
    AddTipoCls "TAMPA", "COM"
    AddTipoCls "TAMPA ESGOTO", "COM"
    AddTipoCls "TAMPAO", "COM"
End Sub

Private Sub CarregarTipoClassifAutoP10()
    AddTipoCls "TAMPA_CANALETA", "COM"
    AddTipoCls "TANQUE", "COM"
    AddTipoCls "TANQUE AGITADOR", "COM"
    AddTipoCls "TAPETE (ANE)", "COM"
    AddTipoCls "TAPETE ISOLANTE", "COM"
    AddTipoCls "TC", "UC"
    AddTipoCls "TC_CX CONCENT", "COM"
    AddTipoCls "TC_MF", "UC"
    AddTipoCls "TE 90° DN200", "UC"
    AddTipoCls "TE 90° DN450", "UC"
    AddTipoCls "TE 90° DN500", "UC"
    AddTipoCls "TE 90° DN800", "UC"
    AddTipoCls "TE DN150", "COM"
    AddTipoCls "TE FLANGE DN700", "UC"
    AddTipoCls "TE FLANGE DN800", "UC"
    AddTipoCls "TE REUCAO DN1000", "UC"
    AddTipoCls "TE REUCAO DN150/75", "COM"
    AddTipoCls "TE REUCAO DN250/100", "COM"
    AddTipoCls "TE REUCAO DN500", "UC"
    AddTipoCls "TE REUCAO DN600", "UC"
    AddTipoCls "TE REUCAO DN700", "UC"
    AddTipoCls "TE REUCAO DN800", "UC"
    AddTipoCls "TECLADO", "COM"
    AddTipoCls "TEE", "COM"
    AddTipoCls "TEE REDUCAO", "COM"
    AddTipoCls "TELA", "COM"
    AddTipoCls "TELA_MALHA", "COM"
    AddTipoCls "TELEFONE", "UC"
    AddTipoCls "TELEFONE MOV/CEL", "COM"
    AddTipoCls "TELEFONE VOIP", "COM"
    AddTipoCls "TERMINAL", "UC"
    AddTipoCls "TERMINAL COBRE", "COM"
    AddTipoCls "TERMINAL MUFLA", "COM"
    AddTipoCls "TERMINAL PINO", "COM"
    AddTipoCls "TERMINAL SERVER", "UC"
    AddTipoCls "TERMINAL_COMPR", "COM"
    AddTipoCls "TERMINAL_ILH", "COM"
    AddTipoCls "TERMINAL_OLHAL", "COM"
    AddTipoCls "TERMINAL_PRE", "COM"
    AddTipoCls "TERMINAL_PRESSAO", "COM"
    AddTipoCls "TERMINAL_TUBULAR", "COM"
    AddTipoCls "TERMO COMUM", "COM"
    AddTipoCls "TERMO NOTIFI", "COM"
    AddTipoCls "TERMO OCORR", "COM"
    AddTipoCls "TERMO REGULAR", "COM"
    AddTipoCls "TERMO VISTORIA", "COM"
    AddTipoCls "TERMOCAPTADOR", "COM"
    AddTipoCls "TERMORESISTENCIA", "COM"
    AddTipoCls "TERMOSTATO", "COM"
    AddTipoCls "TERMOVISOR", "UC"
    AddTipoCls "TERRENO", "UC"
    AddTipoCls "TERROMETRO DIG", "UC"
    AddTipoCls "TIJOLO", "COM"
    AddTipoCls "TINTA", "COM"
    AddTipoCls "TINTA (ANE)", "COM"
    AddTipoCls "TINTA SPRAY", "COM"
    AddTipoCls "TOMADA ODS", "COM"
    AddTipoCls "TOMADAS_DE", "COM"
    AddTipoCls "TOMADA_DUPLA", "COM"
    AddTipoCls "TORA", "COM"
    AddTipoCls "TORRE", "UC"
    AddTipoCls "TP", "UC"
    AddTipoCls "TP_CX CONCENT", "COM"
    AddTipoCls "TP_INDUTIVO", "UC"
    AddTipoCls "TP_TC_SUPORT ACO", "UC"
    AddTipoCls "TP_TC_SUPORT CONC", "UC"
    AddTipoCls "TR", "UC"
    AddTipoCls "TRAFO COMANDO", "COM"
    AddTipoCls "TRAFO CONV/RETIF", "COM"
    AddTipoCls "TRAFO SE", "UC"
    AddTipoCls "TRAFO_MOVEL", "UC"
    AddTipoCls "TRANS BALANC", "UC"
    AddTipoCls "TRANSCEIVER", "COM"
    AddTipoCls "TRANSCEPTOR", "UC"
    AddTipoCls "TRANSDUTOR", "UC"
End Sub

Private Sub CarregarTipoClassifAutoP11()
    AddTipoCls "TRANSMISSOR", "UC"
    AddTipoCls "TRILHO MET", "COM"
    AddTipoCls "TRILHOS", "COM"
    AddTipoCls "TROCADOR DE CALOR", "COM"
    AddTipoCls "TSA", "UC"
    AddTipoCls "TSA_SUPORTE", "UC"
    AddTipoCls "TUBO", "COM"
    AddTipoCls "TUBO 25X6MM", "COM"
    AddTipoCls "TUBO ACO", "COM"
    AddTipoCls "TUBO COLETOR", "COM"
    AddTipoCls "TUBO DEFOFO", "COM"
    AddTipoCls "TUBO DEFOFO DN150", "COM"
    AddTipoCls "TUBO DEFOFO DN200", "COM"
    AddTipoCls "TUBO DEFOFO DN250", "COM"
    AddTipoCls "TUBO DN100", "UC"
    AddTipoCls "TUBO DN200", "UC"
    AddTipoCls "TUBO DN400", "UC"
    AddTipoCls "TUBO DN600", "UC"
    AddTipoCls "TUBO DN700", "UC"
    AddTipoCls "TUBO DN800", "UC"
    AddTipoCls "TUBO ELEVACAO", "COM"
    AddTipoCls "TUBO PBA", "COM"
    AddTipoCls "TUBO PEAD", "COM"
    AddTipoCls "TUBO PVC", "COM"
    AddTipoCls "TUBO PVC DN50/60", "COM"
    AddTipoCls "TUBO PVC DN75/85", "COM"
    AddTipoCls "TUBO RIGIDO", "COM"
    AddTipoCls "TURBIDIMETRO", "UC"
    AddTipoCls "TV", "UC"
    AddTipoCls "TV LED", "UC"
    AddTipoCls "UCC", "UC"
    AddTipoCls "UNIAO_ELETRODUTO", "COM"
    AddTipoCls "UNID CONDENSADORA", "UC"
    AddTipoCls "UNID EVAPORADORA", "UC"
    AddTipoCls "UNIDA TEST", "UC"
    AddTipoCls "UNIDADE CONTROLE", "UC"
    AddTipoCls "UNID_CONC_COMUM", "UC"
    AddTipoCls "UNID_DE_CONTROLE", "UC"
    AddTipoCls "UPDG", "UC"
    AddTipoCls "UPS", "UC"
    AddTipoCls "URTT", "UC"
    AddTipoCls "UTR", "UC"
    AddTipoCls "VALETADEIRA", "UC"
    AddTipoCls "VALVULA", "COM"
    AddTipoCls "VALVULA BORBOLETA", "UC"
    AddTipoCls "VALVULA DN1", "COM"
    AddTipoCls "VALVULA DN2", "COM"
    AddTipoCls "VALVULA DN200", "COM"
    AddTipoCls "VALVULA DN3/4", "COM"
    AddTipoCls "VALVULA DN300", "UC"
    AddTipoCls "VALVULA DN32", "COM"
    AddTipoCls "VALVULA DN350", "UC"
    AddTipoCls "VALVULA DN400", "UC"
    AddTipoCls "VALVULA ESFERA", "COM"
    AddTipoCls "VALVULA GAVETA", "UC"
    AddTipoCls "VALVULA PRESSAO", "COM"
    AddTipoCls "VALVULA RETENTORA", "UC"
    AddTipoCls "VARA MANOBRA", "COM"
    AddTipoCls "VARETA", "COM"
    AddTipoCls "VARETAS", "COM"
    AddTipoCls "VASSOURA", "COM"
    AddTipoCls "VEDACAO", "COM"
    AddTipoCls "VEDANTE", "COM"
    AddTipoCls "VEIC FROT", "UC"
    AddTipoCls "VEIC_FROT", "UC"
    AddTipoCls "VENTILADOR", "UC"
    AddTipoCls "VENTOSA", "COM"
    AddTipoCls "VERGALHAO", "COM"
    AddTipoCls "VERIFICAR", "UC"
    AddTipoCls "VIGA", "UC"
    AddTipoCls "VIGA_CONC", "UC"
    AddTipoCls "VOLTIMETRO", "UC"
    AddTipoCls "ZEOLITA", "COM"
    AddTipoCls "ZZ CA CADASTRO", "UC"
End Sub

Private Sub AddTipoCls(ByVal fam As String, ByVal tipo As String)
    If dTipoCls Is Nothing Then Set dTipoCls = CreateObject("Scripting.Dictionary")
    Dim k As String: k = NormClassif(fam)
    If k <> "" And Not dTipoCls.Exists(k) Then dTipoCls(k) = tipo
End Sub


' De-para CLS2 de SERVICO -> familia de MATERIAL equivalente. Regulariza
' servicos cuja familia nao existe no catalogo de materiais, ligando-os a
' familia equivalente para a aderencia MAT vs SRV (via FamiliaAlias).
' Extensivel sem codigo pela chave EQUIV_SRV_MAT da aba CONFIG ("SRV=MAT;...").
Private Sub CarregarEquivSrvMat()
    Set dFamEquiv = CreateObject("Scripting.Dictionary")
    Dim s As String
    s = "CH FUSIVEL=CH FUS;" & _
        "CONEXAO=CONECTOR;" & _
        "TENSIONAR=COND PROT;" & _
        "ESTRUTURA=CRUZETA;" & _
        "ESTRUT MT RSB=CRUZETA;" & _
        "ESTRUT RDC=ESPACADOR LOSAG;" & _
        "ESTRUT MT S/CRUZ=CRUZETA;" & _
        "CONECTOR=RAMAL;" & _
        "DESLIG E RELIG RM=RAMAL"
    Dim extra As String: extra = CfgTxt("EQUIV_SRV_MAT", "")
    If extra <> "" Then s = s & ";" & extra

    Dim parts() As String, kv() As String, i As Long, key As String
    parts = Split(s, ";")
    For i = 0 To UBound(parts)
        kv = Split(parts(i), "=")
        If UBound(kv) >= 1 Then
            key = NormClassif(kv(0))
            If key <> "" Then dFamEquiv(key) = Trim$(kv(1))
        End If
    Next i
End Sub

' Familias de servico que legitimamente NAO tem material correspondente
' (servico puro): nao devem gerar alerta em SERVICO SEM MATERIAL.
' Extensivel pela chave SRV_PURO da aba CONFIG (familias separadas por ;).
Private Sub CarregarSrvPuro()
    Set dSrvPuro = CreateObject("Scripting.Dictionary")
    Dim s As String
    ' ESCAVACAO e servico puro: na MATERIAL vs SERVICO, a qtd de servico do
    ' POSTE so vem de servicos classificados como POSTE RD / MAO DE OBRA.
    s = "PODA;ESCAVACAO;CIVIL;BASE CONC;FUNDACAO;DEMOLIR;MEIO AMBIENTE;" & _
        "ESTUDO AMBIENTAL;PROJETO;FRETE/TRANSP;TRANSPORTE;DESLOCAMENTO;" & _
        "MOBILIZAR;DESMOBILIZAR;ATV DRT;DISPONIBILIDADE;DISPON_STC;" & _
        "ATEND EMERGENCIA;ABRIR ACESSO/FAIXA_SERV;PUBLICIDADE;TURMA LV;" & _
        "ALUGUEL GERADOR;CAIXA ALVENARIA;CAIXA DE PASSAGEM;CALCADA;" & _
        "CAMARA ABRIGADA;CANO GALV;CANO GALVAN;HOSPEDAGEM;INSPECAO;IP;" & _
        "RECOMP PAVIMENTO;RESTAURAR PASSEIO;TAMPA_CONCRETO;TAXA DNIT;ENTULHOS/LODO;" & _
        "FEC_CAVA"
    Dim extra As String: extra = CfgTxt("SRV_PURO", "")
    If extra <> "" Then s = s & ";" & extra

    Dim parts() As String, i As Long, key As String
    parts = Split(s, ";")
    For i = 0 To UBound(parts)
        key = NormClassif(parts(i))
        If key <> "" Then dSrvPuro(key) = 1
    Next i
End Sub

' Servico puro = familia sem material esperado (nao gera falso alerta)
Private Function EhServicoPuro(ByVal cls2 As String) As Boolean
    If dSrvPuro Is Nothing Then Exit Function
    EhServicoPuro = dSrvPuro.Exists(NormClassif(cls2))
End Function

Private Sub CarregarFamiliaNivel()
    ' Nivel da familia (FAMILIAS_ATUAIS): P=PRIMARIO, S=SECUNDARIO, demais=NA.
    Set dFamNivel = CreateObject("Scripting.Dictionary")
    ' Tabela embutida dividida em 9 blocos (limite de tamanho de procedimento).
    CarregarFamiliaNivelP1
    CarregarFamiliaNivelP2
    CarregarFamiliaNivelP3
    CarregarFamiliaNivelP4
    CarregarFamiliaNivelP5
    CarregarFamiliaNivelP6
    CarregarFamiliaNivelP7
    CarregarFamiliaNivelP8
    CarregarFamiliaNivelP9
End Sub

Private Sub CarregarFamiliaNivelP1()
    dFamNivel(NormClassif("102100067")) = "PRIMARIO"
    dFamNivel(NormClassif("102100160")) = "PRIMARIO"
    dFamNivel(NormClassif("ABERTURA DE ACESSO")) = "NA"
    dFamNivel(NormClassif("ABRIGO_EXTINTOR")) = "NA"
    dFamNivel(NormClassif("ABRIR ACESSO/FAIXA_SERV")) = "NA"
    dFamNivel(NormClassif("ABRIR FAIXA")) = "NA"
    dFamNivel(NormClassif("ABRIR/FECHAR VALA ATERR")) = "NA"
    dFamNivel(NormClassif("ACES P/REGUL")) = "NA"
    dFamNivel(NormClassif("ACES P/RELIG")) = "NA"
    dFamNivel(NormClassif("ACES P/TRAFO")) = "NA"
    dFamNivel(NormClassif("ACRES GRAU")) = "NA"
    dFamNivel(NormClassif("ADD EXEC MANUAL")) = "NA"
    dFamNivel(NormClassif("ADD_RET_BASE_POSTE")) = "NA"
    dFamNivel(NormClassif("ADIC EXEC MANUAL")) = "NA"
    dFamNivel(NormClassif("AJUSTE RELE")) = "NA"
    dFamNivel(NormClassif("ALAMBRADO")) = "NA"
    dFamNivel(NormClassif("ALCA")) = "NA"
    dFamNivel(NormClassif("ALUGUEL")) = "NA"
    dFamNivel(NormClassif("ALUGUEL GERADOR")) = "NA"
    dFamNivel(NormClassif("ALUGUEL MUNK")) = "NA"
    dFamNivel(NormClassif("ALUGUEL SUPORTE")) = "NA"
    dFamNivel(NormClassif("AMORTECEDOR")) = "NA"
    dFamNivel(NormClassif("ANEL CONC")) = "SECUNDARIO"
    dFamNivel(NormClassif("ANTENA")) = "SECUNDARIO"
    dFamNivel(NormClassif("APRUMAR POSTE")) = "SECUNDARIO"
    dFamNivel(NormClassif("AQUIS COMBUSTIVEL")) = "NA"
    dFamNivel(NormClassif("AREIA")) = "NA"
    dFamNivel(NormClassif("ARRASTAR_MAT_C/TRATOR")) = "NA"
    dFamNivel(NormClassif("ARRASTAR_MAT_MANUAL")) = "NA"
    dFamNivel(NormClassif("ART EMISSAO DE DOC")) = "NA"
    dFamNivel(NormClassif("ARTEFATO CONC")) = "NA"
    dFamNivel(NormClassif("ATEND EMERGENCIA")) = "NA"
    dFamNivel(NormClassif("ATERRAMENTO")) = "NA"
    dFamNivel(NormClassif("ATERRO")) = "SECUNDARIO"
    dFamNivel(NormClassif("ATERR_CERCA")) = "NA"
    dFamNivel(NormClassif("ATERR_EQUIP")) = "NA"
    dFamNivel(NormClassif("ATV DIRETA")) = "NA"
    dFamNivel(NormClassif("ATV DRT")) = "NA"
    dFamNivel(NormClassif("AVISO")) = "NA"
    dFamNivel(NormClassif("BANCO BATERIA")) = "PRIMARIO"
    dFamNivel(NormClassif("BANCO CAPACITOR")) = "PRIMARIO"
    dFamNivel(NormClassif("BANCO DUTO")) = "SECUNDARIO"
    dFamNivel(NormClassif("BARRACAO")) = "NA"
    dFamNivel(NormClassif("BARRAMENTO")) = "PRIMARIO"
    dFamNivel(NormClassif("BARRA_SUPORTE")) = "NA"
    dFamNivel(NormClassif("BASE")) = "SECUNDARIO"
    dFamNivel(NormClassif("BASE BCP")) = "NA"
    dFamNivel(NormClassif("BASE CABO FORCA")) = "SECUNDARIO"
    dFamNivel(NormClassif("BASE CH SEC")) = "NA"
    dFamNivel(NormClassif("BASE CH SEC 138KV")) = "SECUNDARIO"
    dFamNivel(NormClassif("BASE CONC")) = "SECUNDARIO"
    dFamNivel(NormClassif("BASE CONC DISJ")) = "SECUNDARIO"
    dFamNivel(NormClassif("BASE CONC TAMBOR")) = "SECUNDARIO"
    dFamNivel(NormClassif("BASE EQUIP")) = "PRIMARIO"
    dFamNivel(NormClassif("BASE TF")) = "SECUNDARIO"
    dFamNivel(NormClassif("BASE TSA")) = "SECUNDARIO"
    dFamNivel(NormClassif("BASE_BANCO_CAPAC")) = "SECUNDARIO"
    dFamNivel(NormClassif("BASE_CHAVE_SEC")) = "SECUNDARIO"
    dFamNivel(NormClassif("BASE_IP_PR_TC_TP")) = "NA"
    dFamNivel(NormClassif("BASE_IP_PR_TC_TP_69_K")) = "SECUNDARIO"
    dFamNivel(NormClassif("BASE_REATOR")) = "SECUNDARIO"
    dFamNivel(NormClassif("BATERIA")) = "SECUNDARIO"
    dFamNivel(NormClassif("BATERIA SOLAR")) = "SECUNDARIO"
    dFamNivel(NormClassif("BC DUTOS")) = "NA"
    dFamNivel(NormClassif("BENZINA")) = "NA"
    dFamNivel(NormClassif("BGAN KIT LEGACY")) = "SECUNDARIO"
    dFamNivel(NormClassif("BIG JUMP LM")) = "SECUNDARIO"
    dFamNivel(NormClassif("BIG JUMP LV")) = "SECUNDARIO"
    dFamNivel(NormClassif("BLINDAGEM_BT")) = "NA"
    dFamNivel(NormClassif("BLOCO CERAMICO")) = "SECUNDARIO"
    dFamNivel(NormClassif("BLOCO CONC ESTAI")) = "SECUNDARIO"
    dFamNivel(NormClassif("BOMBA DAGUA")) = "SECUNDARIO"
    dFamNivel(NormClassif("BOMBA SUBM")) = "SECUNDARIO"
    dFamNivel(NormClassif("BORNE")) = "NA"
    dFamNivel(NormClassif("BRINDE PROMOC - RISCO")) = "NA"
End Sub

Private Sub CarregarFamiliaNivelP2()
    dFamNivel(NormClassif("BRITA")) = "NA"
    dFamNivel(NormClassif("BRITAMENTO DO PATIO")) = "SECUNDARIO"
    dFamNivel(NormClassif("CABLAGEM _BT")) = "NA"
    dFamNivel(NormClassif("CABO COAXIAL")) = "NA"
    dFamNivel(NormClassif("CABO COBRE")) = "NA"
    dFamNivel(NormClassif("CABO CONTROLE")) = "NA"
    dFamNivel(NormClassif("CABO CU")) = "NA"
    dFamNivel(NormClassif("CABO FIB OPT")) = "NA"
    dFamNivel(NormClassif("CABO ISOLADO")) = "NA"
    dFamNivel(NormClassif("CABO NU")) = "NA"
    dFamNivel(NormClassif("CABO RELE MULTI")) = "NA"
    dFamNivel(NormClassif("CABO SOLAR")) = "NA"
    dFamNivel(NormClassif("CABOS_PARA_RAIOS")) = "NA"
    dFamNivel(NormClassif("CABO_ILUMINACAO")) = "NA"
    dFamNivel(NormClassif("CABO_ILUMINACAƑO")) = "NA"
    dFamNivel(NormClassif("CADEADO")) = "SECUNDARIO"
    dFamNivel(NormClassif("CADEIA ISOL")) = "PRIMARIO"
    dFamNivel(NormClassif("CADEIA_ISOLADORES")) = "PRIMARIO"
    dFamNivel(NormClassif("CAD_CLIENTE")) = "NA"
    dFamNivel(NormClassif("CAD_GIS")) = "NA"
    dFamNivel(NormClassif("CAIXA 4MED")) = "NA"
    dFamNivel(NormClassif("CAIXA 8MED")) = "NA"
    dFamNivel(NormClassif("CAIXA ALVENARIA")) = "NA"
    dFamNivel(NormClassif("CAIXA BLIN PDR")) = "SECUNDARIO"
    dFamNivel(NormClassif("CAIXA BLIND PDR")) = "SECUNDARIO"
    dFamNivel(NormClassif("CAIXA DE PASSAGEM")) = "NA"
    dFamNivel(NormClassif("CAIXA LIG")) = "SECUNDARIO"
    dFamNivel(NormClassif("CAIXA PASS")) = "NA"
    dFamNivel(NormClassif("CAIXA PASSAGEM")) = "SECUNDARIO"
    dFamNivel(NormClassif("CAIXA PDR")) = "SECUNDARIO"
    dFamNivel(NormClassif("CAIXA SIGFI")) = "SECUNDARIO"
    dFamNivel(NormClassif("CAIXA TC")) = "NA"
    dFamNivel(NormClassif("CAIXA TP_TC")) = "SECUNDARIO"
    dFamNivel(NormClassif("CAIXA_DERIVACAƑO_TC_TP")) = "NA"
    dFamNivel(NormClassif("CAIXA_INSP_MALHA")) = "NA"
    dFamNivel(NormClassif("CAIXA_SEPARADORA")) = "NA"
    dFamNivel(NormClassif("CAIXA_TC")) = "SECUNDARIO"
    dFamNivel(NormClassif("CALCADA")) = "NA"
    dFamNivel(NormClassif("CAMARA ABRIGADA")) = "NA"
    dFamNivel(NormClassif("CAMINHAO GUIND")) = "NA"
    dFamNivel(NormClassif("CANALETA")) = "NA"
    dFamNivel(NormClassif("CANALETA CONC")) = "NA"
    dFamNivel(NormClassif("CANO ELETRODUTO")) = "NA"
    dFamNivel(NormClassif("CANO GALV")) = "NA"
    dFamNivel(NormClassif("CANO GALVAN")) = "NA"
    dFamNivel(NormClassif("CANTEIRO_DE_OBRAS")) = "NA"
    dFamNivel(NormClassif("CANTERIO DE OBRA")) = "NA"
    dFamNivel(NormClassif("CANTONEIRA SIGFI")) = "NA"
    dFamNivel(NormClassif("CAPACITOR")) = "PRIMARIO"
    dFamNivel(NormClassif("CARGA/DESCARGA")) = "NA"
    dFamNivel(NormClassif("CARRINHO_LIMPEZA")) = "SECUNDARIO"
    dFamNivel(NormClassif("CASA DE COMANDO")) = "PRIMARIO"
    dFamNivel(NormClassif("CAVA")) = "NA"
    dFamNivel(NormClassif("CAVA_ROCHA")) = "SECUNDARIO"
    dFamNivel(NormClassif("CAVA_ROCHA_C/EXP")) = "SECUNDARIO"
    dFamNivel(NormClassif("CAVA_ROCHA_S/EXP")) = "SECUNDARIO"
    dFamNivel(NormClassif("CENSO CADASTRO UC")) = "NA"
    dFamNivel(NormClassif("CERCA")) = "NA"
    dFamNivel(NormClassif("CH AFER")) = "NA"
    dFamNivel(NormClassif("CH FU RELIG")) = "PRIMARIO"
    dFamNivel(NormClassif("CH FUS")) = "PRIMARIO"
    dFamNivel(NormClassif("CH FUS RELIG")) = "PRIMARIO"
    dFamNivel(NormClassif("CH SEC 138KV")) = "PRIMARIO"
    dFamNivel(NormClassif("CH SEC 138KV_BASE")) = "NA"
    dFamNivel(NormClassif("CH SEC 145KV")) = "PRIMARIO"
    dFamNivel(NormClassif("CH SEC 69KV")) = "PRIMARIO"
    dFamNivel(NormClassif("CH SEC OLEO")) = "PRIMARIO"
    dFamNivel(NormClassif("CH SEC SE")) = "PRIMARIO"
    dFamNivel(NormClassif("CH SEC SECO")) = "PRIMARIO"
    dFamNivel(NormClassif("CH SEC TANDEM")) = "PRIMARIO"
    dFamNivel(NormClassif("CH SEC TRI")) = "PRIMARIO"
    dFamNivel(NormClassif("CH SU AT")) = "PRIMARIO"
    dFamNivel(NormClassif("CH SU FACA")) = "PRIMARIO"
    dFamNivel(NormClassif("CH SU FACA AT")) = "PRIMARIO"
    dFamNivel(NormClassif("CH VAC 3F")) = "PRIMARIO"
End Sub

Private Sub CarregarFamiliaNivelP3()
    dFamNivel(NormClassif("CHAPA XADREZ")) = "NA"
    dFamNivel(NormClassif("CHAVE_E_OU_PARA_RAIO")) = "PRIMARIO"
    dFamNivel(NormClassif("CIMENTO")) = "NA"
    dFamNivel(NormClassif("CINTA POSTE")) = "NA"
    dFamNivel(NormClassif("CIVIL")) = "PRIMARIO"
    dFamNivel(NormClassif("CIVIL SUBT")) = "PRIMARIO"
    dFamNivel(NormClassif("COBERTURA ISOLANTES")) = "NA"
    dFamNivel(NormClassif("COMANDO RELIG")) = "SECUNDARIO"
    dFamNivel(NormClassif("COMISSIONAMENTO")) = "NA"
    dFamNivel(NormClassif("COMPONENTE SE")) = "NA"
    dFamNivel(NormClassif("COM_MAT_SERV")) = "NA"
    dFamNivel(NormClassif("COM_MAT_SRV")) = "NA"
    dFamNivel(NormClassif("CONCENTRADOR TP")) = "SECUNDARIO"
    dFamNivel(NormClassif("CONCRETAR_CALCADA")) = "NA"
    dFamNivel(NormClassif("CONCRETO_ARMADO")) = "SECUNDARIO"
    dFamNivel(NormClassif("COND COBRE")) = "PRIMARIO"
    dFamNivel(NormClassif("COND ISOLADO")) = "PRIMARIO"
    dFamNivel(NormClassif("COND ISOLADO/PROT")) = "PRIMARIO"
    dFamNivel(NormClassif("COND NU")) = "PRIMARIO"
    dFamNivel(NormClassif("COND PROT")) = "PRIMARIO"
    dFamNivel(NormClassif("CONDUTOR")) = "SECUNDARIO"
    dFamNivel(NormClassif("CONDUTORES_LD")) = "PRIMARIO"
    dFamNivel(NormClassif("CONECTAR TF")) = "NA"
    dFamNivel(NormClassif("CONECTAR_RAMAL")) = "NA"
    dFamNivel(NormClassif("CONECTOR")) = "NA"
    dFamNivel(NormClassif("CONEXAO")) = "NA"
    dFamNivel(NormClassif("CONEXAO_BC")) = "NA"
    dFamNivel(NormClassif("CONEXA•ES_BARRAMENTO")) = "NA"
    dFamNivel(NormClassif("CONEXOES_BARRAMENTO")) = "NA"
    dFamNivel(NormClassif("CONF FORMAR CONCR")) = "NA"
    dFamNivel(NormClassif("CONJ CUBICULO")) = "PRIMARIO"
    dFamNivel(NormClassif("CONJ MED")) = "SECUNDARIO"
    dFamNivel(NormClassif("CONST_CAIXA PASSAGEM ALV")) = "NA"
    dFamNivel(NormClassif("CONST_IMOVEL")) = "PRIMARIO"
    dFamNivel(NormClassif("CONSULTORIA")) = "NA"
    dFamNivel(NormClassif("CONSULTORIA ADMINISTRATIVA")) = "NA"
    dFamNivel(NormClassif("CONTENCAO_EROSAO")) = "NA"
    dFamNivel(NormClassif("CONTR ELETN")) = "NA"
    dFamNivel(NormClassif("CONTRAPINO_ISOL")) = "NA"
    dFamNivel(NormClassif("CONTROLADOR")) = "SECUNDARIO"
    dFamNivel(NormClassif("CONTROLADOR_SOLAR")) = "NA"
    dFamNivel(NormClassif("CONTROLE")) = "NA"
    dFamNivel(NormClassif("CONTROLE RELIG")) = "SECUNDARIO"
    dFamNivel(NormClassif("CONVERSOR")) = "NA"
    dFamNivel(NormClassif("CORDAO FIBRA")) = "NA"
    dFamNivel(NormClassif("CORDAO_FIBRA")) = "NA"
    dFamNivel(NormClassif("CORDOALHA")) = "SECUNDARIO"
    dFamNivel(NormClassif("CORTA-CIRCUITO")) = "NA"
    dFamNivel(NormClassif("CORTE_MECANIZADO")) = "NA"
    dFamNivel(NormClassif("CP_CS_MD")) = "PRIMARIO"
    dFamNivel(NormClassif("CRUZETA")) = "NA"
    dFamNivel(NormClassif("CUBICULO")) = "PRIMARIO"
    dFamNivel(NormClassif("CX BLIND CMB")) = "NA"
    dFamNivel(NormClassif("CX DISPLAY SMC")) = "NA"
    dFamNivel(NormClassif("CX PAPELAO")) = "NA"
    dFamNivel(NormClassif("CX PORTA DISJ")) = "NA"
    dFamNivel(NormClassif("DEFENSA")) = "PRIMARIO"
    dFamNivel(NormClassif("DEFENSA CONC")) = "SECUNDARIO"
    dFamNivel(NormClassif("DEMARCAR_POSTE")) = "SECUNDARIO"
    dFamNivel(NormClassif("DEMOLIR")) = "NA"
    dFamNivel(NormClassif("DEMOLIR BASE CONCR")) = "NA"
    dFamNivel(NormClassif("DESAT INTERLIG DISJ")) = "NA"
    dFamNivel(NormClassif("DESAT INTERLIG RELIG")) = "NA"
    dFamNivel(NormClassif("DESAT INTERLIG TPTC")) = "SECUNDARIO"
    dFamNivel(NormClassif("DESAT INTERLIGAR TP_TC")) = "NA"
    dFamNivel(NormClassif("DESAT INTERLIGAR TSA")) = "NA"
    dFamNivel(NormClassif("DESCARTE_MATERIAL")) = "NA"
    dFamNivel(NormClassif("DESCONECTAR TF")) = "NA"
    dFamNivel(NormClassif("DESCONECT_RAMAL")) = "NA"
    dFamNivel(NormClassif("DESLIG E RELIG IP")) = "NA"
    dFamNivel(NormClassif("DESLIG E RELIG RM")) = "NA"
    dFamNivel(NormClassif("DESLOCAMENTO")) = "NA"
    dFamNivel(NormClassif("DESMOBILIZAR")) = "NA"
    dFamNivel(NormClassif("DESMONT E REINST IP")) = "NA"
    dFamNivel(NormClassif("DESTOCAMENTO")) = "NA"
End Sub

Private Sub CarregarFamiliaNivelP4()
    dFamNivel(NormClassif("DETECTOR_FALTA")) = "NA"
    dFamNivel(NormClassif("DETRITOS")) = "NA"
    dFamNivel(NormClassif("DIAGNOSTICO")) = "NA"
    dFamNivel(NormClassif("DILIGENCIA FAIXA SERVIDAO")) = "NA"
    dFamNivel(NormClassif("DISJ")) = "PRIMARIO"
    dFamNivel(NormClassif("DISJ BT")) = "NA"
    dFamNivel(NormClassif("DISJ CUBICULO")) = "PRIMARIO"
    dFamNivel(NormClassif("DISJ SE")) = "PRIMARIO"
    dFamNivel(NormClassif("DISJUNTOR")) = "PRIMARIO"
    dFamNivel(NormClassif("DISJUNTOR_SE")) = "PRIMARIO"
    dFamNivel(NormClassif("DISPLAY SMC")) = "NA"
    dFamNivel(NormClassif("DISPON")) = "NA"
    dFamNivel(NormClassif("DISPONIBILIDADE")) = "NA"
    dFamNivel(NormClassif("DISPON_LT")) = "NA"
    dFamNivel(NormClassif("DISPON_LV")) = "NA"
    dFamNivel(NormClassif("DISPON_SE")) = "SECUNDARIO"
    dFamNivel(NormClassif("DISPON_STC")) = "SECUNDARIO"
    dFamNivel(NormClassif("DISTRIBUIDOR")) = "NA"
    dFamNivel(NormClassif("DISTRIBUIDOR_INT")) = "NA"
    dFamNivel(NormClassif("DIST_ESTRUTURAS")) = "NA"
    dFamNivel(NormClassif("DRENO")) = "NA"
    dFamNivel(NormClassif("DSISTRIBUIR_ESTRUTURAS")) = "NA"
    dFamNivel(NormClassif("DUTO")) = "NA"
    dFamNivel(NormClassif("DUTOS")) = "NA"
    dFamNivel(NormClassif("ELETRODUTO")) = "NA"
    dFamNivel(NormClassif("EMENDA")) = "NA"
    dFamNivel(NormClassif("EMERGENCIAL")) = "NA"
    dFamNivel(NormClassif("ENTULHO")) = "NA"
    dFamNivel(NormClassif("ENTULHOS/LODO")) = "NA"
    dFamNivel(NormClassif("EQUIP")) = "PRIMARIO"
    dFamNivel(NormClassif("ESCADA PEDAROLA")) = "NA"
    dFamNivel(NormClassif("ESCADA_CONCRETO")) = "NA"
    dFamNivel(NormClassif("ESCADA_MET_DISJ")) = "NA"
    dFamNivel(NormClassif("ESCAVA MANUAL")) = "NA"
    dFamNivel(NormClassif("ESCAVACAO")) = "NA"
    dFamNivel(NormClassif("ESCAV_ROCHA")) = "NA"
    dFamNivel(NormClassif("ESFERA")) = "NA"
    dFamNivel(NormClassif("ESFERA SINALIZ")) = "NA"
    dFamNivel(NormClassif("ESGOTAMENTO")) = "NA"
    dFamNivel(NormClassif("ESPACADOR")) = "NA"
    dFamNivel(NormClassif("ESPACADOR LOSAG")) = "NA"
    dFamNivel(NormClassif("ESPALHAR_BRITA")) = "NA"
    dFamNivel(NormClassif("ESPUMA POLIURET")) = "NA"
    dFamNivel(NormClassif("ESTAI")) = "NA"
    dFamNivel(NormClassif("ESTAI AEREO")) = "NA"
    dFamNivel(NormClassif("ESTAI ANCORA")) = "NA"
    dFamNivel(NormClassif("ESTAI CONTRAPOSTE")) = "NA"
    dFamNivel(NormClassif("ESTEIO_MADEIRA")) = "NA"
    dFamNivel(NormClassif("ESTRUT")) = "SECUNDARIO"
    dFamNivel(NormClassif("ESTRUT AUX LD")) = "SECUNDARIO"
    dFamNivel(NormClassif("ESTRUT BT")) = "NA"
    dFamNivel(NormClassif("ESTRUT BT RSB")) = "NA"
    dFamNivel(NormClassif("ESTRUT LD")) = "SECUNDARIO"
    dFamNivel(NormClassif("ESTRUT MET")) = "NA"
    dFamNivel(NormClassif("ESTRUT METALICA")) = "NA"
    dFamNivel(NormClassif("ESTRUT MT")) = "NA"
    dFamNivel(NormClassif("ESTRUT MT RSB")) = "NA"
    dFamNivel(NormClassif("ESTRUT MT S/ CRUZ")) = "SECUNDARIO"
    dFamNivel(NormClassif("ESTRUT MT S/CRUZ")) = "NA"
    dFamNivel(NormClassif("ESTRUT RDC")) = "NA"
    dFamNivel(NormClassif("ESTRUT REGUL")) = "NA"
    dFamNivel(NormClassif("ESTRUT SIGFI")) = "NA"
    dFamNivel(NormClassif("ESTRUT SOLAR")) = "NA"
    dFamNivel(NormClassif("ESTRUT SUP CONC")) = "NA"
    dFamNivel(NormClassif("ESTRUT TRAFO/POSTE")) = "SECUNDARIO"
    dFamNivel(NormClassif("ESTRUTURA")) = "NA"
    dFamNivel(NormClassif("ESTRUTURA LD")) = "SECUNDARIO"
    dFamNivel(NormClassif("ESTRUTURA SUPORTE")) = "NA"
    dFamNivel(NormClassif("ESTRUTURAS")) = "NA"
    dFamNivel(NormClassif("ESTRUTURA_ILUMINACAƑO")) = "NA"
    dFamNivel(NormClassif("ESTRUT_CHAVE_13.8KV")) = "PRIMARIO"
    dFamNivel(NormClassif("ESTR_RDC (PR, CF, SU)")) = "PRIMARIO"
    dFamNivel(NormClassif("ESTUDO AMBIENTAL")) = "NA"
    dFamNivel(NormClassif("EXAUSTOR")) = "NA"
    dFamNivel(NormClassif("EXEC_POCO")) = "SECUNDARIO"
End Sub

Private Sub CarregarFamiliaNivelP5()
    dFamNivel(NormClassif("EXTENSAO OTICA")) = "NA"
    dFamNivel(NormClassif("EXTERNALIZACAO PADRAO")) = "SECUNDARIO"
    dFamNivel(NormClassif("EXTINTOR")) = "NA"
    dFamNivel(NormClassif("FAIXA DE SERVIDAO")) = "NA"
    dFamNivel(NormClassif("FASEAMENTO")) = "NA"
    dFamNivel(NormClassif("FATOR K")) = "NA"
    dFamNivel(NormClassif("FEC_CAVA")) = "SECUNDARIO"
    dFamNivel(NormClassif("FERRAGEM")) = "NA"
    dFamNivel(NormClassif("FIBRA OPTICA")) = "NA"
    dFamNivel(NormClassif("FINAC_PDR")) = "NA"
    dFamNivel(NormClassif("FIREWALL")) = "SECUNDARIO"
    dFamNivel(NormClassif("FORNEC MATERIAL")) = "NA"
    dFamNivel(NormClassif("FRETE/TRANSP")) = "NA"
    dFamNivel(NormClassif("FUNDACAO")) = "SECUNDARIO"
    dFamNivel(NormClassif("FUSIVEL")) = "SECUNDARIO"
    dFamNivel(NormClassif("GARANTIA ESTENDIDA")) = "SECUNDARIO"
    dFamNivel(NormClassif("GRAMA")) = "NA"
    dFamNivel(NormClassif("GRAMP COND")) = "NA"
    dFamNivel(NormClassif("GRAMPO")) = "NA"
    dFamNivel(NormClassif("GRAXA")) = "NA"
    dFamNivel(NormClassif("GRAXA (ANE)")) = "SECUNDARIO"
    dFamNivel(NormClassif("GUARDA_BIG_JUMPER")) = "NA"
    dFamNivel(NormClassif("HASTE ANCORA ESTAI")) = "NA"
    dFamNivel(NormClassif("HASTE ATERR")) = "NA"
    dFamNivel(NormClassif("HIDROMETRO")) = "PRIMARIO"
    dFamNivel(NormClassif("HORA EXTRA")) = "NA"
    dFamNivel(NormClassif("HOSPEDAGEM")) = "NA"
    dFamNivel(NormClassif("IDENT CABOS")) = "NA"
    dFamNivel(NormClassif("INSPECAO")) = "NA"
    dFamNivel(NormClassif("INTERLIGAR BC")) = "NA"
    dFamNivel(NormClassif("INTERLIGAR CONTACTOR")) = "NA"
    dFamNivel(NormClassif("INTERLIGAR DISJ")) = "NA"
    dFamNivel(NormClassif("INTERLIGAR PAINEL")) = "NA"
    dFamNivel(NormClassif("INTERLIGAR RELE")) = "NA"
    dFamNivel(NormClassif("INTERLIGAR RELIG")) = "NA"
    dFamNivel(NormClassif("INTERLIGAR TC")) = "NA"
    dFamNivel(NormClassif("INTERLIGAR TF")) = "NA"
    dFamNivel(NormClassif("INTERLIGAR TP")) = "NA"
    dFamNivel(NormClassif("INTERLIGAR TSA")) = "NA"
    dFamNivel(NormClassif("INTERRUPTOR TOM")) = "NA"
    dFamNivel(NormClassif("INTERVENCAO SE/LT LV")) = "NA"
    dFamNivel(NormClassif("INTERVENCAƑO SE/LT LV")) = "NA"
    dFamNivel(NormClassif("INVERSOR")) = "PRIMARIO"
    dFamNivel(NormClassif("INVERSOR SIGFI")) = "SECUNDARIO"
    dFamNivel(NormClassif("IP")) = "NA"
    dFamNivel(NormClassif("IP_69KV")) = "NA"
    dFamNivel(NormClassif("ISOLADO ANC MT")) = "NA"
    dFamNivel(NormClassif("ISOLADOR")) = "NA"
    dFamNivel(NormClassif("ISOLADOR ANC MT")) = "NA"
    dFamNivel(NormClassif("ISOLADOR AT")) = "SECUNDARIO"
    dFamNivel(NormClassif("ISOLADOR BT")) = "NA"
    dFamNivel(NormClassif("ISOLADOR ESPACADOR")) = "NA"
    dFamNivel(NormClassif("ISOLADOR MT")) = "NA"
    dFamNivel(NormClassif("ISOLADOR PEDESTAL")) = "SECUNDARIO"
    dFamNivel(NormClassif("ISOLADOR PINO MT")) = "NA"
    dFamNivel(NormClassif("JUMPER")) = "NA"
    dFamNivel(NormClassif("JUMPERS")) = "NA"
    dFamNivel(NormClassif("KIT BUSHLINK")) = "SECUNDARIO"
    dFamNivel(NormClassif("KIT ELETR SIGFI")) = "NA"
    dFamNivel(NormClassif("KIT PLPT")) = "SECUNDARIO"
    dFamNivel(NormClassif("LACRE")) = "NA"
    dFamNivel(NormClassif("LAMPADA")) = "NA"
    dFamNivel(NormClassif("LANC CONDUTOR")) = "PRIMARIO"
    dFamNivel(NormClassif("LAUDO")) = "SECUNDARIO"
    dFamNivel(NormClassif("LAVAGEM CAMARA TR")) = "NA"
    dFamNivel(NormClassif("LAVAGEM CX")) = "NA"
    dFamNivel(NormClassif("LAVAGEM CX LIG")) = "NA"
    dFamNivel(NormClassif("LICENCA AMBIENTAL")) = "NA"
    dFamNivel(NormClassif("LIG CONSUMIDOR")) = "NA"
    dFamNivel(NormClassif("LIGACAO_PROVISORIA")) = "NA"
    dFamNivel(NormClassif("LIMPEZA DE TELA")) = "NA"
    dFamNivel(NormClassif("LIMPEZA SE")) = "NA"
    dFamNivel(NormClassif("LIMPEZA_TERRENO")) = "NA"
    dFamNivel(NormClassif("LINHA MORTA")) = "NA"
    dFamNivel(NormClassif("LINHA VIVA")) = "NA"
End Sub

Private Sub CarregarFamiliaNivelP6()
    dFamNivel(NormClassif("LNB AUTO")) = "SECUNDARIO"
    dFamNivel(NormClassif("LOC EQUIP")) = "NA"
    dFamNivel(NormClassif("LOC MAT/EQ")) = "NA"
    dFamNivel(NormClassif("LOC. DE VEICULOS")) = "NA"
    dFamNivel(NormClassif("LOCACAO_ESTRUTURAS")) = "NA"
    dFamNivel(NormClassif("LOCACAO_SE")) = "NA"
    dFamNivel(NormClassif("LOC_EQP")) = "NA"
    dFamNivel(NormClassif("LOGISTICA")) = "NA"
    dFamNivel(NormClassif("LUBRIFICANTE")) = "NA"
    dFamNivel(NormClassif("LUMINARIA")) = "NA"
    dFamNivel(NormClassif("LUVA")) = "NA"
    dFamNivel(NormClassif("MALHA ATERR")) = "NA"
    dFamNivel(NormClassif("MALHA TERRA")) = "NA"
    dFamNivel(NormClassif("MANILHA")) = "NA"
    dFamNivel(NormClassif("MANOBRA CHAVE")) = "NA"
    dFamNivel(NormClassif("MANOBRA_AT")) = "NA"
    dFamNivel(NormClassif("MANUT CORRET")) = "NA"
    dFamNivel(NormClassif("MANUT RELE")) = "NA"
    dFamNivel(NormClassif("MANUT_SOFTWARE")) = "SECUNDARIO"
    dFamNivel(NormClassif("MAO FRANCESA")) = "NA"
    dFamNivel(NormClassif("MAPEAM AEROFOTOG")) = "NA"
    dFamNivel(NormClassif("MAP_AEROFOTOGRAMETR")) = "NA"
    dFamNivel(NormClassif("MASSA ISOLANTE")) = "NA"
    dFamNivel(NormClassif("MASTRO ANTENA")) = "NA"
    dFamNivel(NormClassif("MATRIZ_DIODO")) = "NA"
    dFamNivel(NormClassif("MED PROVISORIA")) = "SECUNDARIO"
    dFamNivel(NormClassif("MEDIDOR")) = "PRIMARIO"
    dFamNivel(NormClassif("MEDIDOR E RAMAL")) = "PRIMARIO"
    dFamNivel(NormClassif("MEDIDOR_BALANC")) = "SECUNDARIO"
    dFamNivel(NormClassif("MEDIDOR_FISCAL")) = "PRIMARIO"
    dFamNivel(NormClassif("MEDIDOR_FRONTEIRA")) = "SECUNDARIO"
    dFamNivel(NormClassif("MEIO AMBIENTE")) = "NA"
    dFamNivel(NormClassif("MEIO-AMBIENTE")) = "NA"
    dFamNivel(NormClassif("MEIO_FIO")) = "NA"
    dFamNivel(NormClassif("MOB/DESMOB")) = "NA"
    dFamNivel(NormClassif("MOBILIZAR")) = "NA"
    dFamNivel(NormClassif("MODEM")) = "PRIMARIO"
    dFamNivel(NormClassif("MODULO")) = "SECUNDARIO"
    dFamNivel(NormClassif("MODULO REMOT")) = "SECUNDARIO"
    dFamNivel(NormClassif("MODULO REMOTO")) = "NA"
    dFamNivel(NormClassif("MODULO REPETIDOR")) = "SECUNDARIO"
    dFamNivel(NormClassif("MONT_PDR_PARCIAL")) = "NA"
    dFamNivel(NormClassif("MUFLA")) = "SECUNDARIO"
    dFamNivel(NormClassif("MURO")) = "NA"
    dFamNivel(NormClassif("MURO PRE MOLD")) = "NA"
    dFamNivel(NormClassif("NIVELAR COND")) = "NA"
    dFamNivel(NormClassif("OPERAR CH")) = "NA"
    dFamNivel(NormClassif("OPERAR EQUIP")) = "NA"
    dFamNivel(NormClassif("PADRAO")) = "SECUNDARIO"
    dFamNivel(NormClassif("PAINEL")) = "PRIMARIO"
    dFamNivel(NormClassif("PAINEL CONTR EXAUSTOR")) = "NA"
    dFamNivel(NormClassif("PAINEL CONTROLE")) = "SECUNDARIO"
    dFamNivel(NormClassif("PAINEL MEDICAO")) = "SECUNDARIO"
    dFamNivel(NormClassif("PAINEL MET")) = "PRIMARIO"
    dFamNivel(NormClassif("PAINEL TELECOM")) = "SECUNDARIO"
    dFamNivel(NormClassif("PANO ALGODAO")) = "NA"
    dFamNivel(NormClassif("PARA RAIO")) = "SECUNDARIO"
    dFamNivel(NormClassif("PARA RAIO AT")) = "SECUNDARIO"
    dFamNivel(NormClassif("PARA RAIO BT")) = "NA"
    dFamNivel(NormClassif("PARA RAIO MT")) = "PRIMARIO"
    dFamNivel(NormClassif("PARALELEPIPEDO")) = "NA"
    dFamNivel(NormClassif("PASSAGEM")) = "NA"
    dFamNivel(NormClassif("PAV CONCRETO")) = "SECUNDARIO"
    dFamNivel(NormClassif("PAVIMENTO")) = "NA"
    dFamNivel(NormClassif("PAVIMENTO ASFALTICO")) = "NA"
    dFamNivel(NormClassif("PERNA_MANCA")) = "SECUNDARIO"
    dFamNivel(NormClassif("PINO")) = "NA"
    dFamNivel(NormClassif("PINTURA")) = "NA"
    dFamNivel(NormClassif("PINTURA_POSTE")) = "NA"
    dFamNivel(NormClassif("PISTA_BLC_SEXTAVADO")) = "NA"
    dFamNivel(NormClassif("PISTA_BLOCO")) = "NA"
    dFamNivel(NormClassif("PLACA ADVERT")) = "NA"
    dFamNivel(NormClassif("PLACA DA OBRA")) = "NA"
    dFamNivel(NormClassif("PLACA DE REDE")) = "SECUNDARIO"
    dFamNivel(NormClassif("PLACA ID")) = "NA"
End Sub

Private Sub CarregarFamiliaNivelP7()
    dFamNivel(NormClassif("PLACA ID OBRA")) = "NA"
    dFamNivel(NormClassif("PLACA ID SE")) = "NA"
    dFamNivel(NormClassif("PLACA MONTG SIGFI")) = "NA"
    dFamNivel(NormClassif("PODA")) = "NA"
    dFamNivel(NormClassif("PORTA FUSIVEL")) = "NA"
    dFamNivel(NormClassif("PORTA GATEW")) = "NA"
    dFamNivel(NormClassif("PORTAO_VEICULOS")) = "SECUNDARIO"
    dFamNivel(NormClassif("PORTAƑO_VEICULOS")) = "NA"
    dFamNivel(NormClassif("PORTICO_INDENTIFICACAO")) = "NA"
    dFamNivel(NormClassif("PORTICO_INDENTIFICACAƑO")) = "NA"
    dFamNivel(NormClassif("POSTE")) = "PRIMARIO"
    dFamNivel(NormClassif("POSTE AT")) = "SECUNDARIO"
    dFamNivel(NormClassif("POSTE CAPITEL")) = "PRIMARIO"
    dFamNivel(NormClassif("POSTE DIST_MATERIAL")) = "NA"
    dFamNivel(NormClassif("POSTE LD")) = "PRIMARIO"
    dFamNivel(NormClassif("POSTE PDR ACO")) = "PRIMARIO"
    dFamNivel(NormClassif("POSTE RD")) = "PRIMARIO"
    dFamNivel(NormClassif("POSTE TORRE")) = "PRIMARIO"
    dFamNivel(NormClassif("POSTE/BARRA_SUPORTE")) = "SECUNDARIO"
    dFamNivel(NormClassif("POSTE_TORRE")) = "PRIMARIO"
    dFamNivel(NormClassif("PREGO")) = "SECUNDARIO"
    dFamNivel(NormClassif("PROJETO")) = "NA"
    dFamNivel(NormClassif("PROJETO_AUTOMACAO")) = "NA"
    dFamNivel(NormClassif("PROJ_TERRAPLAN")) = "NA"
    dFamNivel(NormClassif("PROTETOR RETICUL")) = "NA"
    dFamNivel(NormClassif("PUBLICIDADE")) = "NA"
    dFamNivel(NormClassif("QUADRO DE COMANDO")) = "PRIMARIO"
    dFamNivel(NormClassif("QUADRO_ELETR")) = "NA"
    dFamNivel(NormClassif("RADIO")) = "SECUNDARIO"
    dFamNivel(NormClassif("RAMAL")) = "PRIMARIO"
    dFamNivel(NormClassif("RAMAL_SUBTER")) = "PRIMARIO"
    dFamNivel(NormClassif("RAMPA_ACESSO")) = "NA"
    dFamNivel(NormClassif("RD CLANDESTINA")) = "NA"
    dFamNivel(NormClassif("RECOLOCAR BRITA")) = "NA"
    dFamNivel(NormClassif("RECOMP PAVIMENTO")) = "NA"
    dFamNivel(NormClassif("RECUPERAR TRAFO RD")) = "SECUNDARIO"
    dFamNivel(NormClassif("REDE BT CONJUG REDE MT")) = "NA"
    dFamNivel(NormClassif("REFLETOR")) = "SECUNDARIO"
    dFamNivel(NormClassif("REGULADOR")) = "PRIMARIO"
    dFamNivel(NormClassif("REGULAR COND")) = "NA"
    dFamNivel(NormClassif("REINST SE MOVEL")) = "NA"
    dFamNivel(NormClassif("RELE")) = "SECUNDARIO"
    dFamNivel(NormClassif("RELE AUX")) = "NA"
    dFamNivel(NormClassif("RELE PROT")) = "NA"
    dFamNivel(NormClassif("RELIGADOR")) = "PRIMARIO"
    dFamNivel(NormClassif("RELIGADOR/SECCIONADOR")) = "PRIMARIO"
    dFamNivel(NormClassif("RELIG_OU_SEC_TRI")) = "PRIMARIO"
    dFamNivel(NormClassif("RELOCAR_ESPALHAR_BRITA")) = "NA"
    dFamNivel(NormClassif("REMAN_TF")) = "SECUNDARIO"
    dFamNivel(NormClassif("REMOCAO PAVIM")) = "NA"
    dFamNivel(NormClassif("REPARO_EQUIPAMENTOPS")) = "SECUNDARIO"
    dFamNivel(NormClassif("RESISTIVIDADE DO SOLO")) = "NA"
    dFamNivel(NormClassif("RESTAURAR PASSEIO")) = "NA"
    dFamNivel(NormClassif("RETIFICADOR")) = "SECUNDARIO"
    dFamNivel(NormClassif("RETROFIT BC")) = "NA"
    dFamNivel(NormClassif("RETROFIT CAIXA TC")) = "NA"
    dFamNivel(NormClassif("RETROFIT DISJ SE")) = "NA"
    dFamNivel(NormClassif("RETROFIT TC")) = "NA"
    dFamNivel(NormClassif("RETROFIT TF")) = "NA"
    dFamNivel(NormClassif("RETROFIT TP")) = "NA"
    dFamNivel(NormClassif("RETROFIT_RELIGADOR")) = "SECUNDARIO"
    dFamNivel(NormClassif("RETROFIT_TF")) = "NA"
    dFamNivel(NormClassif("RET_ILUM_PUBLICA")) = "NA"
    dFamNivel(NormClassif("REVESTIMENTO_PA”_BRITA")) = "NA"
    dFamNivel(NormClassif("REVESTIMENTO_PO_BRITA")) = "NA"
    dFamNivel(NormClassif("RISCO")) = "SECUNDARIO"
    dFamNivel(NormClassif("ROCA_PODA")) = "NA"
    dFamNivel(NormClassif("ROTEADOR")) = "SECUNDARIO"
    dFamNivel(NormClassif("SACO MEDIDOR")) = "NA"
    dFamNivel(NormClassif("SECCIONADOR")) = "NA"
    dFamNivel(NormClassif("SECCIONAMENTO CERCAS")) = "NA"
    dFamNivel(NormClassif("SEGUR PATRI")) = "NA"
    dFamNivel(NormClassif("SEGURANCA DE OBRAS")) = "NA"
    dFamNivel(NormClassif("SEGURANCA DO TRABALHO")) = "NA"
    dFamNivel(NormClassif("SENSOR")) = "SECUNDARIO"
End Sub

Private Sub CarregarFamiliaNivelP8()
    dFamNivel(NormClassif("SEQ_FASE")) = "NA"
    dFamNivel(NormClassif("SERV ADM")) = "NA"
    dFamNivel(NormClassif("SERV ADM MAT")) = "NA"
    dFamNivel(NormClassif("SERVICO_SE")) = "NA"
    dFamNivel(NormClassif("SERVIDAO")) = "NA"
    dFamNivel(NormClassif("SERVIDAƑO")) = "PRIMARIO"
    dFamNivel(NormClassif("SERV_AUTO")) = "PRIMARIO"
    dFamNivel(NormClassif("SERV_COMB_FRAUDE/R_PERD")) = "SECUNDARIO"
    dFamNivel(NormClassif("SERV_CONSULTORIA")) = "NA"
    dFamNivel(NormClassif("SERV_LT")) = "SECUNDARIO"
    dFamNivel(NormClassif("SERV_MEC")) = "NA"
    dFamNivel(NormClassif("SERV_SUBSTITUCAO")) = "PRIMARIO"
    dFamNivel(NormClassif("SIGFI_PAINEL SOLAR")) = "PRIMARIO"
    dFamNivel(NormClassif("SINALIZACA•ES")) = "NA"
    dFamNivel(NormClassif("SINALIZACOES")) = "NA"
    dFamNivel(NormClassif("SINCRONIZADOR")) = "NA"
    dFamNivel(NormClassif("SIST ILUM BASE RELE")) = "NA"
    dFamNivel(NormClassif("SIST ILUM REFLETOR")) = "NA"
    dFamNivel(NormClassif("SIST ILUM RELE")) = "NA"
    dFamNivel(NormClassif("SIST ILUM SUPORTE")) = "NA"
    dFamNivel(NormClassif("SIST ILUMINACAO")) = "NA"
    dFamNivel(NormClassif("SIST RETIFICACAO")) = "NA"
    dFamNivel(NormClassif("SISTEMA DE ALIMENTACAO")) = "SECUNDARIO"
    dFamNivel(NormClassif("SOFTWARE")) = "NA"
    dFamNivel(NormClassif("SOLDA")) = "NA"
    dFamNivel(NormClassif("SONDAGEM")) = "SECUNDARIO"
    dFamNivel(NormClassif("SUCATA")) = "NA"
    dFamNivel(NormClassif("SUCCAO CAMARA TR")) = "NA"
    dFamNivel(NormClassif("SUCCAO RESID CX")) = "NA"
    dFamNivel(NormClassif("SUPORTE")) = "NA"
    dFamNivel(NormClassif("SUPORTE CONC")) = "SECUNDARIO"
    dFamNivel(NormClassif("SUPORTE CS")) = "NA"
    dFamNivel(NormClassif("SUPORTE MET")) = "NA"
    dFamNivel(NormClassif("SUPORTE TC")) = "NA"
    dFamNivel(NormClassif("SUPORTE TP/TC")) = "NA"
    dFamNivel(NormClassif("SUPORTE TRAFO")) = "NA"
    dFamNivel(NormClassif("SUPORTE_CAPITEIS")) = "NA"
    dFamNivel(NormClassif("SUPORTE_TC")) = "NA"
    dFamNivel(NormClassif("SWITCH")) = "NA"
    dFamNivel(NormClassif("TABUA")) = "NA"
    dFamNivel(NormClassif("TAMPA CX PASSAGEM")) = "SECUNDARIO"
    dFamNivel(NormClassif("TAMPA DE FERRO")) = "NA"
    dFamNivel(NormClassif("TAMPA_CONCRETO")) = "SECUNDARIO"
    dFamNivel(NormClassif("TAXA DNIT")) = "NA"
    dFamNivel(NormClassif("TAXAS")) = "NA"
    dFamNivel(NormClassif("TC")) = "SECUNDARIO"
    dFamNivel(NormClassif("TC/TP")) = "PRIMARIO"
    dFamNivel(NormClassif("TC_COM")) = "NA"
    dFamNivel(NormClassif("TC_CX CONCENT")) = "NA"
    dFamNivel(NormClassif("TC_UC")) = "SECUNDARIO"
    dFamNivel(NormClassif("TELA TAPUME")) = "NA"
    dFamNivel(NormClassif("TELECOM")) = "NA"
    dFamNivel(NormClassif("TELECOMUNICACAO")) = "PRIMARIO"
    dFamNivel(NormClassif("TELECOMUNICACAƑO")) = "NA"
    dFamNivel(NormClassif("TELECONTROLE")) = "NA"
    dFamNivel(NormClassif("TENSIONAR")) = "NA"
    dFamNivel(NormClassif("TENSIONAR_ESTAI")) = "NA"
    dFamNivel(NormClassif("TERMINAL")) = "NA"
    dFamNivel(NormClassif("TERMINAL ANEL/OLHAL")) = "NA"
    dFamNivel(NormClassif("TERMINAL BIMETALICO")) = "NA"
    dFamNivel(NormClassif("TERMINAL CABO")) = "NA"
    dFamNivel(NormClassif("TERMINAL COND")) = "NA"
    dFamNivel(NormClassif("TERMINAL ILHOS")) = "NA"
    dFamNivel(NormClassif("TERMORESISTENCIA")) = "NA"
    dFamNivel(NormClassif("TERRAPLENAGEM")) = "NA"
    dFamNivel(NormClassif("TER_LEITURA")) = "NA"
    dFamNivel(NormClassif("TESTE MEGOMETRO")) = "NA"
    dFamNivel(NormClassif("TF")) = "PRIMARIO"
    dFamNivel(NormClassif("TF FORCA_BASE")) = "SECUNDARIO"
    dFamNivel(NormClassif("TI")) = "SECUNDARIO"
    dFamNivel(NormClassif("TOMADA ODS")) = "NA"
    dFamNivel(NormClassif("TOPOGRAFIA")) = "NA"
    dFamNivel(NormClassif("TORA EUCALIPTO")) = "NA"
    dFamNivel(NormClassif("TORRE")) = "PRIMARIO"
    dFamNivel(NormClassif("TP")) = "SECUNDARIO"
End Sub

Private Sub CarregarFamiliaNivelP9()
    dFamNivel(NormClassif("TP/TC AT")) = "PRIMARIO"
    dFamNivel(NormClassif("TP/TC_SUPORT CONC")) = "NA"
    dFamNivel(NormClassif("TP_CAIXA LIG")) = "NA"
    dFamNivel(NormClassif("TP_CX CONCENT")) = "NA"
    dFamNivel(NormClassif("TP_TC")) = "PRIMARIO"
    dFamNivel(NormClassif("TP_UC")) = "PRIMARIO"
    dFamNivel(NormClassif("TRAFO")) = "PRIMARIO"
    dFamNivel(NormClassif("TRAFO DE FORCA")) = "SECUNDARIO"
    dFamNivel(NormClassif("TRAFO_OU_REGULADOR")) = "PRIMARIO"
    dFamNivel(NormClassif("TRANSCEPTOR")) = "SECUNDARIO"
    dFamNivel(NormClassif("TRANSPORTE")) = "NA"
    dFamNivel(NormClassif("TRAV RIOS")) = "PRIMARIO"
    dFamNivel(NormClassif("TRAVESSIA BT")) = "NA"
    dFamNivel(NormClassif("TSA_SUPORTE")) = "NA"
    dFamNivel(NormClassif("TUBO PVC")) = "NA"
    dFamNivel(NormClassif("TUBO_PVC")) = "NA"
    dFamNivel(NormClassif("TURMA LV")) = "NA"
    dFamNivel(NormClassif("TURMA LV EST AMAR")) = "NA"
    dFamNivel(NormClassif("TURMA LV EST SUSP")) = "NA"
    dFamNivel(NormClassif("TURMA LV EST_VIGA")) = "NA"
    dFamNivel(NormClassif("TURMA LV REGUL")) = "NA"
    dFamNivel(NormClassif("TURMA LV RELIG")) = "NA"
    dFamNivel(NormClassif("TURMA LV TRAFO")) = "NA"
    dFamNivel(NormClassif("TURMA_RD_SUBTER")) = "NA"
    dFamNivel(NormClassif("UNID_CONC_COMUM")) = "NA"
    dFamNivel(NormClassif("UTR/RACK")) = "PRIMARIO"
    dFamNivel(NormClassif("VALETA_ESCOA_AGUA")) = "NA"
    dFamNivel(NormClassif("VEDACAO_DUTO")) = "SECUNDARIO"
    dFamNivel(NormClassif("VERIFICAR SEQ_FASE")) = "NA"
    dFamNivel(NormClassif("VIGA CONC")) = "SECUNDARIO"
    dFamNivel(NormClassif("VIGA MET")) = "NA"
    dFamNivel(NormClassif("VISITA TECNICA")) = "SECUNDARIO"
End Sub

Private Function FamiliaNivel(ByVal fam As String) As String
    FamiliaNivel = ""
    If dFamNivel Is Nothing Then Exit Function
    Dim k As String: k = NormClassif(fam)
    If dFamNivel.Exists(k) Then FamiliaNivel = dFamNivel(k)
End Function

' Normaliza familia/classificacao p/ casar: maiusc, sem acento, espacos colapsados
Private Function NormClassif(ByVal s As String) As String
    If dNormCache Is Nothing Then Set dNormCache = CreateObject("Scripting.Dictionary")
    If dNormCache.Exists(s) Then NormClassif = dNormCache(s): Exit Function
    Dim sOrig As String: sOrig = s
    s = UCase$(SemAcento(Trim$(CStr(s))))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    dNormCache(sOrig) = s
    NormClassif = s
End Function

' TIPO (COM/UC/UAR) a partir da CLASSIFICACAO (familia/CLS2).
' textoMat: texto breve do material (opcional). Usado para distinguir
' CH FUS 34,5kV (UC) de CH FUS em outras tensoes (COM).
Private Function TipoDaClassif(ByVal classif As String, _
                               Optional ByVal textoMat As String = "") As String
    TipoDaClassif = ""
    If dTipoCls Is Nothing Then Exit Function
    Dim k As String: k = NormClassif(classif)
    If dTipoCls.Exists(k) Then TipoDaClassif = dTipoCls(k)
    ' Regra especifica: CH FUS de 34,5 kV e UC (independente do mapeamento geral)
    If k = "CH FUS" And TipoDaClassif = "COM" Then
        Dim tx As String: tx = UCase$(Replace(textoMat, ".", ","))
        If InStr(tx, "34,5") > 0 Then TipoDaClassif = "UC"
    End If
End Function

' Unifica familias equivalentes para a aderencia MATERIAL vs SERVICO.
' COND ISOLADO e COND ISOLADO/PROT -> COND PROT.
' COND NU -> COND PROT: servico de cabo protegido cobre o material cabo nu
' e vice-versa (ambos sao condutores de linha, o servico paga os dois).
Private Function FamiliaAlias(ByVal cls2 As String) As String
    Dim k As String: k = NormClassif(cls2)
    Select Case k
        Case "COND ISOLADO/PROT", "COND ISOLADO": FamiliaAlias = "COND PROT"
        Case Else
            FamiliaAlias = cls2
            ' De-para servico -> material equivalente (CarregarEquivSrvMat)
            If Not dFamEquiv Is Nothing Then
                If dFamEquiv.Exists(k) Then FamiliaAlias = dFamEquiv(k)
            End If
    End Select
End Function

' Familia de cabo/condutor (para a regra do RAMAL cobrir o cabo)
Private Function EhCabo(ByVal cls2 As String) As Boolean
    Dim s As String: s = NormClassif(cls2)
    EhCabo = (Left$(s, 5) = "COND " Or Left$(s, 5) = "CABO " _
              Or s = "RAMAL")
End Function

' Familias que o servico de RELIGADOR torna aderentes (ate 60 m por religador)
Private Function CobertoReligador(ByVal cls2 As String) As Boolean
    Dim s As String: s = NormClassif(cls2)
    CobertoReligador = (s = "COND PROT" Or s = "CABO ISOLADO" Or s = "COND COBRE")
End Function

' Aderencia com margem de 10% (para mais ou para menos), por magnitude.
Private Function DentroMargem(ByVal a As Double, ByVal b As Double) As Boolean
    Dim x As Double, y As Double, base As Double
    x = Abs(a): y = Abs(b)
    base = x: If y > base Then base = y
    If base = 0 Then
        DentroMargem = True            ' ambos zero
    Else
        DentroMargem = (Abs(x - y) <= (CfgNum("MARGEM_ADERENCIA", 10) / 100) * base)
    End If
End Function


'==============================================================================
'  FUNES AUXILIARES (derivaes)
'==============================================================================
Private Function PEP3(ByVal pep As String) As String
    Dim s As String: s = UCase$(Right$(pep, 2))
    If s = ".I" Or s = ".D" Or s = ".M" Then
        PEP3 = Left$(pep, Len(pep) - 2)
    Else
        PEP3 = pep
    End If
End Function
' Cdigo de segmento da regional (3 letras aps "RS-"+7 dgitos). Ex: ...UNR...
Private Function SegmentoPI(ByVal pep As String) As String
    Dim s As String: s = Trim$(pep)
    If Len(s) >= 13 Then
        Dim seg As String: seg = Mid$(s, 11, 3)
        If seg Like "[A-Za-z][A-Za-z][A-Za-z]" Then SegmentoPI = UCase$(seg): Exit Function
    End If
    SegmentoPI = ""
End Function

Private Function GrupoPerc(ByVal pep As String) As String
    If EhPepEmergencia(pep) Then
        GrupoPerc = "EME/EMM (8%)"
    Else
        GrupoPerc = "OUTROS (25%)"
    End If
End Function


Private Function EhMaterial(ByVal classif As String) As Boolean
    Dim c As String: c = UCase$(Trim$(classif))
    EhMaterial = (c = "UC" Or c = "COM" Or c = "UAR" Or InStr(c, "FALTA") > 0)
End Function

Private Function ToNum(v As Variant) As Double
    If IsNumeric(v) Then ToNum = CDbl(v) Else ToNum = 0
End Function


'==============================================================================
'  ABA: RAZAO CJ  (base enriquecida com classificacoes usadas nas analises)
'==============================================================================
Private Sub Gerar_RazaoCJ()
    Dim extra As Long
    extra = 0
    If cCLS1Raw = 0 Then extra = extra + 1
    If cCLS2Raw = 0 Then extra = extra + 1
    If cCLS3Raw = 0 Then extra = extra + 1
    If cTipoAplicRaw = 0 Then extra = extra + 1

    Dim outp() As Variant: ReDim outp(0 To nLin, 1 To rawColCount + extra)
    Dim j As Long, outCol As Long
    For j = 1 To rawColCount
        outp(0, j) = rawHeaders(1, j)
    Next j

    outCol = rawColCount
    If cCLS1Raw = 0 Then outCol = outCol + 1: outp(0, outCol) = "CLS1"
    If cCLS2Raw = 0 Then outCol = outCol + 1: outp(0, outCol) = "CLS2"
    If cCLS3Raw = 0 Then outCol = outCol + 1: outp(0, outCol) = "CLS3"
    If cTipoAplicRaw = 0 Then outCol = outCol + 1: outp(0, outCol) = "TIPO_APLICACAO"

    Dim i As Long, r As Long, cls1 As String, cls2 As String, cls3 As String, ta As String
    For i = 1 To UBound(dados, 1)
        If Trim$(CStr(dados(i, cPEP))) = "" Then GoTo Prox
        r = r + 1
        For j = 1 To rawColCount
            outp(r, j) = dados(i, j)
        Next j

        If EhMaterial(CStr(dados(i, cClassif))) Then
            cls1 = MatInfoLinha(i, 1)
            cls2 = MatInfoLinha(i, 2)
            cls3 = MatInfoLinha(i, 3)
            ta = LinhaTipoAplic(i)
        Else
            cls1 = SrvInfoLinha(i, 0)
            cls2 = SrvInfoLinha(i, 1)
            cls3 = SrvInfoLinha(i, 2)
            ta = SrvInfoLinha(i, 3)
        End If

        outCol = rawColCount
        If cCLS1Raw = 0 Then outCol = outCol + 1: outp(r, outCol) = cls1
        If cCLS2Raw = 0 Then outCol = outCol + 1: outp(r, outCol) = cls2
        If cCLS3Raw = 0 Then outCol = outCol + 1: outp(r, outCol) = cls3
        If cTipoAplicRaw = 0 Then outCol = outCol + 1: outp(r, outCol) = ta
Prox:
    Next i

    EscreverAba "RAZAO CJ", outp
End Sub
'==============================================================================
'  ABA: MATERIAL vs SERVIO  (aderncia por QUANTIDADE, por PEP + CLS2)
'  Lgica oficial: para cada PEP e cada classificao de famlia (CLS2),
'  soma a QTD de material (MAT) e a QTD de servio (SRV). Compara:
'     MAT = SRV  -> ADERENTE      (material instalado tem servio que o instale)
'     MAT <> SRV -> NAO ADERENTE  (sobra material sem servio, ou vice-versa)
'     ambos 0    -> NULO
'  O CLS2 vem do catlogo de MATERIAIS (linhas de material) e do catlogo de
'  SERVIOS (linhas de CA) - os dois usam o mesmo vocabulrio de famlia.
'==============================================================================
Private Sub Gerar_MaterialVsServico()
    Dim dMat As Object, dSrv As Object, dTA As Object, dKeys As Object
    Dim dNeg As Object
    Set dMat = CreateObject("Scripting.Dictionary")   ' PEP|CLS2 -> qtd material
    Set dSrv = CreateObject("Scripting.Dictionary")   ' PEP|CLS2 -> qtd servico
    Set dTA = CreateObject("Scripting.Dictionary")    ' PEP|CLS2 -> TIPO_APLICACAO
    Set dKeys = CreateObject("Scripting.Dictionary")  ' conjunto de chaves PEP|CLS2
    Set dNeg = CreateObject("Scripting.Dictionary")   ' marca chave com qtd negativa

    ' FASE 1.2: dicionarios compartilhados com o PAINEL EXECUTIVO
    Set dMvSVerd = CreateObject("Scripting.Dictionary")
    Set dMvSFamNC = CreateObject("Scripting.Dictionary")
    Set dMvSDif = CreateObject("Scripting.Dictionary")

    Dim i As Long, pep As String, cls2 As String, k As String, q As Double, ta As String
    Dim fat As Double
    ' dPep4TemSrv: marca PEP4NIVEL que possuem pelo menos 1 lancamento de servico.
    ' Usado para bloquear regras de cobertura quando o PEP nao tem nenhum servico.
    Dim dPep4TemSrv As Object: Set dPep4TemSrv = CreateObject("Scripting.Dictionary")
    ' dFamTipo: TIPO (UC/COM) inferido do CLS3 real do catalogo de materiais
    ' (MAT. UC / MAT. COM), por familia. Usado como fallback quando a familia
    ' nao esta na tabela fixa dTipoCls -> evita falso "SEM UC".
    Dim dFamTipo As Object: Set dFamTipo = CreateObject("Scripting.Dictionary")
    ' dFamKey: FAMILIA (campo do catalogo de materiais) por chave PEP|CLS2
    Dim dFamKey As Object: Set dFamKey = CreateObject("Scripting.Dictionary")
    Dim cls3m As String, famN As String
    For i = 1 To UBound(dados, 1)
        pep = Trim$(CStr(dados(i, cPEP))): If pep = "" Then GoTo Prox
        q = ToNum(dados(i, cQtd))
        If EhMaterial(CStr(dados(i, cClassif))) Then
            cls2 = MatInfoLinha(i, 2)                  ' CLS2 do catalogo/base enriquecida
            If cls2 = "" Then cls2 = "(SEM CLS2)"
            cls2 = FamiliaAlias(cls2)                  ' unifica familias equivalentes
            ' TIPO por CLS3 do item (MAT. UC / MAT. COM); UC prevalece
            cls3m = UCase$(MatInfoLinha(i, 3))
            famN = NormClassif(cls2)
            If InStr(cls3m, "UC") > 0 Then
                dFamTipo(famN) = "UC"
            ElseIf InStr(cls3m, "COM") > 0 And dFamTipo(famN) <> "UC" Then
                dFamTipo(famN) = "COM"
            End If
            ' Ajuste fino: cabos em KG -> metros (metros = kg * fator)
            fat = CaboFator(dados(i, cMaterial))
            If fat > 0 Then q = q * fat
            k = pep & "|" & cls2
            dMat(k) = dMat(k) + q
            dKeys(k) = 1
            If q < 0 Then dNeg(k) = 1
            ' FAMILIA do catalogo de materiais (idx 0) - material tem prioridade
            Dim famM As String: famM = MatInfoLinha(i, 0)
            If famM <> "" And famM <> "(SEM FAMILIA)" Then dFamKey(k) = famM
        Else
            cls2 = SrvInfoLinha(i, 1)                  ' CLS2 do catalogo/base enriquecida
            If cls2 = "" Then cls2 = "(SEM CLS2)"
            cls2 = FamiliaAlias(cls2)                  ' unifica familias equivalentes
            ' Ajuste fino: SRV COMBO -> multiplica a qtd (cruzeta dupla=2, km*1000, ...)
            q = q * ComboFator(dados(i, cMaterial))
            k = pep & "|" & cls2
            dSrv(k) = dSrv(k) + q
            dKeys(k) = 1
            If q < 0 Then dNeg(k) = 1
            ta = SrvInfoLinha(i, 3)                    ' TIPO_APLICACAO do servico
            If ta <> "" And Not dTA.Exists(k) Then dTA(k) = ta
            ' marca que este PEP4 tem pelo menos um servico lancado
            If Not dPep4TemSrv.Exists(pep) Then dPep4TemSrv(pep) = 1
        End If
Prox:
    Next i

    Dim ks As Variant: ks = dKeys.Keys

    ' 1a passada: conta as chaves que NAO sao NULO (MAT=0 e SRV=0)
    ' -> linhas NULO nao entram no relatorio (a pedido)
    Dim r As Long, mqx As Double, sqx As Double, nKeep As Long
    For r = 0 To dKeys.Count - 1
        mqx = 0: If dMat.Exists(ks(r)) Then mqx = dMat(ks(r))
        sqx = 0: If dSrv.Exists(ks(r)) Then sqx = dSrv(ks(r))
        If Not (Round(mqx, 2) = 0 And Round(sqx, 2) = 0) Then nKeep = nKeep + 1
    Next r

    ' Veredito por PEP3: TODAS as familias UC (exceto cabos/condutores) com MAT=SRV?
    '   APROVADO = todas as familias UC tem MAT ~ SRV (margem de 10%)
    '   REPROVADO = pelo menos uma familia UC fora da margem
    '   Condutores/cabos sao avaliados normalmente (com conversao de cabo,
    '   regra do ramal e margem de 10%).
    ' Veredito calculado por PEP NIVEL 4 (PEP completo, com .I/.D), nao por PEP3.
    Dim dVerd As Object: Set dVerd = CreateObject("Scripting.Dictionary")
    Dim dUAR As Object: Set dUAR = CreateObject("Scripting.Dictionary")   ' PEP4 -> tem UAR
    Dim dPep4FamRep As Object: Set dPep4FamRep = CreateObject("Scripting.Dictionary")   ' PEP4 -> familias UC reprovadas
    Dim pv As Variant, pep4v As String, clsv As String, mv As Double, sv As Double, tcv As String

    ' Regra do RAMAL: 1 ramal = 30 m. Se SRV(RAMAL)*30 cobrir (>=) o total de cabo
    ' do PEP4, os cabos/condutores desse PEP sao considerados aderentes.
    Dim dRamalM As Object: Set dRamalM = CreateObject("Scripting.Dictionary")
    Dim dCaboMat As Object: Set dCaboMat = CreateObject("Scripting.Dictionary")
    Dim cabCov As Object: Set cabCov = CreateObject("Scripting.Dictionary")
    For r = 0 To dKeys.Count - 1
        pv = Split(ks(r), "|")
        pep4v = CStr(pv(0)): clsv = CStr(pv(1))
        If NormClassif(clsv) = "RAMAL" Then
            sv = 0: If dSrv.Exists(ks(r)) Then sv = dSrv(ks(r))
            dRamalM(pep4v) = dRamalM(pep4v) + Abs(sv) * 30
        End If
        If EhCabo(clsv) Then
            mv = 0: If dMat.Exists(ks(r)) Then mv = dMat(ks(r))
            dCaboMat(pep4v) = dCaboMat(pep4v) + Abs(mv)
        End If
    Next r
    Dim pk As Variant
    For Each pk In dRamalM.Keys
        If dRamalM(pk) > 0 And dRamalM(pk) >= dCaboMat(pk) Then cabCov(pk) = 1
    Next pk

    ' Regra do RELIGADOR: cada servico de religador libera 60 m para tornar
    ' aderentes COND PROT, CABO ISOLADO e COND COBRE (2 servicos = 120 m, etc.).
    Dim dRelSrv As Object: Set dRelSrv = CreateObject("Scripting.Dictionary")
    Dim dRelMat As Object: Set dRelMat = CreateObject("Scripting.Dictionary")
    Dim relCov As Object: Set relCov = CreateObject("Scripting.Dictionary")
    For r = 0 To dKeys.Count - 1
        pv = Split(ks(r), "|")
        pep4v = CStr(pv(0)): clsv = CStr(pv(1))
        If NormClassif(clsv) = "RELIGADOR" Then
            sv = 0: If dSrv.Exists(ks(r)) Then sv = dSrv(ks(r))
            dRelSrv(pep4v) = dRelSrv(pep4v) + Abs(sv)
        End If
        If CobertoReligador(clsv) Then
            mv = 0: If dMat.Exists(ks(r)) Then mv = dMat(ks(r))
            dRelMat(pep4v) = dRelMat(pep4v) + Abs(mv)
        End If
    Next r
    For Each pk In dRelSrv.Keys
        If dRelSrv(pk) > 0 And dRelMat(pk) <= dRelSrv(pk) * 60 Then relCov(pk) = 1
    Next pk

    ' TRAFO / REGULADOR: o servico de TRAFO ou REGULADOR torna aderentes
    ' CH FUS, PARA RAIO MT e PARA RAIO BT no mesmo PEP4 (ODD e ODI).
    ' Tambem: servico TRAFO_OU_REGULADOR cobre material TRAFO/REGULADOR e vice-versa.
    Dim trafRegCov As Object:    Set trafRegCov    = CreateObject("Scripting.Dictionary")
    Dim chFusCov As Object:      Set chFusCov      = CreateObject("Scripting.Dictionary")   ' PEP4 com servico que paga CH fusivel
    Dim dTrafOrRegSrv As Object: Set dTrafOrRegSrv = CreateObject("Scripting.Dictionary")
    Dim dTrafHasMatPep As Object: Set dTrafHasMatPep = CreateObject("Scripting.Dictionary")
    For r = 0 To dKeys.Count - 1
        pv = Split(ks(r), "|")
        pep4v = CStr(pv(0)): clsv = CStr(pv(1))
        Select Case NormClassif(clsv)
            Case "TRAFO", "REGULADOR"
                sv = 0: If dSrv.Exists(ks(r)) Then sv = dSrv(ks(r))
                If Abs(sv) > 0 Then trafRegCov(pep4v) = 1
                mv = 0: If dMat.Exists(ks(r)) Then mv = dMat(ks(r))
                If Abs(mv) > 0 Then dTrafHasMatPep(pep4v) = 1
            Case "TRAFO OU REGULADOR", "TRAFO_OU_REGULADOR"
                sv = 0: If dSrv.Exists(ks(r)) Then sv = dSrv(ks(r))
                If Abs(sv) > 0 Then
                    trafRegCov(pep4v) = 1
                    dTrafOrRegSrv(pep4v) = dTrafOrRegSrv(pep4v) + Abs(sv)
                End If
        End Select
    Next r
    ' Servicos especificos que cobrem acessorios de TRAFO (CH FUS, PARA RAIO MT/BT):
    ' 5500200645 - servico de inst/troca de trafo, paga os acessorios.
    ' d572Fam: pep4 -> familia do servico 5500000572 (para regra SRV>MAT=REPROVADO)
    Dim d572Fam As Object: Set d572Fam = CreateObject("Scripting.Dictionary")
    Dim fam572 As String
    Dim iSE As Long, pepSE As String, codSE As String, qSE As Double
    For iSE = 1 To UBound(dados, 1)
        pepSE = Trim$(CStr(dados(iSE, cPEP))): If pepSE = "" Then GoTo ProxSE
        If EhMaterial(CStr(dados(iSE, cClassif))) Then GoTo ProxSE
        codSE = NormCod(dados(iSE, cMaterial))
        If codSE = "5500200645" Then
            qSE = ToNum(dados(iSE, cQtd))
            If Abs(qSE) > 0 Then trafRegCov(pepSE) = 1
        End If
        ' 5500000618 / 5500200645: servico que paga a CH fusivel (material 105300003)
        If codSE = "5500000618" Or codSE = "5500200645" Then
            qSE = ToNum(dados(iSE, cQtd))
            If Abs(qSE) > 0 Then chFusCov(pepSE) = 1
        End If
        ' 5500000572: registra a familia para checar SRV > MAT depois
        If codSE = "5500000572" Then
            qSE = ToNum(dados(iSE, cQtd))
            If Abs(qSE) > 0 Then
                fam572 = FamiliaAlias(SrvInfoLinha(iSE, 1))
                If fam572 <> "" Then d572Fam(pepSE) = fam572
            End If
        End If
ProxSE:
    Next iSE
    ' COND COBRE: PEP4 que tem PARA RAIO BT/MT, REGULADOR ou TRAFO (MAT ou SRV)
    ' -> COND COBRE aderente (e conexao de aterramento desses equipamentos).
    Dim dEquipPep As Object: Set dEquipPep = CreateObject("Scripting.Dictionary")
    Dim iEQ As Long, pepEQ As String, clsEQ As String
    For iEQ = 1 To UBound(dados, 1)
        pepEQ = Trim$(CStr(dados(iEQ, cPEP))): If pepEQ = "" Then GoTo ProxEQ
        If EhMaterial(CStr(dados(iEQ, cClassif))) Then
            clsEQ = NormClassif(FamiliaAlias(MatInfoLinha(iEQ, 2)))
        Else
            clsEQ = NormClassif(FamiliaAlias(SrvInfoLinha(iEQ, 1)))
        End If
        Select Case clsEQ
            Case "PARA RAIO BT", "PARA RAIO MT", "REGULADOR", "TRAFO"
                dEquipPep(pepEQ) = 1
        End Select
ProxEQ:
    Next iEQ

    ' COND NU + CORDOALHA: o servico de COND NU contempla o material CORDOALHA.
    ' CH SU FACA: a SRV de CH SU FACA fica disponivel para cobrir COND COBRE
    ' (conexao da CH SU FACA), como fallback quando nada mais cobre.
    Dim dCondNuSrv As Object: Set dCondNuSrv = CreateObject("Scripting.Dictionary")
    Dim dCordMat As Object: Set dCordMat = CreateObject("Scripting.Dictionary")
    Dim cordCondCov As Object: Set cordCondCov = CreateObject("Scripting.Dictionary")
    Dim dChFacaSrv As Object: Set dChFacaSrv = CreateObject("Scripting.Dictionary")
    For r = 0 To dKeys.Count - 1
        pv = Split(ks(r), "|")
        pep4v = CStr(pv(0)): clsv = CStr(pv(1))
        Select Case NormClassif(clsv)
            Case "COND NU"
                sv = 0: If dSrv.Exists(ks(r)) Then sv = dSrv(ks(r))
                dCondNuSrv(pep4v) = dCondNuSrv(pep4v) + Abs(sv)
            Case "CORDOALHA"
                mv = 0: If dMat.Exists(ks(r)) Then mv = dMat(ks(r))
                dCordMat(pep4v) = dCordMat(pep4v) + Abs(mv)
            Case "CH SU FACA"
                sv = 0: If dSrv.Exists(ks(r)) Then sv = dSrv(ks(r))
                dChFacaSrv(pep4v) = dChFacaSrv(pep4v) + Abs(sv)
        End Select
    Next r
    For Each pk In dCondNuSrv.Keys
        If dCondNuSrv(pk) > 0 And dCordMat(pk) > 0 Then cordCondCov(pk) = 1
    Next pk

    For r = 0 To dKeys.Count - 1
        pv = Split(ks(r), "|")
        clsv = CStr(pv(1))
        pep4v = CStr(pv(0))                     ' PEP completo (nivel 4)
        tcv = UCase$(TipoDaClassif(clsv))
        ' Fallback: familia sem TIPO na tabela fixa -> usa CLS3 do catalogo
        ' (MAT. UC / MAT. COM). Corrige falso "SEM UC" em familias nao mapeadas.
        If tcv = "" Then
            If dFamTipo.Exists(NormClassif(clsv)) Then tcv = dFamTipo(NormClassif(clsv))
        End If
        If tcv = "UAR" Then dUAR(pep4v) = 1     ' PEP4 possui familia UAR
        If tcv = "UC" Then
            mv = 0: If dMat.Exists(ks(r)) Then mv = dMat(ks(r))
            sv = 0: If dSrv.Exists(ks(r)) Then sv = dSrv(ks(r))
            Dim aderiu As Boolean: aderiu = False
            ' cabos cobertos pelo ramal (>=) ou religador (60 m/un)
            If EhCabo(clsv) And cabCov.Exists(pep4v) And dPep4TemSrv.Exists(pep4v) Then aderiu = True
            If CobertoReligador(clsv) And relCov.Exists(pep4v) And dPep4TemSrv.Exists(pep4v) Then aderiu = True
            ' cabo com quantidade < 15 m -> negligenciavel, aderente
            ' (so aplica se ha SRV lancado E o SRV da linha nao e zero)
            If EhCabo(clsv) And dPep4TemSrv.Exists(pep4v) And sv <> 0 Then
                Dim qmax As Double
                qmax = Abs(mv): If Abs(sv) > qmax Then qmax = Abs(sv)
                If qmax < 15 Then aderiu = True
            End If
            ' RAMAL: cada servico de ramal contempla 75 m do material ramal
            If NormClassif(clsv) = "RAMAL" And sv <> 0 _
               And Abs(mv) <= Abs(sv) * 75 Then aderiu = True
            ' COND NU + CORDOALHA: servico de COND NU contempla a CORDOALHA
            If NormClassif(clsv) = "COND NU" And cordCondCov.Exists(pep4v) Then aderiu = True
            ' PI = ERD: cada servico de religador deixa aderente ate 6 CH SU FACA
            If NormClassif(clsv) = "CH SU FACA" _
               And SegmentoPI(pep4v) = "ERD" _
               And Abs(mv) <= dRelSrv(pep4v) * 6 Then aderiu = True
            ' Fallback: COND COBRE como conexao da CH SU FACA -> servico de CH SU
            ' FACA torna aderente quando nenhuma outra regra cobriu.
            If Not aderiu And NormClassif(clsv) = "COND COBRE" _
               And dChFacaSrv(pep4v) > 0 Then aderiu = True
            ' COND COBRE como aterramento de PARA RAIO BT/MT, REGULADOR ou TRAFO.
            If Not aderiu And NormClassif(clsv) = "COND COBRE" _
               And dEquipPep.Exists(pep4v) Then aderiu = True
            ' TRAFO / REGULADOR: servico de trafo/regulador cobre CH FUS,
            ' PARA RAIO MT e PARA RAIO BT (o servico paga pelos acessorios).
            If Not aderiu And trafRegCov.Exists(pep4v) Then
                Select Case NormClassif(clsv)
                    Case "CH FUS", "PARA RAIO MT", "PARA RAIO BT"
                        aderiu = True
                End Select
            End If
            ' CH FUS paga por servico especifico (5500000618 / 5500200645)
            If Not aderiu And chFusCov.Exists(pep4v) And NormClassif(clsv) = "CH FUS" Then aderiu = True
            ' TRAFO / REGULADOR (MAT) coberto por SRV de TRAFO_OU_REGULADOR no PEP.
            If Not aderiu And dTrafOrRegSrv.Exists(pep4v) And dTrafHasMatPep.Exists(pep4v) _
               And (NormClassif(clsv) = "TRAFO" Or NormClassif(clsv) = "REGULADOR") Then aderiu = True
            ' CH SU FACA: cada servico paga 7 m de conexao de cabo (COND PROT etc.)
            If Not aderiu And EhCabo(clsv) And dChFacaSrv(pep4v) > 0 _
               And Abs(mv) <= Abs(sv) + dChFacaSrv(pep4v) * 7 Then aderiu = True
            If Not dVerd.Exists(pep4v) Then dVerd(pep4v) = "APROVADO"
            If Not aderiu Then
                If Not DentroMargem(mv, sv) Then
                    dVerd(pep4v) = "REPROVADO"
                    Dim fN As String: fN = NormClassif(clsv)
                    Dim curF As String: curF = CStr(dPep4FamRep(pep4v))
                    If curF = "" Then
                        dPep4FamRep(pep4v) = fN
                    ElseIf InStr(", " & curF & ", ", ", " & fN & ", ") = 0 Then
                        dPep4FamRep(pep4v) = curF & ", " & fN
                    End If
                End If
            End If
        End If
    Next r

    ' -----------------------------------------------------------------------
    ' PROPAGACAO DE REPROVADO: se a ODI (.I) de um PEP3NIVEL for REPROVADA,
    ' todos os outros PEP4NIVEL do mesmo PEP3NIVEL tambem ficam REPROVADOS.
    ' Logica: 1) identifica quais PEP3NIVEL tem pelo menos uma ODI reprovada
    '         2) marca todos os PEP4NIVEL desse PEP3NIVEL como REPROVADO
    ' -----------------------------------------------------------------------
    Dim dPep3Reprov As Object: Set dPep3Reprov = CreateObject("Scripting.Dictionary")

    ' 1a passagem: detecta PEP3NIVEL com ODI reprovada OU sem UC.
    ' Regra: ODI SEM UC = nenhuma familia UC no PEP4 -> PEP3 REPROVADO.
    Dim dPep3Motiv As Object: Set dPep3Motiv = CreateObject("Scripting.Dictionary")
    Dim kk As Variant
    For Each kk In dVerd.Keys
        Dim pep4k As String: pep4k = CStr(kk)
        If TipoPEPCodigo(pep4k) = "I" Then          ' e ODI (.I)
            If dVerd(pep4k) = "REPROVADO" Then
                Dim pep3k As String: pep3k = PEP3(pep4k)
                dPep3Reprov(pep3k) = 1               ' marca o PEP3NIVEL
                Dim msgA As String: msgA = "ODI " & pep4k & " reprovada"
                If CStr(dPep4FamRep(pep4k)) <> "" Then _
                    msgA = msgA & " (familia: " & CStr(dPep4FamRep(pep4k)) & ")"
                If Not dPep3Motiv.Exists(pep3k) Then
                    dPep3Motiv(pep3k) = msgA
                ElseIf InStr(CStr(dPep3Motiv(pep3k)), msgA) = 0 Then
                    dPep3Motiv(pep3k) = CStr(dPep3Motiv(pep3k)) & " ; " & msgA
                End If
            End If
        End If
    Next kk
    ' 1b passagem: ODI SEM UC (nunca entrou em dVerd) tambem reprova o PEP3.
    ' Coleta todos os PEP4 ODI que aparecem em dKeys mas nao tem familia UC nem UAR.
    Dim dPep4ODI As Object: Set dPep4ODI = CreateObject("Scripting.Dictionary")
    For r = 0 To dKeys.Count - 1
        pv = Split(ks(r), "|")
        pep4v = CStr(pv(0))
        If TipoPEPCodigo(pep4v) = "I" Then dPep4ODI(pep4v) = 1
    Next r
    For Each kk In dPep4ODI.Keys
        pep4k = CStr(kk)
        If Not dVerd.Exists(pep4k) And Not dUAR.Exists(pep4k) Then
            pep3k = PEP3(pep4k)
            dPep3Reprov(pep3k) = 1                   ' ODI sem UC reprova o PEP3
            Dim msgB As String: msgB = "ODI " & pep4k & " sem familia UC"
            If Not dPep3Motiv.Exists(pep3k) Then
                dPep3Motiv(pep3k) = msgB
            ElseIf InStr(CStr(dPep3Motiv(pep3k)), msgB) = 0 Then
                dPep3Motiv(pep3k) = CStr(dPep3Motiv(pep3k)) & " ; " & msgB
            End If
        End If
    Next kk

    ' 1c passagem: servico 5500000572 com SRV > MAT na familia -> PEP3 REPROVADO.
    Dim k572 As String, srv572 As Double, mat572 As Double, msg572 As String
    For Each kk In d572Fam.Keys
        k572 = CStr(kk) & "|" & CStr(d572Fam(kk))
        srv572 = 0: If dSrv.Exists(k572) Then srv572 = dSrv(k572)
        mat572 = 0: If dMat.Exists(k572) Then mat572 = dMat(k572)
        If Abs(srv572) > Abs(mat572) Then
            pep3k = PEP3(CStr(kk))
            dPep3Reprov(pep3k) = 1
            msg572 = "SRV 5500000572 excede MAT no PEP " & CStr(kk)
            If Not dPep3Motiv.Exists(pep3k) Then
                dPep3Motiv(pep3k) = msg572
            ElseIf InStr(CStr(dPep3Motiv(pep3k)), "5500000572") = 0 Then
                dPep3Motiv(pep3k) = CStr(dPep3Motiv(pep3k)) & " ; " & msg572
            End If
        End If
    Next kk

    ' 2a passagem: propaga REPROVADO para todos os PEP4NIVEL do PEP3NIVEL marcado
    If dPep3Reprov.Count > 0 Then
        For Each kk In dVerd.Keys
            pep4k = CStr(kk)
            pep3k = PEP3(pep4k)
            If dPep3Reprov.Exists(pep3k) Then
                dVerd(pep4k) = "REPROVADO"
            End If
        Next kk
        ' Garante que PEP4NIVEL sem entrada em dVerd (SEM UC / UAR) tambem
        ' sejam marcados, iterando as chaves do conjunto completo
        For r = 0 To dKeys.Count - 1
            pv = Split(ks(r), "|")
            pep4v = CStr(pv(0))
            pep3k = PEP3(pep4v)
            If dPep3Reprov.Exists(pep3k) Then
                dVerd(pep4v) = "REPROVADO"
            End If
        Next r
    End If

    Dim outp() As Variant: ReDim outp(0 To nKeep, 1 To 15)
    outp(0, 1) = "STATUS_PEP4": outp(0, 2) = "PEP3NIVEL": outp(0, 3) = "PEP4NIVEL"
    outp(0, 4) = "CLS2": outp(0, 5) = "TIPO"
    outp(0, 6) = "MAT": outp(0, 7) = "SRV": outp(0, 8) = "DIFERENCA"
    outp(0, 9) = "SITUACAO": outp(0, 10) = "TIPO_PEP": outp(0, 11) = "PI"
    outp(0, 12) = "OBS1": outp(0, 13) = "OBS2": outp(0, 14) = "FAMILIA": outp(0, 15) = "NVL_FAM"

    Dim rr As Long, p As Variant, pp As String, mq As Double, sq As Double, dif As Double
    Dim p3MvS As String
    rr = 0
    For r = 0 To dKeys.Count - 1
        mq = 0: If dMat.Exists(ks(r)) Then mq = dMat(ks(r))
        sq = 0: If dSrv.Exists(ks(r)) Then sq = dSrv(ks(r))
        ' Pula NULO: nao traz a informacao da linha
        If Round(mq, 2) = 0 And Round(sq, 2) = 0 Then GoTo PulaNulo
        rr = rr + 1
        p = Split(ks(r), "|")
        pp = CStr(p(0))
        dif = Round(Abs(Abs(mq) - Abs(sq)), 2)   ' compara magnitude (estornos negativos)
        ' col 1: veredito do PEP NIVEL 4 (PEP completo)
        ' Regra de propagacao: se a ODI do PEP3NIVEL for REPROVADA, todos os
        ' PEP4NIVEL do mesmo PEP3NIVEL (incluindo ODD) ficam REPROVADOS.
        If dPep3Reprov.Exists(PEP3(pp)) Then
            outp(rr, 1) = "REPROVADO"  ' ODI reprovada contamina todo o PEP3NIVEL
        ElseIf TipoPEPCodigo(pp) = "D" Then
            outp(rr, 1) = "APROVADO"   ' ODD (.D) nao exige aderencia de UC
        ElseIf dVerd.Exists(pp) Then
            outp(rr, 1) = dVerd(pp)
        ElseIf dUAR.Exists(pp) Then
            outp(rr, 1) = "APROVADO"   ' PEP4 sem UC mas com UAR -> aprovado
        Else
            outp(rr, 1) = "SEM UC"     ' PEP4 sem nenhuma familia UC
        End If
        outp(rr, 2) = PEP3(pp)
        outp(rr, 3) = pp
        outp(rr, 4) = p(1)
        ' col 5: TIPO (COM/UC/UAR); vazio -> SERV
        Dim tcls As String: tcls = TipoDaClassif(CStr(p(1)))
        If tcls = "" Then tcls = "SERV"
        outp(rr, 5) = tcls
        outp(rr, 6) = Round(mq, 2)
        outp(rr, 7) = Round(sq, 2)
        ' SITUACAO: reflete a realidade dos numeros MAT vs SRV para todas as linhas,
        ' incluindo ODD. O STATUS_PEP4 ja trata a isencao da ODD separadamente.
        ' Regras de cobertura (cabo, ramal, religador, etc.) so se aplicam se o
        ' PEP4 tiver pelo menos um lancamento de servico (dPep4TemSrv).
        Dim pepTemSrv As Boolean: pepTemSrv = dPep4TemSrv.Exists(pp)
        If pepTemSrv And EhCabo(CStr(p(1))) And cabCov.Exists(pp) Then
            outp(rr, 8) = 0
            outp(rr, 9) = "ADERENTE"
        ElseIf pepTemSrv And CobertoReligador(CStr(p(1))) And relCov.Exists(pp) Then
            outp(rr, 8) = 0
            outp(rr, 9) = "ADERENTE"
        ElseIf pepTemSrv And EhCabo(CStr(p(1))) _
               And sq <> 0 _
               And IIf(Abs(mq) > Abs(sq), Abs(mq), Abs(sq)) < 15 Then
            outp(rr, 8) = dif
            outp(rr, 9) = "ADERENTE"   ' cabo < 15 m (so aplica se ha SRV lancado)
        ElseIf pepTemSrv And NormClassif(CStr(p(1))) = "RAMAL" And sq <> 0 _
               And Abs(mq) <= Abs(sq) * 75 Then
            outp(rr, 8) = dif
            outp(rr, 9) = "ADERENTE"   ' RAMAL coberto pelo proprio servico (75 m/un)
        ElseIf pepTemSrv And (NormClassif(CStr(p(1))) = "COND NU" _
                Or NormClassif(CStr(p(1))) = "CORDOALHA") _
               And cordCondCov.Exists(pp) Then
            outp(rr, 8) = dif
            outp(rr, 9) = "ADERENTE"   ' COND NU + CORDOALHA: servico contempla
        ElseIf pepTemSrv And NormClassif(CStr(p(1))) = "CH SU FACA" _
               And SegmentoPI(pp) = "ERD" _
               And Abs(mq) <= dRelSrv(pp) * 6 Then
            outp(rr, 8) = dif
            outp(rr, 9) = "ADERENTE"   ' PI ERD: religador cobre ate 6 CH SU FACA por servico
        ElseIf pepTemSrv And NormClassif(CStr(p(1))) = "COND COBRE" _
               And dChFacaSrv(pp) > 0 Then
            outp(rr, 8) = dif
            outp(rr, 9) = "ADERENTE"   ' COND COBRE eh conexao da CH SU FACA
        ElseIf NormClassif(CStr(p(1))) = "COND COBRE" _
               And dEquipPep.Exists(pp) Then
            outp(rr, 8) = dif
            outp(rr, 9) = "ADERENTE"   ' COND COBRE eh aterramento de PARA RAIO/TRAFO/REG
        ElseIf pepTemSrv And EhCabo(CStr(p(1))) And dChFacaSrv(pp) > 0 _
               And Abs(mq) <= Abs(sq) + dChFacaSrv(pp) * 7 Then
            outp(rr, 8) = dif
            outp(rr, 9) = "ADERENTE"   ' cabo de conexao da CH SU FACA (7 m por servico)
        ElseIf trafRegCov.Exists(pp) And _
               (NormClassif(CStr(p(1))) = "CH FUS" Or _
                NormClassif(CStr(p(1))) = "PARA RAIO MT" Or _
                NormClassif(CStr(p(1))) = "PARA RAIO BT") Then
            outp(rr, 8) = dif
            outp(rr, 9) = "ADERENTE"   ' acessorio coberto por SRV de TRAFO/REGULADOR
        ElseIf chFusCov.Exists(pp) And NormClassif(CStr(p(1))) = "CH FUS" Then
            outp(rr, 8) = dif
            outp(rr, 9) = "ADERENTE"   ' CH fusivel paga por servico (5500000618/5500200645)
        ElseIf dTrafOrRegSrv.Exists(pp) And dTrafHasMatPep.Exists(pp) And _
               (NormClassif(CStr(p(1))) = "TRAFO" Or _
                NormClassif(CStr(p(1))) = "REGULADOR") Then
            outp(rr, 8) = dif
            outp(rr, 9) = "ADERENTE"   ' SRV TRAFO_OU_REGULADOR cobre MAT de TRAFO/REGULADOR
        ElseIf dTrafHasMatPep.Exists(pp) And _
               (NormClassif(CStr(p(1))) = "TRAFO OU REGULADOR" Or _
                NormClassif(CStr(p(1))) = "TRAFO_OU_REGULADOR") Then
            outp(rr, 8) = dif
            outp(rr, 9) = "ADERENTE"   ' MAT de TRAFO/REGULADOR existe -> SRV TRAFO_OU_REG aderente
        ElseIf DentroMargem(mq, sq) Then
            outp(rr, 8) = dif
            outp(rr, 9) = "ADERENTE"
        Else
            outp(rr, 8) = dif
            outp(rr, 9) = "NAO ADERENTE"
        End If
        outp(rr, 10) = TipoPEPCodigo(pp)
        outp(rr, 11) = SegmentoPI(pp)

        ' OBS1: motivo do REPROVADO
        Dim obs1 As String: obs1 = ""
        Dim p3obs As String: p3obs = PEP3(pp)
        ' a) linha propria nao aderente (UC fora da margem)
        If CStr(outp(rr, 9)) = "NAO ADERENTE" And _
           UCase$(TipoDaClassif(CStr(p(1)))) = "UC" Then
            obs1 = "Familia UC fora da margem (MAT=" & _
                   Format$(Round(mq, 2), "0.##") & _
                   " / SRV=" & Format$(Round(sq, 2), "0.##") & ")"
        End If
        ' b) reprovado por propagacao (ODI de outro PEP4 ou SEM UC)
        If CStr(outp(rr, 1)) = "REPROVADO" And dPep3Motiv.Exists(p3obs) Then
            Dim motProp As String: motProp = CStr(dPep3Motiv(p3obs))
            If obs1 = "" Then
                obs1 = "Reprovado por: " & motProp
            ElseIf InStr(obs1, motProp) = 0 Then
                obs1 = obs1 & " | Reprovado por: " & motProp
            End If
        End If
        outp(rr, 12) = obs1

        ' OBS2: motivo do veredito do PEP3 (mesmo texto em TODAS as linhas do
        ' PEP3, para leitura do porque aprovou/reprovou).
        Dim obs2 As String
        If dPep3Reprov.Exists(p3obs) Then
            If dPep3Motiv.Exists(p3obs) Then
                obs2 = "PEP3 REPROVADO: " & CStr(dPep3Motiv(p3obs))
            Else
                obs2 = "PEP3 REPROVADO: familia UC fora da margem"
            End If
        Else
            obs2 = "PEP3 APROVADO: nenhuma familia UC do PEP3 fora da margem de " & _
                   Format$(CfgNum("MARGEM_ADERENCIA", 10), "0.##") & "%"
        End If
        outp(rr, 13) = obs2
        ' FAMILIA (campo do catalogo de materiais); vazio p/ chave so de servico
        outp(rr, 14) = IIf(dFamKey.Exists(ks(r)), CStr(dFamKey(ks(r))), "")
        outp(rr, 15) = FamiliaNivel(CStr(p(1)))   ' nivel da familia (FAMILIAS_ATUAIS)

        ' -------------------------------------------------------------------
        ' FASE 1.2: alimenta os vereditos ODI em memoria, replicando a regra
        ' que o PAINEL EXECUTIVO aplicava ao ler a planilha:
        '   - somente linhas com TIPO_PEP = "I" (ODI)
        '   - qualquer ODI REPROVADA -> PEP3 REPROVADO
        '   - SITUACAO = "NAO ADERENTE" -> conta familia e soma DIFERENCA
        ' -------------------------------------------------------------------
        If CStr(outp(rr, 10)) = "I" Then
            p3MvS = CStr(outp(rr, 2))
            If CStr(outp(rr, 1)) = "REPROVADO" Then
                dMvSVerd(p3MvS) = "REPROVADO"
            ElseIf Not dMvSVerd.Exists(p3MvS) Then
                dMvSVerd(p3MvS) = "APROVADO"
            End If
            If CStr(outp(rr, 9)) = "NAO ADERENTE" Then
                dMvSFamNC(p3MvS) = dMvSFamNC(p3MvS) + 1
                dMvSDif(p3MvS) = dMvSDif(p3MvS) + ToNum(outp(rr, 8))
            End If
        End If
PulaNulo:
    Next r
    EscreverAba "MATERIAL vs SERVICO", outp
End Sub


'==============================================================================
'  ABA: ANALISE DE CA  (CA por PEP x Descrio SA)
'==============================================================================
Private Sub Gerar_AnaliseCA()
    Dim cats As Variant
    cats = Array("JOA", "LOGISTICA", "TRIBUTOS", "SUPORTE", "FRETE_TRANSP", _
                 "PUBLICIDADE", "MEIO AMBIENTE", "APOIO A OBRA", "COMISSIONAMENTO", _
                 "FISCALIZACAO", "PROJETO", "PROJETO_DRT_OBRA", "ATIVACAO DIRETA", _
                 "MAO DE OBRA CIVIL", "DISPONIBILIDADE", "MAO DE OBRA", "MOP", _
                 "RISCO", "MAT UC", "MAT COM", "VALOR MATERIAL", "OUTROS", _
                 "BAIXA ODD ODS", "CLASSIFICAR")

    Dim dPEP As Object, dGrp As Object, dPepTot As Object, dEmp As Object, dMaoCLS3 As Object
    Set dPEP = CreateObject("Scripting.Dictionary")
    Set dGrp = CreateObject("Scripting.Dictionary")
    Set dPepTot = CreateObject("Scripting.Dictionary")
    Set dEmp = CreateObject("Scripting.Dictionary")
    Set dMaoCLS3 = CreateObject("Scripting.Dictionary")

    Dim i As Long, pep As String, cat As String, val As Double, k As String
    For i = 1 To UBound(dados, 1)
        pep = Trim$(CStr(dados(i, cPEP))): If pep = "" Then GoTo P1
        val = ToNum(dados(i, cValor))
        If Not dPEP.Exists(pep) Then dPEP.Add pep, dPEP.Count
        If Not dEmp.Exists(pep) Then dEmp(pep) = TextoCampo(i, cEmpresa)
        dPepTot(pep) = dPepTot(pep) + val
        ' MAO DE OBRA pela CLS3 (fonte direta da classificacao do servico)
        If UCase$(SemAcento(Trim$(LinhaCLS3(i)))) = "MAO DE OBRA" Then _
            dMaoCLS3(pep) = dMaoCLS3(pep) + val

        cat = CategoriaAnaliseCA(i)
        k = pep & "|" & cat
        dGrp(k) = dGrp(k) + val

        If cat = "MAT UC" Or cat = "MAT COM" Then
            dGrp(pep & "|VALOR MATERIAL") = dGrp(pep & "|VALOR MATERIAL") + val
        End If

        ' Acumula ATIVACAO DIRETA quando o servico tem CLS2 = "ATV DRT"
        ' (mesmo quando ja foi classificado em FISCALIZACAO / APOIO A OBRA /
        '  COMISSIONAMENTO / PROJETO pelo CLS3).
        If Not EhMaterial(CStr(dados(i, cClassif))) Then
            Dim cls2srv As String
            cls2srv = UCase$(SemAcento(SrvInfoLinha(i, 1)))
            If (InStr(cls2srv, "ATV DRT") > 0 Or InStr(cls2srv, "ATIVACAO") > 0) _
               And cat <> "ATIVACAO DIRETA" Then
                dGrp(pep & "|ATIVACAO DIRETA") = dGrp(pep & "|ATIVACAO DIRETA") + val
            End If
        End If
P1:
    Next i

    If dPEP.Count = 0 Then Exit Sub
    Dim peps As Variant: peps = dPEP.Keys
    Dim outp() As Variant: ReDim outp(0 To dPEP.Count, 1 To 39)
    outp(0, 1) = "EMPRESA": outp(0, 2) = "PEP4NIVEL"
    Dim j As Long
    For j = 0 To UBound(cats): outp(0, 3 + j) = cats(j): Next j
    outp(0, 27) = "VALOR TOTAL": outp(0, 28) = "PORC ATV DRT 25%"
    outp(0, 29) = "PORC MOP": outp(0, 30) = "UC MENOR 10%"
    outp(0, 31) = "TIPO PEP"
    outp(0, 32) = "ATV PREVISTA": outp(0, 33) = "DIF ATV DRT"
    outp(0, 34) = "REF FRETE": outp(0, 35) = "PI"
    outp(0, 36) = "TOTAL CA SEM MOP": outp(0, 37) = "CALCULO DO MOP"
    outp(0, 38) = "DIF MOP"
    outp(0, 39) = "MAO DE OBRA CLS3"

    Dim r As Long, tot As Double, v As Double, mao As Double, atv As Double
    Dim mop As Double, totalSemMop As Double, matUc As Double, atvPrev As Double, calcMop As Double
    For r = 0 To dPEP.Count - 1
        pep = peps(r)
        outp(r + 1, 1) = dEmp(pep)
        outp(r + 1, 2) = pep
        For j = 0 To UBound(cats)
            k = pep & "|" & cats(j): v = 0
            If dGrp.Exists(k) Then v = dGrp(k)
            outp(r + 1, 3 + j) = Round(v, 2)
        Next j
        ' VALOR TOTAL: soma real de todos os lancamentos do PEP (fonte unica)
        tot = dPepTot(pep)
        mao = ValorCat(dGrp, pep, "MAO DE OBRA")
        atv = ValorCat(dGrp, pep, "ATIVACAO DIRETA")
        mop = ValorCat(dGrp, pep, "MOP")
        matUc = ValorCat(dGrp, pep, "MAT UC")
        totalSemMop = tot - mop
        ' ATV PREVISTA centralizada (8% EME/EMM, 25% demais) - FASE 1.1
        atvPrev = AtvPrevista(pep, mao)
        calcMop = Round(totalSemMop * (CfgNum("PERC_MOP", 5.483) / 100), 2)

        outp(r + 1, 27) = Round(tot, 2)
        If mao <> 0 Then outp(r + 1, 28) = Round(atv / mao * 100, 2) & "%"
        If totalSemMop <> 0 Then outp(r + 1, 29) = Round(mop / totalSemMop * 100, 2) & "%"
        If tot <> 0 Then outp(r + 1, 30) = Round(matUc / tot * 100, 2) & "%"
        outp(r + 1, 31) = TipoPEPCodigo(pep)
        outp(r + 1, 32) = atvPrev
        outp(r + 1, 33) = Round(atv - atvPrev, 2)
        outp(r + 1, 34) = ValorCat(dGrp, pep, "FRETE_TRANSP")
        outp(r + 1, 35) = SegmentoPI(pep)
        outp(r + 1, 36) = Round(totalSemMop, 2)
        outp(r + 1, 37) = calcMop
        outp(r + 1, 38) = Round(mop - calcMop, 2)
        outp(r + 1, 39) = Round(IIf(dMaoCLS3.Exists(pep), dMaoCLS3(pep), 0), 2)
    Next r
    EscreverAba "ANALISE DE CA", outp
End Sub


Private Function ValorCat(dGrp As Object, ByVal pep As String, ByVal cat As String) As Double
    ValorCat = 0
    If dGrp.Exists(pep & "|" & cat) Then ValorCat = dGrp(pep & "|" & cat)
End Function

Private Function CategoriaAnaliseCA(ByVal lin As Long) As String
    Dim cls2 As String, cls3 As String, c As String, cl As String, cat As String
    cl = UCase$(SemAcento(CStr(dados(lin, cClassif))))

    ' A classe de custo e a fonte mais confiavel para indiretos da ANALISE DE CA.
    cat = CategoriaPorClasseCusto(lin)
    If cat <> "" Then CategoriaAnaliseCA = cat: Exit Function

    If EhMaterial(CStr(dados(lin, cClassif))) Then
        cls3 = UCase$(SemAcento(MatInfoLinha(lin, 3)))
        If InStr(cl, "UC") > 0 Or InStr(cls3, "UC") > 0 Then CategoriaAnaliseCA = "MAT UC": Exit Function
        If InStr(cl, "COM") > 0 Or InStr(cls3, "COM") > 0 Then CategoriaAnaliseCA = "MAT COM": Exit Function
        CategoriaAnaliseCA = "OUTROS": Exit Function
    End If

    cls2 = UCase$(SemAcento(SrvInfoLinha(lin, 1)))
    cls3 = UCase$(SemAcento(SrvInfoLinha(lin, 2)))

    cat = MapCategoriaCA(cls3)
    If cat <> "" Then CategoriaAnaliseCA = cat: Exit Function
    cat = MapCategoriaCA(cls2)
    If cat <> "" Then CategoriaAnaliseCA = cat: Exit Function

    c = cls3 & " " & cls2 & " " & UCase$(SemAcento(TextoCampo(lin, cDescSA)))

    If InStr(c, "CLASSIFICAR") > 0 Then CategoriaAnaliseCA = "CLASSIFICAR": Exit Function
    If InStr(c, "RISCO") > 0 Then CategoriaAnaliseCA = "RISCO": Exit Function
    If InStr(c, "MOP") > 0 Then CategoriaAnaliseCA = "MOP": Exit Function
    If InStr(c, "FRETE") > 0 Or InStr(c, "TRANSP") > 0 Then CategoriaAnaliseCA = "FRETE_TRANSP": Exit Function
    If InStr(c, "FISCAL") > 0 Then CategoriaAnaliseCA = "FISCALIZACAO": Exit Function
    If InStr(c, "PROJETO DRT") > 0 Then CategoriaAnaliseCA = "PROJETO_DRT_OBRA": Exit Function
    If InStr(c, "PROJETO") > 0 Then CategoriaAnaliseCA = "PROJETO": Exit Function
    If InStr(c, "ATIVACAO") > 0 Or InStr(c, "ATV DRT") > 0 Then CategoriaAnaliseCA = "ATIVACAO DIRETA": Exit Function
    If InStr(c, "APOIO") > 0 Then CategoriaAnaliseCA = "APOIO A OBRA": Exit Function
    If InStr(c, "COMISSION") > 0 Then CategoriaAnaliseCA = "COMISSIONAMENTO": Exit Function
    If InStr(c, "LOGIST") > 0 Then CategoriaAnaliseCA = "LOGISTICA": Exit Function
    If InStr(c, "TRIBUT") > 0 Then CategoriaAnaliseCA = "TRIBUTOS": Exit Function
    If InStr(c, "SUPORTE") > 0 Then CategoriaAnaliseCA = "SUPORTE": Exit Function
    If InStr(c, "PUBLIC") > 0 Then CategoriaAnaliseCA = "PUBLICIDADE": Exit Function
    If InStr(c, "MEIO AMBIENTE") > 0 Or InStr(c, "AMBIENT") > 0 Then CategoriaAnaliseCA = "MEIO AMBIENTE": Exit Function
    If InStr(c, "CIVIL") > 0 Then CategoriaAnaliseCA = "MAO DE OBRA CIVIL": Exit Function
    If InStr(c, "DISPONIBIL") > 0 Then CategoriaAnaliseCA = "DISPONIBILIDADE": Exit Function
    If InStr(c, "MAO DE OBRA") > 0 Or InStr(c, "M O") > 0 Then CategoriaAnaliseCA = "MAO DE OBRA": Exit Function
    If InStr(c, "BAIXA") > 0 Then CategoriaAnaliseCA = "BAIXA ODD ODS": Exit Function
    If InStr(c, "JOA") > 0 Then CategoriaAnaliseCA = "JOA": Exit Function
    CategoriaAnaliseCA = "OUTROS"
End Function

Private Function CategoriaPorClasseCusto(ByVal lin As Long) As String
    Dim cls2 As String, cls3 As String
    If ClasseCustoDadosOutros(ValorCampo(lin, cClasse)) Then
        CategoriaPorClasseCusto = "OUTROS"
        Exit Function
    End If

    cls3 = CCInfo(ValorCampo(lin, cClasse), 2)
    CategoriaPorClasseCusto = MapCategoriaCA(cls3)
    If CategoriaPorClasseCusto <> "" Then Exit Function

    cls2 = CCInfo(ValorCampo(lin, cClasse), 1)
    CategoriaPorClasseCusto = MapCategoriaCA(cls2)
End Function

Private Function ClasseCustoDadosOutros(codCC As Variant) As Boolean
    ClasseCustoDadosOutros = _
        (NormCod(codCC) = CfgTxt("CLASSE_COMBUSTIVEL", "8119980000"))
End Function


Private Function MapCategoriaCA(ByVal valor As String) As String
    Dim s As String
    s = UCase$(SemAcento(Trim$(valor)))
    s = Replace(s, "/", "_")
    s = Replace(s, "-", "_")
    s = Replace(s, ".", " ")     ' "MAT. UC"/"MAT.COM" (catalogos ATUAIS) -> "MAT UC"/"MAT COM"
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    s = Trim$(s)

    If s = "" Then Exit Function
    If InStr(s, "CLASSIFICAR") > 0 Then MapCategoriaCA = "CLASSIFICAR": Exit Function
    If InStr(s, "RISCO") > 0 Then MapCategoriaCA = "RISCO": Exit Function
    If InStr(s, "MOP") > 0 Then MapCategoriaCA = "MOP": Exit Function
    If InStr(s, "JOA") > 0 Then MapCategoriaCA = "JOA": Exit Function
    If InStr(s, "TRIBUT") > 0 Or InStr(s, "ODC_MATERIAL") > 0 Then MapCategoriaCA = "TRIBUTOS": Exit Function
    If InStr(s, "FRETE") > 0 Or InStr(s, "TRANSP") > 0 Then MapCategoriaCA = "FRETE_TRANSP": Exit Function
    If InStr(s, "SUPORTE") > 0 Or InStr(s, "ALIMENTACAO") > 0 _
       Or InStr(s, "HOSPEDAGEM") > 0 Or InStr(s, "PASSAGEM") > 0 _
       Or InStr(s, "MOBILIDADE") > 0 Then MapCategoriaCA = "SUPORTE": Exit Function
    If InStr(s, "PUBLIC") > 0 Then MapCategoriaCA = "PUBLICIDADE": Exit Function
    If InStr(s, "FISCAL") > 0 Then MapCategoriaCA = "FISCALIZACAO": Exit Function
    If InStr(s, "PROJETO DRT") > 0 Or InStr(s, "PROJETO_DRT") > 0 Then MapCategoriaCA = "PROJETO_DRT_OBRA": Exit Function
    If InStr(s, "PROJETO") > 0 Then MapCategoriaCA = "PROJETO": Exit Function
    If InStr(s, "ATIVACAO") > 0 Or InStr(s, "ATV DRT") > 0 Then MapCategoriaCA = "ATIVACAO DIRETA": Exit Function
    If InStr(s, "APOIO") > 0 Then MapCategoriaCA = "APOIO A OBRA": Exit Function
    If InStr(s, "COMISSION") > 0 Then MapCategoriaCA = "COMISSIONAMENTO": Exit Function
    If InStr(s, "LOGIST") > 0 Then MapCategoriaCA = "LOGISTICA": Exit Function
    If InStr(s, "MEIO AMBIENTE") > 0 Or InStr(s, "AMBIENT") > 0 Then MapCategoriaCA = "MEIO AMBIENTE": Exit Function
    If InStr(s, "CIVIL") > 0 Then MapCategoriaCA = "MAO DE OBRA CIVIL": Exit Function
    If InStr(s, "DISPONIBIL") > 0 Then MapCategoriaCA = "DISPONIBILIDADE": Exit Function
    If InStr(s, "MAO DE OBRA") > 0 Or InStr(s, "M O") > 0 Then MapCategoriaCA = "MAO DE OBRA": Exit Function
    If InStr(s, "BAIXA") > 0 Then MapCategoriaCA = "BAIXA ODD ODS": Exit Function
    If InStr(s, "MAT UC") > 0 Then MapCategoriaCA = "MAT UC": Exit Function
    If InStr(s, "MAT COM") > 0 Then MapCategoriaCA = "MAT COM": Exit Function
    If InStr(s, "VALOR MATERIAL") > 0 Then MapCategoriaCA = "VALOR MATERIAL"
End Function


'==============================================================================
'  ABA: CLASSE DE CUSTO  (por PEP x classe)
'==============================================================================
Private Sub Gerar_ClasseDeCusto()
    Dim dGrp As Object, dQtd As Object, dCnt As Object, dDesc As Object, dEmpresa As Object
    Set dGrp = CreateObject("Scripting.Dictionary")
    Set dQtd = CreateObject("Scripting.Dictionary")
    Set dCnt = CreateObject("Scripting.Dictionary")
    Set dDesc = CreateObject("Scripting.Dictionary")
    Set dEmpresa = CreateObject("Scripting.Dictionary")

    Dim i As Long, pep As String, cc As String, k As String
    For i = 1 To UBound(dados, 1)
        pep = Trim$(CStr(dados(i, cPEP))): If pep = "" Then GoTo Prox
        cc = TextoCampo(i, cClasse)
        k = pep & "|" & cc
        dGrp(k) = dGrp(k) + ToNum(dados(i, cValor))
        dQtd(k) = dQtd(k) + ToNum(dados(i, cQtd))
        dCnt(k) = dCnt(k) + 1
        If cDescClasse > 0 And Not dDesc.Exists(k) Then dDesc(k) = TextoCampo(i, cDescClasse)
        If Not dEmpresa.Exists(k) Then dEmpresa(k) = TextoCampo(i, cEmpresa)
Prox:
    Next i

    Dim ks As Variant: ks = dGrp.Keys
    Dim outp() As Variant: ReDim outp(0 To dGrp.Count, 1 To 11)
    outp(0, 1) = "EMPRESA": outp(0, 2) = "PEP4NIVEL": outp(0, 3) = "TIPO_PEP"
    outp(0, 4) = "CLASSE_CUSTO": outp(0, 5) = "DESC_CLASSE_CUSTO"
    outp(0, 6) = "QTD_ENTRADA": outp(0, 7) = "CLS1": outp(0, 8) = "CLS2"
    outp(0, 9) = "CLS3": outp(0, 10) = "VALOR_MOEDA": outp(0, 11) = "LANCAMENTOS"
    Dim r As Long, p As Variant
    For r = 0 To dGrp.Count - 1
        p = Split(ks(r), "|")
        outp(r + 1, 1) = dEmpresa(ks(r))
        outp(r + 1, 2) = p(0)
        outp(r + 1, 3) = TipoPEPANEEL(CStr(p(0)))
        outp(r + 1, 4) = p(1)
        If dDesc.Exists(ks(r)) Then outp(r + 1, 5) = dDesc(ks(r))
        outp(r + 1, 6) = Round(dQtd(ks(r)), 2)
        outp(r + 1, 7) = CCInfo(p(1), 0)
        outp(r + 1, 8) = CCInfo(p(1), 1)
        outp(r + 1, 9) = CCInfo(p(1), 2)
        outp(r + 1, 10) = Round(dGrp(ks(r)), 2)
        outp(r + 1, 11) = dCnt(ks(r))
    Next r
    EscreverAba "CLASSE DE CUSTO", outp
End Sub


'==============================================================================
'  ABA: MATERIAL  (detalhe UC/COM/UAR)
'==============================================================================
Private Sub Gerar_Material()
    ' Consolida por PEP + MATERIAL + CLASSIFICACAO: soma QTD e VALOR numa unica
    ' linha (entradas e estornos), mostrando o liquido correto de entrada/saida.
    Dim dQ As Object, dV As Object, dFirst As Object
    Set dQ = CreateObject("Scripting.Dictionary")
    Set dV = CreateObject("Scripting.Dictionary")
    Set dFirst = CreateObject("Scripting.Dictionary")   ' chave -> indice da 1a linha

    Dim i As Long, pep As String, cod As String, cl As String, k As String
    For i = 1 To UBound(dados, 1)
        pep = Trim$(CStr(dados(i, cPEP))): If pep = "" Then GoTo Prox
        If Not EhMaterial(CStr(dados(i, cClassif))) Then GoTo Prox
        cod = NormCod(dados(i, cMaterial))
        cl = UCase$(Trim$(CStr(dados(i, cClassif))))
        k = pep & "|" & cod & "|" & cl
        If Not dFirst.Exists(k) Then dFirst(k) = i
        dQ(k) = dQ(k) + ToNum(dados(i, cQtd))
        dV(k) = dV(k) + ToNum(dados(i, cValor))
Prox:
    Next i

    Dim ks As Variant: ks = dFirst.Keys
    Dim r As Long, fi As Long, q As Double, val As Double

    ' 1a passada: conta as linhas que ficam (exclui QTD=0 e VALOR=0)
    Dim nKeep As Long
    For r = 0 To dFirst.Count - 1
        fi = dFirst(ks(r))
        q = Round(dQ(ks(r)), 2): val = Round(dV(ks(r)), 2)
        If q = 0 And val = 0 Then GoTo PulaConta
        ' Exclui linhas com classificacao "FALTA" e QTD=0/VALOR<>0
        If InStr(UCase$(CStr(dados(fi, cClassif))), "FALTA") > 0 _
           And q = 0 And val <> 0 Then GoTo PulaConta
        nKeep = nKeep + 1
PulaConta:
    Next r

    Dim outp() As Variant: ReDim outp(0 To nKeep, 1 To 15)
    outp(0, 1) = "PEP4NIVEL": outp(0, 2) = "PEP3": outp(0, 3) = "TIPO_PEP"
    outp(0, 4) = "CLASSE_CUSTO": outp(0, 5) = "MATERIAL": outp(0, 6) = "TEXTO_MATERIAL"
    outp(0, 7) = "UML": outp(0, 8) = "QTD_ENTRADA": outp(0, 9) = "VALOR_MOEDA"
    outp(0, 10) = "CLASSIFICACAO": outp(0, 11) = "CLS2"
    outp(0, 12) = "PRECO_UNITARIO": outp(0, 13) = "ADERENCIA"
    outp(0, 14) = "SITUACAO": outp(0, 15) = "OBS_SITUACAO"

    Dim rr As Long: rr = 0
    For r = 0 To dFirst.Count - 1
        k = ks(r): fi = dFirst(k)
        pep = Trim$(CStr(dados(fi, cPEP)))
        q = Round(dQ(k), 2): val = Round(dV(k), 2)
        If q = 0 And val = 0 Then GoTo PulaLin
        If InStr(UCase$(CStr(dados(fi, cClassif))), "FALTA") > 0 _
           And q = 0 And val <> 0 Then GoTo PulaLin
        rr = rr + 1
        outp(rr, 1) = pep: outp(rr, 2) = PEP3(pep): outp(rr, 3) = TipoPEPANEEL(pep)
        outp(rr, 4) = ValorCampo(fi, cClasse)   ' CLASSE_CUSTO
        outp(rr, 5) = NormCod(dados(fi, cMaterial)): outp(rr, 6) = ValorCampo(fi, cTexto)
        outp(rr, 7) = ValorCampo(fi, cUML)
        outp(rr, 8) = q: outp(rr, 9) = val
        outp(rr, 10) = dados(fi, cClassif)
        outp(rr, 11) = MatInfoLinha(fi, 2)   ' CLS2
        ' PRECO_UNITARIO = VALOR_MOEDA / QTD_ENTRADA (somas brutas, 4 casas)
        If q <> 0 Then outp(rr, 12) = Round(dV(k) / dQ(k), 4)
        ' ADERENCIA:
        '   1) QTD=0 e VALOR<>0                        -> NAO ADERENTE
        '   2) sinais opostos (QTD e VALOR)             -> NAO ADERENTE
        '   3) VALOR=0 e QTD<>0                         -> NAO ADERENTE
        '   4) por tipo de PEP ANEEL (ODD/ODI/ODM/ODS):
        '      ODD (.D): QTD_ENTRADA ou VALOR_MOEDA positivo   -> NAO ADERENTE
        '      ODI/ODM/ODS: QTD_ENTRADA ou VALOR_MOEDA negativo -> NAO ADERENTE
        If q = 0 And val <> 0 Then
            outp(rr, 13) = "NAO ADERENTE"
        ElseIf (q > 0 And val < 0) Or (q < 0 And val > 0) Then
            outp(rr, 13) = "NAO ADERENTE"
        ElseIf val = 0 And q <> 0 Then
            outp(rr, 13) = "NAO ADERENTE"
        ElseIf TipoPEPCodigo(pep) = "D" Then
            outp(rr, 13) = IIf(q > 0 Or val > 0, "NAO ADERENTE", "ADERENTE")
        Else
            outp(rr, 13) = IIf(q < 0 Or val < 0, "NAO ADERENTE", "ADERENTE")
        End If
        ' SITUACAO de correcao (1..9) por QTD/VALOR e tipo de PEP.
        ' Retorna "N|descricao"; col 14 = numero, col 15 = observacao.
        Dim sitS As String: sitS = SituacaoCorrecao(q, val, TipoPEPCodigo(pep))
        If sitS <> "" Then
            Dim sitP() As String: sitP = Split(sitS, "|")
            outp(rr, 14) = CLng(sitP(0))
            outp(rr, 15) = sitP(1)
        End If
PulaLin:
    Next r
    EscreverAba "MATERIAL", outp
End Sub

' Classifica a linha de material em uma das 9 SITUACOES de correcao, conforme
' o sinal de QTD e VALOR e o tipo de PEP. tp = codigo do PEP (I/D/M/S):
'   D = ODD ; I/M/S = ODI/ODM/ODS. Combos nao previstos retornam "".
'   Retorno: "N|descricao" (vazio quando nao ha situacao).
'   Situacoes 4 e 5 exigem PU > 0 ; situacao 8 exige PU < 0.
Private Function SituacaoCorrecao(ByVal q As Double, ByVal v As Double, ByVal tp As String) As String
    q = Round(q, 2): v = Round(v, 2)
    Dim odd As Boolean: odd = (tp = "D")
    Dim outros As Boolean: outros = Not odd          ' ODI / ODM / ODS
    SituacaoCorrecao = ""
    If q = 0 And v = 0 Then Exit Function            ' NULO (nem entra na aba)

    If q < 0 And v < 0 Then
        If outros Then SituacaoCorrecao = "1|QTD NEGATIVA/VLR NEGATIVO"
    ElseIf q = 0 And v < 0 Then
        SituacaoCorrecao = "2|QTD ZERO/VLR NEGATIVO"
    ElseIf q < 0 And v > 0 Then
        SituacaoCorrecao = IIf(odd, "8|QTD NEGATIVA/VLR POSITIVA - ODD (PU < 0)", _
                                    "3|QTD NEGATIVA/VLR POSITIVA")
    ElseIf q > 0 And v = 0 Then
        SituacaoCorrecao = IIf(odd, "9|QTD POSITIVA/VLR ZERO - ODD", _
                                    "4|QTD POSITIVA/VLR ZERO (PU > 0)")
    ElseIf q > 0 And v < 0 Then
        If outros Then SituacaoCorrecao = "5|QTD POSITIVA/VLR NEGATIVO (PU > 0)"
    ElseIf q = 0 And v > 0 Then
        SituacaoCorrecao = "6|QTD ZERO/VLR POSITIVA"
    ElseIf q < 0 And v = 0 Then
        SituacaoCorrecao = "7|QTD NEGATIVA/VLR ZERO"
    End If
    ' q > 0 e v > 0 -> lancamento normal, sem situacao de correcao
End Function



'==============================================================================
'  ABA: SERVIÇOS  (somente linhas classificadas como MAO DE OBRA na CLS3)
'  Consolida por PEP4 + COD_SERVICO: soma QTD e VALOR numa unica linha.
'  Filtro: NOT EhMaterial  AND  LinhaCLS3 = "MAO DE OBRA"
'==============================================================================

'==============================================================================
'  ABA: SERVIO  (detalhe das linhas que NO so material: CA/servios)
'==============================================================================
Private Sub Gerar_Servico()
    ' Consolida por PEP + COD_SERVICO: soma QTD e VALOR numa unica linha.
    ' Linhas sem COD_SERVICO nao sao trazidas.
    Dim dQ As Object, dV As Object, dFirst As Object
    Set dQ = CreateObject("Scripting.Dictionary")
    Set dV = CreateObject("Scripting.Dictionary")
    Set dFirst = CreateObject("Scripting.Dictionary")   ' chave -> indice da 1a linha

    ' Qtd de CAVA (servico ESCAVACAO) e de POSTE por PEP4 -> regra de aderencia:
    ' mais servico de cava do que poste no PEP marca a cava como NAO ADERENTE.
    Dim dCava As Object:     Set dCava = CreateObject("Scripting.Dictionary")
    Dim dPosteSrv As Object: Set dPosteSrv = CreateObject("Scripting.Dictionary")
    Dim dPosteMat As Object: Set dPosteMat = CreateObject("Scripting.Dictionary")

    Dim i As Long, pep As String, codSrv As String, k As String, famSv As String
    For i = 1 To UBound(dados, 1)
        pep = Trim$(CStr(dados(i, cPEP))): If pep = "" Then GoTo Prox
        If EhMaterial(CStr(dados(i, cClassif))) Then
            ' poste (material) do PEP -> numero de postes
            If NormClassif(FamiliaAlias(MatInfoLinha(i, 2))) = "POSTE RD" Then _
                dPosteMat(pep) = dPosteMat(pep) + Abs(ToNum(dados(i, cQtd)))
            GoTo Prox
        End If
        codSrv = NormCod(dados(i, cMaterial))   ' codigo do servico (coluna Material)
        ' acumula CAVA e POSTE (servico) por PEP (independe de ter COD_SERVICO)
        famSv = NormClassif(FamiliaAlias(SrvInfoLinha(i, 1)))
        If famSv = "ESCAVACAO" Then dCava(pep) = dCava(pep) + Abs(ToNum(dados(i, cQtd)))
        If famSv = "POSTE RD" Then dPosteSrv(pep) = dPosteSrv(pep) + Abs(ToNum(dados(i, cQtd)))
        If codSrv = "" Or codSrv = "0" Then GoTo Prox   ' sem COD_SERVICO -> nao traz
        k = pep & "|" & codSrv
        If Not dFirst.Exists(k) Then dFirst(k) = i
        dQ(k) = dQ(k) + ToNum(dados(i, cQtd)) * ComboFator(dados(i, cMaterial))
        dV(k) = dV(k) + ToNum(dados(i, cValor))
Prox:
    Next i

    Dim ks As Variant: ks = dFirst.Keys
    Dim r As Long, ki As Long, fi As Long

    ' conta linhas que ficam (exclui QTD=0 e VALOR=0)
    Dim nKeep As Long
    For r = 0 To dFirst.Count - 1
        If Not (Round(dQ(ks(r)), 2) = 0 And Round(dV(ks(r)), 2) = 0) Then nKeep = nKeep + 1
    Next r

    Dim outp() As Variant: ReDim outp(0 To nKeep, 1 To 17)
    outp(0, 1) = "PEP4NIVEL": outp(0, 2) = "PEP3": outp(0, 3) = "TIPO_PEP"
    outp(0, 4) = "COD_SERVICO": outp(0, 5) = "DESCRICAO_SERVICO"
    outp(0, 6) = "CLASSE_CUSTO"
    outp(0, 7) = "UND": outp(0, 8) = "QTD_ENTRADA": outp(0, 9) = "VALOR_MOEDA"
    outp(0, 10) = "CLS1": outp(0, 11) = "CLS2": outp(0, 12) = "CLS3"
    outp(0, 13) = "TIPO_APLICACAO": outp(0, 14) = "GRUPO_PERC"
    outp(0, 15) = "ALERTA_VALOR": outp(0, 16) = "APROPRIACAO"
    outp(0, 17) = "ADERENCIA"

    ki = 0
    Dim codSrvOut As String, und As String, taSrv As String, tpPep As String
    Dim descOut As String, descU As String
    For r = 0 To dFirst.Count - 1
        k = ks(r): fi = dFirst(k)
        If Round(dQ(k), 2) = 0 And Round(dV(k), 2) = 0 Then GoTo PulaZero   ' QTD=0 e VALOR=0
        ki = ki + 1
        pep = Trim$(CStr(dados(fi, cPEP)))
        codSrvOut = NormCod(dados(fi, cMaterial))
        outp(ki, 1) = pep: outp(ki, 2) = PEP3(pep): outp(ki, 3) = TipoPEPANEEL(pep)
        outp(ki, 4) = codSrvOut
        ' Descricao: catalogo embutido/externo; fallback TEXTO BREVE da base
        descOut = DescServico(codSrvOut)
        If descOut = "" Then descOut = Trim$(TextoCampo(fi, cTexto))
        outp(ki, 5) = descOut
        outp(ki, 6) = ValorCampo(fi, cClasse)
        ' UND: unidade de medida (M, UN, M3...). Preferencia: coluna UML do
        ' proprio lancamento SAP; fallback: catalogo de servicos (UN).
        und = Trim$(TextoCampo(fi, cUML))
        If und = "" Then und = SrvInfo(codSrvOut, 5)
        outp(ki, 7) = und
        outp(ki, 8) = Round(dQ(k), 2)
        outp(ki, 9) = Round(dV(k), 2)
        outp(ki, 10) = SrvInfoLinha(fi, 0)    ' CLS1
        outp(ki, 11) = SrvInfoLinha(fi, 1)    ' CLS2
        outp(ki, 12) = SrvInfoLinha(fi, 2)    ' CLS3
        outp(ki, 13) = SrvInfoLinha(fi, 3)    ' TIPO_APLICACAO
        outp(ki, 14) = GrupoPerc(pep)

        ' ALERTA_VALOR: automatiza o "CHECK DE VALOR DO SERVICO" do checklist
        If Round(dV(k), 2) < 0 And Round(dQ(k), 2) < 0 Then
            outp(ki, 15) = "VALOR E QTD NEGATIVOS"
        ElseIf Round(dV(k), 2) < 0 Then
            outp(ki, 15) = "VALOR NEGATIVO"
        ElseIf Round(dQ(k), 2) < 0 Then
            outp(ki, 15) = "QTD NEGATIVA"
        Else
            outp(ki, 15) = "OK"
        End If

        ' APROPRIACAO: automatiza o "CHECK DE APROPRIACAO CORRETA DE SERVICO".
        ' 1) TIPO_APLICACAO do catalogo (ODI*/ODD*) vs tipo do PEP (.I/.D);
        ' 2) fallback por palavra-chave na descricao (INST/RET/DESAT).
        taSrv = UCase$(SrvInfoLinha(fi, 3))
        tpPep = TipoPEPCodigo(pep)
        descU = UCase$(SemAcento(descOut))
        If Left$(taSrv, 3) = "ODI" And tpPep = "D" Then
            outp(ki, 16) = "VERIFICAR: SERVICO ODI LANCADO NA ODD"
        ElseIf Left$(taSrv, 3) = "ODD" And tpPep = "I" Then
            outp(ki, 16) = "VERIFICAR: SERVICO ODD LANCADO NA ODI"
        ElseIf tpPep = "I" And (Left$(descU, 4) = "RET " Or InStr(descU, "RETIRADA") > 0 _
               Or InStr(descU, "DESATIV") > 0 Or InStr(descU, "DESMONT") > 0) Then
            outp(ki, 16) = "VERIFICAR: RETIRADA/DESATIVACAO NA ODI"
        ElseIf tpPep = "D" And (Left$(descU, 5) = "INST " Or InStr(descU, "INSTALACAO") > 0 _
               Or InStr(descU, "IMPLANT") > 0) Then
            outp(ki, 16) = "VERIFICAR: INSTALACAO NA ODD"
        Else
            outp(ki, 16) = "OK"
        End If

        ' ADERENCIA (mesma regra da aba MATERIAL):
        '   1) QTD=0 com VALOR (positivo ou negativo)     -> NAO ADERENTE
        '   2) ODD (.D): QTD ou VALOR positivo            -> NAO ADERENTE
        '   3) ODI/ODM/ODS: QTD ou VALOR negativo         -> NAO ADERENTE
        If Round(dQ(k), 2) = 0 And Round(dV(k), 2) <> 0 Then
            outp(ki, 17) = "NAO ADERENTE"
        ElseIf tpPep = "D" Then
            outp(ki, 17) = IIf(Round(dQ(k), 2) > 0 Or Round(dV(k), 2) > 0, "NAO ADERENTE", "ADERENTE")
        Else
            outp(ki, 17) = IIf(Round(dQ(k), 2) < 0 Or Round(dV(k), 2) < 0, "NAO ADERENTE", "ADERENTE")
        End If

        ' Regra CAVA x POSTE: servico de CAVA (ESCAVACAO) com quantidade maior
        ' que a de poste no PEP -> NAO ADERENTE (a cava nao e ligada ao poste,
        ' apenas sinalizada). POSTE = qtd de servico de poste; se 0, usa o material.
        If NormClassif(FamiliaAlias(SrvInfoLinha(fi, 1))) = "ESCAVACAO" Then
            Dim qCava As Double, qPoste As Double
            qCava = IIf(dCava.Exists(pep), CDbl(dCava(pep)), 0)
            qPoste = IIf(dPosteSrv.Exists(pep), CDbl(dPosteSrv(pep)), 0)
            If qPoste = 0 Then qPoste = IIf(dPosteMat.Exists(pep), CDbl(dPosteMat(pep)), 0)
            If Round(qCava, 2) > Round(qPoste, 2) Then
                outp(ki, 17) = "NAO ADERENTE"
                outp(ki, 16) = "CAVA (" & Format$(qCava, "0.##") & ") > POSTE (" & _
                               Format$(qPoste, "0.##") & ")"
            End If
        End If
PulaZero:
    Next r
    EscreverAba "SERVICO", outp
End Sub


'==============================================================================
'  ABA: ALERTAS E PONTOS CRITICOS
'  Secoes:
'    A) PEPs sem nenhuma familia UC (nao possuem material de investimento)
'    B) Material com QTD positiva e VALOR zero ou negativo
'    D) Material com QTD negativa e VALOR positivo ou zero
'==============================================================================
Private Sub Gerar_AlertasCriticos()

    ' -----------------------------------------------------------------------
    ' Pre-calculo: quais PEPs possuem pelo menos uma familia UC
    ' (reutiliza a mesma logica de Gerar_MaterialVsServico)
    ' -----------------------------------------------------------------------
    gEtapa = "Alertas: pre-calc UC/UAR"
    Dim dPepTemUC As Object
    Set dPepTemUC = CreateObject("Scripting.Dictionary")

    Dim dPepTodos As Object
    Set dPepTodos = CreateObject("Scripting.Dictionary")   ' todos os PEPs encontrados

    Dim i As Long, pep As String, cls2 As String, q As Double, val As Double

    ' Pre-scan: PEPs (PEP4NIVEL) que possuem material da familia UAR.
    ' Esses PEPs sao excluidos de TODAS as secoes dos alertas criticos.
    Dim dPepUAR As Object: Set dPepUAR = CreateObject("Scripting.Dictionary")
    Dim cls2u As String
    For i = 1 To UBound(dados, 1)
        pep = Trim$(CStr(dados(i, cPEP))): If pep = "" Then GoTo ProxUAR
        If Not EhMaterial(CStr(dados(i, cClassif))) Then GoTo ProxUAR
        cls2u = FamiliaAlias(MatInfoLinha(i, 2))
        If UCase$(TipoDaClassif(cls2u, TextoCampo(i, cTexto))) = "UAR" Then dPepUAR(pep) = 1
ProxUAR:
    Next i
    ' dPepTemSrv: PEP4 que possuem pelo menos um servico lancado (com COD_SERVICO)
    Dim dPepTemSrv As Object: Set dPepTemSrv = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(dados, 1)
        pep = Trim$(CStr(dados(i, cPEP))): If pep = "" Or dPepUAR.Exists(pep) Then GoTo PreCalc
        ' ODD (.D) nao exige material UC -> nao entra na secao A
        If TipoPEPCodigo(pep) = "D" Then GoTo PreCalc
        If Not dPepTodos.Exists(pep) Then dPepTodos(pep) = 1
        If EhMaterial(CStr(dados(i, cClassif))) Then
            cls2 = FamiliaAlias(MatInfoLinha(i, 2))
            Dim tUC As String: tUC = UCase$(TipoDaClassif(cls2, TextoCampo(i, cTexto)))
            ' Fallback: familia sem TIPO -> usa CLS3 do catalogo (MAT. UC)
            If tUC = "" And InStr(UCase$(MatInfoLinha(i, 3)), "UC") > 0 Then tUC = "UC"
            If tUC = "UC" Then dPepTemUC(pep) = 1
        Else
            ' servico com COD_SERVICO valido
            Dim codSv As String: codSv = NormCod(dados(i, cMaterial))
            If codSv <> "" And codSv <> "0" Then dPepTemSrv(pep) = 1
        End If
PreCalc:
    Next i

    ' -----------------------------------------------------------------------
    ' Monta a aba (design colorido: banda de titulo verde, cards com fundo,
    ' secoes com faixa de cor cheia e zebra na tinta de cada secao)
    ' -----------------------------------------------------------------------
    gEtapa = "Alertas: criar aba"
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("ALERTAS CRITICOS")
    If Not ws Is Nothing Then ws.Delete
    On Error GoTo 0
    Set ws = ActiveWorkbook.Worksheets.Add( _
        After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    ws.Name = "ALERTAS CRITICOS"
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 9

    ' --- Paleta -------------------------------------------------------------
    Dim corInk As Long, corMut As Long
    Dim corA As Long, corAcl As Long, corAzb As Long
    Dim corB As Long, corBcl As Long, corBzb As Long
    Dim corC As Long, corCcl As Long, corCzb As Long
    Dim corE As Long, corEcl As Long, corEzb As Long
    corInk = RGB(45, 52, 56)        ' texto principal
    corMut = COR_UI_MUDO            ' texto secundario
    corA = RGB(21, 87, 148):  corAcl = RGB(232, 240, 248): corAzb = RGB(246, 250, 253)
    corB = RGB(168, 32, 40):  corBcl = RGB(253, 235, 236): corBzb = RGB(253, 247, 247)
    corC = RGB(90, 62, 138):  corCcl = RGB(238, 233, 246): corCzb = RGB(249, 247, 252)
    corE = RGB(159, 96, 12):  corEcl = RGB(253, 243, 224): corEzb = RGB(254, 251, 244)

    ' Tipografia base da aba inteira (evita mistura de fontes do Excel)
    On Error Resume Next
    ws.Cells.Font.Name = FONTE_UI
    ws.Cells.Font.Size = 9.5
    ws.Cells.VerticalAlignment = xlCenter
    On Error GoTo 0

    Dim row As Long: row = 1
    Dim contA As Long, contB As Long, contC As Long, contE As Long
    Dim rowCards As Long

    gEtapa = "Alertas: cabecalho"
    ' === CABECALHO DA ABA (banda verde institucional) ===
    With ws.Range(ws.Cells(1, 1), ws.Cells(2, 10))
        .Interior.Color = COR_UI_HDR
    End With
    ws.Cells(row, 1).Value = "ALERTAS CRITICOS"
    With ws.Cells(row, 1)
        .Font.Size = 20: .Font.Bold = True: .Font.Color = vbWhite
        .IndentLevel = 1
    End With
    ws.Rows(row).RowHeight = 34
    row = row + 1
    ws.Cells(row, 1).Value = "Analise de custo / CKCP  -  gerado em " & Format(Now, "dd/mm/yyyy hh:nn")
    With ws.Cells(row, 1)
        .Font.Size = 9: .Font.Color = RGB(195, 235, 212)
        .IndentLevel = 1
    End With
    ws.Rows(row).RowHeight = 16
    row = row + 1
    ' faixa de acento verde-claro sob a banda
    With ws.Range(ws.Cells(row, 1), ws.Cells(row, 10))
        .Interior.Color = RGB(0, 176, 102)
    End With
    ws.Rows(row).RowHeight = 4
    row = row + 2

    ' === CARDS DE RESUMO (numeros preenchidos no final, com contagens reais) ===
    ws.Cells(row, 1).Value = "RESUMO"
    With ws.Cells(row, 1)
        .Font.Size = 8: .Font.Bold = True: .Font.Color = COR_UI_MUDO
        .IndentLevel = 1
    End With
    row = row + 1
    rowCards = row
    ws.Rows(rowCards).RowHeight = 18
    ws.Rows(rowCards + 1).RowHeight = 40
    row = row + 3

    gEtapa = "Alertas: secao A"
    ' -----------------------------------------------------------------------
    ' SECAO A - PEPs sem nenhuma familia UC
    ' -----------------------------------------------------------------------
    row = EscreverCabecalhoAlerta(ws, row, _
        "A  |  PEPs SEM FAMILIA UC  (nenhum material de investimento classificado como UC)", _
        Array("PEP4NIVEL", "PEP3", "TIPO_PEP", "TOTAL_LANCAMENTOS", "VALOR_TOTAL", "OBSERVACAO"), _
        corA, corAcl)

    Dim dPepValor As Object: Set dPepValor = CreateObject("Scripting.Dictionary")
    Dim dPepCnt As Object: Set dPepCnt = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(dados, 1)
        pep = Trim$(CStr(dados(i, cPEP))): If pep = "" Or dPepUAR.Exists(pep) Then GoTo SomaPep
        dPepValor(pep) = dPepValor(pep) + ToNum(dados(i, cValor))
        dPepCnt(pep) = dPepCnt(pep) + 1
SomaPep:
    Next i

    Dim pk As Variant
    Dim pepKeys As Variant: pepKeys = dPepTodos.Keys
    For Each pk In pepKeys
        pep = CStr(pk)
        If Not dPepTemUC.Exists(pep) Then
            contA = contA + 1
            ws.Cells(row, 1).Value = pep
            ws.Cells(row, 2).Value = PEP3(pep)
            ws.Cells(row, 3).Value = TipoPEPANEEL(pep)
            ws.Cells(row, 4).Value = IIf(dPepCnt.Exists(pep), dPepCnt(pep), 0)
            ws.Cells(row, 5).Value = Round(IIf(dPepValor.Exists(pep), dPepValor(pep), 0), 2)
            ws.Cells(row, 6).Value = "PEP sem material UC - verificar se ha servico sem material correspondente"
            With ws.Range(ws.Cells(row, 1), ws.Cells(row, 6))
                .Font.Color = corInk
                If (contA Mod 2) = 0 Then .Interior.Color = corAzb
            End With
            ws.Cells(row, 5).NumberFormat = "#,##0.00"
            With ws.Cells(row, 3)
                .Font.Color = corA: .Font.Bold = True
            End With
            With ws.Cells(row, 6)
                .Interior.Color = corAcl
                .Font.Color = corA: .Font.Size = 8.5
            End With
            row = row + 1
        End If
    Next pk
    If contA = 0 Then
        ws.Cells(row, 1).Value = "(nenhum PEP sem UC encontrado)"
        ws.Cells(row, 1).Font.Italic = True: ws.Cells(row, 1).Font.Color = corMut
        row = row + 1
    End If
    row = row + 1

    gEtapa = "Alertas: secao C"
    ' -----------------------------------------------------------------------
    ' SECAO C - PEP4NIVEL com material UC e SEM servico lancado
    ' -----------------------------------------------------------------------
    row = EscreverCabecalhoAlerta(ws, row, _
        "C  |  PEP4 COM MATERIAL UC E SEM SERVICO  (material de investimento sem mao de obra correspondente)", _
        Array("PEP4NIVEL", "PEP3", "TIPO_PEP", "TOTAL_LANCAMENTOS", "VALOR_TOTAL", "OBSERVACAO"), _
        corC, corCcl)

    Dim pkC As Variant
    For Each pkC In dPepTemUC.Keys
        pep = CStr(pkC)
        If Not dPepTemSrv.Exists(pep) Then          ' tem UC mas nao tem servico
            contC = contC + 1
            ws.Cells(row, 1).Value = pep
            ws.Cells(row, 2).Value = PEP3(pep)
            ws.Cells(row, 3).Value = TipoPEPANEEL(pep)
            ws.Cells(row, 4).Value = IIf(dPepCnt.Exists(pep), dPepCnt(pep), 0)
            ws.Cells(row, 5).Value = Round(IIf(dPepValor.Exists(pep), dPepValor(pep), 0), 2)
            ws.Cells(row, 6).Value = "Material UC lancado sem servico - verificar mao de obra faltante"
            With ws.Range(ws.Cells(row, 1), ws.Cells(row, 6))
                .Font.Color = corInk
                If (contC Mod 2) = 0 Then .Interior.Color = corCzb
            End With
            ws.Cells(row, 5).NumberFormat = "#,##0.00"
            With ws.Cells(row, 3)
                .Font.Color = corC: .Font.Bold = True
            End With
            row = row + 1
        End If
    Next pkC
    If contC = 0 Then
        ws.Cells(row, 1).Value = "(nenhum PEP com UC sem servico)"
        ws.Cells(row, 1).Font.Italic = True: ws.Cells(row, 1).Font.Color = corMut
        row = row + 1
    End If
    row = row + 1

    gEtapa = "Alertas: secao B"
    ' -----------------------------------------------------------------------
    ' SECAO B - MATERIAL COM QTD/VALOR INCOERENTES / NAO ADERENTE
    '   Consolida por PEP4NIVEL + COD_MATERIAL + CLASSIFICACAO (igual aba
    '   MATERIAL) e lista as linhas NAO ADERENTES pela regra ODD/ODI/ODM/ODS:
    '     ODD (.D): QTD ou VALOR positivo   -> NAO ADERENTE
    '     ODI/ODM/ODS: QTD ou VALOR negativo -> NAO ADERENTE
    ' -----------------------------------------------------------------------
    row = EscreverCabecalhoAlerta(ws, row, _
        "B  |  MATERIAL NAO ADERENTE  (QTD/VALOR incoerentes - espelha a aba MATERIAL)", _
        Array("PEP4NIVEL", "TIPO_PEP", "MATERIAL", "TEXTO_MATERIAL", "CLASSIFICACAO", _
              "CLASSE_CUSTO", "CLS2", "QTD_ENTRADA", "VALOR_MOEDA", "MOTIVO"), _
        corB, corBcl)

    Dim dGQ As Object, dGV As Object, dGFirst As Object
    Set dGQ = CreateObject("Scripting.Dictionary")
    Set dGV = CreateObject("Scripting.Dictionary")
    Set dGFirst = CreateObject("Scripting.Dictionary")
    Dim codG As String, clG As String, kG As String
    For i = 1 To UBound(dados, 1)
        pep = Trim$(CStr(dados(i, cPEP))): If pep = "" Then GoTo ProxBcalc
        If Not EhMaterial(CStr(dados(i, cClassif))) Then GoTo ProxBcalc
        codG = NormCod(dados(i, cMaterial))
        clG = UCase$(Trim$(CStr(dados(i, cClassif))))
        kG = pep & "|" & codG & "|" & clG
        If Not dGFirst.Exists(kG) Then dGFirst(kG) = i
        dGQ(kG) = dGQ(kG) + ToNum(dados(i, cQtd))
        dGV(kG) = dGV(kG) + ToNum(dados(i, cValor))
ProxBcalc:
    Next i

    Dim kvG As Variant, fiG As Long, qG As Double, vG As Double
    Dim motG As String, tipoG As String, tpAneel As String
    For Each kvG In dGFirst.Keys
        fiG = dGFirst(kvG)
        qG = Round(dGQ(kvG), 2): vG = Round(dGV(kvG), 2)
        pep = Trim$(CStr(dados(fiG, cPEP)))
        ' Mesmas exclusoes da aba MATERIAL
        If qG = 0 And vG = 0 Then GoTo ProxB
        If InStr(UCase$(CStr(dados(fiG, cClassif))), "FALTA") > 0 _
           And qG = 0 And vG <> 0 Then GoTo ProxB
        ' Regras de ADERENCIA (mesma ordem/precedencia da aba MATERIAL):
        '   1) QTD=0 com VALOR<>0; 2) sinais opostos; 3) VALOR=0 com QTD<>0;
        '   4) sinal por tipo de PEP (ODD positivo / ODI-ODM-ODS negativo)
        tipoG = TipoPEPCodigo(pep)
        tpAneel = TipoPEPANEEL(pep)
        If tipoG = "S" Then tpAneel = "ODS"
        motG = ""
        If qG = 0 And vG <> 0 Then
            motG = "QTD=0 COM VALOR"
        ElseIf (qG > 0 And vG < 0) Or (qG < 0 And vG > 0) Then
            motG = "SINAIS OPOSTOS QTD x VALOR"
        ElseIf vG = 0 And qG <> 0 Then
            motG = "VALOR=0 COM QTD"
        ElseIf tipoG = "D" Then
            If qG > 0 Then motG = "QTD+"
            If vG > 0 Then motG = Trim$(motG & " VALOR+")
        Else
            If qG < 0 Then motG = "QTD-"
            If vG < 0 Then motG = Trim$(motG & " VALOR-")
        End If
        If motG = "" Then GoTo ProxB   ' ADERENTE -> ignora
        motG = tpAneel & ": " & Trim$(motG)
        contB = contB + 1
        ws.Cells(row, 1).Value = pep
        ws.Cells(row, 2).Value = tpAneel
        ws.Cells(row, 3).Value = NormCod(dados(fiG, cMaterial))
        ws.Cells(row, 4).Value = ValorCampo(fiG, cTexto)
        ws.Cells(row, 5).Value = dados(fiG, cClassif)
        ws.Cells(row, 6).Value = ValorCampo(fiG, cClasse)
        ws.Cells(row, 7).Value = MatInfoLinha(fiG, 2)
        ws.Cells(row, 8).Value = qG
        ws.Cells(row, 9).Value = vG
        ws.Cells(row, 10).Value = motG
        With ws.Range(ws.Cells(row, 1), ws.Cells(row, 10))
            .Font.Color = corInk
            If (contB Mod 2) = 0 Then .Interior.Color = corBzb
        End With
        ws.Cells(row, 8).NumberFormat = "#,##0.00"
        ws.Cells(row, 9).NumberFormat = "#,##0.00"
        With ws.Cells(row, 2)
            .Font.Color = corB: .Font.Bold = True
        End With
        If qG < 0 Or (tipoG = "D" And qG > 0) Then ws.Cells(row, 8).Font.Color = corB
        If vG < 0 Or (tipoG = "D" And vG > 0) Then ws.Cells(row, 9).Font.Color = corB
        With ws.Cells(row, 10)
            .Interior.Color = corBcl
            .Font.Color = corB: .Font.Bold = True: .Font.Size = 8.5
        End With
        row = row + 1
ProxB:
    Next kvG
    If contB = 0 Then
        ws.Cells(row, 1).Value = "(nenhuma ocorrencia encontrada)"
        ws.Cells(row, 1).Font.Italic = True: ws.Cells(row, 1).Font.Color = corMut
        row = row + 1
    End If

    gEtapa = "Alertas: secao E"
    ' -----------------------------------------------------------------------
    ' SECAO E - Classes de custo de VIAGEM com valor consolidado > 0
    ' Monitora: 8111290000 Alimentacao Viagem
    '           8210390000 Servicos de Passagens
    '           8210550000 Servicos de Hospedagem
    ' Consolida por PEP4NIVEL + CLASSE_CUSTO.
    ' -----------------------------------------------------------------------
    row = EscreverCabecalhoAlerta(ws, row + 1, _
        "E  |  CLASSES DE CUSTO VIAGEM  (Alimentacao / Passagem / Hospedagem com valor consolidado > 0)", _
        Array("PEP4NIVEL", "PEP3", "TIPO_PEP", "CLASSE_CUSTO", "DESC_CLASSE_CUSTO", _
              "QTD_LANCAMENTOS", "VALOR_TOTAL"), _
        corE, corEcl)

    Dim dEVal As Object: Set dEVal = CreateObject("Scripting.Dictionary")
    Dim dEQtd As Object: Set dEQtd = CreateObject("Scripting.Dictionary")
    Dim dEDesc As Object: Set dEDesc = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(dados, 1)
        pep = Trim$(CStr(dados(i, cPEP))): If pep = "" Or dPepUAR.Exists(pep) Then GoTo ProxE
        Dim ccEv As String: ccEv = NormCod(ValorCampo(i, cClasse))
        If Not EhClasseViagem(ccEv) Then GoTo ProxE
        val = ToNum(dados(i, cValor))
        Dim chEv As String: chEv = pep & "|" & ccEv
        dEVal(chEv) = dEVal(chEv) + val
        dEQtd(chEv) = dEQtd(chEv) + 1
        If Not dEDesc.Exists(chEv) Then
            Dim dscEv As String: dscEv = TextoCampo(i, cDescClasse)
            If dscEv = "" Then dscEv = DescClasseViagem(ccEv)
            dEDesc(chEv) = dscEv
        End If
ProxE:
    Next i

    Dim pepE As Variant
    For Each pepE In dEVal.Keys
        If CDbl(dEVal(pepE)) > 0 Then
            contE = contE + 1
            Dim pvEv() As String: pvEv = Split(CStr(pepE), "|")
            pep = CStr(pvEv(0))
            ws.Cells(row, 1).Value = pep
            ws.Cells(row, 2).Value = PEP3(pep)
            ws.Cells(row, 3).Value = TipoPEPANEEL(pep)
            ws.Cells(row, 4).Value = CStr(pvEv(1))
            ws.Cells(row, 5).Value = IIf(dEDesc.Exists(CStr(pepE)), dEDesc(CStr(pepE)), "")
            ws.Cells(row, 6).Value = dEQtd(CStr(pepE))
            ws.Cells(row, 7).Value = Round(CDbl(dEVal(CStr(pepE))), 2)
            With ws.Range(ws.Cells(row, 1), ws.Cells(row, 7))
                .Font.Color = corInk
                If (contE Mod 2) = 0 Then .Interior.Color = corEzb
            End With
            With ws.Cells(row, 3)
                .Font.Color = corE: .Font.Bold = True
            End With
            ws.Cells(row, 7).NumberFormat = "#,##0.00"
            With ws.Cells(row, 7)
                .Interior.Color = corEcl
                .Font.Color = corE: .Font.Bold = True
            End With
            row = row + 1
        End If
    Next pepE
    If contE = 0 Then
        ws.Cells(row, 1).Value = "(nenhuma ocorrencia encontrada)"
        ws.Cells(row, 1).Font.Italic = True: ws.Cells(row, 1).Font.Color = corMut
        row = row + 1
    End If

    ' -----------------------------------------------------------------------
    ' SECAO F - Classes de custo marcadas como RISCO (servicos de obra,
    ' saldo remanescente, ligacao nova) com valor diferente de zero.
    ' *** DESABILITADA TEMPORARIAMENTE - logica em desenvolvimento ***
    ' -----------------------------------------------------------------------
    ' (secao F reservada para implementacao futura)

    gEtapa = "Alertas: cards resumo"
    ' -----------------------------------------------------------------------
    ' Cards de resumo (contagens reais de cada secao)
    ' -----------------------------------------------------------------------
    EscreverCardAlerta ws, rowCards, 1, "PEPS SEM UC", contA, corA, corAcl
    EscreverCardAlerta ws, rowCards, 3, "MATERIAL NAO ADERENTE", contB, corB, corBcl
    EscreverCardAlerta ws, rowCards, 5, "UC SEM SERVICO", contC, corC, corCcl
    EscreverCardAlerta ws, rowCards, 7, "CLASSES VIAGEM", contE, corE, corEcl

    gEtapa = "Alertas: formatacao final"
    ' -----------------------------------------------------------------------
    ' Formatacao final da aba
    ' -----------------------------------------------------------------------
    ws.Columns("A").ColumnWidth = 36
    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 16
    ws.Columns("E").ColumnWidth = 18
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 14
    ws.Columns("H").ColumnWidth = 32
    ws.Columns("I").ColumnWidth = 24
    ws.Columns("J").ColumnWidth = 14

    On Error Resume Next
    ws.Tab.Color = CorAba("ALERTAS CRITICOS")
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 90
    On Error GoTo 0
    PrepararImpressao ws, "ALERTAS CRITICOS", row, 10
    ' Congela cabecalho + cards de resumo (sticky ao rolar as secoes)
    AplicarFreeze ws, "A" & (rowCards + 3), congelar:=True
End Sub

' Helper: card de resumo (bloco de 3 colunas, borda completa, barra de acento
' no topo, rotulo em cima e numero grande centralizado).
Private Sub EscreverCardAlerta(ws As Worksheet, ByVal r As Long, ByVal c As Long, _
        ByVal rotulo As String, ByVal valor As Long, _
        ByVal cor As Long, ByVal corFundo As Long)
    Dim bloco As Range
    Set bloco = ws.Range(ws.Cells(r, c), ws.Cells(r + 1, c + 1))
    With bloco
        .Interior.Color = corFundo
        ' borda completa fina na cor da secao
        .Borders(xlEdgeLeft).LineStyle = xlContinuous:   .Borders(xlEdgeLeft).Color = cor
        .Borders(xlEdgeRight).LineStyle = xlContinuous:  .Borders(xlEdgeRight).Color = cor
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Color = cor
        ' barra de acento grossa no topo
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThick: .Borders(xlEdgeTop).Color = cor
    End With
    ' rotulo (linha de cima, mesclado nas 2 colunas do card)
    With ws.Range(ws.Cells(r, c), ws.Cells(r, c + 1))
        .Merge
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    With ws.Cells(r, c)
        .Value = UCase$(rotulo)
        .Font.Name = FONTE_UI
        .Font.Size = 8: .Font.Bold = True: .Font.Color = cor
        .IndentLevel = 1
    End With
    ' numero grande centralizado no bloco inferior (mescla c..c+1)
    With ws.Range(ws.Cells(r + 1, c), ws.Cells(r + 1, c + 1))
        .Merge
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    With ws.Cells(r + 1, c)
        .Value = valor
        .Font.Name = FONTE_UI
        .Font.Size = 26: .Font.Bold = True
        .Font.Color = IIf(valor > 0, cor, COR_UI_MUDO)
    End With
    ws.Rows(r).RowHeight = 18
    ws.Rows(r + 1).RowHeight = 40
End Sub

' Helper: faixa de titulo da secao (cor cheia + texto branco) + cabecalho tintado
Private Function EscreverCabecalhoAlerta(ws As Worksheet, ByVal startRow As Long, _
        ByVal titulo As String, cabecalhos As Variant, _
        ByVal corForte As Long, ByVal corClara As Long) As Long
    Dim nCols As Long: nCols = UBound(cabecalhos) - LBound(cabecalhos) + 1
    Dim r As Long: r = startRow

    ' Faixa de titulo da secao: cor cheia + texto branco
    ws.Cells(r, 1).Value = titulo
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, nCols))
        .Merge
        .Font.Name = FONTE_UI
        .Font.Bold = True: .Font.Size = 11: .Font.Color = vbWhite
        .Interior.Color = corForte
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    ws.Rows(r).RowHeight = 26
    r = r + 1

    ' Cabecalho das colunas: tinta clara da secao + texto na cor forte
    Dim j As Long
    For j = 0 To nCols - 1
        ws.Cells(r, j + 1).Value = cabecalhos(j)
    Next j
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, nCols))
        .Font.Name = FONTE_UI
        .Font.Bold = True: .Font.Size = 8.5
        .Font.Color = corForte
        .Interior.Color = corClara
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeBottom).Color = corForte
    End With
    ws.Rows(r).RowHeight = 20
    r = r + 1

    EscreverCabecalhoAlerta = r
End Function


'==============================================================================
'  ESCRITA + FORMATAO
'==============================================================================
Private Sub Gerar_Regras()
    ' Aba REGRAS: documenta as regras de negocio aplicadas pelo relatorio.
    Dim s As String
    s = ""
    s = s & "CLASSIFICACAO TIPO|De-para familia -> TIPO|Cada familia (CLS2) e mapeada para COM, UC ou UAR por uma tabela fixa. Onde a mesma familia aparece com TIPO diferente, vale o primeiro (igual PROCV). Comparacao ignora acento e maiusculas." & vbLf
    s = s & "CLASSIFICACAO TIPO|CH FUS 34,5 kV = UC|CH FUS e COM por padrao. Quando o texto do material contem '34,5' (34,5 kV) a familia CH FUS passa a UC. CH FUS RELIG e sempre UC." & vbLf
    s = s & "ADERENCIA|Veredito so sobre familias UC|O veredito APROVADO/REPROVADO do PEP considera apenas as familias classificadas como UC. Familias COM e UAR nao reprovam o PEP." & vbLf
    s = s & "ADERENCIA|Margem de 10%|Uma familia UC e ADERENTE quando a soma de MATERIAL (MAT) e a soma de SERVICO (SRV) do PEP ficam dentro de 10% de diferenca." & vbLf
    s = s & "ADERENCIA|Cabo em KG -> metros|Materiais de cabo lancados em KG sao convertidos para metros (metros = kg x fator) antes de comparar MAT com SRV." & vbLf
    s = s & "ADERENCIA|Servico COMBO|Alguns servicos tem fator multiplicador (ex.: cruzeta dupla = 2, por km x 1000). A QTD de servico e multiplicada por esse fator antes da comparacao." & vbLf
    s = s & "ADERENCIA|Cabo < 15 m negligenciavel|Se ha servico lancado e o maior valor entre MAT e SRV de cabo for menor que 15 m, a familia de cabo e considerada ADERENTE." & vbLf
    s = s & "ADERENCIA|RAMAL (75 m por servico)|Cada servico de RAMAL contempla ate 75 m de material RAMAL. Se MAT <= SRV x 75, o RAMAL fica ADERENTE." & vbLf
    s = s & "ADERENCIA|RAMAL cobre cabo (30 m por servico)|Cada servico de RAMAL libera 30 m. Se SRV(RAMAL) x 30 cobrir (>=) o total de cabo do PEP, os cabos/condutores ficam ADERENTES." & vbLf
    s = s & "ADERENCIA|RELIGADOR (60 m por servico)|Cada servico de RELIGADOR libera 60 m para tornar aderentes COND PROT, CABO ISOLADO e COND COBRE (2 servicos = 120 m, etc.)." & vbLf
    s = s & "ADERENCIA|COND NU + CORDOALHA|O servico de COND NU contempla o material CORDOALHA no mesmo PEP." & vbLf
    s = s & "ADERENCIA|CH SU FACA (PI = ERD)|Em PI do tipo ERD, cada servico de RELIGADOR deixa aderentes ate 6 CH SU FACA." & vbLf
    s = s & "ADERENCIA|COND COBRE como conexao da CH SU FACA|Quando ha servico de CH SU FACA no PEP, o COND COBRE (conexao) fica ADERENTE." & vbLf
    s = s & "ADERENCIA|CH SU FACA paga conexao de cabo|Cada servico de CH SU FACA abona 7 m de cabo (COND PROT etc.). O cabo fica ADERENTE se Abs(MAT) <= Abs(SRV) + SRV_CHSUFACA x 7." & vbLf
    s = s & "UNIFICACAO DE FAMILIAS|COND NU / ISOLADO -> COND PROT|As familias COND NU, COND ISOLADO e COND ISOLADO/PROT sao somadas junto com COND PROT na mesma linha da aba MATERIAL vs SERVICO, e comparadas como uma so familia." & vbLf
    s = s & "TRAFO / REGULADOR|Servico de TRAFO/REGULADOR cobre acessorios|Quando ha servico de TRAFO ou REGULADOR no PEP4, as familias CH FUS, PARA RAIO MT e PARA RAIO BT ficam ADERENTES (o servico paga os acessorios). Vale para ODD e ODI." & vbLf
    s = s & "TRAFO / REGULADOR|Servico 5500200645 cobre acessorios|O codigo de servico 5500200645 (instalacao/troca de trafo) tambem torna aderentes CH FUS, PARA RAIO MT e PARA RAIO BT no PEP4, independente da familia dele." & vbLf
    s = s & "CH FUSIVEL|Servico que paga CH fusivel|Os servicos 5500000618 (equivale a uma CH fusivel - material 105300003) e 5500200645 tornam a familia CH FUS ADERENTE quando presentes no PEP." & vbLf
    s = s & "TRAFO / REGULADOR|TRAFO_OU_REGULADOR x TRAFO/REGULADOR|Sao complementares: se o PEP tem servico TRAFO_OU_REGULADOR e material de TRAFO ou REGULADOR (ou vice-versa), ambas as linhas ficam ADERENTES mesmo com familias diferentes." & vbLf
    s = s & "COND COBRE|Aterramento de equipamentos|Se o PEP4 tem PARA RAIO BT, PARA RAIO MT, REGULADOR ou TRAFO (material ou servico), o COND COBRE fica ADERENTE (e o aterramento desses equipamentos)." & vbLf
    s = s & "VEREDITO PEP|ODD (.D) nao exige UC|PEP do tipo ODD (.D) e sempre APROVADO; nao exige aderencia de familias UC." & vbLf
    s = s & "VEREDITO PEP|PEP com UAR|PEP sem familia UC mas com familia UAR e APROVADO." & vbLf
    s = s & "VEREDITO PEP|PEP sem UC|PEP sem nenhuma familia UC recebe o status SEM UC." & vbLf
    s = s & "PROPAGACAO PEP3|ODI reprovada reprova o PEP3|Se a ODI (.I) de um PEP3NIVEL for REPROVADA, todos os PEP4NIVEL do mesmo PEP3 (inclusive ODD) ficam REPROVADOS." & vbLf
    s = s & "PROPAGACAO PEP3|ODI sem UC reprova o PEP3|Se a ODI (.I) de um PEP3NIVEL estiver SEM UC, todo o PEP3 fica REPROVADO." & vbLf
    s = s & "ADERENCIA ANEEL (aba MATERIAL)|Sinal por tipo de PEP|Na aba MATERIAL: ODD (.D) com QTD ou VALOR positivo = NAO ADERENTE; ODI/ODM/ODS com QTD ou VALOR negativo = NAO ADERENTE." & vbLf
    s = s & "OBSERVACOES|Motivo do REPROVADO (OBS1)|A coluna OBS1 da aba MATERIAL vs SERVICO traz o motivo da linha: familia UC fora da margem (com MAT/SRV) e, na propagacao, qual ODI reprovou e quais familias (ou ODI sem UC)." & vbLf
    s = s & "OBSERVACOES|Veredito do PEP3 (OBS2)|A coluna OBS2 repete em TODAS as linhas do PEP3 o motivo do veredito: APROVADO (nenhuma familia UC fora da margem) ou REPROVADO (com o motivo agregado)." & vbLf
    s = s & "FILTRAGEM|Linhas NULO ignoradas|Linhas com MAT = 0 e SRV = 0 (NULO) nao sao trazidas para a aba MATERIAL vs SERVICO." & vbLf
    s = s & "ABA MATERIAL|SITUACAO (numero) + OBS_SITUACAO|Duas colunas: SITUACAO traz o numero (1..9) e OBS_SITUACAO a descricao. Classifica por QTD/VALOR e tipo de PEP. 1:QTD<0/VLR<0 (ODI/ODM/ODS); 2:QTD=0/VLR<0 (todos); 3:QTD<0/VLR>0 (ODI/ODM/ODS); 4:QTD>0/VLR=0 (ODI/ODM/ODS, PU>0); 5:QTD>0/VLR<0 (ODI/ODM/ODS, PU>0); 6:QTD=0/VLR>0 (todos); 7:QTD<0/VLR=0 (todos); 8:QTD<0/VLR>0 (ODD, PU<0); 9:QTD>0/VLR=0 (ODD). QTD>0/VLR>0 = lancamento normal (sem situacao)." & vbLf
    s = s & "ABA SERVICO|Unidade de medida (UND)|Preferencia: UML do proprio lancamento SAP; se vazio, usa a unidade (UN) do catalogo SERVICOS_ATUAIS pelo COD_SERVICO." & vbLf
    s = s & "ABA SERVICO|ALERTA_VALOR|Check automatico de valor do servico: VALOR e/ou QTD negativos sao sinalizados; caso contrario OK." & vbLf
    s = s & "ABA SERVICO|APROPRIACAO|Check automatico de apropriacao: TIPO_APLICACAO ODI* lancado em PEP .D (ou ODD* em .I) -> VERIFICAR; fallback por palavra-chave (INST em ODD, RET/DESATIV em ODI)." & vbLf
    s = s & "ABA SERVICO|Descricao com fallback|DESCRICAO_SERVICO usa o catalogo (embutido + SERVICOS_ATUAIS); se nao encontrado, usa o TEXTO BREVE do proprio lancamento SAP." & vbLf
    s = s & "ABA SERVICO|ADERENCIA|QTD=0 com VALOR (+ ou -) -> NAO ADERENTE; ODD (.D) com QTD ou VALOR positivo -> NAO ADERENTE; ODI/ODM/ODS com QTD ou VALOR negativo -> NAO ADERENTE." & vbLf
    s = s & "ABA SERVICO|CAVA x POSTE|Servico de CAVA (ESCAVACAO) com quantidade maior que a de poste no PEP -> NAO ADERENTE. A cava NAO e ligada ao poste (nao entra na MATERIAL vs SERVICO), apenas sinalizada. POSTE = qtd de servico de poste; se 0, usa a qtd de material poste." & vbLf

    Dim linhas() As String: linhas = Split(s, vbLf)
    Dim i As Long, cnt As Long
    For i = 0 To UBound(linhas)
        If Trim$(linhas(i)) <> "" Then cnt = cnt + 1
    Next i

    Dim outp() As Variant: ReDim outp(0 To cnt, 1 To 3)
    outp(0, 1) = "GRUPO": outp(0, 2) = "REGRA": outp(0, 3) = "DESCRICAO"
    Dim rr As Long: rr = 0
    Dim parts() As String
    For i = 0 To UBound(linhas)
        If Trim$(linhas(i)) <> "" Then
            parts = Split(linhas(i), "|")
            rr = rr + 1
            outp(rr, 1) = parts(0)
            outp(rr, 2) = parts(1)
            outp(rr, 3) = parts(2)
        End If
    Next i

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("REGRAS")
    If Not ws Is Nothing Then ws.Delete
    On Error GoTo 0
    Set ws = ActiveWorkbook.Worksheets.Add( _
        After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    ws.Name = "REGRAS"
    ws.Range(ws.Cells(1, 1), ws.Cells(cnt + 1, 3)).Value = outp

    ' Visual padrao (cabecalho verde, bordas) sem reordenar as linhas
    FormatarVisualAba ws, "REGRAS", cnt + 1, 3

    ' Ajusta a coluna DESCRICAO: larga e com quebra de texto
    ws.Columns(1).ColumnWidth = 28
    ws.Columns(2).ColumnWidth = 38
    ws.Columns(3).ColumnWidth = 95
    ws.Range(ws.Cells(2, 1), ws.Cells(cnt + 1, 3)).WrapText = True
    ws.Range(ws.Cells(2, 1), ws.Cells(cnt + 1, 3)).VerticalAlignment = xlTop
    ws.Range(ws.Cells(2, 1), ws.Cells(cnt + 1, 3)).EntireRow.AutoFit
    On Error Resume Next
    ws.PageSetup.Orientation = xlPortrait
    On Error GoTo 0
End Sub

Private Sub EscreverAba(nome As String, outp() As Variant)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(nome)
    If Not ws Is Nothing Then ws.Delete
    On Error GoTo 0
    Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    ws.Name = nome

    Dim nR As Long, nC As Long
    nR = UBound(outp, 1) + 1: nC = UBound(outp, 2)
    ws.Range(ws.Cells(1, 1), ws.Cells(nR, nC)).Value = outp

    ' --- Ordenacao por aba ---------------------------------------------------
    If nR > 2 Then Call OrdenarAba(ws, nome, nR, nC)

    ' --- Visual padrao (FASE 5): guia, cabecalho, zebra, bordas, formatos ----
    FormatarVisualAba ws, nome, nR, nC

    ' Colore TODAS as colunas de veredito (verde = bom / vermelho = problema /
    ' cinza = neutro), em blocos contiguos. Aplicado APOS o visual: as colunas
    ' de veredito ficam fora da zebra, entao as cores prevalecem.
    If nR >= 2 Then
        Dim jc As Long, hh As String
        For jc = 1 To nC
            hh = UCase$(CStr(ws.Cells(1, jc).Value))
            If EhColunaVeredito(hh) Then ColorirColunaVeredito ws, jc, nR
        Next jc
    End If
End Sub
'==============================================================================
'  ORDENACAO POR ABA
'  Aplica Sort de ate 3 niveis conforme a estrutura de colunas de cada aba.
'  Chamado por EscreverAba logo apos a escrita dos dados, antes da formatacao.
'
'  Regra geral: PEP3NIVEL > PEP4NIVEL > classificacao/material/classe
'  (espelha o criterio mostrado na imagem de referencia)
'==============================================================================
Private Sub OrdenarAba(ws As Worksheet, ByVal nome As String, _
                        ByVal nR As Long, ByVal nC As Long)
    On Error GoTo SemSort

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(nR, nC))

    ' Localiza coluna pelo cabecalho (busca exata, case-insensitive)
    Dim c1 As Long, c2 As Long, c3 As Long
    c1 = 0: c2 = 0: c3 = 0

    Select Case UCase$(Trim$(nome))

        ' ------------------------------------------------------------------
        ' MATERIAL vs SERVICO
        '   PEP3NIVEL (col 2) > PEP4NIVEL (col 3) > CLASSIFICACAO (col 4)
        ' ------------------------------------------------------------------
        Case "MATERIAL VS SERVICO"
            c1 = 2: c2 = 3: c3 = 4

        ' ------------------------------------------------------------------
        ' MATERIAL
        '   PEP4NIVEL (col 1) > PEP3 (col 2) > MATERIAL (col 5)
        ' ------------------------------------------------------------------
        Case "MATERIAL"
            c1 = 1: c2 = 2: c3 = 5

        ' ------------------------------------------------------------------
        ' SERVICO
        '   PEP4NIVEL (col 1) > PEP3 (col 2) > COD_SERVICO (col 4)
        ' ------------------------------------------------------------------
        Case "SERVICO"
            c1 = 1: c2 = 2: c3 = 4

        ' ------------------------------------------------------------------
        ' ANALISE DE CA
        '   EMPRESA (col 1) > PEP4NIVEL (col 2)
        ' ------------------------------------------------------------------
        Case "ANALISE DE CA"
            c1 = 1: c2 = 2

        ' ------------------------------------------------------------------
        ' CLASSE DE CUSTO
        '   PEP4NIVEL (col 2) > TIPO_PEP (col 3) > CLASSE_CUSTO (col 4)
        ' ------------------------------------------------------------------
        Case "CLASSE DE CUSTO"
            c1 = 2: c2 = 3: c3 = 4

        ' ------------------------------------------------------------------
        ' RAZAO CJ  - usa indices das colunas mapeadas da base crua
        '   PEP (cPEP) > CLASSE_CUSTO (cClasse) > CLASSIFICACAO (cClassif)
        ' ------------------------------------------------------------------
        Case "RAZAO CJ"
            If cPEP > 0 Then c1 = cPEP
            If cClasse > 0 Then c2 = cClasse
            If cClassif > 0 Then c3 = cClassif

        Case "SERVICO SEM MATERIAL"
            c1 = 1: c2 = 3   ' COD_SERVICO > FAMILIA_CLS2

        Case "PORTFOLIO OBRA"
            c1 = 1: c2 = 2   ' EMPRESA > PEP3NIVEL

        Case "NAO CLASSIFICADOS"
            c1 = 1: c2 = 4   ' PEP4NIVEL > TIPO_LANCAMENTO

        Case "PAINEL EXECUTIVO"
            c1 = 1: c2 = 2   ' EMPRESA > PEP3NIVEL

        Case "RACIONALIZACAO COM"
            c1 = 1: c2 = 3   ' PEP4NIVEL > COD_MATERIAL (substitui bubble sort)

        Case Else
            Exit Sub   ' aba sem regra definida -> nao ordena
    End Select

    If c1 = 0 Then Exit Sub

    Dim dataRng As Range
    Set dataRng = ws.Range(ws.Cells(1, 1), ws.Cells(nR, nC))

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Cells(1, c1), Order:=xlAscending
        If c2 > 0 Then .SortFields.Add Key:=ws.Cells(1, c2), Order:=xlAscending
        If c3 > 0 Then .SortFields.Add Key:=ws.Cells(1, c3), Order:=xlAscending
        .SetRange dataRng
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    Exit Sub
SemSort:
    ' Falha silenciosa: a aba e gerada sem ordenacao
End Sub


'==============================================================================
'  ABA: PAINEL EXECUTIVO
'  Uma linha por PEP3NIVEL com os principais indicadores consolidados:
'
'  FINANCEIRO
'    VALOR_TOTAL       : soma de todos os lancamentos do PEP3
'    VALOR_MAT         : soma dos lancamentos de material (UC + COM + UAR)
'    VALOR_CA          : soma dos lancamentos de servico/CA
'    PERC_MAT          : VALOR_MAT / VALOR_TOTAL %
'    PERC_CA           : VALOR_CA  / VALOR_TOTAL %
'
'  ADERENCIA (da aba MATERIAL vs SERVICO)
'    STATUS_ODI        : APROVADO / REPROVADO / SEM UC da ODI do PEP3
'    QTD_FAM_NAO_ADER  : quantas familias UC estao NAO ADERENTES na ODI
'    DIFERENCA_TOTAL   : soma das diferencas MAT-SRV das familias NAO ADERENTES
'
'  ALERTAS
'    TEM_ALIMENTACAO   : S/N — tem lancamento na classe 8111290000 > 0
'    TEM_RISCO         : S/N — tem lancamento em classe marcada como RISCO
'    TEM_QTD_SEM_VALOR : S/N — tem material com QTD+ e VALOR=0
'    TEM_VALOR_SEM_QTD : S/N — tem material com QTD- e VALOR>=0
'
'  COMPOSICAO CA
'    MOP, FRETE, TRIBUTOS, SUPORTE, JOA  (valores consolidados por categoria)
'==============================================================================
Private Sub Gerar_PainelExecutivo()

    ' -----------------------------------------------------------------------
    ' Passagem unica: acumula tudo por PEP3NIVEL
    ' -----------------------------------------------------------------------
    Dim dOrdem    As Object: Set dOrdem    = CreateObject("Scripting.Dictionary") ' preserva ordem
    Dim dVTot     As Object: Set dVTot     = CreateObject("Scripting.Dictionary") ' valor total
    Dim dVMat     As Object: Set dVMat     = CreateObject("Scripting.Dictionary") ' valor material
    Dim dVCA      As Object: Set dVCA      = CreateObject("Scripting.Dictionary") ' valor CA/servico
    Dim dVMOP     As Object: Set dVMOP     = CreateObject("Scripting.Dictionary")
    Dim dVFrete   As Object: Set dVFrete   = CreateObject("Scripting.Dictionary")
    Dim dVTrib    As Object: Set dVTrib    = CreateObject("Scripting.Dictionary")
    Dim dVSup     As Object: Set dVSup     = CreateObject("Scripting.Dictionary")
    Dim dVJOA     As Object: Set dVJOA     = CreateObject("Scripting.Dictionary")
    Dim dAlimS    As Object: Set dAlimS    = CreateObject("Scripting.Dictionary") ' flag alimentacao
    Dim dRiscoS   As Object: Set dRiscoS   = CreateObject("Scripting.Dictionary") ' flag risco
    Dim dQtdSemV  As Object: Set dQtdSemV  = CreateObject("Scripting.Dictionary") ' flag qtd+/val=0
    Dim dValSemQ  As Object: Set dValSemQ  = CreateObject("Scripting.Dictionary") ' flag qtd-/val>=0
    Dim dEmpresa  As Object: Set dEmpresa  = CreateObject("Scripting.Dictionary")

    Dim i As Long, pepPn As String, p3Pn As String, valPn As Double, qPn As Double
    Dim catPn As String, cls3ccPn As String

    For i = 1 To UBound(dados, 1)
        pepPn = Trim$(CStr(dados(i, cPEP))): If pepPn = "" Then GoTo Prox
        p3Pn = PEP3(pepPn)
        valPn = ToNum(dados(i, cValor))
        qPn   = ToNum(dados(i, cQtd))

        If Not dOrdem.Exists(p3Pn) Then
            dOrdem(p3Pn) = dOrdem.Count
            If Not dEmpresa.Exists(p3Pn) Then dEmpresa(p3Pn) = TextoCampo(i, cEmpresa)
        End If

        dVTot(p3Pn) = dVTot(p3Pn) + valPn

        If EhMaterial(CStr(dados(i, cClassif))) Then
            dVMat(p3Pn) = dVMat(p3Pn) + valPn
            If qPn > 0 And valPn <= 0 Then dQtdSemV(p3Pn) = "S"
            If qPn < 0 And valPn >= 0 Then dValSemQ(p3Pn) = "S"
        Else
            dVCA(p3Pn) = dVCA(p3Pn) + valPn
            catPn = UCase$(CategoriaAnaliseCA(i))
            Select Case catPn
                Case "MOP":          dVMOP(p3Pn)   = dVMOP(p3Pn)   + valPn
                Case "FRETE_TRANSP": dVFrete(p3Pn) = dVFrete(p3Pn) + valPn
                Case "TRIBUTOS":     dVTrib(p3Pn)  = dVTrib(p3Pn)  + valPn
                Case "SUPORTE":      dVSup(p3Pn)   = dVSup(p3Pn)   + valPn
                Case "JOA":          dVJOA(p3Pn)   = dVJOA(p3Pn)   + valPn
            End Select
        End If

        If EhClasseViagem(NormCod(ValorCampo(i, cClasse))) And valPn > 0 Then dAlimS(p3Pn) = "S"
        cls3ccPn = UCase$(CCInfo(ValorCampo(i, cClasse), 2))
        If cls3ccPn = "RISCO" And valPn <> 0 Then dRiscoS(p3Pn) = "S"
Prox:
    Next i

    ' -----------------------------------------------------------------------
    ' FASE 1.2 - Aderencia ODI: REUTILIZA os vereditos calculados em
    ' Gerar_MaterialVsServico via dicionarios em memoria. Fonte unica de
    ' verdade, sem dependencia de indices de coluna nem da ordem de geracao.
    ' Guard: se a MvS nao tiver rodado, segue com dicionarios vazios
    ' (todos os PEP3 ficam "SEM ODI", mesmo efeito de antes sem a aba).
    ' -----------------------------------------------------------------------
    Dim dODIVerd  As Object
    Dim dODIFamNC As Object
    Dim dODIDif   As Object
    If dMvSVerd Is Nothing Then Set dMvSVerd = CreateObject("Scripting.Dictionary")
    If dMvSFamNC Is Nothing Then Set dMvSFamNC = CreateObject("Scripting.Dictionary")
    If dMvSDif Is Nothing Then Set dMvSDif = CreateObject("Scripting.Dictionary")
    Set dODIVerd = dMvSVerd
    Set dODIFamNC = dMvSFamNC
    Set dODIDif = dMvSDif

    ' -----------------------------------------------------------------------
    ' Monta output
    ' -----------------------------------------------------------------------
    Dim nP As Long: nP = dOrdem.Count
    Dim outp() As Variant: ReDim outp(0 To nP, 1 To 22)

    outp(0, 1)  = "EMPRESA"
    outp(0, 2)  = "PEP3NIVEL"
    outp(0, 3)  = "VALOR_TOTAL"
    outp(0, 4)  = "VALOR_MAT"
    outp(0, 5)  = "VALOR_CA"
    outp(0, 6)  = "PERC_MAT_%"
    outp(0, 7)  = "PERC_CA_%"
    outp(0, 8)  = "STATUS_ODI"
    outp(0, 9)  = "FAM_NAO_ADERENTES"
    outp(0, 10) = "DIFERENCA_MAT_SRV"
    outp(0, 11) = "ALERTA_ALIMENTACAO"
    outp(0, 12) = "ALERTA_RISCO"
    outp(0, 13) = "ALERTA_QTD_SEM_VALOR"
    outp(0, 14) = "ALERTA_VALOR_SEM_QTD"
    outp(0, 15) = "QTD_ALERTAS"
    outp(0, 16) = "MOP"
    outp(0, 17) = "FRETE_TRANSP"
    outp(0, 18) = "TRIBUTOS"
    outp(0, 19) = "SUPORTE"
    outp(0, 20) = "JOA"
    outp(0, 21) = "PERC_MOP_%"
    outp(0, 22) = "PERC_CA_SEM_MOP_%"

    Dim pepsPn As Variant: pepsPn = dOrdem.Keys
    Dim rPn As Long, vTotPn As Double, vMatPn As Double, vCAPn As Double, vMopPn As Double
    Dim nAlertasPn As Long, statusODIPn As String, p3outPn As String

    For rPn = 0 To nP - 1
        p3outPn = CStr(pepsPn(rPn))
        vTotPn = IIf(dVTot.Exists(p3outPn), CDbl(dVTot(p3outPn)), 0)
        vMatPn = IIf(dVMat.Exists(p3outPn), CDbl(dVMat(p3outPn)), 0)
        vCAPn  = IIf(dVCA.Exists(p3outPn),  CDbl(dVCA(p3outPn)),  0)
        vMopPn = IIf(dVMOP.Exists(p3outPn), CDbl(dVMOP(p3outPn)), 0)

        statusODIPn = IIf(dODIVerd.Exists(p3outPn), CStr(dODIVerd(p3outPn)), "SEM ODI")

        nAlertasPn = 0
        If dAlimS.Exists(p3outPn)   Then nAlertasPn = nAlertasPn + 1
        If dRiscoS.Exists(p3outPn)  Then nAlertasPn = nAlertasPn + 1
        If dQtdSemV.Exists(p3outPn) Then nAlertasPn = nAlertasPn + 1
        If dValSemQ.Exists(p3outPn) Then nAlertasPn = nAlertasPn + 1
        If statusODIPn = "REPROVADO" Then nAlertasPn = nAlertasPn + 1

        outp(rPn + 1, 1)  = IIf(dEmpresa.Exists(p3outPn), dEmpresa(p3outPn), "")
        outp(rPn + 1, 2)  = p3outPn
        outp(rPn + 1, 3)  = Round(vTotPn, 2)
        outp(rPn + 1, 4)  = Round(vMatPn, 2)
        outp(rPn + 1, 5)  = Round(vCAPn, 2)
        If vTotPn <> 0 Then
            outp(rPn + 1, 6) = Round(vMatPn / vTotPn * 100, 1)
            outp(rPn + 1, 7) = Round(vCAPn  / vTotPn * 100, 1)
        End If
        outp(rPn + 1, 8)  = statusODIPn
        outp(rPn + 1, 9)  = IIf(dODIFamNC.Exists(p3outPn), dODIFamNC(p3outPn), 0)
        outp(rPn + 1, 10) = IIf(dODIDif.Exists(p3outPn),   Round(CDbl(dODIDif(p3outPn)), 2), 0)
        outp(rPn + 1, 11) = IIf(dAlimS.Exists(p3outPn),   "S", "")
        outp(rPn + 1, 12) = IIf(dRiscoS.Exists(p3outPn),  "S", "")
        outp(rPn + 1, 13) = IIf(dQtdSemV.Exists(p3outPn), "S", "")
        outp(rPn + 1, 14) = IIf(dValSemQ.Exists(p3outPn), "S", "")
        outp(rPn + 1, 15) = nAlertasPn
        outp(rPn + 1, 16) = IIf(dVMOP.Exists(p3outPn),   Round(CDbl(dVMOP(p3outPn)),   2), 0)
        outp(rPn + 1, 17) = IIf(dVFrete.Exists(p3outPn), Round(CDbl(dVFrete(p3outPn)), 2), 0)
        outp(rPn + 1, 18) = IIf(dVTrib.Exists(p3outPn),  Round(CDbl(dVTrib(p3outPn)),  2), 0)
        outp(rPn + 1, 19) = IIf(dVSup.Exists(p3outPn),   Round(CDbl(dVSup(p3outPn)),   2), 0)
        outp(rPn + 1, 20) = IIf(dVJOA.Exists(p3outPn),   Round(CDbl(dVJOA(p3outPn)),   2), 0)
        Dim vCASemMopPn As Double: vCASemMopPn = vCAPn - vMopPn
        If vCASemMopPn <> 0 Then outp(rPn + 1, 21) = Round(vMopPn / vCASemMopPn * 100, 1)
        If vTotPn <> 0 Then outp(rPn + 1, 22) = Round(vCASemMopPn / vTotPn * 100, 1)
    Next rPn

    ' -----------------------------------------------------------------------
    ' Escreve e formata
    ' -----------------------------------------------------------------------
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("PAINEL EXECUTIVO")
    If Not ws Is Nothing Then ws.Delete
    On Error GoTo 0
    Set ws = ActiveWorkbook.Worksheets.Add( _
        After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    ws.Name = "PAINEL EXECUTIVO"

    ws.Range(ws.Cells(1, 1), ws.Cells(nP + 1, 22)).Value = outp

    ' FASE 5: visual padrao (guia, zebra, formatos, bordas); grupos de
    ' cabecalho e vereditos sao recoloridos logo abaixo, por cima.
    FormatarVisualAba ws, "PAINEL EXECUTIVO", nP + 1, 22, "C2"


    ' Grupos de colunas com cores distintas no cabecalho
    ' Financeiro (3-7): azul
    With ws.Range(ws.Cells(1, 3), ws.Cells(1, 7))
        .Interior.Color = RGB(28, 64, 99)
    End With
    ' Aderencia (8-10): verde escuro
    With ws.Range(ws.Cells(1, 8), ws.Cells(1, 10))
        .Interior.Color = COR_UI_HDR_ACENTO
    End With
    ' Alertas (11-15): vermelho escuro
    With ws.Range(ws.Cells(1, 11), ws.Cells(1, 15))
        .Interior.Color = RGB(122, 26, 32)
    End With
    ' CA detalhe (16-22): roxo escuro
    With ws.Range(ws.Cells(1, 16), ws.Cells(1, 22))
        .Interior.Color = RGB(72, 47, 102)
    End With

    ' Colore linhas por STATUS_ODI e QTD_ALERTAS
    If nP > 0 Then
        Dim rrPn As Long
        For rrPn = 2 To nP + 1
            Dim stODIPn As String: stODIPn = UCase$(Trim$(CStr(ws.Cells(rrPn, 8).Value)))
            Dim nAlPn As Long: nAlPn = CLng(ws.Cells(rrPn, 15).Value)
            Select Case stODIPn
                Case "REPROVADO"
                    ws.Cells(rrPn, 8).Interior.Color = COR_UI_BAD_BG
                    ws.Cells(rrPn, 8).Font.Color = COR_UI_BAD_FG
                    ws.Cells(rrPn, 8).Font.Bold = True
                Case "APROVADO"
                    ws.Cells(rrPn, 8).Interior.Color = COR_UI_OK_BG
                    ws.Cells(rrPn, 8).Font.Color = COR_UI_OK_FG
                    ws.Cells(rrPn, 8).Font.Bold = True
                Case Else
                    ws.Cells(rrPn, 8).Interior.Color = COR_UI_NEU_BG
                    ws.Cells(rrPn, 8).Font.Color = COR_UI_NEU_FG
            End Select
            Select Case True
                Case nAlPn = 0:
                Case nAlPn = 1
                    ws.Cells(rrPn, 15).Interior.Color = COR_UI_WARN_BG
                    ws.Cells(rrPn, 15).Font.Color = COR_UI_WARN_FG
                    ws.Cells(rrPn, 15).Font.Bold = True
                Case nAlPn >= 2
                    ws.Cells(rrPn, 15).Interior.Color = COR_UI_BAD_BG
                    ws.Cells(rrPn, 15).Font.Color = COR_UI_BAD_FG
                    ws.Cells(rrPn, 15).Font.Bold = True
            End Select
            Dim jcPn As Long
            For jcPn = 11 To 14
                If UCase$(Trim$(CStr(ws.Cells(rrPn, jcPn).Value))) = "S" Then
                    ws.Cells(rrPn, jcPn).Interior.Color = COR_UI_BAD_BG
                    ws.Cells(rrPn, jcPn).Font.Color = COR_UI_BAD_FG
                    ws.Cells(rrPn, jcPn).Font.Bold = True
                End If
            Next jcPn
        Next rrPn
    End If

End Sub

'==============================================================================
'  HELPERS DE CLASSES DE VIAGEM
'  Centraliza a lista de classes monitoradas na secao E dos alertas.
'==============================================================================
Private Function EhClasseViagem(ByVal cod As String) As Boolean
    If dClsViagem Is Nothing Then
        ' fallback: CONFIG nao carregada -> lista padrao
        EhClasseViagem = (Trim$(cod) = "8111290000" Or Trim$(cod) = "8210390000" _
                          Or Trim$(cod) = "8210550000")
    Else
        EhClasseViagem = dClsViagem.Exists(Trim$(cod))
    End If
End Function


Private Function DescClasseViagem(ByVal cod As String) As String
    Select Case Trim$(cod)
        Case "8111290000": DescClasseViagem = "Alimentacao Viagem"
        Case "8210390000": DescClasseViagem = "Servicos de Passagens"
        Case "8210550000": DescClasseViagem = "Servicos de Hospedagem"
        Case Else:         DescClasseViagem = ""
    End Select
End Function


'==============================================================================
'  ABA: SERVICO SEM MATERIAL
'  Detecta servicos (CA) que nao possuem material correspondente no mesmo PEP4.
'  Indica possivel servico lancado sem compra de material — risco regulatorio.
'  Consolida por PEP4NIVEL + CLS2 onde SRV > 0 e MAT = 0.
'==============================================================================
Private Sub Gerar_ServicoSemMaterial()
    ' Consolida por CODIGO de servico (sem PEP): lista os servicos lancados em
    ' familias que ficaram sem material correspondente, com codigo e descricao.
    Dim dMatSM As Object: Set dMatSM = CreateObject("Scripting.Dictionary")  ' PEP|familia -> MAT
    Dim dQ As Object:     Set dQ     = CreateObject("Scripting.Dictionary")  ' pep4|cod -> QTD
    Dim dVv As Object:    Set dVv    = CreateObject("Scripting.Dictionary")  ' pep4|cod -> VALOR
    Dim dLanc As Object:  Set dLanc  = CreateObject("Scripting.Dictionary")  ' pep4|cod -> n lancamentos
    Dim dFam As Object:   Set dFam   = CreateObject("Scripting.Dictionary")  ' pep4|cod -> familia
    Dim dPep4K As Object: Set dPep4K = CreateObject("Scripting.Dictionary")  ' pep4|cod -> pep4
    Dim dLinSM As Object: Set dLinSM = CreateObject("Scripting.Dictionary")  ' pep4|cod -> 1a linha (p/ TEXTO BREVE)

    Dim i As Long, pepSM As String, cls2SM As String, kSM As String, cod As String

    ' 1a passada: total de material por PEP|familia
    For i = 1 To UBound(dados, 1)
        pepSM = Trim$(CStr(dados(i, cPEP))): If pepSM = "" Then GoTo P1
        If EhMaterial(CStr(dados(i, cClassif))) Then
            cls2SM = FamiliaAlias(MatInfoLinha(i, 2))
            If cls2SM = "" Then cls2SM = "(SEM CLS2)"
            kSM = pepSM & "|" & cls2SM
            dMatSM(kSM) = dMatSM(kSM) + Abs(ToNum(dados(i, cQtd)))
        End If
P1:
    Next i

    ' 2a passada: servicos sem material -> agrega por pep4|codigo
    Dim kPepCod As String
    For i = 1 To UBound(dados, 1)
        pepSM = Trim$(CStr(dados(i, cPEP))): If pepSM = "" Then GoTo P2
        If EhMaterial(CStr(dados(i, cClassif))) Then GoTo P2
        cls2SM = FamiliaAlias(SrvInfoLinha(i, 1))
        If cls2SM = "" Then cls2SM = "(SEM CLS2)"
        kSM = pepSM & "|" & cls2SM
        If dMatSM.Exists(kSM) Then If dMatSM(kSM) > 0 Then GoTo P2  ' tem material -> ignora
        cod = NormCod(dados(i, cMaterial))
        If cod = "" Or cod = "0" Then GoTo P2
        kPepCod = pepSM & "|" & cod
        dQ(kPepCod) = dQ(kPepCod) + Abs(ToNum(dados(i, cQtd)) * ComboFator(dados(i, cMaterial)))
        dVv(kPepCod) = dVv(kPepCod) + ToNum(dados(i, cValor))
        dLanc(kPepCod) = dLanc(kPepCod) + 1
        If Not dFam.Exists(kPepCod) Then dFam(kPepCod) = cls2SM
        If Not dPep4K.Exists(kPepCod) Then dPep4K(kPepCod) = pepSM
        If Not dLinSM.Exists(kPepCod) Then dLinSM(kPepCod) = i
P2:
    Next i

    Dim nSM As Long, kv As Variant
    For Each kv In dQ.Keys
        If CDbl(dQ(kv)) > 0 Then nSM = nSM + 1
    Next kv

    Dim outpSM() As Variant: ReDim outpSM(0 To nSM, 1 To 10)
    outpSM(0, 1) = "PEP3NIVEL": outpSM(0, 2) = "PEP4NIVEL"
    outpSM(0, 3) = "COD_SERVICO": outpSM(0, 4) = "DESCRICAO_SERVICO"
    outpSM(0, 5) = "FAMILIA_CLS2": outpSM(0, 6) = "TIPO_FAMILIA"
    outpSM(0, 7) = "QTD_SERVICO": outpSM(0, 8) = "VALOR_SERVICO"
    outpSM(0, 9) = "QTD_LANCAMENTOS": outpSM(0, 10) = "RISCO"

    Dim rrSM As Long: rrSM = 0
    For Each kv In dQ.Keys
        If CDbl(dQ(kv)) > 0 Then
            rrSM = rrSM + 1
            Dim kvStr As String: kvStr = CStr(kv)
            Dim pep4v As String: pep4v = CStr(dPep4K(kvStr))
            ' extrai codigo: tudo apos o ultimo "|"
            Dim sepPos As Long: sepPos = InStrRev(kvStr, "|")
            cod = Mid$(kvStr, sepPos + 1)
            Dim famX As String: famX = CStr(dFam(kvStr))
            Dim tipoSM As String: tipoSM = TipoDaClassif(famX)
            If tipoSM = "" Then tipoSM = "SERV"
            Dim riscoSM As String
            If EhServicoPuro(famX) Then
                riscoSM = "N/A (SERVICO PURO)"   ' familia sem material esperado
            ElseIf tipoSM = "UC" Then
                riscoSM = "ALTO"
            ElseIf tipoSM = "COM" Then
                riscoSM = "MEDIO"
            Else
                riscoSM = "BAIXO"
            End If
            outpSM(rrSM, 1) = PEP3(pep4v)
            outpSM(rrSM, 2) = pep4v
            outpSM(rrSM, 3) = cod
            ' Descricao: catalogo; fallback TEXTO BREVE do proprio lancamento
            Dim descSM As String: descSM = DescServico(cod)
            If descSM = "" And dLinSM.Exists(kvStr) Then _
                descSM = Trim$(TextoCampo(CLng(dLinSM(kvStr)), cTexto))
            outpSM(rrSM, 4) = descSM
            outpSM(rrSM, 5) = famX
            outpSM(rrSM, 6) = tipoSM
            outpSM(rrSM, 7) = Round(CDbl(dQ(kvStr)), 2)
            outpSM(rrSM, 8) = Round(CDbl(dVv(kvStr)), 2)
            outpSM(rrSM, 9) = dLanc(kvStr)
            outpSM(rrSM, 10) = riscoSM
        End If
    Next kv
    EscreverAba "SERVICO SEM MATERIAL", outpSM
End Sub

'==============================================================================
'  ABA: PORTFOLIO OBRA
'  Uma linha por PEP3NIVEL com composicao financeira completa:
'  valor UC, COM, UAR, CA por categoria, % de cada um, status de aderencia
'  e indicadores de risco consolidados. Visao de portfólio para priorizacao.
'==============================================================================
Private Sub Gerar_PortfolioObra()
    Dim dPO As Object: Set dPO = CreateObject("Scripting.Dictionary") ' ordem
    Dim dEmpPO As Object: Set dEmpPO = CreateObject("Scripting.Dictionary")
    Dim dVTotPO As Object: Set dVTotPO = CreateObject("Scripting.Dictionary")
    Dim dVUC As Object: Set dVUC = CreateObject("Scripting.Dictionary")
    Dim dVCOM As Object: Set dVCOM = CreateObject("Scripting.Dictionary")
    Dim dVUAR As Object: Set dVUAR = CreateObject("Scripting.Dictionary")
    Dim dVCA As Object: Set dVCA = CreateObject("Scripting.Dictionary")
    Dim dVMOP As Object: Set dVMOP = CreateObject("Scripting.Dictionary")
    Dim dVATV As Object: Set dVATV = CreateObject("Scripting.Dictionary")
    Dim dVMAO As Object: Set dVMAO = CreateObject("Scripting.Dictionary")
    Dim dNLanc As Object: Set dNLanc = CreateObject("Scripting.Dictionary")
    Dim dNPEP4 As Object: Set dNPEP4 = CreateObject("Scripting.Dictionary") ' PEP4 distintos
    Dim dReprov As Object: Set dReprov = CreateObject("Scripting.Dictionary") ' PEP3 reprovados

    Dim i As Long, pepPO As String, p3PO As String, valPO As Double
    Dim cls3PO As String, catPO As String

    For i = 1 To UBound(dados, 1)
        pepPO = Trim$(CStr(dados(i, cPEP))): If pepPO = "" Then GoTo ProxPO
        p3PO = PEP3(pepPO)
        valPO = ToNum(dados(i, cValor))

        If Not dPO.Exists(p3PO) Then
            dPO(p3PO) = dPO.Count
            dEmpPO(p3PO) = TextoCampo(i, cEmpresa)
        End If
        If Not dNPEP4.Exists(p3PO & "|" & pepPO) Then dNPEP4(p3PO & "|" & pepPO) = 1

        dVTotPO(p3PO) = dVTotPO(p3PO) + valPO
        dNLanc(p3PO) = dNLanc(p3PO) + 1

        If EhMaterial(CStr(dados(i, cClassif))) Then
            Dim tipoPO As String: tipoPO = UCase$(Trim$(CStr(dados(i, cClassif))))
            Select Case tipoPO
                Case "UC":  dVUC(p3PO)  = dVUC(p3PO)  + valPO
                Case "COM": dVCOM(p3PO) = dVCOM(p3PO) + valPO
                Case "UAR": dVUAR(p3PO) = dVUAR(p3PO) + valPO
            End Select
        Else
            dVCA(p3PO) = dVCA(p3PO) + valPO
            catPO = UCase$(CategoriaAnaliseCA(i))
            If catPO = "MOP" Then dVMOP(p3PO) = dVMOP(p3PO) + valPO
            If catPO = "ATIVACAO DIRETA" Then dVATV(p3PO) = dVATV(p3PO) + valPO
            If catPO = "MAO DE OBRA" Then dVMAO(p3PO) = dVMAO(p3PO) + valPO
        End If
ProxPO:
    Next i

    ' Conta PEP4 distintos por PEP3
    Dim dCntPEP4 As Object: Set dCntPEP4 = CreateObject("Scripting.Dictionary")
    Dim kvPO As Variant
    For Each kvPO In dNPEP4.Keys
        Dim pvPO() As String: pvPO = Split(CStr(kvPO), "|")
        Dim p3cPO As String: p3cPO = CStr(pvPO(0))
        dCntPEP4(p3cPO) = dCntPEP4(p3cPO) + 1
    Next kvPO

    Dim nPO As Long: nPO = dPO.Count
    Dim outpPO() As Variant: ReDim outpPO(0 To nPO, 1 To 20)
    outpPO(0, 1)  = "EMPRESA":       outpPO(0, 2)  = "PEP3NIVEL"
    outpPO(0, 3)  = "QTD_PEP4":      outpPO(0, 4)  = "QTD_LANCAMENTOS"
    outpPO(0, 5)  = "VALOR_TOTAL":   outpPO(0, 6)  = "VALOR_UC"
    outpPO(0, 7)  = "VALOR_COM":     outpPO(0, 8)  = "VALOR_UAR"
    outpPO(0, 9)  = "VALOR_CA":      outpPO(0, 10) = "VALOR_MOP"
    outpPO(0, 11) = "VALOR_ATV_DRT": outpPO(0, 12) = "VALOR_MAO_OBRA"
    outpPO(0, 13) = "PERC_UC_%":     outpPO(0, 14) = "PERC_COM_%"
    outpPO(0, 15) = "PERC_CA_%":     outpPO(0, 16) = "PERC_MOP_CA_%"
    outpPO(0, 17) = "PERC_ATV_MAO_%":outpPO(0, 18) = "ATV_PREVISTA"
    outpPO(0, 19) = "DIF_ATV":       outpPO(0, 20) = "CLASSIFICACAO_OBRA"

    Dim pepsPO As Variant: pepsPO = dPO.Keys
    Dim rPO As Long
    For rPO = 0 To nPO - 1
        p3PO = CStr(pepsPO(rPO))
        Dim vtPO As Double: vtPO = IIf(dVTotPO.Exists(p3PO), CDbl(dVTotPO(p3PO)), 0)
        Dim vucPO As Double: vucPO = IIf(dVUC.Exists(p3PO), CDbl(dVUC(p3PO)), 0)
        Dim vcomPO As Double: vcomPO = IIf(dVCOM.Exists(p3PO), CDbl(dVCOM(p3PO)), 0)
        Dim vuarPO As Double: vuarPO = IIf(dVUAR.Exists(p3PO), CDbl(dVUAR(p3PO)), 0)
        Dim vcaPO As Double: vcaPO = IIf(dVCA.Exists(p3PO), CDbl(dVCA(p3PO)), 0)
        Dim vmopPO As Double: vmopPO = IIf(dVMOP.Exists(p3PO), CDbl(dVMOP(p3PO)), 0)
        Dim vatvPO As Double: vatvPO = IIf(dVATV.Exists(p3PO), CDbl(dVATV(p3PO)), 0)
        Dim vmaoPO As Double: vmaoPO = IIf(dVMAO.Exists(p3PO), CDbl(dVMAO(p3PO)), 0)
        ' FASE 1.1 [BUG corrigido]: antes era fixo vmaoPO * 0.25;
        ' agora respeita a regra EME/EMM = 8% (mesma da ANALISE DE CA).
        Dim atvPrevPO As Double: atvPrevPO = AtvPrevista(p3PO, vmaoPO)
        Dim difAtvPO As Double: difAtvPO = Round(vatvPO - atvPrevPO, 2)
        Dim vcaSemMopPO As Double: vcaSemMopPO = vcaPO - vmopPO

        ' Classificacao da obra pelo perfil de custo
        Dim classPO As String
        If vtPO = 0 Then
            classPO = "SEM VALOR"
        ElseIf vucPO / vtPO >= 0.5 Then
            classPO = "OBRA UC INTENSIVA"
        ElseIf vcaPO / vtPO >= 0.6 Then
            classPO = "CA INTENSIVA"
        ElseIf vmopPO / IIf(vcaSemMopPO <> 0, vcaSemMopPO, 1) > 0.08 Then
            classPO = "MOP ELEVADO"
        Else
            classPO = "BALANCEADA"
        End If

        outpPO(rPO + 1, 1)  = IIf(dEmpPO.Exists(p3PO), dEmpPO(p3PO), "")
        outpPO(rPO + 1, 2)  = p3PO
        outpPO(rPO + 1, 3)  = IIf(dCntPEP4.Exists(p3PO), dCntPEP4(p3PO), 0)
        outpPO(rPO + 1, 4)  = IIf(dNLanc.Exists(p3PO), dNLanc(p3PO), 0)
        outpPO(rPO + 1, 5)  = Round(vtPO, 2)
        outpPO(rPO + 1, 6)  = Round(vucPO, 2)
        outpPO(rPO + 1, 7)  = Round(vcomPO, 2)
        outpPO(rPO + 1, 8)  = Round(vuarPO, 2)
        outpPO(rPO + 1, 9)  = Round(vcaPO, 2)
        outpPO(rPO + 1, 10) = Round(vmopPO, 2)
        outpPO(rPO + 1, 11) = Round(vatvPO, 2)
        outpPO(rPO + 1, 12) = Round(vmaoPO, 2)
        If vtPO <> 0 Then
            outpPO(rPO + 1, 13) = Round(vucPO  / vtPO * 100, 1)
            outpPO(rPO + 1, 14) = Round(vcomPO / vtPO * 100, 1)
            outpPO(rPO + 1, 15) = Round(vcaPO  / vtPO * 100, 1)
        End If
        If vcaSemMopPO <> 0 Then outpPO(rPO + 1, 16) = Round(vmopPO / vcaSemMopPO * 100, 1)
        If vmaoPO <> 0 Then outpPO(rPO + 1, 17) = Round(vatvPO / vmaoPO * 100, 1)
        outpPO(rPO + 1, 18) = atvPrevPO
        outpPO(rPO + 1, 19) = difAtvPO
        outpPO(rPO + 1, 20) = classPO
    Next rPO
    EscreverAba "PORTFOLIO OBRA", outpPO

    ' --- Realce semantico pos-formatacao (CLASSIFICACAO_OBRA e DIF_ATV) -------
    Dim wsPO As Worksheet
    On Error Resume Next
    Set wsPO = ActiveWorkbook.Worksheets("PORTFOLIO OBRA")
    On Error GoTo 0
    If Not wsPO Is Nothing And nPO > 0 Then
        Dim rp As Long, clsTxt As String, bgc As Long, fgc As Long, difc As Double
        For rp = 2 To nPO + 1
            ' col 20 = CLASSIFICACAO_OBRA -> badge por perfil
            clsTxt = UCase$(Trim$(CStr(wsPO.Cells(rp, 20).Value)))
            Select Case True
                Case InStr(clsTxt, "UC INTENSIVA") > 0
                    bgc = RGB(198, 239, 206): fgc = RGB(0, 97, 0)       ' verde
                Case InStr(clsTxt, "BALANCEADA") > 0
                    bgc = RGB(214, 236, 238): fgc = RGB(11, 107, 115)   ' teal
                Case InStr(clsTxt, "CA INTENSIVA") > 0
                    bgc = RGB(234, 224, 244): fgc = RGB(96, 50, 150)    ' roxo
                Case InStr(clsTxt, "MOP ELEVADO") > 0
                    bgc = RGB(253, 233, 217): fgc = RGB(180, 95, 6)     ' laranja (atencao)
                Case InStr(clsTxt, "SEM VALOR") > 0
                    bgc = RGB(230, 232, 235): fgc = RGB(89, 89, 89)     ' cinza
                Case Else
                    bgc = RGB(255, 242, 204): fgc = RGB(140, 90, 0)
            End Select
            With wsPO.Cells(rp, 20)
                .Interior.Color = bgc: .Font.Color = fgc
                .Font.Bold = True: .HorizontalAlignment = xlCenter
            End With
            ' col 19 = DIF_ATV -> vermelho (fora do previsto) / verde (ok ~0)
            difc = ToNum(wsPO.Cells(rp, 19).Value)
            If Abs(difc) > 0.005 Then
                wsPO.Cells(rp, 19).Font.Color = RGB(156, 0, 6)
                wsPO.Cells(rp, 19).Font.Bold = True
            Else
                wsPO.Cells(rp, 19).Font.Color = RGB(0, 97, 0)
            End If
        Next rp
    End If
End Sub


'==============================================================================
'  ABA: NAO CLASSIFICADOS
'  Consolida lançamentos cuja familia/CLS2 = "CLASSIFICAR" ou vazia.
'  Mostra o valor financeiro em risco de nao-classificacao por PEP e familia.
'  Util para priorizar saneamento do catalogo de materiais/servicos.
'==============================================================================
Private Sub Gerar_NaoClassificados()
    Dim dNC As Object: Set dNC = CreateObject("Scripting.Dictionary") ' chave PEP|CLS2 -> valor
    Dim dNCQtd As Object: Set dNCQtd = CreateObject("Scripting.Dictionary")
    Dim dNCFrst As Object: Set dNCFrst = CreateObject("Scripting.Dictionary")
    Dim dNCTipo As Object: Set dNCTipo = CreateObject("Scripting.Dictionary") ' MAT ou SRV

    Dim i As Long, pepNC As String, cls2NC As String, kNC As String
    For i = 1 To UBound(dados, 1)
        pepNC = Trim$(CStr(dados(i, cPEP))): If pepNC = "" Then GoTo ProxNC
        Dim ehMatNC As Boolean: ehMatNC = EhMaterial(CStr(dados(i, cClassif)))
        If ehMatNC Then
            cls2NC = MatInfoLinha(i, 2)
        Else
            cls2NC = SrvInfoLinha(i, 1)
        End If
        ' So traz se nao classificado
        Dim cls2NormNC As String: cls2NormNC = UCase$(SemAcento(Trim$(cls2NC)))
        If cls2NormNC <> "" And cls2NormNC <> "CLASSIFICAR" And cls2NormNC <> "(SEM CLS2)" Then GoTo ProxNC

        kNC = pepNC & "|" & IIf(cls2NC = "", "(SEM CLS2)", cls2NC)
        dNC(kNC) = dNC(kNC) + ToNum(dados(i, cValor))
        dNCQtd(kNC) = dNCQtd(kNC) + 1
        If Not dNCFrst.Exists(kNC) Then
            dNCFrst(kNC) = i
            dNCTipo(kNC) = IIf(ehMatNC, "MATERIAL", "SERVICO")
        End If
ProxNC:
    Next i

    Dim nNC As Long: nNC = dNC.Count
    Dim outpNC() As Variant: ReDim outpNC(0 To nNC, 1 To 11)
    outpNC(0, 1) = "PEP4NIVEL":      outpNC(0, 2) = "PEP3NIVEL"
    outpNC(0, 3) = "TIPO_PEP":       outpNC(0, 4) = "TIPO_LANCAMENTO"
    outpNC(0, 5) = "CLASSE_CUSTO":   outpNC(0, 6) = "CLS2_ATUAL"
    outpNC(0, 7) = "COD_MATERIAL":   outpNC(0, 8) = "TEXTO_MATERIAL"
    outpNC(0, 9) = "QTD_LANCAMENTOS": outpNC(0, 10) = "VALOR_TOTAL"
    outpNC(0, 11) = "ACAO_RECOMENDADA"

    Dim rrNC As Long: rrNC = 0
    Dim kvNC As Variant
    For Each kvNC In dNC.Keys
        rrNC = rrNC + 1
        Dim pvNC() As String: pvNC = Split(CStr(kvNC), "|")
        Dim p4NC As String: p4NC = CStr(pvNC(0))
        Dim fiNC As Long: fiNC = dNCFrst(kvNC)
        Dim acaoNC As String
        If dNCTipo(kvNC) = "MATERIAL" Then
            acaoNC = "Cadastrar material no catalogo MATERIAS_ATUAIS com CLS1/CLS2/CLS3"
        Else
            acaoNC = "Cadastrar servico no catalogo SERVICOS_ATUAIS com CLS1/CLS2/CLS3"
        End If
        outpNC(rrNC, 1) = p4NC
        outpNC(rrNC, 2) = PEP3(p4NC)
        outpNC(rrNC, 3) = TipoPEPANEEL(p4NC)
        outpNC(rrNC, 4) = dNCTipo(kvNC)
        outpNC(rrNC, 5) = ValorCampo(fiNC, cClasse)           ' CLASSE_CUSTO (1a linha do grupo)
        outpNC(rrNC, 6) = IIf(UBound(pvNC) >= 1, CStr(pvNC(1)), "")
        outpNC(rrNC, 7) = NormCod(dados(fiNC, cMaterial))
        outpNC(rrNC, 8) = ValorCampo(fiNC, cTexto)
        outpNC(rrNC, 9) = dNCQtd(kvNC)
        outpNC(rrNC, 10) = Round(CDbl(dNC(kvNC)), 2)
        outpNC(rrNC, 11) = acaoNC
    Next kvNC
    EscreverAba "NAO CLASSIFICADOS", outpNC
End Sub


'==============================================================================
'  ABA: RACIONALIZACAO COM
'  Valida materiais COM contra gabaritos de proporcao NT.006 (13,8kV).
'  Para cada PEP4 + material COM, verifica se a quantidade esta dentro da
'  faixa esperada em relacao a sua ancora (ex.: 2-3 isoladores por cruzeta).
'
'  Status possiveis:
'    ANCORA          — material de referencia (ex.: CRUZETA)
'    OK              — dentro da faixa NT.006 (±10%)
'    EXCESSO         — acima da faixa maxima
'    INSUFICIENTE    — abaixo da faixa minima
'    SEM ANCORA      — ancora nao lancada no mesmo PEP
'    ESTORNO SEM ENTRADA — qtd liquida negativa
'    QTD ZERO        — qtd liquida = 0
'    SEM REFERENCIA  — material nao mapeado no NT.006
'==============================================================================
Private Sub Gerar_RacionalizacaoCOM()

    ' --- Carrega mapa NT.006 ---
    Dim nt006RC As Object: Set nt006RC = CriarMapaNT006_RC()

    ' --- Consolida COMs por PEP4 + COD_MATERIAL (qtd liquida) ---
    Dim dQRC As Object: Set dQRC = CreateObject("Scripting.Dictionary")
    Dim dInfoRC As Object: Set dInfoRC = CreateObject("Scripting.Dictionary")

    Dim i As Long, pepRC As String, matRC As String, kRC As String, qRC As Double
    For i = 1 To UBound(dados, 1)
        pepRC = Trim$(CStr(dados(i, cPEP))): If pepRC = "" Then GoTo ProxRC
        If UCase$(Trim$(CStr(dados(i, cClassif)))) <> "COM" Then GoTo ProxRC
        matRC = NormCod(dados(i, cMaterial))
        If matRC = "" Or matRC = "0" Then GoTo ProxRC
        qRC = ToNum(dados(i, cQtd))
        kRC = pepRC & "|" & matRC
        dQRC(kRC) = dQRC(kRC) + qRC
        If Not dInfoRC.Exists(kRC) Then
            dInfoRC(kRC) = Array(pepRC, ValorCampo(i, cTexto), ValorCampo(i, cUML))
        End If
ProxRC:
    Next i

    ' --- Calcula ancoras por PEP4 + familia ---
    Dim dAncRC As Object: Set dAncRC = CreateObject("Scripting.Dictionary")
    Dim kvRC As Variant
    For Each kvRC In dQRC.Keys
        Dim pvRC() As String: pvRC = Split(CStr(kvRC), "|")
        Dim p4RC As String: p4RC = CStr(pvRC(0))
        Dim mRC As String: mRC = CStr(pvRC(1))
        If nt006RC.Exists(mRC) Then
            Dim tmRC As Object: Set tmRC = nt006RC(mRC)
            If CBool(tmRC("EhAncora")) Then
                Dim aKRC As String: aKRC = p4RC & "|" & CStr(tmRC("Familia"))
                dAncRC(aKRC) = dAncRC(aKRC) + CDbl(dQRC(kvRC))
            End If
        End If
    Next kvRC

    ' --- Monta output (sem pre-ordenacao; OrdenarAba cuida do sort) ----------
    Dim nRC As Long: nRC = dQRC.Count
    Dim outpRC() As Variant: ReDim outpRC(0 To nRC, 1 To 15)
    outpRC(0, 1)  = "PEP4NIVEL":         outpRC(0, 2)  = "PEP3NIVEL"
    outpRC(0, 3)  = "COD_MATERIAL":      outpRC(0, 4)  = "DESCRICAO_MATERIAL"
    outpRC(0, 5)  = "UML":               outpRC(0, 6)  = "QTD_LIQUIDA"
    outpRC(0, 7)  = "FAMILIA_NT006":     outpRC(0, 8)  = "COD_NT006"
    outpRC(0, 9)  = "DESCRICAO_NT006":   outpRC(0, 10) = "ANCORA_FAMILIA"
    outpRC(0, 11) = "QTD_ANCORA":        outpRC(0, 12) = "FAIXA_MIN"
    outpRC(0, 13) = "FAIXA_MAX":         outpRC(0, 14) = "REGRA_NT006"
    outpRC(0, 15) = "STATUS_VALIDACAO"

    Dim rrRC As Long: rrRC = 0
    For Each kvRC In dQRC.Keys
        kRC = CStr(kvRC)
        pvRC = Split(kRC, "|")
        p4RC = CStr(pvRC(0)): mRC = CStr(pvRC(1))
        Dim qtdLRC As Double: qtdLRC = CDbl(dQRC(kRC))
        Dim infRC As Variant: infRC = dInfoRC(kRC)

        Dim famRC As String: famRC = "SEM REFERENCIA"
        Dim codNTRC As String: codNTRC = "-"
        Dim descrNTRC As String: descrNTRC = ""
        Dim ancFamRC As String: ancFamRC = ""
        Dim qtdAncRC As Double: qtdAncRC = 0
        Dim fMinRC As Double: fMinRC = 0
        Dim fMaxRC As Double: fMaxRC = 0
        Dim regraRC As String: regraRC = ""
        Dim statusRC As String: statusRC = "SEM REFERENCIA"
        Dim ehAncRC As Boolean: ehAncRC = False

        If nt006RC.Exists(mRC) Then
            Dim tmRCv As Object: Set tmRCv = nt006RC(mRC)
            famRC    = CStr(tmRCv("Familia"))
            codNTRC  = CStr(tmRCv("CodNT006"))
            descrNTRC = CStr(tmRCv("DescrNT006"))
            ehAncRC  = CBool(tmRCv("EhAncora"))
            regraRC  = CStr(tmRCv("DescrRegra"))
            ancFamRC = CStr(tmRCv("AncoraDep"))

            If ehAncRC Then
                If qtdLRC < 0 Then
                    statusRC = "ESTORNO SEM ENTRADA"
                ElseIf qtdLRC = 0 Then
                    statusRC = "QTD ZERO"
                Else
                    statusRC = "ANCORA"
                End If
            Else
                Dim ancKRC As String: ancKRC = p4RC & "|" & ancFamRC
                If dAncRC.Exists(ancKRC) Then
                    qtdAncRC = CDbl(dAncRC(ancKRC))
                    fMinRC = qtdAncRC * CDbl(tmRCv("RazaoMin")) * 0.9
                    fMaxRC = qtdAncRC * CDbl(tmRCv("RazaoMax")) * 1.1
                    If qtdLRC < 0 Then
                        statusRC = "ESTORNO SEM ENTRADA"
                    ElseIf qtdAncRC <= 0 Then
                        statusRC = "SEM ANCORA"
                    ElseIf qtdLRC < fMinRC Then
                        statusRC = "INSUFICIENTE"
                    ElseIf qtdLRC > fMaxRC Then
                        statusRC = "EXCESSO"
                    Else
                        statusRC = "OK"
                    End If
                Else
                    statusRC = "SEM ANCORA"
                    ancFamRC = ancFamRC & " (nao lancada)"
                End If
            End If
        End If

        rrRC = rrRC + 1
        outpRC(rrRC, 1)  = p4RC
        outpRC(rrRC, 2)  = PEP3(p4RC)
        outpRC(rrRC, 3)  = mRC
        outpRC(rrRC, 4)  = CStr(infRC(1))
        outpRC(rrRC, 5)  = CStr(infRC(2))
        outpRC(rrRC, 6)  = Round(qtdLRC, 3)
        outpRC(rrRC, 7)  = famRC
        outpRC(rrRC, 8)  = codNTRC
        outpRC(rrRC, 9)  = descrNTRC
        outpRC(rrRC, 10) = ancFamRC
        outpRC(rrRC, 11) = IIf(qtdAncRC > 0, Round(qtdAncRC, 3), "")
        outpRC(rrRC, 12) = IIf(fMinRC > 0, Round(fMinRC, 2), "")
        outpRC(rrRC, 13) = IIf(fMaxRC > 0, Round(fMaxRC, 2), "")
        outpRC(rrRC, 14) = regraRC
        outpRC(rrRC, 15) = statusRC
    Next kvRC

    EscreverAba "RACIONALIZACAO COM", outpRC

    ' --- Colore coluna STATUS apos escrita (FASE 3.1: em blocos) -------------
    ' Le a coluna ja ORDENADA da planilha e agrupa statuses contiguos.
    Dim wsRC As Worksheet: Set wsRC = ActiveWorkbook.Worksheets("RACIONALIZACAO COM")
    If rrRC >= 1 Then
        Dim ultRC As Long: ultRC = rrRC + 1
        Dim arrSt As Variant
        If ultRC = 2 Then
            ReDim arrSt(1 To 1, 1 To 1)
            arrSt(1, 1) = wsRC.Cells(2, 15).Value
        Else
            arrSt = wsRC.Range(wsRC.Cells(2, 15), wsRC.Cells(ultRC, 15)).Value
        End If
        Dim stAnt As String, stCur As String, iniSt As Long, rr2 As Long
        stAnt = Chr$(1): iniSt = 1
        For rr2 = 1 To UBound(arrSt, 1)
            stCur = UCase$(Trim$(CStr(arrSt(rr2, 1))))
            If stCur <> stAnt Then
                If rr2 > 1 Then PintarStatusRC wsRC, iniSt + 1, rr2, stAnt
                stAnt = stCur: iniSt = rr2
            End If
        Next rr2
        PintarStatusRC wsRC, iniSt + 1, UBound(arrSt, 1) + 1, stAnt
    End If
End Sub


'==============================================================================
'  MAPA NT.006 — retorna Dictionary: codSAP -> Dictionary com campos do material
'  Usa Dictionary aninhado em vez de Type para compatibilidade com o modulo.
'==============================================================================
Private Function CriarMapaNT006_RC() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")

    ' Helper interno
    ' AddMatRC d, cod, familia, nt006, descr, ehAncora, ancDep, rMin, rMax, regra

    ' === CRUZETAS (ancora) ===
    AddMatRC d, "133100007", "CRUZETA", "R-02", "Cruzeta concreto T 1900mm",    True,  "",             0,   0,   ""
    AddMatRC d, "133100001", "CRUZETA", "R-02", "Cruzeta concreto L 1700mm",    True,  "",             0,   0,   ""
    AddMatRC d, "133100006", "CRUZETA", "R-02", "Cruzeta concreto T 2200mm",    True,  "",             0,   0,   ""
    AddMatRC d, "133400012", "CRUZETA", "R-02", "Cruzeta PRFV 90x112,5 2,4m",  True,  "",             0,   0,   ""
    AddMatRC d, "133400003", "CRUZETA", "R-02", "Cruzeta PRFV",                 True,  "",             0,   0,   ""
    AddMatRC d, "133400004", "CRUZETA", "R-02", "Cruzeta PRFV",                 True,  "",             0,   0,   ""

    ' === ISOLADOR PILAR: 2-3 por cruzeta ===
    AddMatRC d, "123140003", "ISOLADOR PILAR", "I-05", "Isolador pilar 15kV M16",    False, "CRUZETA", 2, 3.5, "2-3 iso. pilar por cruzeta (N1=3)"
    AddMatRC d, "123140016", "ISOLADOR PILAR", "I-05", "Isolador pilar 24,2kV M16",  False, "CRUZETA", 2, 3.5, "2-3 iso. pilar por cruzeta"
    AddMatRC d, "123140015", "ISOLADOR PILAR", "I-05", "Isolador pilar polim. 25kV", False, "CRUZETA", 2, 3.5, "2-3 iso. pilar por cruzeta"
    AddMatRC d, "123140014", "ISOLADOR PILAR", "I-05", "Isolador pilar",             False, "CRUZETA", 2, 3.5, "2-3 iso. pilar por cruzeta"

    ' === ISOLADOR SUSPENSAO (ancora) ===
    AddMatRC d, "123230001", "ISOL SUSPENSAO", "I-06", "Isolador suspensao polim. 15kV", True, "", 0, 0, ""
    AddMatRC d, "123230002", "ISOL SUSPENSAO", "I-06", "Isolador suspensao",             True, "", 0, 0, ""

    ' === ARRUELA: 2-8 por cruzeta ===
    AddMatRC d, "134830013", "ARRUELA", "A-02", "Arruela quad. 38x38x3mm F18", False, "CRUZETA", 2, 9, "2-8 arruelas por cruzeta"
    AddMatRC d, "134830014", "ARRUELA", "A-02", "Arruela quad. lis 18x50x3mm", False, "CRUZETA", 2, 9, "2-8 arruelas por cruzeta"
    AddMatRC d, "134830051", "ARRUELA", "A-02", "Arruela red pres M18",        False, "CRUZETA", 2, 9, "2-8 arruelas por cruzeta"

    ' === PARAFUSO: 1-8 por cruzeta ===
    AddMatRC d, "134700040", "PARAFUSO", "F-30", "Parafuso cab qd 125mm M16x2", False, "CRUZETA", 1, 8, "1-8 parafusos por cruzeta"
    AddMatRC d, "134700043", "PARAFUSO", "F-30", "Parafuso cab qd 200mm M16x2", False, "CRUZETA", 1, 8, "1-8 parafusos por cruzeta"
    AddMatRC d, "134700046", "PARAFUSO", "F-30", "Parafuso cab qd 250mm M16x2", False, "CRUZETA", 1, 8, "1-8 parafusos por cruzeta"
    AddMatRC d, "134700047", "PARAFUSO", "F-30", "Parafuso cab qd 300mm M16x2", False, "CRUZETA", 1, 8, "1-8 parafusos por cruzeta"
    AddMatRC d, "134700049", "PARAFUSO", "F-30", "Parafuso cab qd 400mm M16x2", False, "CRUZETA", 1, 8, "1-8 parafusos por cruzeta"
    AddMatRC d, "134700028", "PARAFUSO", "F-30", "Parafuso cab abaul 16x45mm",  False, "CRUZETA", 1, 8, "1-8 parafusos por cruzeta"
    AddMatRC d, "134700030", "PARAFUSO", "F-30", "Parafuso cab abaul 16x150mm", False, "CRUZETA", 1, 8, "1-8 parafusos por cruzeta"
    AddMatRC d, "134700082", "PARAFUSO", "F-30", "Parafuso rosca dupla 16x500", False, "CRUZETA", 1, 8, "1-8 parafusos por cruzeta"

    ' === PINO: 2-3 por cruzeta ===
    AddMatRC d, "134280005", "PINO", "F-36", "Pino iso pilar autotrav M16x2", False, "CRUZETA", 2, 3.5, "~3 pinos por cruzeta (1 por isolador pilar)"
    AddMatRC d, "134280002", "PINO", "F-37", "Pino curto suporte topo",       False, "CRUZETA", 1, 2.5, "1-2 pinos curtos por cruzeta"

    ' === PORCA: 2-6 por cruzeta ===
    AddMatRC d, "134800002", "PORCA", "A-21", "Porca quad. M16x2", False, "CRUZETA", 2, 6.5, "2-6 porcas por cruzeta"

    ' === SELA DE CRUZETA: 2-4 por cruzeta ===
    AddMatRC d, "134380004", "SELA CRUZETA", "-", "Sela cruzeta 110x116mm", False, "CRUZETA", 2, 4, "2-3 selas por cruzeta trifasica"
    AddMatRC d, "134380005", "SELA CRUZETA", "-", "Sela cruzeta",           False, "CRUZETA", 2, 4, "2-3 selas por cruzeta"

    ' === MAO FRANCESA: 1-2 por cruzeta ===
    AddMatRC d, "134100006", "MAO FRANCESA", "-", "Mao francesa plana 726x38x5mm", False, "CRUZETA", 0.5, 2.5, "1-2 maos-francesas por cruzeta"

    ' === GANCHO OLHAL (ancora suspensao) ===
    AddMatRC d, "134250015", "GANCHO OLHAL", "F-13", "Gancho olhal 5000daN", True, "", 0, 0, ""

    ' === MANILHA / OLHAL: 1:1 com gancho ===
    AddMatRC d, "134200006", "MANILHA",        "F-22", "Manilha sapatilha 5000daN",  False, "GANCHO OLHAL", 0.8, 1.2, "1 manilha por ponto de suspensao"
    AddMatRC d, "134250023", "OLHAL PARAFUSO", "-",    "Olhal parafuso M16 5000daN", False, "GANCHO OLHAL", 0.8, 1.2, "1 olhal por ponto de suspensao"
    AddMatRC d, "134740023", "PARAFUSO OLHAL", "F-34", "Parafuso olhal M16x250mm",   False, "GANCHO OLHAL", 0.8, 1.2, "1 parafuso olhal por ponto de suspensao"

    ' === HASTE DE ATERRAMENTO (ancora AT) ===
    AddMatRC d, "134600010", "HASTE TERRA", "F-17", "Haste aco-cobreado 14,3mm 2,4m", True, "", 0, 0, ""
    AddMatRC d, "134600004", "HASTE TERRA", "F-17", "Haste aco-cobreado 12,7mm 2,4m", True, "", 0, 0, ""

    ' === CONECTOR HASTE: 1:1 com haste ===
    AddMatRC d, "124140026", "CONEC HASTE", "M-10", "Conector cunha haste 6-16mm",  False, "HASTE TERRA", 0.8, 1.2, "1 conector por haste de aterramento"
    AddMatRC d, "124140078", "CONEC HASTE", "M-10", "Conector aterramento p/haste", False, "HASTE TERRA", 0.8, 1.2, "1 conector por haste de aterramento"
    AddMatRC d, "124140011", "CONEC HASTE", "M-10", "Conector cunha haste",         False, "HASTE TERRA", 0.8, 1.2, "1 conector por haste de aterramento"

    ' === SUPORTE PARA-RAIOS (ancora PR) ===
    AddMatRC d, "134190064", "SUP PARA-RAIO", "F-47", "Suporte L para-raios 38x205", True, "", 0, 0, ""

    ' === PARA-RAIOS: 1:1 com suporte ===
    AddMatRC d, "104010001", "PARA-RAIO", "E-29", "Para-raios ZnO 12kV 10kA", False, "SUP PARA-RAIO", 0.8, 1.2, "1 para-raios por suporte (1:1)"
    AddMatRC d, "104010004", "PARA-RAIO", "E-29", "Para-raios ZnO 15kV",      False, "SUP PARA-RAIO", 0.8, 1.2, "1 para-raios por suporte (1:1)"

    ' === CHAVE FUSIVEL (ancora) ===
    AddMatRC d, "105300003", "CHAVE FUSIVEL", "E-09", "Chave fusivel 15kV 100A base C", True, "", 0, 0, ""

    ' === TRANSFORMADOR (ancora) ===
    AddMatRC d, "102100035", "TRAFO", "E-45", "Trafo trifasico 13,8kV 500kVA", True, "", 0, 0, ""
    AddMatRC d, "102100036", "TRAFO", "E-45", "Trafo trifasico 13,8kV",        True, "", 0, 0, ""
    AddMatRC d, "102100030", "TRAFO", "E-45", "Trafo monofasico",               True, "", 0, 0, ""

    ' === CONECTOR RAMAL (ancora informativo) ===
    AddMatRC d, "124010010", "CONEC RAMAL", "O-02", "Conector cunha CuEst tipo II",  True, "", 0, 0, ""
    AddMatRC d, "124010012", "CONEC RAMAL", "O-02", "Conector cunha CuEst tipo III", True, "", 0, 0, ""

    Set CriarMapaNT006_RC = d
End Function

Private Sub AddMatRC(d As Object, ByVal cod As String, ByVal familia As String, _
                     ByVal nt006 As String, ByVal descr As String, ByVal ehAnc As Boolean, _
                     ByVal ancDep As String, ByVal rMin As Double, ByVal rMax As Double, _
                     ByVal regra As String)
    If d.Exists(cod) Then Exit Sub   ' primeiro registro vence
    Dim m As Object: Set m = CreateObject("Scripting.Dictionary")
    m("Familia")    = familia
    m("CodNT006")   = nt006
    m("DescrNT006") = descr
    m("EhAncora")   = ehAnc
    m("AncoraDep")  = ancDep
    m("RazaoMin")   = rMin
    m("RazaoMax")   = rMax
    m("DescrRegra") = regra
    d.Add cod, m
End Sub


'==============================================================================
'  FUNCOES ADICIONADAS NAS FASES 1, 3 e 4 (refatoracao)
'==============================================================================

Private Function EhPepEmergencia(ByVal pep As String) As Boolean
    Dim u As String: u = UCase$(pep)
    EhPepEmergencia = (InStr(u, "EME") > 0 Or InStr(u, "EMM") > 0)
End Function

Private Function AtvPrevista(ByVal pep As String, ByVal valorMaoObra As Double) As Double
    Dim perc As Double
    If EhPepEmergencia(pep) Then
        perc = CfgNum("PERC_ATV_EME", 8) / 100
    Else
        perc = CfgNum("PERC_ATV_PADRAO", 25) / 100
    End If
    AtvPrevista = Round(valorMaoObra * perc, 2)
End Function

Private Function ColExata(ws As Worksheet, frags As Variant) As Long
    Dim ult As Long, j As Long, i As Long, hdr As String, fr As String
    ult = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For i = LBound(frags) To UBound(frags)
        fr = SemAcento(UCase$(CStr(frags(i))))
        For j = 1 To ult
            hdr = SemAcento(UCase$(Trim$(CStr(ws.Cells(1, j).Value))))
            If hdr = fr Then ColExata = j: Exit Function
        Next j
    Next i
    ColExata = 0
End Function

Private Sub AplicarFreeze(ws As Worksheet, ByVal celula As String, _
                          Optional ByVal congelar As Boolean = True)
    ' Best-effort (so formatacao visual): falha aqui nao altera a analise.
    On Error Resume Next
    Dim su As Boolean: su = Application.ScreenUpdating
    Application.ScreenUpdating = True
    ws.Activate
    ActiveWindow.FreezePanes = False
    If congelar Then
        ws.Range(celula).Select
        ActiveWindow.FreezePanes = True
    Else
        ws.Range(celula).Select
    End If
    Application.ScreenUpdating = su
    On Error GoTo 0
End Sub

Private Function CategoriaVeredito(ByVal v As String) As Long
    v = UCase$(Trim$(v))
    If v = "APROVADO" Or v = "ADERENTE" Or v = "OK" Then
        CategoriaVeredito = 1
    ElseIf InStr(v, "REPROVAD") > 0 Or v = "NAO ADERENTE" _
           Or InStr(v, "NEGATIV") > 0 Or InStr(v, "NAO CADASTR") > 0 _
           Or InStr(v, "ZERAD") > 0 Or InStr(v, "ABAIXO") > 0 _
           Or InStr(v, "ACIMA") > 0 Or InStr(v, "VERIFICAR") > 0 Then
        CategoriaVeredito = 2
    ElseIf v = "NULO" Or v = "SEM UC" Then
        CategoriaVeredito = 3
    Else
        CategoriaVeredito = 0
    End If
End Function

Private Sub ColorirColunaVeredito(ws As Worksheet, ByVal jc As Long, ByVal nR As Long)
    Dim arr As Variant
    If nR = 2 Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = ws.Cells(2, jc).Value
    Else
        arr = ws.Range(ws.Cells(2, jc), ws.Cells(nR, jc)).Value
    End If

    Dim rr As Long, cat As Long, catAnt As Long, ini As Long
    catAnt = -1
    For rr = 1 To UBound(arr, 1)
        cat = CategoriaVeredito(CStr(arr(rr, 1)))
        If cat <> catAnt Then
            ' fecha o bloco anterior (linhas de planilha: ini+1 ate rr)
            If catAnt > 0 Then PintarRunVeredito ws, jc, ini + 1, rr, catAnt
            catAnt = cat: ini = rr
        End If
    Next rr
    If catAnt > 0 Then PintarRunVeredito ws, jc, ini + 1, UBound(arr, 1) + 1, catAnt
End Sub

Private Sub PintarRunVeredito(ws As Worksheet, ByVal jc As Long, _
                              ByVal linIni As Long, ByVal linFim As Long, _
                              ByVal cat As Long)
    With ws.Range(ws.Cells(linIni, jc), ws.Cells(linFim, jc))
        .Font.Name = FONTE_UI
        Select Case cat
            Case 1
                .Interior.Color = COR_UI_OK_BG
                .Font.Color = COR_UI_OK_FG
                .Font.Bold = True
            Case 2
                .Interior.Color = COR_UI_BAD_BG
                .Font.Color = COR_UI_BAD_FG
                .Font.Bold = True
            Case 3
                .Interior.Color = COR_UI_NEU_BG
                .Font.Color = COR_UI_NEU_FG
        End Select
    End With
End Sub

Private Sub PintarStatusRC(ws As Worksheet, ByVal linIni As Long, _
                           ByVal linFim As Long, ByVal st As String)
    Dim bgRC As Long, fgRC As Long
    Select Case st
        Case "OK":                  bgRC = COR_UI_OK_BG:   fgRC = COR_UI_OK_FG
        Case "ANCORA":              bgRC = COR_UI_INFO_BG: fgRC = COR_UI_INFO_FG
        Case "EXCESSO":             bgRC = COR_UI_WARN_BG: fgRC = COR_UI_WARN_FG
        Case "INSUFICIENTE":        bgRC = COR_UI_BAD_BG:  fgRC = COR_UI_BAD_FG
        Case "ESTORNO SEM ENTRADA": bgRC = COR_UI_BAD_BG:  fgRC = COR_UI_BAD_FG
        Case "QTD ZERO":            bgRC = COR_UI_BAD_BG:  fgRC = COR_UI_BAD_FG
        Case "SEM ANCORA":          bgRC = COR_UI_WARN_BG: fgRC = COR_UI_NEU_FG
        Case Else:                  bgRC = COR_UI_NEU_BG:  fgRC = COR_UI_NEU_FG
    End Select
    With ws.Range(ws.Cells(linIni, 15), ws.Cells(linFim, 15))
        .Interior.Color = bgRC
        .Font.Color = fgRC
        .Font.Name = FONTE_UI
        .Font.Bold = True
    End With
End Sub

Private Sub GarantirConfig()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("CONFIG")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add( _
            After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        ws.Name = "CONFIG"
        ws.Cells(1, 1).Value = "CHAVE"
        ws.Cells(1, 2).Value = "VALOR"
        ws.Cells(1, 3).Value = "DESCRICAO"
        ws.Range("A1:C1").Font.Bold = True
    End If

    ' chaves padrao (chave, valor, descricao)
    Dim def As Variant
    def = Array( _
        Array("CAT_MATERIAIS", "%USERPROFILE%\Downloads\MATERIAS_ATUAIS_4.xlsx;%USERPROFILE%\Downloads\MATERIAS_ATUAIS (2).xlsx;%USERPROFILE%\Downloads\MATERIAS_ATUAIS.xlsx", "Caminhos do catalogo de materiais (separar alternativas por ;)"), _
        Array("CAT_SERVICOS", "%USERPROFILE%\Downloads\SERVICOS_ATUAIS_3.xlsx;%USERPROFILE%\Downloads\SERVICOS_ATUAIS_2.xlsx;%USERPROFILE%\Downloads\SERVICOS_ATUAIS.xlsx", "Caminho do catalogo de servicos"), _
        Array("CAT_CLASSE", "%USERPROFILE%\Downloads\CLASSE_CUSTO_ATUAIS_2.xlsx;%USERPROFILE%\Downloads\CLASSE_CUSTO_ATUAIS.xlsx", "Caminhos do catalogo de classe de custo"), _
        Array("CAT_CABO", "%USERPROFILE%\Downloads\CONVERSOES_CABO_ATUAIS.xlsx", "Caminho da conversao de cabo KG->m (opcional)"), _
        Array("CAT_COMBO", "%USERPROFILE%\Downloads\SRV_COMBO_ATUAIS.xlsx", "Caminho do catalogo SRV COMBO"), _
        Array("PERC_ATV_EME", "8", "Percentual ATV PREVISTA para PEP EME/EMM (%)"), _
        Array("PERC_ATV_PADRAO", "25", "Percentual ATV PREVISTA para os demais PEPs (%)"), _
        Array("PERC_MOP", "5.483", "Percentual do CALCULO DO MOP sobre o total sem MOP (%)"), _
        Array("MARGEM_ADERENCIA", "10", "Margem da aderencia MAT vs SRV (%)"), _
        Array("CLASSES_VIAGEM", "8111290000;8210390000;8210550000", "Classes de custo de viagem (Alimentacao;Passagem;Hospedagem)"), _
        Array("CLASSE_COMBUSTIVEL", "8119980000", "Classe de custo de combustiveis -> categoria OUTROS na ANALISE DE CA"), _
        Array("EQUIV_SRV_MAT", "", "De-para extra CLS2 servico=familia material (ex.: JUMPER=ALCA;X=Y)"), _
        Array("SRV_PURO", "", "Familias extra de servico sem material esperado (separadas por ;)"), _
        Array("FAM_UAR", "", "Familias UAR extra (equivalentes a UC; nao geram alerta SEM UC). Separadas por ;"))

    ' indexa chaves ja existentes
    Dim ult As Long: ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim dExist As Object: Set dExist = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To ult
        Dim kEx As String: kEx = UCase$(Trim$(CStr(ws.Cells(r, 1).Value)))
        If kEx <> "" And Not dExist.Exists(kEx) Then dExist(kEx) = 1
    Next r

    ' acrescenta apenas as ausentes
    Dim i As Long, lin As Long: lin = ult
    For i = LBound(def) To UBound(def)
        If Not dExist.Exists(UCase$(CStr(def(i)(0)))) Then
            lin = lin + 1
            ws.Cells(lin, 1).Value = def(i)(0)
            ws.Cells(lin, 2).Value = def(i)(1)
            ws.Cells(lin, 3).Value = def(i)(2)
        End If
    Next i

    ws.Columns("A:C").AutoFit
    ws.Visible = xlSheetHidden   ' oculta (reexibir: botao direito na guia > Reexibir)
End Sub

Private Sub CarregarConfig()
    Set dCfg = CreateObject("Scripting.Dictionary")
    Set dClsViagem = CreateObject("Scripting.Dictionary")

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("CONFIG")
    On Error GoTo 0
    If Not ws Is Nothing Then
        Dim ult As Long: ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim r As Long, k As String
        For r = 2 To ult
            k = UCase$(Trim$(CStr(ws.Cells(r, 1).Value)))
            If k <> "" And Not dCfg.Exists(k) Then dCfg(k) = Trim$(CStr(ws.Cells(r, 2).Value))
        Next r
    End If

    ' cache das classes de viagem (parse uma vez; EhClasseViagem roda por linha)
    Dim lista As String: lista = CfgTxt("CLASSES_VIAGEM", "8111290000;8210390000;8210550000")
    Dim parts() As String, p As Long
    parts = Split(lista, ";")
    For p = LBound(parts) To UBound(parts)
        Dim cod As String: cod = Trim$(parts(p))
        If cod <> "" And Not dClsViagem.Exists(cod) Then dClsViagem(cod) = 1
    Next p
End Sub

Private Function CfgTxt(ByVal chave As String, ByVal padrao As String) As String
    CfgTxt = padrao
    If dCfg Is Nothing Then GoTo Expande
    Dim k As String: k = UCase$(Trim$(chave))
    If dCfg.Exists(k) Then
        If Trim$(CStr(dCfg(k))) <> "" Then CfgTxt = Trim$(CStr(dCfg(k)))
    End If
Expande:
    CfgTxt = Replace(CfgTxt, "%USERPROFILE%", Environ$("USERPROFILE"), 1, -1, vbTextCompare)
End Function

Private Function CfgNum(ByVal chave As String, ByVal padrao As Double) As Double
    CfgNum = padrao
    If dCfg Is Nothing Then Exit Function
    Dim k As String: k = UCase$(Trim$(chave))
    If dCfg.Exists(k) Then
        Dim s As String: s = Replace(Trim$(CStr(dCfg(k))), ",", ".")
        If IsNumeric(s) Then CfgNum = Val(s)
    End If
End Function

Private Function CaminhoCatalogo(ByVal chave As String, ByVal padrao As String) As String
    Dim lista As String: lista = CfgTxt(chave, padrao)
    Dim parts() As String, p As Long, cam As String
    parts = Split(lista, ";")
    For p = LBound(parts) To UBound(parts)
        cam = Trim$(parts(p))
        If cam <> "" Then
            If Dir(cam) <> "" Then CaminhoCatalogo = cam: Exit Function
        End If
    Next p
    CaminhoCatalogo = ""
End Function


'==============================================================================
'  FASE 5 - CAMADA VISUAL DAS ABAS
'==============================================================================
' Identifica colunas de veredito/alertas: ficam FORA da zebra para que as
' cores semanticas (verde/vermelho/cinza) prevalecam.
Private Function EhColunaVeredito(ByVal hh As String) As Boolean
    hh = UCase$(Trim$(hh))
    EhColunaVeredito = (InStr(hh, "STATUS") > 0 Or InStr(hh, "APROV") > 0 _
        Or InStr(hh, "SITUACAO") > 0 Or hh = "OBS" Or InStr(hh, "ALERTA") > 0 _
        Or InStr(hh, "APROPRIACAO") > 0 Or hh = "ADERENCIA")
End Function

' Cor da guia por grupo funcional.
Private Function CorAba(ByVal nome As String) As Long
    Select Case UCase$(Trim$(nome))
        Case "PAINEL EXECUTIVO", "PORTFOLIO OBRA"
            CorAba = RGB(46, 125, 50)      ' verde  - visao gerencial
        Case "ALERTAS CRITICOS", "NAO CLASSIFICADOS"
            CorAba = RGB(198, 40, 40)      ' vermelho - exige acao
        Case "MATERIAL VS SERVICO", "MATERIAL", "SERVICO", "SERVICO SEM MATERIAL"
            CorAba = RGB(21, 101, 192)     ' azul - aderencia MAT/SRV
        Case "ANALISE DE CA"
            CorAba = RGB(106, 27, 154)     ' roxo - composicao de custo
        Case "RACIONALIZACAO COM"
            CorAba = RGB(239, 108, 0)      ' laranja - NT.006
        Case Else
            CorAba = RGB(96, 125, 139)     ' cinza - dados de apoio
    End Select
End Function

' Formato numerico pelo cabecalho. "" = nao mexer.
Private Function FormatoColuna(ByVal hh As String) As String
    hh = UCase$(Trim$(hh))
    If InStr(hh, "PERC") > 0 Or InStr(hh, "PORC") > 0 Or InStr(hh, "%") > 0 _
       Or InStr(hh, "MENOR 10") > 0 Then
        FormatoColuna = "0.0"
    ElseIf hh = "SITUACAO" Then
        FormatoColuna = "0"                     ' numero da situacao (1..9)
    ElseIf InStr(hh, "LANCAMENTO") > 0 Or Left$(hh, 2) = "N " _
       Or InStr(hh, "QTD FAM") > 0 Then
        FormatoColuna = "#,##0"                 ' contagens inteiras
    ElseIf hh = "MAT" Or hh = "SRV" _
       Or Left$(hh, 3) = "QTD" Or Left$(hh, 5) = "FAIXA" _
       Or InStr(hh, "VALOR") > 0 Or InStr(hh, "TOTAL") > 0 _
       Or InStr(hh, "DIF") > 0 Or hh = "MOP" Or InStr(hh, "CALCULO DO MOP") > 0 _
       Or InStr(hh, "FRETE") > 0 Or InStr(hh, "TRIBUTOS") > 0 _
       Or hh = "SUPORTE" Or hh = "JOA" Or InStr(hh, "ATV") > 0 _
       Or InStr(hh, "ATIVACAO") > 0 Or InStr(hh, "MAO DE OBRA") > 0 _
       Or InStr(hh, "PRECO") > 0 Then
        FormatoColuna = "#,##0.00"
    Else
        FormatoColuna = ""
    End If
End Function

' Visual padrao de uma aba: cor de guia, cabecalho, zebra (formatacao
' condicional - 1 operacao por coluna, sobrevive a filtro/ordenacao),
' bordas finas, formatos numericos, autofit com teto e freeze.
Private Sub FormatarVisualAba(ws As Worksheet, ByVal nome As String, _
                              ByVal nR As Long, ByVal nC As Long, _
                              Optional ByVal celFreeze As String = "A2")
    Dim jc As Long, hh As String

    ' Paleta institucional (verde Equatorial) + tons de apoio
    Dim corHdr As Long, corHdrLn As Long, corZebra As Long, corBorda As Long
    Dim corTinta As Long, corChave As Long, corGrade As Long
    corHdr = COR_UI_HDR             ' verde institucional (cabecalho)
    corHdrLn = COR_UI_HDR_ACENTO    ' verde escuro (linha de acento)
    corZebra = COR_UI_ZEBRA         ' zebra verde muito clara
    corBorda = COR_UI_BORDA         ' borda cinza clara
    corTinta = COR_UI_TINTA         ' texto do corpo
    corChave = COR_UI_CHAVE         ' 1a coluna (chave)
    corGrade = COR_UI_GRADE         ' linhas internas

    On Error Resume Next
    ws.Tab.Color = CorAba(nome)
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.Zoom = 90
    On Error GoTo 0

    ' Cabecalho (verde, branco, negrito) + linha de acento por baixo
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, nC))
        .Interior.Color = corHdr
        .Font.Color = vbWhite: .Font.Bold = True
        .Font.Name = FONTE_UI: .Font.Size = 9.5
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
        .WrapText = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeBottom).Color = corHdrLn
    End With
    ws.Rows(1).RowHeight = 34

    If nR >= 2 Then
        Dim corpo As Range
        Set corpo = ws.Range(ws.Cells(2, 1), ws.Cells(nR, nC))
        corpo.Font.Name = FONTE_UI
        corpo.Font.Size = 9.5
        corpo.VerticalAlignment = xlCenter
        corpo.Font.Color = corTinta             ' texto grafite (menos duro que preto)

        ' Zebra por coluna, pulando colunas de veredito/alerta (as cores
        ' semanticas dessas colunas devem prevalecer sobre a zebra).
        ' FormatConditions.Add exige a formula no IDIOMA LOCAL do Excel
        ' (pt-BR: =MOD(LIN();2)=0). Escreve a formula em en-US numa celula
        ' temporaria e le FormulaLocal para obter a traducao correta.
        Dim fZebra As String
        On Error Resume Next
        ws.Cells(1, nC + 2).Formula = "=MOD(ROW(),2)=0"
        fZebra = ws.Cells(1, nC + 2).FormulaLocal
        ws.Cells(1, nC + 2).ClearContents
        On Error GoTo 0
        If fZebra = "" Then fZebra = "=MOD(ROW(),2)=0"
        For jc = 1 To nC
            hh = UCase$(CStr(ws.Cells(1, jc).Value))
            If Not EhColunaVeredito(hh) Then
                ' zebra e cosmetica: nunca pode derrubar a geracao
                On Error Resume Next
                With ws.Range(ws.Cells(2, jc), ws.Cells(nR, jc)).FormatConditions.Add( _
                        Type:=xlExpression, Formula1:=fZebra)
                    .Interior.Color = corZebra
                End With
                On Error GoTo 0
            End If
            ' Formato numerico + alinhamento por tipo de coluna
            Dim fmt As String: fmt = FormatoColuna(hh)
            Dim colRg As Range: Set colRg = ws.Range(ws.Cells(2, jc), ws.Cells(nR, jc))
            If fmt <> "" Then colRg.NumberFormat = FormatoNegativo(fmt)
            If EhColunaVeredito(hh) Then
                colRg.HorizontalAlignment = xlCenter
                ws.Cells(1, jc).HorizontalAlignment = xlCenter
                ws.Cells(1, jc).IndentLevel = 0
            ElseIf fmt <> "" Then
                colRg.HorizontalAlignment = xlRight    ' numeros alinhados a direita
                colRg.IndentLevel = 1
                ws.Cells(1, jc).HorizontalAlignment = xlRight
                ws.Cells(1, jc).IndentLevel = 1
            Else
                colRg.HorizontalAlignment = xlLeft
                colRg.IndentLevel = 1
            End If

            ' Data bars sutis em colunas de valor (leitura rapida de magnitude)
            If fmt = "#,##0.00" And (InStr(hh, "VALOR") > 0 Or InStr(hh, "DIF") > 0 _
               Or InStr(hh, "TOTAL") > 0) And Not EhColunaVeredito(hh) Then
                ' Best-effort (so formatacao visual): falha aqui nao altera a analise.
                On Error Resume Next
                With ws.Range(ws.Cells(2, jc), ws.Cells(nR, jc)).FormatConditions.AddDatabar
                    .BarColor.Color = COR_UI_BARRA
                    .BarFillType = xlDataBarFillSolid
                    .BarBorder.Type = xlDataBarBorderNone
                    .Direction = xlContext
                    .NegativeBarFormat.ColorType = xlDataBarColor
                    .NegativeBarFormat.Color.Color = COR_UI_BARRA_NEG
                    .ShowValue = True
                End With
                On Error GoTo 0
            End If

            ' Data bars em colunas de percentual (0..100) - leitura de proporcao
            If fmt = "0.0" And InStr(hh, "PERC") > 0 And Not EhColunaVeredito(hh) Then
                ' Best-effort (so formatacao visual): falha aqui nao altera a analise.
                On Error Resume Next
                With ws.Range(ws.Cells(2, jc), ws.Cells(nR, jc)).FormatConditions.AddDatabar
                    .BarColor.Color = COR_UI_BARRA_PCT
                    .BarFillType = xlDataBarFillSolid
                    .BarBorder.Type = xlDataBarBorderNone
                    .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
                    .MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=100
                    .ShowValue = True
                End With
                On Error GoTo 0
            End If
        Next jc

        ' Bordas finas claras (uma operacao no range inteiro)
        With corpo.Borders
            .LineStyle = xlContinuous
            .Weight = xlHairline
            .Color = corBorda
        End With
        ' Linhas horizontais um pouco mais visiveis (leitura por linha)
        corpo.Borders(xlInsideHorizontal).Color = corGrade
        corpo.Borders(xlInsideVertical).Color = corBorda

        ' 1a coluna em destaque (chave/PEP): negrito e cor institucional
        With ws.Range(ws.Cells(2, 1), ws.Cells(nR, 1))
            .Font.Bold = True
            .Font.Color = corChave
        End With

        ' Altura de linha uniforme (visual mais limpo que o default variavel)
        On Error Resume Next
        ws.Range(ws.Cells(2, 1), ws.Cells(nR, 1)).EntireRow.RowHeight = 19
        On Error GoTo 0

        ' Contorno da tabela (fecha o bloco visualmente)
        On Error Resume Next
        ws.Range(ws.Cells(1, 1), ws.Cells(nR, nC)).BorderAround _
            LineStyle:=xlContinuous, Weight:=xlThin, Color:=corHdrLn
        On Error GoTo 0
    End If

    ws.Rows(1).AutoFilter
    ' AutoFit amostrado: cabecalho + primeiras linhas (rapido em abas grandes)
    Dim amostraR As Long: amostraR = nR
    If amostraR > 200 Then amostraR = 200
    ws.Range(ws.Cells(2, 1), ws.Cells(amostraR, nC)).Columns.AutoFit
    ' Teto e piso de largura por tipo de coluna (descricoes podem ser largas;
    ' colunas de numero/veredito ficam compactas e nao se espremem)
    For jc = 1 To nC
        hh = UCase$(CStr(ws.Cells(1, jc).Value))
        Dim wMax As Double, wMin As Double, wAtual As Double
        wMax = LarguraMaxColuna(hh): wMin = 10
        If EhColunaVeredito(hh) Then wMin = 16
        wAtual = ws.Columns(jc).ColumnWidth
        If wAtual > wMax Then ws.Columns(jc).ColumnWidth = wMax
        If wAtual < wMin Then ws.Columns(jc).ColumnWidth = wMin
    Next jc
    ' Congela cabecalho + 1a coluna quando o freeze padrao (A2) e usado
    Dim fCong As String: fCong = celFreeze
    If UCase$(fCong) = "A2" And nC > 3 Then fCong = "B2"
    AplicarFreeze ws, fCong
    PrepararImpressao ws, nome, nR, nC
End Sub

' Largura maxima da coluna conforme o tipo de conteudo do cabecalho.
Private Function LarguraMaxColuna(ByVal hh As String) As Double
    hh = UCase$(Trim$(hh))
    If InStr(hh, "DESC") > 0 Or InStr(hh, "TEXTO") > 0 Or InStr(hh, "DENOM") > 0 _
       Or InStr(hh, "OBSERV") > 0 Or InStr(hh, "REGRA") > 0 Or InStr(hh, "ACAO") > 0 Then
        LarguraMaxColuna = 52
    ElseIf InStr(hh, "PEP") > 0 Then
        LarguraMaxColuna = 34
    ElseIf EhColunaVeredito(hh) Then
        LarguraMaxColuna = 24
    Else
        LarguraMaxColuna = 28
    End If
End Function

' Acrescenta o tratamento de negativo (vermelho) aos formatos numericos.
Private Function FormatoNegativo(ByVal fmt As String) As String
    Select Case fmt
        Case "#,##0.00": FormatoNegativo = "#,##0.00;[Red]-#,##0.00;""-"""
        Case "#,##0":    FormatoNegativo = "#,##0;[Red]-#,##0;""-"""
        Case "0.0":      FormatoNegativo = "0.0;[Red]-0.0;""-"""
        Case Else:       FormatoNegativo = fmt
    End Select
End Function

' Impressao pronta: paisagem, ajuste a largura, cabecalho repetido e rodape.
Private Sub PrepararImpressao(ws As Worksheet, ByVal nome As String, _
                              ByVal nR As Long, ByVal nC As Long)
    ' Best-effort (so formatacao visual): falha aqui nao altera a analise.
    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$1:$1"
        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.3)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.45)
        .HeaderMargin = Application.InchesToPoints(0.2)
        .FooterMargin = Application.InchesToPoints(0.2)
        .CenterHeader = "&""" & FONTE_UI & ",Bold""&11" & UCase$(nome)
        .LeftFooter = "&""" & FONTE_UI & """&8CKCP RS2  |  Analise de Custos"
        .RightFooter = "&""" & FONTE_UI & """&8Pagina &P de &N"
        .PrintGridlines = False
    End With
    Application.PrintCommunication = True
    On Error GoTo 0
End Sub

' Ordena as guias por fluxo de leitura (gerencial -> detalhe -> apoio)
' e deixa o PAINEL EXECUTIVO ativo ao final.
Private Sub OrganizarAbas()
    Dim ordem As Variant, i As Long, ws As Worksheet, pos As Long
    ordem = Array("PAINEL EXECUTIVO", "CHECKLIST CKCP", "PORTFOLIO OBRA", _
                  "ALERTAS CRITICOS", "MATERIAL vs SERVICO", "MATERIAL", "SERVICO", _
                  "SERVICO SEM MATERIAL", "ANALISE DE CA", _
                  "NAO CLASSIFICADOS", "RACIONALIZACAO COM", "RAZAO CJ", "REGRAS")
    pos = 0
    For i = 0 To UBound(ordem)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ActiveWorkbook.Worksheets(CStr(ordem(i)))
        On Error GoTo 0
        If Not ws Is Nothing Then
            pos = pos + 1
            If ws.Index <> pos Then ws.Move Before:=ActiveWorkbook.Worksheets(pos)
        End If
    Next i
    On Error Resume Next
    ActiveWorkbook.Worksheets("PAINEL EXECUTIVO").Activate
    ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = 1
    On Error GoTo 0
End Sub


Private Sub Gerar_MatVsServAT()
    ' Pipeline MAT vs SERV AT integrado ao GerarRelatorio.
    ' Reutiliza dados() ja carregado; de-para vem da aba CORRESP.
    Call CarregarCorresp
    If nCorr = 0 Then Exit Sub
    Call CarregarDados_AT
    If nItens = 0 Then Exit Sub
    Call AplicarRegrasPreAgrupamento
    Call AgruparItens
    Call AplicarRegrasPosAgrupamento
    Call PadronizarCls2
    Call CalcularMatSrv
    Call CalcularAderencia
    Call CalcularTipoCusto
    Call OrdenarPorGrupo
    Call CalcularPctMop
    Call EscreverAbaAT
    Call CriarPremissas
End Sub

Private Sub CarregarDados_AT()
    Dim i As Long, n As Long
    n = UBound(dados, 1)
    nItens = 0
    ReDim aItens(n)
    For i = 1 To n
        Dim vPep As String
        vPep = Trim$(CStr(dados(i, cPEP)))
        If vPep = "" Then GoTo NextRow
        Dim vMat As String
        vMat = NormCod(dados(i, cMaterial))
        If vMat = "" Then vMat = NormCod(ValorCampo(i, cClasse))
        Dim vTxt As String
        vTxt = TextoCampo(i, cTexto)
        If vTxt = "" Then vTxt = TextoCampo(i, cDescClasse)
        Dim eMat As Boolean
        eMat = EhMaterial(CStr(dados(i, cClassif)))
        Dim vC2 As String
        If eMat Then vC2 = MatInfoLinha(i, 2) Else vC2 = SrvInfoLinha(i, 1)
        With aItens(nItens)
            .Empresa = TextoCampo(i, cEmpresa)
            .Pep = vPep
            .Pep3Nivel = PEP3(vPep)
            .Tipo = Right$(vPep, 1)
            .Segmento = ""                ' nao ha segmento de obra no dataset CKCP
            .TipoObraAneel = ""           ' idem (coluna fica vazia)
            .Material = vMat
            .TextoMaterial = vTxt
            .Uml = TextoCampo(i, cUML)
            .QtdEntrada = ToNum(dados(i, cQtd))
            .ValorMoeda = ToNum(dados(i, cValor))
            .Cls1 = IIf(eMat, "MATERIAL", "SERVI" & Chr(199) & "O")
            .Cls2 = vC2
            .Cls2Orig = vC2
            .Inconformidade = ""
        End With
        nItens = nItens + 1
NextRow:
    Next i

    ' Preencher UML vazio com UML de outro lancamento do mesmo codigo
    Dim dUml As Object: Set dUml = CreateObject("Scripting.Dictionary")
    Dim iu As Long
    For iu = 0 To nItens - 1
        If aItens(iu).Uml <> "" And Not dUml.Exists(aItens(iu).Material) Then
            dUml(aItens(iu).Material) = aItens(iu).Uml
        End If
    Next iu
    For iu = 0 To nItens - 1
        If aItens(iu).Uml = "" Then
            If dUml.Exists(aItens(iu).Material) Then aItens(iu).Uml = dUml(aItens(iu).Material)
        End If
    Next iu
End Sub

' ============================================================
' 1. CARREGAR CORRESP
' ============================================================
Private Sub CarregarCorresp()
    Dim ws As Worksheet
    Dim lr As Long, i As Long
    Dim colMat As Integer, colSrv As Integer, colTipo As Integer

    Set ws = AcharAbaCorresp()

    If ws Is Nothing Then
        Dim wsLst As Worksheet, lst As String
        For Each wsLst In ActiveWorkbook.Worksheets
            lst = lst & " - " & wsLst.Name & vbCrLf
        Next wsLst
        MsgBox "Aba CORRESP nao encontrada!" & vbCrLf & vbCrLf & _
               "A analise MAT vs SERV AT precisa de uma aba cujo nome CONTENHA " & _
               "'CORRESP' (de-para MATERIAL x SERVICO)." & vbCrLf & vbCrLf & _
               "Abas existentes neste arquivo:" & vbCrLf & lst, vbExclamation, _
               "MAT vs SERV AT"
        Exit Sub
    End If

    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lr < 2 Then MsgBox "CORRESP vazia!": Exit Sub

    ' Detectar colunas
    colMat = 0: colSrv = 0: colTipo = 0
    Dim h As Integer
    For h = 1 To 10
        Select Case UCase(Trim(CStr(ws.Cells(1, h).Value)))
            Case "MATERIAL": colMat = h
            Case "SERVI" & Chr(199) & "O", "SERVICO", "SERVICO": colSrv = h
            Case "TIPO": colTipo = h
        End Select
    Next h

    If colMat = 0 Or colSrv = 0 Then
        MsgBox "Colunas MATERIAL/SERVICO nao encontradas na aba CORRESP"
        Exit Sub
    End If

    nCorr = 0
    ReDim aMatCorr(lr)
    ReDim aSrvCorr(lr)
    ReDim aTipoCorr(lr)

    For i = 2 To lr
        Dim vMat As String, vSrv As String, vTipo As String
        vMat  = CleanCod(CStr(ws.Cells(i, colMat).Value))
        vSrv  = CleanCod(CStr(ws.Cells(i, colSrv).Value))
        vTipo = "D"
        If colTipo > 0 Then
            vTipo = UCase(Trim(CStr(ws.Cells(i, colTipo).Value)))
            If vTipo = "" Then vTipo = "D"
        End If
        If vMat <> "" And vSrv <> "" Then
            aMatCorr(nCorr) = vMat
            aSrvCorr(nCorr) = vSrv
            aTipoCorr(nCorr) = vTipo
            nCorr = nCorr + 1
        End If
    Next i
End Sub

' Localiza a aba CORRESP de forma tolerante: ignora caixa, espacos comuns,
' espacos nao-separaveis (Chr 160) e tabs, e aceita nomes que CONTENHAM
' "CORRESP" (ex.: "CORRESP 2025", "DE-PARA CORRESP"). Procura primeiro no
' workbook ativo, depois em qualquer workbook aberto.
Private Function AcharAbaCorresp() As Worksheet
    Dim wbx As Workbook
    On Error Resume Next
    Set AcharAbaCorresp = AchaCorrespNoWb(ActiveWorkbook)
    If Not AcharAbaCorresp Is Nothing Then Exit Function
    For Each wbx In Application.Workbooks
        Set AcharAbaCorresp = AchaCorrespNoWb(wbx)
        If Not AcharAbaCorresp Is Nothing Then Exit Function
    Next wbx
    On Error GoTo 0
End Function

Private Function AchaCorrespNoWb(ByVal wb As Workbook) As Worksheet
    Dim wsx As Worksheet
    ' 1a preferencia: nome normalizado == CORRESP
    For Each wsx In wb.Worksheets
        If NomeNorm(wsx.Name) = "CORRESP" Then
            Set AchaCorrespNoWb = wsx
            Exit Function
        End If
    Next wsx
    ' 2a preferencia: nome CONTEM CORRESP
    For Each wsx In wb.Worksheets
        If InStr(NomeNorm(wsx.Name), "CORRESP") > 0 Then
            Set AchaCorrespNoWb = wsx
            Exit Function
        End If
    Next wsx
End Function

Private Function NomeNorm(ByVal s As String) As String
    s = Replace(s, Chr(160), "")
    s = Replace(s, " ", "")
    s = Replace(s, Chr(9), "")
    NomeNorm = UCase$(Trim$(s))
End Function

' ============================================================
' 2. CARREGAR DADOS DO RAZAO CJ
' ============================================================

' ============================================================
' 3. REGRAS PRE-AGRUPAMENTO (R1 e R2)
' ============================================================
Private Sub AplicarRegrasPreAgrupamento()
    Dim i As Long
    For i = 0 To nItens - 1
        With aItens(i)
            If .Cls1 <> "SERVI" & Chr(199) & "O" And .Cls1 <> "SERVICO" And .Cls1 <> "SERVICO" Then GoTo NextItem
            If Not TemSaldo(.ValorMoeda, .QtdEntrada) Then GoTo NextItem
            If Right(UCase(.Pep), 2) = ".D" Then GoTo NextItem

            Dim txt As String
            txt = UCase(.TextoMaterial)

            ' R1 - retirada/demolicao/desativacao
            Dim nomeAcao As String
            nomeAcao = ""
            If ContemPalavra(txt, "RETIRADA") Or Left(txt, 4) = "RET." Or Left(txt, 4) = "RET " Then
                nomeAcao = "retirada"
            ElseIf ContemPalavra(txt, "REMOCAO") Or ContemPalavra(txt, "REMO" & Chr(199) & Chr(195) & "O") Or Left(txt, 4) = "REM." Or Left(txt, 4) = "REM " Then
                nomeAcao = "remo" & Chr(231) & Chr(227) & "o"
            ElseIf ContemPalavra(txt, "DEMOLICAO") Or ContemPalavra(txt, "DEMOLI" & Chr(199) & Chr(195) & "O") Or Left(txt, 4) = "DEM." Or Left(txt, 4) = "DEM " Then
                nomeAcao = "demoli" & Chr(231) & Chr(227) & "o"
            ElseIf ContemPalavra(txt, "DESMONTAGEM") Or Left(txt, 5) = "DESM." Or Left(txt, 5) = "DESM " Then
                nomeAcao = "desmontagem"
            ElseIf ContemPalavra(txt, "DESINSTALACAO") Or ContemPalavra(txt, "DESINSTALA") Or Left(txt, 5) = "DESI." Or Left(txt, 5) = "DESI " Then
                nomeAcao = "desinstala" & Chr(231) & Chr(227) & "o"
            ElseIf ContemPalavra(txt, "DESATIVAR") Or ContemPalavra(txt, "DESATIVACAO") Or Left(txt, 6) = "DESAT." Or Left(txt, 7) = "DESATIV" Then
                nomeAcao = "desativa" & Chr(231) & Chr(227) & "o"
            End If

            If nomeAcao <> "" Then
                .Inconformidade = "Servi" & Chr(231) & "o de " & nomeAcao & " indevido em PEP de instala" & Chr(231) & Chr(227) & "o"
                GoTo NextItem
            End If

            ' R2 - substituicao
            If ContemPalavra(txt, "SUBSTITUICAO") Or ContemPalavra(txt, "SUBSTITUI" & Chr(199) & Chr(195) & "O") Or _
               ContemPalavra(txt, "SUBSTITUIR") Or _
               ContemPalavra(txt, "SUBST.") Or Left(txt, 6) = "SUBST " Then
                .Inconformidade = "Servi" & Chr(231) & "o de substitui" & Chr(231) & Chr(227) & "o deve ser reclassificado para o OPEX"
            End If
        End With
NextItem:
    Next i
End Sub

' ============================================================
' 4. AGRUPAR ITENS (soma por chave)
' ============================================================
Private Sub AgruparItens()
    Dim dictKey  As Object
    Dim dictVal  As Object
    Dim dictInc  As Object
    Set dictKey = CreateObject("Scripting.Dictionary")
    Set dictVal = CreateObject("Scripting.Dictionary")
    Set dictInc = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 0 To nItens - 1
        With aItens(i)
            Dim k As String
            k = .Empresa & "|" & .Segmento & "|" & .TipoObraAneel & "|" & _
                .Pep3Nivel & "|" & .Pep & "|" & .Tipo & "|" & .Material & "|" & _
                .TextoMaterial & "|" & .Uml & "|" & .Cls1 & "|" & .Cls2 & "|" & .Cls2Orig

            If Not dictKey.Exists(k) Then
                dictKey(k) = i
                dictVal(k) = Array(.ValorMoeda, .QtdEntrada)
                dictInc(k) = .Inconformidade
            Else
                Dim arr As Variant
                arr = dictVal(k)
                arr(0) = arr(0) + .ValorMoeda
                arr(1) = arr(1) + .QtdEntrada
                dictVal(k) = arr
                If dictInc(k) = "" And .Inconformidade <> "" Then
                    dictInc(k) = .Inconformidade
                End If
            End If
        End With
    Next i

    ' Reconstruir array agrupado
    Dim nNovo As Long
    nNovo = dictKey.Count
    Dim aNovo() As tItem
    ReDim aNovo(nNovo)

    Dim j As Long
    j = 0
    Dim vk As Variant
    For Each vk In dictKey.Keys
        Dim idx As Long
        idx = dictKey(vk)
        Dim arr2 As Variant
        arr2 = dictVal(vk)
        aNovo(j) = aItens(idx)
        aNovo(j).ValorMoeda  = arr2(0)
        aNovo(j).QtdEntrada  = arr2(1)
        aNovo(j).Inconformidade = dictInc(vk)
        ' Near-zero
        If Abs(aNovo(j).ValorMoeda) < 0.000000001 Then aNovo(j).ValorMoeda = 0
        If Abs(aNovo(j).QtdEntrada) < 0.000000001 Then aNovo(j).QtdEntrada = 0
        ' Limpar inconformidade se sem saldo apos agrupamento
        If aNovo(j).ValorMoeda = 0 And aNovo(j).QtdEntrada = 0 Then
            aNovo(j).Inconformidade = ""
        End If
        j = j + 1
    Next vk

    nItens = nNovo
    ReDim aItens(nItens)
    For i = 0 To nItens - 1
        aItens(i) = aNovo(i)
    Next i
End Sub

' ============================================================
' 5. REGRAS POS-AGRUPAMENTO (R3 a R7)
' ============================================================
Private Sub AplicarRegrasPosAgrupamento()
    Dim i As Long
    For i = 0 To nItens - 1
        With aItens(i)
            If .Inconformidade <> "" Then GoTo NextI

            Dim eSrv As Boolean, eMat As Boolean
            eSrv = (.Cls1 = "SERVI" & Chr(199) & "O" Or .Cls1 = "SERVICO" Or .Cls1 = "SERVICO")
            eMat = (.Cls1 = "MATERIAL")
            Dim temS As Boolean
            temS = TemSaldo(.ValorMoeda, .QtdEntrada)

            ' R3: saldo negativo servico
            If eSrv And temS Then
                If .ValorMoeda < 0 Or .QtdEntrada < 0 Then
                    .Inconformidade = "Servi" & Chr(231) & "o com saldo negativo"
                    GoTo NextI
                End If
            End If

            ' R4: genericos
            If eSrv And temS Then
                Dim txt As String
                txt = UCase(.TextoMaterial)
                Dim motivo As String
                motivo = ""
                If (ContemPalavra(txt, "EXEC") And ContemPalavra(txt, "INTEGRACAO") And ContemPalavra(txt, "AUTOMACAO")) Or _
                   (ContemPalavra(txt, "EXEC") And ContemPalavra(txt, "INTEGRA") And ContemPalavra(txt, "AUTOMA")) Then
                    motivo = "sem especificar objeto ou escopo da integra" & Chr(231) & Chr(227) & "o/automa" & Chr(231) & Chr(227) & "o"
                ElseIf ContemPalavra(txt, "TRABALHOS EXCEPCIONAIS") Or ContemPalavra(txt, "TRABALHO EXCEPCIONAL") Then
                    motivo = "termo ""excepcionais"" n" & Chr(227) & "o define o servi" & Chr(231) & "o executado"
                ElseIf ContemPalavra(txt, "MATERIAIS MENORES") Or ContemPalavra(txt, "MATERIAL MENOR") Then
                    motivo = "sem identificar qual material foi fornecido"
                ElseIf ContemPalavra(txt, "INSTALACAO E MANUTENCAO DE EQUIPAMENTO") Or _
                       ContemPalavra(txt, "INSTALA" & Chr(199) & Chr(195) & "O E MANUTEN" & Chr(199) & Chr(195) & "O DE EQUIPAMENTO") Then
                    motivo = "n" & Chr(227) & "o especifica qual equipamento"
                ElseIf ContemPalavra(txt, "ADMINISTRACAO DE OBRA") Or ContemPalavra(txt, "ADMINISTRA" & Chr(199) & Chr(195) & "O DE OBRA") Then
                    motivo = "custo administrativo sem detalhamento t" & Chr(233) & "cnico"
                ElseIf ContemPalavra(txt, "MAO DE OBRA DE ADMINISTRA") Or ContemPalavra(txt, "M" & Chr(195) & "O DE OBRA DE ADMINISTRA") Then
                    motivo = "custo administrativo sem detalhamento t" & Chr(233) & "cnico"
                ElseIf ContemPalavra(txt, "FISCALIZACAO DE ENGENHARIA") Or ContemPalavra(txt, "FISCALIZA" & Chr(199) & Chr(195) & "O DE ENGENHARIA") Then
                    motivo = "n" & Chr(227) & "o identifica qual servi" & Chr(231) & "o foi fiscalizado"
                ElseIf txt = "CONSULTORIA" Then
                    motivo = "sem especificar objeto ou escopo"
                ElseIf ContemPalavra(txt, "DIFERENCA DE VALOR") And ContemPalavra(txt, "FATOR K") Then
                    motivo = "ajuste de valor sem descri" & Chr(231) & Chr(227) & "o de servi" & Chr(231) & "o executado"
                ElseIf ContemPalavra(txt, "LOCACAO DE VEICULO") Or ContemPalavra(txt, "LOCA" & Chr(199) & Chr(195) & "O DE VE" & Chr(205) & "CULO") Then
                    motivo = "sem identificar finalidade ou v" & Chr(237) & "nculo com a obra"
                ElseIf ContemPalavra(txt, "TRANSPORTE DE EQUIPAMENTOS MENORES") Or ContemPalavra(txt, "TRANSPORTE DE EQUIPAMENTO MENOR") Then
                    motivo = "sem especificar o equipamento transportado"
                ElseIf ContemPalavra(txt, "TRANSPORTE DE EQUIPAMENTO") And (ContemPalavra(txt, "/MATERIAL") Or ContemPalavra(txt, "E MATERIAL")) Then
                    motivo = "sem especificar o equipamento ou material transportado"
                ElseIf (ContemPalavra(txt, "LIMPEZA") And ContemPalavra(txt, "OBRA")) Or ContemPalavra(txt, "LIMPEZA INICIAL") Then
                    motivo = "servi" & Chr(231) & "o de apoio sem v" & Chr(237) & "nculo com ativo da obra"
                End If
                If motivo <> "" Then
                    .Inconformidade = "Servi" & Chr(231) & "o gen" & Chr(233) & "rico (" & motivo & ")"
                    GoTo NextI
                End If
            End If

            ' R5: mobilizacao sem par
            If eSrv And temS Then
                Dim ehMob As Boolean
                ehMob = ContemPalavra(txt, "MOB") Or ContemPalavra(txt, "MOBILIZAR") Or _
                        ContemPalavra(txt, "MOBILIZACAO") Or ContemPalavra(txt, "MOBILIZA" & Chr(199) & Chr(195) & "O") Or _
                        ContemPalavra(txt, "DESMOB") Or ContemPalavra(txt, "DESMOBILIZAR")
                If ehMob Then
                    Dim sufixo As String
                    sufixo = Right(UCase(.Pep), 2)
                    Dim pep3 As String
                    pep3 = .Pep3Nivel
                    Dim temD As Boolean, temI As Boolean
                    Dim temMobD As Boolean, temMobI As Boolean
                    temD = PepExisteComSufixo(pep3, "D")
                    temI = PepExisteComSufixo(pep3, "I")
                    If temD And temI Then
                        temMobD = PepTemMob(pep3, "D")
                        temMobI = PepTemMob(pep3, "I")
                        If sufixo = ".I" And Not temMobD Then
                            .Inconformidade = "Custo de mob. n" & Chr(227) & "o identificado no PEP .D"
                            GoTo NextI
                        ElseIf sufixo = ".D" And Not temMobI Then
                            .Inconformidade = "Custo de mob. n" & Chr(227) & "o identificado no PEP .I"
                            GoTo NextI
                        End If
                    End If
                End If
            End If

            ' R6: viagem/hospedagem/alimentacao (valor ? 0, exceto caixa de passagem)
            If .ValorMoeda <> 0 Then
                Dim txtR6 As String
                txtR6 = UCase(.TextoMaterial) & " " & UCase(.Cls2)
                Dim ehCaixaPassagem As Boolean
                ehCaixaPassagem = ContemPalavra(txtR6, "CAIXA") And ContemPalavra(txtR6, "PASSAGEM")
                If Not ehCaixaPassagem Then
                    If ContemPalavra(txtR6, "HOSPEDAGEM") Or ContemPalavra(txtR6, "HOSPEDAGEM") Or _
                       ContemPalavra(txtR6, "ALIMENTACAO") Or ContemPalavra(txtR6, "ALIMENTA" & Chr(199) & Chr(195) & "O") Or _
                       ContemPalavra(txtR6, "VIAGEM") Or ContemPalavra(txtR6, "PASSAGEM") Or _
                       ContemPalavra(txtR6, "MOBILIDADE") Or ContemPalavra(txtR6, "CONDUCAO") Or _
                       ContemPalavra(txtR6, "CONDU" & Chr(199) & Chr(195) & "O") Then
                        .Inconformidade = "Custo de viagem deve ser reclassificado para Opex"
                        GoTo NextI
                    End If
                End If
            End If

            ' R7: material saldo negativo fora de .D
            If eMat Then
                If Right(UCase(.Pep), 2) <> ".D" Then
                    If .ValorMoeda < 0 Or .QtdEntrada < 0 Then
                        .Inconformidade = "Material com saldo negativo indevido em PEP de instala" & Chr(231) & Chr(227) & "o"
                    End If
                End If
            End If

            ' R8: poda em PEP .I do segmento MANUTENCAO AT
            If eSrv And .ValorMoeda <> 0 Then
                If Right(UCase(.Pep), 2) = ".I" Then
                    If ContemPalavra(UCase(.Segmento), "MANUT") Then
                        If ContemPalavra(UCase(.TextoMaterial), "PODA") Then
                            .Inconformidade = "Servi" & Chr(231) & "o de poda " & Chr(233) & " indevido em obra de manuten" & Chr(231) & Chr(227) & "o deve ser reclassificado para Opex"
                        End If
                    End If
                End If
            End If
        End With
NextI:
    Next i
End Sub

' ============================================================
' 6. PADRONIZAR CLS2 DOS SERVICOS
' ============================================================
Private Sub PadronizarCls2()
    ' Lookup: (PEP|MATERIAL) -> CLS2 do material
    Dim dictMatCls2 As Object
    Set dictMatCls2 = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 0 To nItens - 1
        If aItens(i).Cls1 = "MATERIAL" Then
            Dim k As String
            k = aItens(i).Pep & "|" & aItens(i).Material
            If Not dictMatCls2.Exists(k) Then
                dictMatCls2(k) = aItens(i).Cls2
            End If
        End If
    Next i

    ' Padronizar CLS2 dos servicos
    For i = 0 To nItens - 1
        If aItens(i).Cls1 = "SERVI" & Chr(199) & "O" Or aItens(i).Cls1 = "SERVICO" Or aItens(i).Cls1 = "SERVICO" Then
            Dim srvCode As String
            srvCode = aItens(i).Material
            ' Buscar materiais correspondentes
            Dim j As Integer
            For j = 0 To nCorr - 1
                If aSrvCorr(j) = srvCode Then
                    Dim matCode As String
                    matCode = aMatCorr(j)
                    Dim kMat As String
                    kMat = aItens(i).Pep & "|" & matCode
                    If dictMatCls2.Exists(kMat) Then
                        Dim matCls2 As String
                        matCls2 = dictMatCls2(kMat)
                        If matCls2 <> "" And matCls2 <> aItens(i).Cls2 Then
                            aItens(i).Cls2 = matCls2
                        End If
                        Exit For
                    End If
                End If
            Next j
        End If
    Next i
End Sub

' ============================================================
' 7. CALCULAR MAT e SRV DO RAZAO CJ
' ============================================================
Private Sub CalcularMatSrv()
    ' MAT = soma QTD MATERIAL por (PEP, CLS2) com abs() no saldo final
    ' SRV do servico = QTD SERVICO por (PEP, CLS2 padronizado)
    ' SRV do material = soma QTD dos servicos D correspondentes por CLS2 original

    Dim dictMat As Object, dictSrv As Object
    Set dictMat = CreateObject("Scripting.Dictionary")
    Set dictSrv = CreateObject("Scripting.Dictionary")
    Dim dictSrvD As Object
    Set dictSrvD = CreateObject("Scripting.Dictionary") ' apenas tipo D


    Dim i As Long
    For i = 0 To nItens - 1
        With aItens(i)
            Dim k As String
            k = .Pep & "|" & .Cls2
            If .Cls1 = "MATERIAL" Then
                If dictMat.Exists(k) Then
                    dictMat(k) = dictMat(k) + .QtdEntrada
                Else
                    dictMat(k) = .QtdEntrada
                End If
            End If
            If .Cls1 = "SERVI" & Chr(199) & "O" Or .Cls1 = "SERVICO" Or .Cls1 = "SERVICO" Then
                If dictSrv.Exists(k) Then
                    dictSrv(k) = dictSrv(k) + .QtdEntrada
                Else
                    dictSrv(k) = .QtdEntrada
                End If
                ' Acumular apenas tipo D para SRV do material
                If GetTipoServico(.Material) <> "C" Then
                    If dictSrvD.Exists(k) Then
                        dictSrvD(k) = dictSrvD(k) + .QtdEntrada
                    Else
                        dictSrvD(k) = .QtdEntrada
                    End If
                End If
            End If
        End With
    Next i

    ' Aplicar abs() no saldo final de MAT
    Dim vk As Variant
    For Each vk In dictMat.Keys
        If dictMat(vk) < 0 Then dictMat(vk) = Abs(dictMat(vk))
    Next vk

    ' Lookup CLS2_ORIG por (PEP|MATERIAL)
    Dim dictOrigCls2 As Object
    Set dictOrigCls2 = CreateObject("Scripting.Dictionary")
    For i = 0 To nItens - 1
        Dim ko As String
        ko = aItens(i).Pep & "|" & aItens(i).Material
        If Not dictOrigCls2.Exists(ko) Then
            dictOrigCls2(ko) = aItens(i).Cls2Orig
        End If
    Next i

    ' Atribuir MAT e SRV
    For i = 0 To nItens - 1
        With aItens(i)
            Dim kCls2 As String
            kCls2 = .Pep & "|" & .Cls2

            ' MAT = qtd total de material da familia CLS2 neste PEP
            If dictMat.Exists(kCls2) Then
                .Mat = dictMat(kCls2)
            Else
                .Mat = 0
            End If

            ' SRV
            If .Cls1 = "MATERIAL" Then
                ' SRV do material = soma QTD dos servicos D correspondentes (CLS2 padronizado)
                Dim srvTotal As Double
                srvTotal = 0
                Dim visitados As Object
                Set visitados = CreateObject("Scripting.Dictionary")
                Dim j As Integer
                For j = 0 To nCorr - 1
                    If aMatCorr(j) = .Material And aTipoCorr(j) = "D" Then
                        Dim srvCod As String
                        srvCod = aSrvCorr(j)
                        ' CLS2 padronizado do servico = CLS2 do material
                        Dim kSrv As String
                        kSrv = .Pep & "|" & .Cls2
                        If Not visitados.Exists(kSrv) Then
                            visitados(kSrv) = 1
                            If dictSrvD.Exists(kSrv) Then
                                srvTotal = srvTotal + dictSrvD(kSrv)
                            End If
                        End If
                    End If
                Next j
                .Srv = srvTotal
            Else
                ' SRV do servico = apenas servicos TIPO D da familia CLS2 padronizada
                If dictSrvD.Exists(kCls2) Then
                    .Srv = dictSrvD(kCls2)
                Else
                    .Srv = 0
                End If
            End If
        End With
    Next i
End Sub

' ============================================================
' 8. CALCULAR ADERENCIA
' ============================================================
Private Sub CalcularAderencia()
    ' Lookup QTD por (PEP|MATERIAL)
    Dim dictQtd As Object
    Set dictQtd = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 0 To nItens - 1
        With aItens(i)
            Dim k As String
            k = .Pep & "|" & .Material
            Dim qtd As Double
            qtd = .QtdEntrada
            If .Cls1 = "MATERIAL" And Right(UCase(.Pep), 2) = ".D" Then
                qtd = Abs(qtd)
            End If
            dictQtd(k) = qtd
        End With
    Next i

    For i = 0 To nItens - 1
        With aItens(i)
            ' Sem saldo ? em branco
            If .ValorMoeda = 0 And .QtdEntrada = 0 Then
                .Aderencia = ""
                GoTo NextAd
            End If

            Dim code As String
            code = .Material
            Dim pep As String
            pep = .Pep

            ' Auto-correspondente
            If EhAutoCorrespondente(code) Then
                Dim qAuto As Double
                qAuto = 0
                Dim kAuto As String
                kAuto = pep & "|" & code
                If dictQtd.Exists(kAuto) Then qAuto = dictQtd(kAuto)
                If qAuto <> 0 Then
                    .Aderencia = "ADERENTE"
                Else
                    .Aderencia = ""
                End If
                GoTo NextAd
            End If

            If .Cls1 = "MATERIAL" Then
                ' Servicos D encontrados no PEP
                Dim foundD() As String, nFoundD As Integer
                nFoundD = 0
                ReDim foundD(nCorr)
                Dim foundC() As String, nFoundC As Integer
                nFoundC = 0
                ReDim foundC(nCorr)

                Dim j As Integer
                For j = 0 To nCorr - 1
                    If aMatCorr(j) = code Then
                        Dim kSrv As String
                        kSrv = pep & "|" & aSrvCorr(j)
                        If dictQtd.Exists(kSrv) Then
                            If aTipoCorr(j) = "D" Then
                                foundD(nFoundD) = aSrvCorr(j)
                                nFoundD = nFoundD + 1
                            Else
                                foundC(nFoundC) = aSrvCorr(j)
                                nFoundC = nFoundC + 1
                            End If
                        End If
                    End If
                Next j

                If nFoundD = 0 And nFoundC = 0 Then
                    .Aderencia = ""
                ElseIf nFoundD = 0 And nFoundC > 0 Then
                    .Aderencia = "SEM SERVI" & Chr(199) & "O CORRESPONDENTE"
                Else
                    ' Soma mats D do grupo
                    Dim allMatsD As Object
                    Set allMatsD = CreateObject("Scripting.Dictionary")
                    Dim jj As Integer
                    For jj = 0 To nFoundD - 1
                        Dim m As Integer
                        For m = 0 To nCorr - 1
                            If aSrvCorr(m) = foundD(jj) And aTipoCorr(m) = "D" Then
                                allMatsD(aMatCorr(m)) = 1
                            End If
                        Next m
                    Next jj

                    Dim qtdMatD As Double, qtdSrvD As Double
                    qtdMatD = 0: qtdSrvD = 0
                    Dim vMat As Variant
                    For Each vMat In allMatsD.Keys
                        Dim kM As String
                        kM = pep & "|" & CStr(vMat)
                        If dictQtd.Exists(kM) Then qtdMatD = qtdMatD + dictQtd(kM)
                    Next vMat
                    For jj = 0 To nFoundD - 1
                        Dim kSD As String
                        kSD = pep & "|" & foundD(jj)
                        If dictQtd.Exists(kSD) Then qtdSrvD = qtdSrvD + dictQtd(kSD)
                    Next jj

                    If Round(qtdMatD, 4) = Round(qtdSrvD, 4) Then
                        .Aderencia = "ADERENTE"
                    Else
                        .Aderencia = "QTD DIVERGENTE"
                    End If
                End If

            ElseIf .Cls1 = "SERVI" & Chr(199) & "O" Or .Cls1 = "SERVICO" Or .Cls1 = "SERVICO" Then
                Dim tipoCodigo As String
                tipoCodigo = GetTipoServico(code)

                If tipoCodigo = "C" Then
                    ' C: qtd servico == qtd material correspondente
                    Dim foundMatsC() As String, nFMC As Integer
                    nFMC = 0
                    ReDim foundMatsC(nCorr)
                    For j = 0 To nCorr - 1
                        If aSrvCorr(j) = code And aTipoCorr(j) = "C" Then
                            Dim kFC As String
                            kFC = pep & "|" & aMatCorr(j)
                            If dictQtd.Exists(kFC) Then
                                foundMatsC(nFMC) = aMatCorr(j)
                                nFMC = nFMC + 1
                            End If
                        End If
                    Next j
                    If nFMC = 0 Then
                        .Aderencia = "SEM MATERIAL CORRESPONDENTE"
                    Else
                        Dim qtdSrvC As Double, qtdMatC As Double
                        qtdSrvC = 0: qtdMatC = 0
                        Dim kSC As String
                        kSC = pep & "|" & code
                        If dictQtd.Exists(kSC) Then qtdSrvC = dictQtd(kSC)
                        For jj = 0 To nFMC - 1
                            Dim kMC As String
                            kMC = pep & "|" & foundMatsC(jj)
                            If dictQtd.Exists(kMC) Then qtdMatC = qtdMatC + dictQtd(kMC)
                        Next jj
                        If Round(qtdSrvC, 4) = Round(qtdMatC, 4) Then
                            .Aderencia = "ADERENTE"
                        Else
                            .Aderencia = "QTD DIVERGENTE"
                        End If
                    End If
                Else
                    ' D: soma srvs D == soma mats D
                    Dim foundMatsD() As String, nFMD As Integer
                    nFMD = 0
                    ReDim foundMatsD(nCorr)
                    For j = 0 To nCorr - 1
                        If aSrvCorr(j) = code And aTipoCorr(j) = "D" Then
                            Dim kFD As String
                            kFD = pep & "|" & aMatCorr(j)
                            If dictQtd.Exists(kFD) Then
                                foundMatsD(nFMD) = aMatCorr(j)
                                nFMD = nFMD + 1
                            End If
                        End If
                    Next j
                    If nFMD = 0 Then
                        .Aderencia = "SEM MATERIAL CORRESPONDENTE"
                    Else
                        Dim allSrvsD As Object
                        Set allSrvsD = CreateObject("Scripting.Dictionary")
                        For jj = 0 To nFMD - 1
                            For m = 0 To nCorr - 1
                                If aMatCorr(m) = foundMatsD(jj) And aTipoCorr(m) = "D" Then
                                    allSrvsD(aSrvCorr(m)) = 1
                                End If
                            Next m
                        Next jj

                        Dim qtdMatDD As Double, qtdSrvDD As Double
                        qtdMatDD = 0: qtdSrvDD = 0
                        For jj = 0 To nFMD - 1
                            Dim kMDD As String
                            kMDD = pep & "|" & foundMatsD(jj)
                            If dictQtd.Exists(kMDD) Then qtdMatDD = qtdMatDD + dictQtd(kMDD)
                        Next jj
                        Dim vSrv As Variant
                        For Each vSrv In allSrvsD.Keys
                            Dim kSDD As String
                            kSDD = pep & "|" & CStr(vSrv)
                            If dictQtd.Exists(kSDD) Then qtdSrvDD = qtdSrvDD + dictQtd(kSDD)
                        Next vSrv

                        If Round(qtdMatDD, 4) = Round(qtdSrvDD, 4) Then
                            .Aderencia = "ADERENTE"
                        Else
                            .Aderencia = "QTD DIVERGENTE"
                        End If
                    End If
                End If
            End If
        End With
NextAd:
    Next i
End Sub

' ============================================================
' 9. CALCULAR TIPO CUSTO
' ============================================================
Private Sub CalcularTipoCusto()
    Dim i As Long
    For i = 0 To nItens - 1
        With aItens(i)
            If .Cls1 = "SERVI" & Chr(199) & "O" Or .Cls1 = "SERVICO" Or .Cls1 = "SERVICO" Then
                .TipoCusto = GetTipoServico(.Material)
            Else
                .TipoCusto = ""
            End If
        End With
    Next i
End Sub

' ============================================================
' 10. CALCULAR %MOP
' ============================================================
Private Sub CalcularPctMop()
    Dim dictMop  As Object, dictSrv2 As Object
    Set dictMop  = CreateObject("Scripting.Dictionary")
    Set dictSrv2 = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 0 To nItens - 1
        With aItens(i)
            If .Cls2 = "MOP" Then
                If dictMop.Exists(.Pep) Then
                    dictMop(.Pep) = dictMop(.Pep) + .ValorMoeda
                Else
                    dictMop(.Pep) = .ValorMoeda
                End If
            End If
            If .Cls1 = "SERVI" & Chr(199) & "O" Or .Cls1 = "SERVICO" Or .Cls1 = "SERVICO" Then
                If dictSrv2.Exists(.Pep) Then
                    dictSrv2(.Pep) = dictSrv2(.Pep) + .ValorMoeda
                Else
                    dictSrv2(.Pep) = .ValorMoeda
                End If
            End If
        End With
    Next i

    For i = 0 To nItens - 1
        With aItens(i)
            If dictMop.Exists(.Pep) And dictSrv2.Exists(.Pep) Then
                If dictSrv2(.Pep) <> 0 Then
                    .PctMop = dictMop(.Pep) / dictSrv2(.Pep)
                End If
            End If
        End With
    Next i
End Sub

' ============================================================
' 11. ORDENAR POR GRUPO
' ============================================================
Private Sub OrdenarPorGrupo()
    Dim i As Long
    For i = 0 To nItens - 1
        Dim code As String
        code = aItens(i).Material
        If EhAutoCorrespondente(code) Then
            aItens(i).GrupoKey = code
        ElseIf EhNaCorresp(code) Then
            aItens(i).GrupoKey = GetGrupoKey(code)
        Else
            aItens(i).GrupoKey = ""
        End If
    Next i
    If nItens < 2 Then Exit Sub

    ' Chave composta (PEP > GRUPO > TIPO_ORDEM > CLS2 > MATERIAL) + quicksort
    ' indexado. Substitui o bubble sort O(n^2) que travaria com muitas linhas.
    Dim idx() As Long, keys() As String
    ReDim idx(0 To nItens - 1)
    ReDim keys(0 To nItens - 1)
    For i = 0 To nItens - 1
        idx(i) = i
        Dim g As String
        g = aItens(i).GrupoKey
        If g = "" Then g = "~~~~~~~~~~"
        keys(i) = aItens(i).Pep & Chr(9) & g & Chr(9) & _
                  CStr(TipoOrdem(aItens(i))) & Chr(9) & _
                  aItens(i).Cls2 & Chr(9) & aItens(i).Material
    Next i
    QuickSortIdx keys, idx, 0, nItens - 1

    Dim aOrd() As tItem
    ReDim aOrd(nItens)
    For i = 0 To nItens - 1
        aOrd(i) = aItens(idx(i))
    Next i
    For i = 0 To nItens - 1
        aItens(i) = aOrd(i)
    Next i
End Sub

Private Sub QuickSortIdx(keys() As String, idx() As Long, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, tmp As Long
    Dim pivot As String
    i = lo: j = hi
    pivot = keys(idx((lo + hi) \ 2))
    Do While i <= j
        Do While keys(idx(i)) < pivot
            i = i + 1
        Loop
        Do While keys(idx(j)) > pivot
            j = j - 1
        Loop
        If i <= j Then
            tmp = idx(i): idx(i) = idx(j): idx(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSortIdx keys, idx, lo, j
    If i < hi Then QuickSortIdx keys, idx, i, hi
End Sub

Private Function DeveOrdenar(a As tItem, b As tItem) As Boolean
    ' PEP
    If a.Pep > b.Pep Then DeveOrdenar = True: Exit Function
    If a.Pep < b.Pep Then DeveOrdenar = False: Exit Function
    ' Grupo: sem grupo vai para o final
    Dim aGrp As String, bGrp As String
    aGrp = IIf(a.GrupoKey = "", "~", a.GrupoKey)
    bGrp = IIf(b.GrupoKey = "", "~", b.GrupoKey)
    If aGrp > bGrp Then DeveOrdenar = True: Exit Function
    If aGrp < bGrp Then DeveOrdenar = False: Exit Function
    ' Dentro do grupo: MATERIAL(0) > SERVICO D(1) > SERVICO C(2) > outros(3)
    If TipoOrdem(a) > TipoOrdem(b) Then DeveOrdenar = True: Exit Function
    If TipoOrdem(a) < TipoOrdem(b) Then DeveOrdenar = False: Exit Function
    ' CLS2
    If a.Cls2 > b.Cls2 Then DeveOrdenar = True: Exit Function
    If a.Cls2 < b.Cls2 Then DeveOrdenar = False: Exit Function
    ' MATERIAL
    If a.Material > b.Material Then DeveOrdenar = True: Exit Function
    DeveOrdenar = False
End Function

Private Function TipoOrdem(a As tItem) As Integer
    If a.Cls1 = "MATERIAL" Then TipoOrdem = 0: Exit Function
    If a.TipoCusto = "D" Then TipoOrdem = 1: Exit Function
    If a.TipoCusto = "C" Then TipoOrdem = 2: Exit Function
    TipoOrdem = 3
End Function

' ============================================================
' 12. ESCREVER ABA MAT vs SERV AT
' ============================================================
Private Sub EscreverAbaAT()
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ActiveWorkbook.Sheets("MAT vs SERV AT").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = ActiveWorkbook.Sheets.Add(Before:=ActiveWorkbook.Sheets(1))
    ws.Name = "MAT vs SERV AT"
    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ' Cabecalho
    Dim headers As Variant
    headers = Array("EMPRESA", "SEGMENTO", "TIPO OBRA ANEEL", "PEP 3 NIVEL", "PEP", "TIPO", _
                    "MATERIAL", "TEXTO MATERIAL", "UML", "VALOR MOEDA", "QTD ENTRADA", _
                    "CLS1", "CLS2", "TIPO CUSTO", "MAT", "SRV", "ADERENCIA MAT X SERV", _
                    "INCONFORMIDADE", "%MOP")

    Dim nCols As Integer
    nCols = UBound(headers) + 1

    Dim c As Integer
    For c = 1 To nCols
        With ws.Cells(1, c)
            .Value = headers(c - 1)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Font.Size = 9.5
            .Font.Name = FONTE_UI
            .Interior.Color = COR_UI_HDR
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .IndentLevel = 1
            .WrapText = True
        End With
    Next c
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, nCols))
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeBottom).Color = COR_UI_HDR_ACENTO
    End With
    ws.Rows(1).RowHeight = 34

    ' Dados
    Dim prevPep As String, prevGrp As String, grpIdx As Integer
    prevPep = "": prevGrp = "": grpIdx = 0
    Dim bgColor As Long

    Dim i As Long
    For i = 0 To nItens - 1
        Dim r As Long
        r = i + 2

        With aItens(i)
            ' Cor de fundo por grupo
            If .Pep <> prevPep Then
                prevPep = .Pep
                prevGrp = .GrupoKey
                grpIdx = 0
            ElseIf .GrupoKey <> prevGrp Then
                prevGrp = .GrupoKey
                grpIdx = 1 - grpIdx
            End If

            If .GrupoKey <> "" Then
                bgColor = IIf(grpIdx = 0, COR_GRUPO_A, COR_GRUPO_B)
            Else
                bgColor = IIf(i Mod 2 = 0, COR_SEM_GRUPO_A, COR_GRUPO_B)
            End If

            ' Escrever celulas
            ws.Cells(r, 1).Value  = .Empresa
            ws.Cells(r, 2).Value  = .Segmento
            ws.Cells(r, 3).Value  = .TipoObraAneel
            ws.Cells(r, 4).Value  = .Pep3Nivel
            ws.Cells(r, 5).Value  = .Pep
            ws.Cells(r, 6).Value  = .Tipo
            ws.Cells(r, 7).Value  = .Material
            ws.Cells(r, 8).Value  = .TextoMaterial
            ws.Cells(r, 9).Value  = .Uml
            ws.Cells(r, 10).Value = .ValorMoeda
            ws.Cells(r, 11).Value = .QtdEntrada
            ws.Cells(r, 12).Value = .Cls1
            ws.Cells(r, 13).Value = .Cls2
            ws.Cells(r, 14).Value = .TipoCusto
            ws.Cells(r, 15).Value = IIf(.Mat = 0, Empty, .Mat)
            ws.Cells(r, 16).Value = IIf(.Srv = 0, Empty, .Srv)
            ws.Cells(r, 17).Value = .Aderencia
            ws.Cells(r, 18).Value = .Inconformidade
            If .PctMop <> 0 Then ws.Cells(r, 19).Value = .PctMop

            ' Formatar linha
            Dim rng As Range
            Set rng = ws.Range(ws.Cells(r, 1), ws.Cells(r, nCols))
            rng.Interior.Color = bgColor
            rng.Font.Size = 9
            rng.Font.Name = FONTE_UI
            rng.Font.Color = COR_UI_TINTA
            rng.VerticalAlignment = xlCenter
            rng.RowHeight = 17

            ' Formatos especificos
            ws.Cells(r, 10).NumberFormat = "#,##0.00"
            ws.Cells(r, 10).HorizontalAlignment = xlRight
            ws.Cells(r, 11).NumberFormat = "#,##0.00"
            ws.Cells(r, 11).HorizontalAlignment = xlRight
            ws.Cells(r, 15).NumberFormat = "#,##0.00"
            ws.Cells(r, 15).HorizontalAlignment = xlRight
            ws.Cells(r, 16).NumberFormat = "#,##0.00"
            ws.Cells(r, 16).HorizontalAlignment = xlRight
            ws.Cells(r, 19).NumberFormat = "0.0%"
            ws.Cells(r, 19).HorizontalAlignment = xlCenter

            ' Centralizados
            ws.Cells(r, 1).HorizontalAlignment = xlCenter
            ws.Cells(r, 6).HorizontalAlignment = xlCenter
            ws.Cells(r, 9).HorizontalAlignment = xlCenter
            ws.Cells(r, 12).HorizontalAlignment = xlCenter

            ' TIPO CUSTO
            If .TipoCusto = "D" Then
                ws.Cells(r, 14).Interior.Color = COR_TIPO_D_BG
                ws.Cells(r, 14).Font.Color = COR_TIPO_D_FG
                ws.Cells(r, 14).Font.Bold = True
                ws.Cells(r, 14).HorizontalAlignment = xlCenter
            ElseIf .TipoCusto = "C" Then
                ws.Cells(r, 14).Interior.Color = COR_TIPO_C_BG
                ws.Cells(r, 14).Font.Color = COR_TIPO_C_FG
                ws.Cells(r, 14).Font.Bold = True
                ws.Cells(r, 14).HorizontalAlignment = xlCenter
            End If

            ' ADERENCIA
            Select Case .Aderencia
                Case "ADERENTE"
                    ws.Cells(r, 17).Interior.Color = COR_UI_OK_BG
                    ws.Cells(r, 17).Font.Color = COR_UI_OK_FG
                    ws.Cells(r, 17).Font.Bold = True
                Case "QTD DIVERGENTE"
                    ws.Cells(r, 17).Interior.Color = COR_UI_WARN_BG
                    ws.Cells(r, 17).Font.Color = COR_UI_WARN_FG
                    ws.Cells(r, 17).Font.Bold = True
                Case "SEM SERVI" & Chr(199) & "O CORRESPONDENTE", "SEM MATERIAL CORRESPONDENTE"
                    ws.Cells(r, 17).Interior.Color = COR_UI_BAD_BG
                    ws.Cells(r, 17).Font.Color = COR_UI_BAD_FG
                    ws.Cells(r, 17).Font.Bold = True
            End Select
            ws.Cells(r, 17).HorizontalAlignment = xlCenter

            ' INCONFORMIDADE
            If .Inconformidade <> "" Then
                ws.Cells(r, 18).Interior.Color = COR_INCONF_BG
                ws.Cells(r, 18).Font.Color = COR_INCONF_FG
                ws.Cells(r, 18).Font.Bold = True
            End If
        End With
    Next i

    ' Bordas
    Dim dataRng As Range
    Set dataRng = ws.Range(ws.Cells(1, 1), ws.Cells(nItens + 1, nCols))
    With dataRng.Borders
        .LineStyle = xlContinuous
        .Color = COR_UI_BORDA
        .Weight = xlHairline
    End With
    dataRng.Borders(xlInsideHorizontal).Color = COR_UI_GRADE

    ' Auto-fit largura
    ws.Columns.AutoFit
    ' Limitar largura maxima
    For c = 1 To nCols
        If ws.Columns(c).ColumnWidth > 50 Then ws.Columns(c).ColumnWidth = 50
    Next c

    ' Freeze e filtro
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Range(ws.Cells(1, 1), ws.Cells(1, nCols)).AutoFilter
    ActiveWindow.Zoom = 90
    On Error Resume Next
    ws.Tab.Color = CorAba("MATERIAL VS SERVICO")
    On Error GoTo 0
    PrepararImpressao ws, "MAT vs SERV AT", nItens + 1, nCols
End Sub

' ============================================================
' FUNCOES AUXILIARES
' ============================================================
Private Function CleanCod(v As String) As String
    v = Trim(v)
    If v = "" Or v = "nan" Or LCase(v) = "none" Then CleanCod = "": Exit Function
    ' Remover .0 de numeros
    If Right(v, 2) = ".0" Then v = Left(v, Len(v) - 2)
    CleanCod = v
End Function

Private Function TemSaldo(val As Double, qtd As Double) As Boolean
    TemSaldo = (val <> 0 And qtd <> 0)
End Function

Private Function ContemPalavra(txt As String, palavra As String) As Boolean
    ContemPalavra = (InStr(1, txt, palavra, vbTextCompare) > 0)
End Function

Private Function EhAutoCorrespondente(code As String) As Boolean
    Dim j As Integer
    For j = 0 To nCorr - 1
        If aMatCorr(j) = code And aSrvCorr(j) = code Then
            EhAutoCorrespondente = True: Exit Function
        End If
    Next j
    EhAutoCorrespondente = False
End Function

Private Function EhNaCorresp(code As String) As Boolean
    Dim j As Integer
    For j = 0 To nCorr - 1
        If aMatCorr(j) = code Or aSrvCorr(j) = code Then
            EhNaCorresp = True: Exit Function
        End If
    Next j
    EhNaCorresp = False
End Function

Private Function GetTipoServico(code As String) As String
    Dim temD As Boolean, temC As Boolean
    temD = False: temC = False
    Dim j As Integer
    For j = 0 To nCorr - 1
        If aSrvCorr(j) = code Then
            If aTipoCorr(j) = "D" Then temD = True
            If aTipoCorr(j) = "C" Then temC = True
        End If
    Next j
    If temD Then GetTipoServico = "D": Exit Function
    If temC Then GetTipoServico = "C": Exit Function
    GetTipoServico = ""
End Function

Private Function GetGrupoKey(code As String) As String
    ' BFS para encontrar todos os codigos do grupo e retornar o minimo
    Dim visited As Object
    Set visited = CreateObject("Scripting.Dictionary")
    Dim queue() As String
    ReDim queue(nCorr * 2)
    Dim head As Integer, tail As Integer
    head = 0: tail = 0
    queue(tail) = code: tail = tail + 1
    visited(code) = 1

    Do While head < tail
        Dim cur As String
        cur = queue(head): head = head + 1
        Dim j As Integer
        For j = 0 To nCorr - 1
            If aMatCorr(j) = cur And Not visited.Exists(aSrvCorr(j)) Then
                visited(aSrvCorr(j)) = 1
                queue(tail) = aSrvCorr(j): tail = tail + 1
            End If
            If aSrvCorr(j) = cur And Not visited.Exists(aMatCorr(j)) Then
                visited(aMatCorr(j)) = 1
                queue(tail) = aMatCorr(j): tail = tail + 1
            End If
        Next j
    Loop

    Dim minKey As String
    minKey = code
    Dim vk As Variant
    For Each vk In visited.Keys
        If CStr(vk) < minKey Then minKey = CStr(vk)
    Next vk
    GetGrupoKey = minKey
End Function

Private Function PepExisteComSufixo(pep3 As String, sufixo As String) As Boolean
    Dim i As Long
    For i = 0 To nItens - 1
        If aItens(i).Pep3Nivel = pep3 And UCase(Right(aItens(i).Pep, 1)) = UCase(sufixo) Then
            PepExisteComSufixo = True: Exit Function
        End If
    Next i
    PepExisteComSufixo = False
End Function

Private Function PepTemMob(pep3 As String, sufixo As String) As Boolean
    Dim i As Long
    For i = 0 To nItens - 1
        With aItens(i)
            If .Pep3Nivel = pep3 And UCase(Right(.Pep, 1)) = UCase(sufixo) Then
                Dim txt As String
                txt = UCase(.TextoMaterial)
                If ContemPalavra(txt, "MOB") Or ContemPalavra(txt, "MOBILIZAR") Or _
                   ContemPalavra(txt, "DESMOB") Or ContemPalavra(txt, "MOBILIZACAO") Then
                    If TemSaldo(.ValorMoeda, .QtdEntrada) Then
                        PepTemMob = True: Exit Function
                    End If
                End If
            End If
        End With
    Next i
    PepTemMob = False
End Function

' ============================================================
' 13. CRIAR ABA PREMISSAS
' ============================================================
Private Sub CriarPremissas()
    Dim ws As Worksheet

    On Error Resume Next
    Application.DisplayAlerts = False
    ActiveWorkbook.Sheets("PREMISSAS").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets("MAT vs SERV AT"))
    ws.Name = "PREMISSAS"
    ws.Activate
    ActiveWindow.DisplayGridlines = False

    Const COR_HDR     As Long = &H660301
    Const COR_WARN_BG As Long = &HD6E4FC
    Const COR_WARN_FG As Long = &H115AC5
    Const COR_OK_BG   As Long = &HCEEFCE
    Const COR_OK_FG   As Long = &H375623

    ws.Columns(1).ColumnWidth = 5
    ws.Columns(2).ColumnWidth = 40
    ws.Columns(3).ColumnWidth = 58
    ws.Columns(4).ColumnWidth = 38
    ws.Columns(5).ColumnWidth = 22

    ' TITULO
    ws.Range("A1:E1").Merge
    With ws.Range("A1")
        .Value = "PREMISSAS DE VALIDACAO - ABA MAT vs SERV AT"
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Font.Size = 14: .Font.Name = FONTE_UI
        .Interior.Color = COR_UI_HDR
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 32

    ws.Range("A2:E2").Merge
    With ws.Range("A2")
        .Value = "Regras aplicadas automaticamente nas colunas INCONFORMIDADE e ADERENCIA MAT X SERV"
        .Font.Italic = True: .Font.Color = RGB(85, 85, 85)
        .Font.Size = 9: .Font.Name = FONTE_UI
        .Interior.Color = RGB(242, 244, 248)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    ws.Rows(2).RowHeight = 16
    ws.Rows(3).RowHeight = 8

    ' SECAO 1: REGRAS DE INCONFORMIDADE
    Call SecaoTitulo(ws, 4, "REGRAS DE INCONFORMIDADE (coluna INCONFORMIDADE)")
    Call TabelaCabecalho(ws, 5, Array("Nr", "REGRA", "CRITERIO / DESCRICAO", "MENSAGEM", "APLICA-SE A"))

    Dim reg(7, 4) As String
    reg(0, 0) = "R1": reg(0, 1) = "Servico de retirada/demolicao em PEP de instalacao"
    reg(0, 2) = "Servicos com retirada, remocao, demolicao, desmontagem, desinstalacao ou desativacao nao sao permitidos em PEPs != .D. Condicao: com saldo."
    reg(0, 3) = "Servico de [acao] indevido em PEP de instalacao"
    reg(0, 4) = "SERVICO / PEP != .D / Com saldo"

    reg(1, 0) = "R2": reg(1, 1) = "Servico de substituicao"
    reg(1, 2) = "Servicos com texto contendo 'substituicao', 'substituir' ou 'SUBST' caracterizam manutencao corretiva (OPEX) e nao devem compor o custo de investimento. Condicao: com saldo apos agrupamento."
    reg(1, 3) = "Servico de substituicao deve ser reclassificado para o OPEX"
    reg(1, 4) = "SERVICO / Todos os PEPs / Com saldo"

    reg(2, 0) = "R3": reg(2, 1) = "Servico com saldo negativo"
    reg(2, 2) = "Apos agrupamento, VALOR MOEDA ou QTD ENTRADA negativos sao sinalizados. Valores < 1e-9 sao tratados como zero."
    reg(2, 3) = "Servico com saldo negativo"
    reg(2, 4) = "SERVICO / Todos os PEPs / Saldo < 0"

    reg(3, 0) = "R4": reg(3, 1) = "Servico generico (descricao vaga)"
    reg(3, 2) = "Servicos que nao identificam o objeto: Trabalhos Excepcionais, Materiais Menores, Adm. Obra, MO Administracao, Fiscalizacao Engenharia, Consultoria, Diferenca Fator K, Locacao Veiculos, Transporte Equipamentos Menores, Limpeza, Exec Servico Integracao Automacao. Condicao: com saldo."
    reg(3, 3) = "Servico generico ([motivo])"
    reg(3, 4) = "SERVICO / Todos os PEPs / Com saldo"

    reg(4, 0) = "R5": reg(4, 1) = "Mobilizacao sem correspondencia no PEP par"
    reg(4, 2) = "Quando PEP 3 NIVEL possui .D e .I, custos de mob/desmob devem estar em ambos. So se aplica quando a obra tem os dois sufixos. Condicao: com saldo."
    reg(4, 3) = "Custo de mob. nao identificado no PEP .D" & Chr(10) & "Custo de mob. nao identificado no PEP .I"
    reg(4, 4) = "SERVICO / PEP 3 NIVEL com .D e .I / Com saldo"

    reg(5, 0) = "R6": reg(5, 1) = "Custo de viagem, hospedagem ou alimentacao"
    reg(5, 2) = "Itens com hospedagem, alimentacao, viagem, passagem, mobilidade ou conducao no texto/CLS2 devem ser reclassificados para OPEX. Excecao: 'caixa de passagem'. Condicao: VALOR MOEDA <> 0."
    reg(5, 3) = "Custo de viagem deve ser reclassificado para Opex"
    reg(5, 4) = "Todos / Todos os PEPs / Valor <> 0"

    reg(6, 0) = "R7": reg(6, 1) = "Material com saldo negativo em PEP de instalacao"
    reg(6, 2) = "Materiais com saldo negativo so sao admitidos em PEPs '.D'. Em PEPs de instalacao indica lancamento incorreto."
    reg(6, 3) = "Material com saldo negativo indevido em PEP de instalacao"
    reg(6, 4) = "MATERIAL / PEP != .D / Saldo < 0"

    reg(7, 0) = "R8": reg(7, 1) = "Servico de poda em obra de manutencao"
    reg(7, 2) = "Servicos com 'poda' no texto em PEPs '.I' do segmento MANUTENCAO AT sao OPEX. Condicao: VALOR MOEDA <> 0."
    reg(7, 3) = "Servico de poda e indevido em obra de manutencao deve ser reclassificado para Opex"
    reg(7, 4) = "SERVICO / PEP .I / Seg. MANUTENCAO AT / Valor <> 0"

    Dim r As Integer
    For r = 0 To 7
        Call LinhaDados(ws, r + 6, reg(r, 0), reg(r, 1), reg(r, 2), reg(r, 3), reg(r, 4), _
            COR_HDR, IIf(r Mod 2 = 0, RGB(242, 244, 248), RGB(255, 255, 255)), _
            COR_WARN_BG, COR_WARN_FG, 52)
    Next r
    Call AplicarBordas(ws, 5, 13, 5)

    ' SECAO 2: LOGICA DE ADERENCIA
    ws.Rows(14).RowHeight = 10
    Call SecaoTitulo(ws, 15, "LOGICA DE ADERENCIA MAT X SERV (coluna ADERENCIA MAT X SERV)")
    Call TabelaCabecalho(ws, 16, Array("Tipo", "DESCRICAO", "CRITERIO", "RESULTADO", "COR"))

    Dim adh(4, 4) As String
    adh(0, 0) = "D - Direto"
    adh(0, 1) = "Custo direto de instalacao (coluna TIPO CUSTO = D)"
    adh(0, 2) = "Soma QTD materiais do grupo D == Soma QTD servicos D, no mesmo PEP. Para PEPs '.D' aplica abs() no saldo total da familia."
    adh(0, 3) = "ADERENTE / QTD DIVERGENTE / SEM SERVICO CORRESPONDENTE"
    adh(0, 4) = "Azul (DDEBF7)"

    adh(1, 0) = "C - Complementar"
    adh(1, 1) = "Custo complementar (coluna TIPO CUSTO = C)"
    adh(1, 2) = "QTD do servico C == QTD do material correspondente (verificacao individual). Nao entra na soma dos servicos D."
    adh(1, 3) = "ADERENTE / QTD DIVERGENTE / SEM MATERIAL CORRESPONDENTE"
    adh(1, 4) = "Verde (E2EFDA)"

    adh(2, 0) = "Auto-corresp."
    adh(2, 1) = "MATERIAL = SERVICO na aba CORRESP"
    adh(2, 2) = "Aderente se tiver saldo (QTD <> 0). Sem verificacao de par."
    adh(2, 3) = "ADERENTE"
    adh(2, 4) = "Verde (C6EFCE)"

    adh(3, 0) = "Sem par"
    adh(3, 1) = "Na CORRESP mas sem correspondente no PEP"
    adh(3, 2) = "Material existe na CORRESP mas nenhum servico correspondente foi encontrado no PEP (ou vice-versa)."
    adh(3, 3) = "SEM SERVICO CORRESPONDENTE" & Chr(10) & "SEM MATERIAL CORRESPONDENTE"
    adh(3, 4) = "Vermelho (FFC7CE)"

    adh(4, 0) = "Sem saldo"
    adh(4, 1) = "VALOR = 0 E QTD = 0 apos agrupamento"
    adh(4, 2) = "Nenhuma aderencia e calculada e nenhuma inconformidade e atribuida."
    adh(4, 3) = "(em branco)"
    adh(4, 4) = "Sem cor"

    Dim a As Integer
    For a = 0 To 4
        Call LinhaDados(ws, a + 17, adh(a, 0), adh(a, 1), adh(a, 2), adh(a, 3), adh(a, 4), _
            COR_HDR, IIf(a Mod 2 = 0, RGB(242, 244, 248), RGB(255, 255, 255)), _
            IIf(a < 3, COR_OK_BG, IIf(a = 3, &HD6E4FC, RGB(240, 240, 240))), _
            IIf(a < 3, COR_OK_FG, IIf(a = 3, &H115AC5, RGB(100, 100, 100))), 44)
    Next a
    Call AplicarBordas(ws, 16, 21, 5)

    ' SECAO 3: TRATAMENTOS GERAIS
    ws.Rows(22).RowHeight = 10
    Call SecaoTitulo(ws, 23, "TRATAMENTOS GERAIS APLICADOS AOS DADOS")
    ws.Cells(24, 2).Value = "TRATAMENTO"
    ws.Cells(24, 3).Value = "DESCRICAO"
    ws.Range("C24:E24").Merge
    Call TabelaCabecalho(ws, 24, Array("", "TRATAMENTO", "DESCRICAO", "", ""))

    Dim trat(8, 1) As String
    trat(0, 0) = "Codigo vazio":       trat(0, 1) = "MATERIAL vazio e preenchido com CLASSE_CUSTO."
    trat(1, 0) = "Descricao vazia":    trat(1, 1) = "TEXTO MATERIAL vazio e preenchido com DESC_CLASSE_CUSTO."
    trat(2, 0) = "Sufixo .0":          trat(2, 1) = "Codigos numericos lidos como float removem o '.0' (ex: 103120010.0 -> 103120010)."
    trat(3, 0) = "UML vazio":          trat(3, 1) = "Lancamentos com UML vazio herdam o UML de outro lancamento do mesmo codigo antes do agrupamento."
    trat(4, 0) = "Agrupamento":        trat(4, 1) = "Multiplos lancamentos do mesmo item (MATERIAL+TEXTO+PEP+CLS1+CLS2+UML) sao agrupados somando VALOR MOEDA e QTD ENTRADA."
    trat(5, 0) = "Near-zero":          trat(5, 1) = "Valores com modulo < 1e-9 apos agrupamento sao convertidos para zero."
    trat(6, 0) = "Inconformidade zero": trat(6, 1) = "Itens com VALOR=0 E QTD=0 apos agrupamento nao recebem mensagem de inconformidade."
    trat(7, 0) = "Sinal .D":           trat(7, 1) = "Para materiais em PEPs '.D', abs() e aplicado no saldo TOTAL da familia (nao por lancamento)."
    trat(8, 0) = "CLS2 padronizado":   trat(8, 1) = "CLS2 dos servicos e padronizado para o CLS2 do material (exibicao). MAT usa CLS2 exibido; SRV usa CLS2 original do servico."

    Dim t As Integer
    For t = 0 To 8
        Dim rowT As Integer: rowT = t + 25
        Dim bgT As Long: bgT = IIf(t Mod 2 = 0, RGB(242, 244, 248), RGB(255, 255, 255))
        ws.Cells(rowT, 1).Value = Chr(149): ws.Cells(rowT, 1).Font.Color = COR_HDR
        ws.Cells(rowT, 1).Font.Bold = True: ws.Cells(rowT, 1).HorizontalAlignment = xlCenter
        ws.Cells(rowT, 1).Interior.Color = bgT
        ws.Cells(rowT, 2).Value = trat(t, 0): ws.Cells(rowT, 2).Font.Bold = True
        ws.Cells(rowT, 2).Font.Size = 8: ws.Cells(rowT, 2).Font.Name = FONTE_UI
        ws.Cells(rowT, 2).Interior.Color = bgT: ws.Cells(rowT, 2).VerticalAlignment = xlCenter
        ws.Range(ws.Cells(rowT, 3), ws.Cells(rowT, 5)).Merge
        ws.Cells(rowT, 3).Value = trat(t, 1): ws.Cells(rowT, 3).Font.Size = 8
        ws.Cells(rowT, 3).Font.Name = FONTE_UI: ws.Cells(rowT, 3).Interior.Color = bgT
        ws.Cells(rowT, 3).VerticalAlignment = xlCenter: ws.Cells(rowT, 3).WrapText = True
        ws.Rows(rowT).RowHeight = 24
    Next t
    Call AplicarBordas(ws, 24, 33, 5)

    ' SECAO 4: DESCRICAO DAS COLUNAS
    ws.Rows(34).RowHeight = 10
    Call SecaoTitulo(ws, 35, "DESCRICAO DAS COLUNAS - ABA MAT vs SERV AT")
    Call TabelaCabecalho(ws, 36, Array("Nr", "COLUNA", "DESCRICAO", "", ""))

    Dim cols(18, 1) As String
    cols(0,  0) = "EMPRESA":              cols(0,  1) = "Codigo da empresa responsavel pela obra."
    cols(1,  0) = "SEGMENTO":             cols(1,  1) = "Segmento da obra (ex: EXPANSAO AT, MANUTENCAO AT)."
    cols(2,  0) = "TIPO OBRA ANEEL":      cols(2,  1) = "Tipo da obra conforme classificacao ANEEL."
    cols(3,  0) = "PEP 3 NIVEL":          cols(3,  1) = "PEP ate o terceiro nivel (sem sufixo .D/.I/.E)."
    cols(4,  0) = "PEP":                  cols(4,  1) = "Codigo completo do PEP com sufixo."
    cols(5,  0) = "TIPO":                 cols(5,  1) = "Sufixo do PEP: D (desativacao), I (instalacao), E (especial)."
    cols(6,  0) = "MATERIAL":             cols(6,  1) = "Codigo do material ou classe de custo."
    cols(7,  0) = "TEXTO MATERIAL":       cols(7,  1) = "Descricao do material ou classe de custo."
    cols(8,  0) = "UML":                  cols(8,  1) = "Unidade de medida. UML vazio herda o de outro lancamento do mesmo codigo."
    cols(9,  0) = "VALOR MOEDA":          cols(9,  1) = "Valor financeiro total apos agrupamento."
    cols(10, 0) = "QTD ENTRADA":          cols(10, 1) = "Quantidade total apos agrupamento."
    cols(11, 0) = "CLS1":                 cols(11, 1) = "Classificacao nivel 1: MATERIAL, SERVICO ou OUTROS."
    cols(12, 0) = "CLS2":                 cols(12, 1) = "Familia do item. Para servicos, padronizado para o CLS2 do material correspondente."
    cols(13, 0) = "TIPO CUSTO":           cols(13, 1) = "D = custo direto (azul). C = custo complementar (verde). Vazio = nao esta na CORRESP."
    cols(14, 0) = "MAT":                  cols(14, 1) = "Soma QTD de materiais da familia CLS2 no PEP (do RAZAO CJ)."
    cols(15, 0) = "SRV":                  cols(15, 1) = "Soma QTD de servicos tipo D da familia no PEP (do RAZAO CJ, usando CLS2 original)."
    cols(16, 0) = "ADERENCIA MAT X SERV": cols(16, 1) = "Resultado da verificacao de aderencia conforme aba CORRESP."
    cols(17, 0) = "INCONFORMIDADE":       cols(17, 1) = "Mensagem da regra R1-R8 quando identificado problema. Vazio se sem saldo ou sem inconformidade."
    cols(18, 0) = "%MOP":                 cols(18, 1) = "Percentual MOP do PEP: Soma VALOR (CLS2=MOP) / Soma VALOR (CLS1=SERVICO)."

    Dim cl As Integer
    For cl = 0 To 18
        Dim rowC As Integer: rowC = cl + 37
        Dim bgC As Long: bgC = IIf(cl Mod 2 = 0, RGB(242, 244, 248), RGB(255, 255, 255))
        ws.Cells(rowC, 1).Value = cl + 1: ws.Cells(rowC, 1).Font.Bold = True
        ws.Cells(rowC, 1).Font.Color = RGB(255, 255, 255): ws.Cells(rowC, 1).Interior.Color = COR_HDR
        ws.Cells(rowC, 1).HorizontalAlignment = xlCenter: ws.Cells(rowC, 1).VerticalAlignment = xlCenter
        ws.Cells(rowC, 2).Value = cols(cl, 0): ws.Cells(rowC, 2).Font.Bold = True
        ws.Cells(rowC, 2).Font.Size = 8: ws.Cells(rowC, 2).Font.Name = FONTE_UI
        ws.Cells(rowC, 2).Interior.Color = bgC: ws.Cells(rowC, 2).VerticalAlignment = xlCenter
        ws.Range(ws.Cells(rowC, 3), ws.Cells(rowC, 5)).Merge
        ws.Cells(rowC, 3).Value = cols(cl, 1): ws.Cells(rowC, 3).Font.Size = 8
        ws.Cells(rowC, 3).Font.Name = FONTE_UI: ws.Cells(rowC, 3).Interior.Color = bgC
        ws.Cells(rowC, 3).VerticalAlignment = xlCenter: ws.Cells(rowC, 3).WrapText = True
        ws.Rows(rowC).RowHeight = 20
    Next cl
    Call AplicarBordas(ws, 36, 55, 5)

    ws.Range("A1").Select
End Sub

' ?? HELPERS PREMISSAS ????????????????????????????????????????
Private Sub SecaoTitulo(ws As Worksheet, rowN As Integer, titulo As String)
    ws.Range(ws.Cells(rowN, 1), ws.Cells(rowN, 5)).Merge
    With ws.Cells(rowN, 1)
        .Value = titulo
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Font.Size = 10: .Font.Name = FONTE_UI
        .Interior.Color = RGB(46, 117, 182)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    ws.Rows(rowN).RowHeight = 20
End Sub

Private Sub TabelaCabecalho(ws As Worksheet, rowN As Integer, cols As Variant)
    Dim c As Integer
    For c = 1 To 5
        With ws.Cells(rowN, c)
            .Value = cols(c - 1)
            .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
            .Font.Size = 9: .Font.Name = FONTE_UI
            .Interior.Color = RGB(68, 114, 196)
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .WrapText = True
        End With
    Next c
    ws.Rows(rowN).RowHeight = 20
End Sub

Private Sub LinhaDados(ws As Worksheet, rowN As Integer, _
    col1 As String, col2 As String, col3 As String, col4 As String, col5 As String, _
    corNr As Long, bgRow As Long, msgBg As Long, msgFg As Long, altura As Integer)

    With ws.Cells(rowN, 1)
        .Value = col1: .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Font.Size = 9: .Font.Name = FONTE_UI: .Interior.Color = corNr
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    With ws.Cells(rowN, 2)
        .Value = col2: .Font.Bold = True: .Font.Size = 9: .Font.Name = FONTE_UI
        .Interior.Color = bgRow: .VerticalAlignment = xlCenter: .WrapText = True
    End With
    With ws.Cells(rowN, 3)
        .Value = col3: .Font.Size = 8: .Font.Name = FONTE_UI
        .Interior.Color = bgRow: .VerticalAlignment = xlCenter: .WrapText = True
    End With
    With ws.Cells(rowN, 4)
        .Value = col4: .Font.Size = 8: .Font.Bold = True: .Font.Name = FONTE_UI
        .Interior.Color = msgBg: .Font.Color = msgFg
        .VerticalAlignment = xlCenter: .WrapText = True
    End With
    With ws.Cells(rowN, 5)
        .Value = col5: .Font.Size = 8: .Font.Name = FONTE_UI
        .Interior.Color = bgRow: .VerticalAlignment = xlCenter: .WrapText = True
    End With
    ws.Rows(rowN).RowHeight = altura
End Sub

Private Sub AplicarBordas(ws As Worksheet, r1 As Integer, r2 As Integer, c2 As Integer)
    With ws.Range(ws.Cells(r1, 1), ws.Cells(r2, c2)).Borders
        .LineStyle = xlContinuous
        .Color = RGB(217, 217, 217)
        .Weight = xlThin
    End With
End Sub



'==============================================================================
'  ABA: CHECKLIST CKCP  (confirmacao manual das etapas do processo)
'  Macro AVULSA (rode via Alt+F8 > GerarChecklistCKCP). Nao faz parte de
'  GerarRelatorio para nao apagar as marcacoes ao reprocessar a base.
'  Nao-destrutiva: se a aba ja existe, preserva o que foi preenchido.
'==============================================================================
Public Sub GerarChecklistCKCP(Optional ByVal silencioso As Boolean = False)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("CHECKLIST CKCP")
    On Error GoTo 0
    If Not ws Is Nothing Then
        ' Aba ja existe: preserva as marcacoes. Silencioso no fluxo automatico.
        If Not silencioso Then
            ws.Activate
            MsgBox "A aba CHECKLIST CKCP ja existe e suas marcacoes foram mantidas." & vbLf & _
                   "Para recria-la do zero, exclua a aba e rode novamente.", _
                   vbInformation, "Checklist CKCP"
        End If
        Exit Sub
    End If

    Set ws = ActiveWorkbook.Worksheets.Add( _
        After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    ws.Name = "CHECKLIST CKCP"

    Dim corHdr As Long, corBanda As Long
    corHdr = RGB(0, 105, 65): corBanda = RGB(0, 60, 38)

    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ' --- Banda de titulo ------------------------------------------------------
    With ws.Range("A1:H1")
        .Merge: .Interior.Color = corHdr
        .Font.Color = vbWhite: .Font.Bold = True: .Font.Size = 16
        .Font.Name = "Segoe UI": .VerticalAlignment = xlCenter
    End With
    ws.Cells(1, 1).Value = "CHECKLIST CKCP  -  CONFIRMACAO DO PROCESSO"
    ws.Cells(1, 1).IndentLevel = 1
    ws.Rows(1).RowHeight = 38

    ' Linha de apoio (CORTE ID + % concluido) sobre fundo cinza-claro
    With ws.Range("A2:H2")
        .Interior.Color = RGB(240, 244, 242)
    End With
    ws.Cells(2, 1).Value = "CORTE ID:": ws.Cells(2, 1).Font.Bold = True
    ws.Cells(2, 1).Font.Name = "Segoe UI": ws.Cells(2, 1).Font.Size = 10
    With ws.Cells(2, 2)
        .Interior.Color = RGB(255, 249, 196)
        .Font.Bold = True: .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(191, 175, 90)
    End With
    ws.Cells(2, 4).Value = "% CONCLUIDO:": ws.Cells(2, 4).Font.Bold = True
    ws.Cells(2, 4).Font.Name = "Segoe UI": ws.Cells(2, 4).Font.Size = 10
    ws.Cells(2, 4).HorizontalAlignment = xlRight
    ws.Rows(2).RowHeight = 24

    ' Faixa de acento verde-claro sob a banda
    With ws.Range("A3:H3")
        .Interior.Color = RGB(0, 176, 102)
    End With
    ws.Rows(3).RowHeight = 4

    ' --- Cabecalho da tabela --------------------------------------------------
    Dim hdr As Variant, c As Long
    hdr = Array("N", "CATEGORIA", "ETAPA", "ORIENTACAO", "STATUS", "RESPONSAVEL", "DATA", "OBS")
    For c = 0 To UBound(hdr)
        ws.Cells(4, c + 1).Value = hdr(c)
    Next c
    With ws.Range("A4:H4")
        .Interior.Color = corHdr: .Font.Color = vbWhite: .Font.Bold = True
        .Font.Name = "Segoe UI": .Font.Size = 10
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium: .Borders(xlEdgeBottom).Color = corBanda
    End With
    ws.Rows(4).RowHeight = 24

    ' --- Linhas (uma chamada AddChk por etapa; evita limite de continuacao) ---
    Dim linha As Long: linha = 4
    AddChk ws, linha, "ANALISE CKCP", "DADOS CADASTRAIS", "ANALISE CKCP - Check Dados Cadastrais"
    AddChk ws, linha, "ANALISE CKCP", "STATUS DESCASADO", "ANALISE CKCP - Check Status Descasados 3o Nivel"
    AddChk ws, linha, "SAP", "1o CHECK DE COMPROMISSO DE SERVICO", "SAP - TRANSACAO CJI5 - Layout /DANILO"
    AddChk ws, linha, "SAP", "CHECK DE MATERIAL ATRELADO", "SAP - TRANSACAO MBBS - Layout /ESTOQUE"
    AddChk ws, linha, "CUSTO DE OBRAS", "GERAR INVENTARIO DE MATERIAL", "CUSTO DE OBRAS - MODULO FGM - INVENTARIO DE MATERIAL"
    AddChk ws, linha, "SAP", "GERAR ANALISE DE CUSTO", "SAP - TRANSACAO CJI3 - Layout /BPR_RAW - ANALISE DE CUSTO"
    AddChk ws, linha, "ANALISE DE CUSTO", "CLASSIFICAR OS CUSTOS", "ANALISE DE CUSTO - Classificar os custos da base"
    AddChk ws, linha, "ANALISE DE CUSTO", "AJUSTES CONTABIL DE MATERIAL", "ANALISE DE CUSTO - MATERIAL > AJUSTE CONTABIL DE MATERIAL"
    AddChk ws, linha, "ANALISE DE CUSTO", "ANALISE DE CUSTO ""RISCO""", "ANALISE DE CUSTO - RISCO - SAP - TRANSACAO CJI3 > AREA SOLUCIONADORA"
    AddChk ws, linha, "ANALISE DE CUSTO", "CHECK DE VALOR DO SERVICO", "ANALISE DE CUSTO - SERVICO > VERIFICAR DISCREPANCIAS E VALORES NEGATIVOS"
    AddChk ws, linha, "ANALISE DE CUSTO", "CHECK DE APROPRIACAO CORRETA DE SERVICO", "ANALISE DE CUSTO - SERVICO > IDENTIFICAR SE OS SERVICOS DE INSTALACAO ESTAO NA ODI E DE DESATIVACAO NA ODD"
    AddChk ws, linha, "ANALISE DE CUSTO", "ANALISE ADERENCIA MAT X SERV", "ANALISE DE CUSTO - MATERIAL vs SERVICO > ORGANIZAR A PLANILHA PARA MELHOR ANALISE"
    AddChk ws, linha, "ANALISE DE CUSTO", "CHECK DE ATIVACAO DIRETA", "ANALISE DE CUSTO - ANALISE DE CA > CUSTO APENAS NA ODI - SENDO 25% (+- 5%) - SENDO 8% NO IOP"
    AddChk ws, linha, "ANALISE DE CUSTO", "CHECK DE FRETE", "ANALISE DE CUSTO - Verificar frete"
    AddChk ws, linha, "ANALISE DE CUSTO", "CHECK DO PRECO UNITARIO DE MATERIAL", "ANALISE DE CUSTO - PRECO UNITARIO > PRIORIZAR UC/UAR"
    AddChk ws, linha, "ANALISE DE CUSTO", "VERIFICAR VALOR UNITARIO E VALOR DE MATERIAIS", "ANALISE DE CUSTO - MATERIAL > IDENTIFICAR VALOR UNITARIO E VALOR DE MATERIAIS ABAIXO OU ACIMA DO ESPERADO"
    AddChk ws, linha, "ANALISE DE CUSTO", "CHECK PEP SEM UC", "ANALISE DE CUSTO - ANALISE DE CA > REVISAR COLUNA MAT.UC"
    AddChk ws, linha, "ANALISE DE CUSTO", "CHECK DE CUSTO DE SUPORTE", "ANALISE DE CUSTO - RAZAO CJ > VERIFICAR CUSTOS DE SUPORTE (VIAGEM, ALIMENTACAO E HOSPEDAGEM) (2026)"
    AddChk ws, linha, "ANALISE DE CUSTO", "ZERAR MOP", "RETIRAR O MOP PARA REDISTRIBUICAO NAS OBRAS"
    AddChk ws, linha, "ANALISE DE CUSTO", "INSERIR MOP", "TRANSFERIR DA ORDEM DE AJUSTE DE MOP CONFORME PORCENTAGEM ACERTADA * EXCETO ODS PLPT"
    AddChk ws, linha, "CUSTO DE OBRAS", "CHECK OBRAS DEVOLVIDAS PELO CADASTRO", "DEVOLVER AS OBRAS NAO CADASTRADAS"
    AddChk ws, linha, "SAP", "CONFIRMAR DIAGRAMA", "SAP - TRANSACAO CN53N - LSMW - Confirmar diagrama"
    AddChk ws, linha, "SAP", "2o CHECK DE COMPROMISSO DE SERVICO", "SAP - TRANSACAO CJI5 - Layout /DANILO"
    AddChk ws, linha, "SAP", "COLOCAR PEP APROVADO NA EXCECAO", "SAP - TRANSACAO ZCHECK00 - ZCHECK11 - PRIMEIRO NIVEL DO PEP"
    AddChk ws, linha, "SAP", "ALTERACAO DE STATUS", "SAP - LSMW > PS_CJ20N - ENTE_ENTE_2 - PS_STAT - ENTECKCP"
    AddChk ws, linha, "SAP", "CHECK FINAL DOS STATUS", "SAP - TRANSACAO CN43N - Layout /PEPS"
    AddChk ws, linha, "CAPITALIZACAO", "ENVIAR EMAIL PARA ATIVOS IMOB.", "CAPITALIZACAO - CARREGAR INFORMACOES"
    AddChk ws, linha, "SIS.CONFORMIDADE", "CARGA DA BASE NO SIS.CONFORMIDADE", "ANALISE CKCP > CARREGAR BASE DE ANALISE NO SIS"

    Dim ultLin As Long: ultLin = linha
    Dim nItens As Long: nItens = ultLin - 4

    ' % concluido com barra de progresso (data bar) na propria celula
    With ws.Cells(2, 5)
        .Formula = "=COUNTIF(E5:E" & ultLin & ",""CONCLUIDO"")/" & nItens
        .NumberFormat = "0%"
        .Font.Bold = True: .Font.Size = 11: .Font.Color = corHdr
        .HorizontalAlignment = xlCenter
    End With
    ' Best-effort (so formatacao visual): falha aqui nao altera a analise.
    On Error Resume Next
    With ws.Cells(2, 5).FormatConditions.AddDatabar
        .BarColor.Color = COR_UI_BARRA
        .BarFillType = xlDataBarFillSolid
        .BarBorder.Type = xlDataBarBorderNone
        .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
        .MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=1
        .ShowValue = True
    End With
    On Error GoTo 0

    ' --- Validacao (dropdown) na coluna STATUS --------------------------------
    ' Formula1 em VBA usa SEMPRE virgula como separador de lista,
    ' independente do idioma do Excel (usar ";" quebra em pt-BR).
    With ws.Range(ws.Cells(5, 5), ws.Cells(ultLin, 5)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="CONCLUIDO,PENDENTE,N/A"
        .IgnoreBlank = True: .InCellDropdown = True
    End With

    ' --- Formatacao condicional do STATUS -------------------------------------
    Dim rgS As Range: Set rgS = ws.Range(ws.Cells(5, 5), ws.Cells(ultLin, 5))
    rgS.HorizontalAlignment = xlCenter: rgS.Font.Bold = True
    With rgS.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""CONCLUIDO""")
        .Interior.Color = COR_UI_OK_BG: .Font.Color = COR_UI_OK_FG
    End With
    With rgS.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""PENDENTE""")
        .Interior.Color = COR_UI_WARN_BG: .Font.Color = COR_UI_WARN_FG
    End With
    With rgS.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""N/A""")
        .Interior.Color = COR_UI_NEU_BG: .Font.Color = COR_UI_NEU_FG
    End With

    ' --- Corpo: fonte, bordas, alinhamento ------------------------------------
    With ws.Range(ws.Cells(5, 1), ws.Cells(ultLin, 8))
        .Font.Name = FONTE_UI: .Font.Size = 9.5
        .Font.Color = COR_UI_TINTA
        .VerticalAlignment = xlCenter: .WrapText = False
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlHairline: .Borders.Color = COR_UI_BORDA
    End With
    ' Wrap apenas onde o texto e longo (ORIENTACAO e OBS), com altura fixa
    ws.Range(ws.Cells(5, 4), ws.Cells(ultLin, 4)).WrapText = True
    ws.Range(ws.Cells(5, 8), ws.Cells(ultLin, 8)).WrapText = True
    Dim rAlt As Long
    For rAlt = 5 To ultLin
        ws.Rows(rAlt).RowHeight = 26
    Next rAlt
    ' ORIENTACAO em cinza discreto (informacao de apoio)
    With ws.Range(ws.Cells(5, 4), ws.Cells(ultLin, 4)).Font
        .Color = RGB(105, 117, 126): .Size = 8
    End With
    ' N em cinza pequeno, centralizado
    With ws.Range(ws.Cells(5, 1), ws.Cells(ultLin, 1))
        .HorizontalAlignment = xlCenter
        .Font.Color = RGB(150, 158, 165): .Font.Size = 8
    End With
    ' DATA centralizada
    ws.Range(ws.Cells(5, 7), ws.Cells(ultLin, 7)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(1, 1), ws.Cells(ultLin, 8)).BorderAround _
        LineStyle:=xlContinuous, Weight:=xlThin, Color:=corBanda

    ' --- Larguras -------------------------------------------------------------
    ws.Columns(1).ColumnWidth = 4:  ws.Columns(2).ColumnWidth = 18
    ws.Columns(3).ColumnWidth = 40: ws.Columns(4).ColumnWidth = 58
    ws.Columns(5).ColumnWidth = 14: ws.Columns(6).ColumnWidth = 16
    ws.Columns(7).ColumnWidth = 12: ws.Columns(8).ColumnWidth = 26

    ws.Tab.Color = corHdr
    PrepararImpressao ws, "CHECKLIST CKCP", ultLin, 8
    ws.Rows(4).AutoFilter
    ws.Range("A5").Select
    ActiveWindow.FreezePanes = True

    If Not silencioso Then _
        MsgBox "Aba CHECKLIST CKCP criada com " & nItens & " etapas." & vbLf & _
               "Preencha CORTE ID e marque o STATUS de cada etapa (dropdown)." & vbLf & _
               "O % CONCLUIDO atualiza automaticamente.", vbInformation, "Checklist CKCP"
End Sub

' Escreve uma linha do checklist. lin (ByRef) e incrementado a cada chamada.
Private Sub AddChk(ws As Worksheet, ByRef lin As Long, _
                   ByVal categoria As String, ByVal etapa As String, ByVal orientacao As String)
    lin = lin + 1
    Dim n As Long: n = lin - 4
    ' Zebra primeiro (a cor da categoria e aplicada por cima)
    If (n Mod 2) = 0 Then _
        ws.Range(ws.Cells(lin, 1), ws.Cells(lin, 8)).Interior.Color = RGB(246, 249, 247)
    ws.Cells(lin, 1).Value = n
    ws.Cells(lin, 2).Value = categoria
    ws.Cells(lin, 3).Value = etapa
    ws.Cells(lin, 4).Value = orientacao
    ws.Cells(lin, 3).Font.Bold = True
    ws.Cells(lin, 7).NumberFormat = "dd/mm/yyyy"
    ' Badge de categoria: fundo claro + fonte escura por area do processo
    Dim bg As Long, fg As Long
    CorCategoriaChk categoria, bg, fg
    With ws.Cells(lin, 2)
        .Interior.Color = bg
        .Font.Color = fg: .Font.Bold = True: .Font.Size = 8
        .HorizontalAlignment = xlCenter
    End With
End Sub

' Paleta por categoria do checklist (fundo claro / fonte escura).
Private Sub CorCategoriaChk(ByVal categoria As String, ByRef bg As Long, ByRef fg As Long)
    Select Case UCase$(Trim$(categoria))
        Case "ANALISE CKCP"
            bg = RGB(214, 238, 223): fg = RGB(0, 97, 50)        ' verde
        Case "SAP"
            bg = RGB(217, 231, 247): fg = RGB(21, 79, 146)      ' azul
        Case "ANALISE DE CUSTO"
            bg = RGB(214, 236, 238): fg = RGB(11, 107, 115)     ' teal
        Case "CUSTO DE OBRAS"
            bg = RGB(253, 233, 217): fg = RGB(180, 95, 6)       ' laranja
        Case "CAPITALIZACAO"
            bg = RGB(234, 224, 244): fg = RGB(96, 50, 150)      ' roxo
        Case "SIS.CONFORMIDADE"
            bg = RGB(230, 232, 235): fg = RGB(70, 80, 90)       ' cinza
        Case Else
            bg = RGB(240, 244, 242): fg = RGB(60, 72, 66)
    End Select
End Sub

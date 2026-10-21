Attribute VB_Name = "Tripe"
'==============================================================================
'  TRIPE  -  Projeto x Baixa SAP x Servico pago
'  "A obra so fica de pe se as tres pernas baterem."
'
'  Junta as duas macros existentes numa conclusao unica:
'    - inventario.bas  (GerarInventario) -> aba ANALISE SAP x PRJ
'        perna 1: MAT PRJ CAD (projetado)  x  perna 2: MAT LIB SAP (baixado)
'    - AnaliseCKCP.bas (GerarRelatorio)  -> aba MATERIAL vs SERVICO
'        perna 2: material lancado          x  perna 3: SERVICO pago
'  e cruza as duas por PEP4 + FAMILIA.
'
'  PLANILHA PADRAO DE UPLOAD (ex.: TESTE1.xlsx):
'   - 1a aba INVENTARIO : PEP3NIVEL, PEP4NIVEL, NOTA, CLASSE_CUSTO, MATERIAL,
'                         VALOR, DESC_MATERIAL, UNIDADE, ATENDIDO SAP,
'                         TOMBADO PROJ, TIPO, FAMILIA, SITUACAO_DO_MATERIAL
'   - 2a aba RAZAO      : export do SAP (Elemento PEP, Classe de custo,
'                         Material, Qtd.total entrada, Valor/moeda objeto,
'                         Classificacao ...)
'
'  COMO USAR:
'   1. Importe os 3 modulos (inventario.bas, AnaliseCKCP.bas, Tripe.bas) num
'      arquivo de macros .xlsm (ex.: TRIPE.xlsm).
'   2. Abra a planilha de upload (ou deixe que o TRIPE pergunte qual e).
'   3. Alt+F8 > GerarTripe > Executar. Os resultados sao gravados na
'      propria planilha de upload.
'
'  ABAS GERADAS (alem das abas normais de cada macro):
'   - TRIPE CONCLUSAO : 1 linha por obra (PEP3) com o VEREDITO FINAL
'   - TRIPE DETALHE   : 1 linha por PEP4 + FAMILIA com as 3 quantidades
'
'  AJUSTES (aba CONFIG, opcionais):
'   - TRIPE_ALIAS_FAMILIA : de-para extra de familia, "DE=PARA;DE2=PARA2"
'   - TRIPE_MIN_MATCH_PEP4: % minimo de PEP4 em comum para cruzar por PEP4
'                           (abaixo disso o cruzamento e feito por PEP3)
'==============================================================================
Option Explicit

Private Const ABA_UP_INV As String = "INVENTARIO"     ' upload: base do inventario
Private Const ABA_UP_RAZAO As String = "RAZAO"        ' upload: razao de custo (CKCP)
Private Const ABA_INV As String = "ANALISE SAP x PRJ"
Private Const ABA_CKCP As String = "MATERIAL vs SERVICO"
Private Const ABA_CONC As String = "TRIPE CONCLUSAO"
Private Const ABA_DET As String = "TRIPE DETALHE"
Private Const ABA_RC As String = "RACIONALIZACAO COM"
Private Const ABA_RC_TMP As String = "RACIONALIZACAO COM TMP"
Private Const ABA_RC_CKCP As String = "RACIONALIZACAO COM CKCP"

' Paleta (r + g*256 + b*65536)
Private Const COR_HDR As Long = 4286720        ' RGB(0,105,65)  verde institucional
Private Const COR_TIT As Long = 2505728        ' RGB(0,60,38)
Private Const COR_OK_BG As Long = 15531448     ' RGB(232,245,236)
Private Const COR_OK_FG As Long = 4681237      ' RGB(21,110,71)
Private Const COR_BAD_BG As Long = 15526909    ' RGB(253,235,236)
Private Const COR_BAD_FG As Long = 2629800     ' RGB(168,32,40)
Private Const COR_WARN_BG As Long = 14087423   ' RGB(255,244,214)
Private Const COR_WARN_FG As Long = 26258      ' RGB(146,102,0)
Private Const COR_NEU_BG As Long = 15922161    ' RGB(241,243,242)
Private Const COR_NEU_FG As Long = 8159352     ' RGB(120,128,124)

' Uma linha da conclusao detalhada (PEP4 + FAMILIA)
Private Type tGrupo
    Pep3Key As String
    Pep3 As String
    Pep4 As String
    Familia As String
    Materiais As String     ' "COD - DESCRICAO | ..." (materiais do inventario)
    NMat As Long            ' qtd de materiais distintos no grupo
    Tipo As String          ' UC / COM / UAR / SERV
    QPrj As Double          ' MAT PRJ CAD (inventario)
    QSap As Double          ' MAT LIB SAP (inventario)
    VSap As Double          ' VALOR (inventario)
    TemInv As Boolean
    InvNaoAder As Boolean   ' alguma linha do grupo NAO ADERENTE no SAP x PRJ
    QMatCk As Double        ' MAT (CKCP, ja com conversao de cabo)
    QSrv As Double          ' SRV (CKCP, ja com fator combo)
    TemCk As Boolean
    CkNaoAder As Boolean    ' SITUACAO NAO ADERENTE no MAT x SRV
    Situacao As String
    Perna As String         ' qual perna quebrou
End Type

Private g() As tGrupo
Private nG As Long
Private dG As Object        ' chave PEP4|FAMILIA -> indice em g()
Private dAlias As Object    ' familia normalizada -> familia canonica
Private dMatSeen As Object  ' indice do grupo|COD MAT -> 1 (sem repetir material)
Private Const MAX_MAT_DESC As Long = 15   ' materiais listados por linha do detalhe

' Vereditos por obra (chave = PEP3 normalizado)
Private dP3Nome As Object, dInvVer As Object, dInvMot As Object
Private dCkVer As Object, dCkMot As Object

Private mEtapa As String
Private mOrigemCKCP As String   ' arquivo onde a macro do CKCP foi encontrada


'==============================================================================
'  ROTINA PRINCIPAL
'==============================================================================
Public Sub GerarTripe()
    Dim t0 As Double: t0 = Timer
    Dim falhas As String, msgUp As String

    ' 0) Planilha de upload: abas INVENTARIO e RAZAO
    Dim wb As Workbook: Set wb = AcharPlanilhaUpload()
    If wb Is Nothing Then Exit Sub
    msgUp = ValidarUpload(wb)
    If msgUp <> "" Then
        MsgBox "A planilha '" & wb.Name & "' nao esta no padrao de upload:" & vbCrLf & vbCrLf & _
               msgUp, vbExclamation, "TRIPE"
        Exit Sub
    End If

    On Error GoTo Falha
    wb.Activate

    ' 1) Perna PRJ x SAP (inventario)
    mEtapa = "Inventario (GerarInventario)"
    If Not RodarOuAvisar("GerarInventario", "inventario.bas", falhas) Then Exit Sub

    ' As duas macros criam "RACIONALIZACAO COM": guarda a do inventario
    ' com nome temporario para o CKCP nao apaga-la.
    ApagarAba wb, ABA_RC_TMP
    RenomearAba wb, ABA_RC, ABA_RC_TMP

    ' 2) Perna MATERIAL x SERVICO (CKCP)
    wb.Activate
    mEtapa = "CKCP (GerarRelatorio)"
    If Not RodarOuAvisar("GerarRelatorio", "AnaliseCKCP.bas", falhas) Then
        RenomearAba wb, ABA_RC_TMP, ABA_RC
        Exit Sub
    End If
    LimparSplashCKCP wb

    ApagarAba wb, ABA_RC_CKCP
    RenomearAba wb, ABA_RC, ABA_RC_CKCP
    RenomearAba wb, ABA_RC_TMP, ABA_RC

    ' 3) Cruzamento das tres pernas
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim wsInv As Worksheet, wsCk As Worksheet
    Set wsInv = PegarAba(wb, ABA_INV)
    Set wsCk = PegarAba(wb, ABA_CKCP)
    If wsInv Is Nothing Or wsCk Is Nothing Then
        Restaurar
        MsgBox "As duas macros rodaram, mas uma das abas de resultado nao foi criada:" & vbCrLf & vbCrLf & _
               "  - " & ABA_INV & IIf(wsInv Is Nothing, "  (NAO GERADA)", "  (ok)") & vbCrLf & _
               "  - " & ABA_CKCP & IIf(wsCk Is Nothing, "  (NAO GERADA)", "  (ok)") & vbCrLf & vbCrLf & _
               "Veja a mensagem de erro que a propria macro mostrou." & _
               IIf(falhas <> "", vbCrLf & vbCrLf & "Falhas:" & falhas, ""), _
               vbExclamation, "TRIPE"
        Exit Sub
    End If

    mEtapa = "Carregar de-para de familias": CarregarAlias wb
    Set dG = CreateObject("Scripting.Dictionary")
    Set dMatSeen = CreateObject("Scripting.Dictionary")
    Set dP3Nome = CreateObject("Scripting.Dictionary")
    Set dInvVer = CreateObject("Scripting.Dictionary")
    Set dInvMot = CreateObject("Scripting.Dictionary")
    Set dCkVer = CreateObject("Scripting.Dictionary")
    Set dCkMot = CreateObject("Scripting.Dictionary")
    nG = 0: ReDim g(1 To 64)

    ' Le as duas abas para memoria
    Dim aInv As Variant, aCk As Variant, ci As Object, cc As Object
    mEtapa = "Ler " & ABA_INV: aInv = LerTabelaAba(wsInv, "PEP4NIVEL", ci)
    mEtapa = "Ler " & ABA_CKCP: aCk = LerTabelaAba(wsCk, "PEP4NIVEL", cc)

    ' Nivel do cruzamento: PEP4 quando os PEP4 das duas bases batem; senao PEP3
    mEtapa = "Nivel do cruzamento"
    Dim porPep4 As Boolean, pctMatch As Double
    pctMatch = PercPep4EmComum(aInv, ci, aCk, cc)
    porPep4 = (pctMatch >= CfgNumTripe(wb, "TRIPE_MIN_MATCH_PEP4", 20) / 100)

    mEtapa = "Agregar inventario": AgregarInventario aInv, ci, porPep4
    mEtapa = "Agregar CKCP": AgregarCKCP aCk, cc, porPep4
    mEtapa = "Classificar": ClassificarGrupos

    mEtapa = "Escrever " & ABA_DET
    EscreverDetalhe wb, porPep4, pctMatch
    mEtapa = "Escrever " & ABA_CONC
    Dim nAprov As Long, nReprov As Long, nVerif As Long, nSem As Long
    EscreverConclusao wb, porPep4, pctMatch, nAprov, nReprov, nVerif, nSem

    mEtapa = "Organizar abas": OrganizarAbasFinal wb

    Restaurar
    On Error Resume Next
    wb.Worksheets(ABA_CONC).Activate
    wb.Worksheets(ABA_CONC).Range("A1").Select
    On Error GoTo Falha

    MsgBox "TRIPE concluido em " & Format$(Timer - t0, "0") & " s." & vbCrLf & vbCrLf & _
           "Obras (PEP3): " & (nAprov + nReprov + nVerif + nSem) & vbCrLf & _
           "  APROVADAS:  " & nAprov & vbCrLf & _
           "  REPROVADAS: " & nReprov & vbCrLf & _
           "  VERIFICAR:  " & nVerif & vbCrLf & _
           "  SEM CRUZAMENTO (so numa das abas): " & nSem & vbCrLf & vbCrLf & _
           "Cruzamento por " & IIf(porPep4, "PEP4", "PEP3") & _
           " (" & Format$(pctMatch, "0%") & " dos PEP4 da base menor em comum)." & _
           IIf(falhas <> "", vbCrLf & vbCrLf & "ATENCAO - falhas:" & falhas, ""), _
           IIf(falhas <> "", vbExclamation, vbInformation), "TRIPE"
    Exit Sub

Falha:
    Restaurar
    MsgBox "Erro " & Err.Number & " na etapa [" & mEtapa & "]:" & vbCrLf & _
           Err.Description, vbCritical, "TRIPE"
End Sub

Private Sub Restaurar()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub


'==============================================================================
'  PLANILHA DE UPLOAD
'==============================================================================
' Usa a pasta ativa se tiver INVENTARIO e RAZAO; senao procura nas pastas
' abertas; senao pede o arquivo.
Private Function AcharPlanilhaUpload() As Workbook
    If EhUpload(ActiveWorkbook) Then Set AcharPlanilhaUpload = ActiveWorkbook: Exit Function
    Dim w As Workbook
    For Each w In Application.Workbooks
        If EhUpload(w) Then Set AcharPlanilhaUpload = w: Exit Function
    Next w

    Dim arq As Variant
    arq = Application.GetOpenFilename("Planilha de upload (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", , _
                                      "TRIPE - escolha a planilha de upload (abas INVENTARIO e RAZAO)")
    If VarType(arq) = vbBoolean Then Exit Function          ' cancelou
    Set w = Workbooks.Open(CStr(arq))
    If Not EhUpload(w) Then
        MsgBox "O arquivo '" & w.Name & "' nao tem as abas " & ABA_UP_INV & " e " & _
               ABA_UP_RAZAO & ".", vbExclamation, "TRIPE"
        Exit Function
    End If
    Set AcharPlanilhaUpload = w
End Function

Private Function EhUpload(w As Workbook) As Boolean
    If w Is Nothing Then Exit Function
    EhUpload = Not PegarAba(w, ABA_UP_INV) Is Nothing And Not PegarAba(w, ABA_UP_RAZAO) Is Nothing
End Function

' Confere os cabecalhos minimos das duas abas; "" = ok
Private Function ValidarUpload(wb As Workbook) As String
    Dim falta As String
    falta = FaltamCabecalhos(PegarAba(wb, ABA_UP_INV), _
              Array("PEP3NIVEL", "PEP4NIVEL", "MATERIAL|COD MAT", "VALOR", "TIPO", "FAMILIA", _
                    "ATENDIDO SAP|MAT LIB SAP", "TOMBADO PROJ|MAT PRJ CAD", _
                    "SITUACAO DO MATERIAL|SIT MAT"))
    If falta <> "" Then ValidarUpload = "  Aba " & ABA_UP_INV & " sem: " & falta & vbCrLf
    falta = FaltamCabecalhos(PegarAba(wb, ABA_UP_RAZAO), _
              Array("ELEMENTO PEP", "MATERIAL", "QTD.TOTAL ENTRADA", "VALOR/MOEDA OBJETO", "CLASSIFICACAO"))
    If falta <> "" Then ValidarUpload = ValidarUpload & "  Aba " & ABA_UP_RAZAO & " sem: " & falta & vbCrLf
    If PegarAba(wb, ABA_UP_INV).Cells(2, 1).Value = "" Then _
        ValidarUpload = ValidarUpload & "  Aba " & ABA_UP_INV & " sem dados (linha 2 vazia)" & vbCrLf
    If PegarAba(wb, ABA_UP_RAZAO).Cells(2, 2).Value = "" Then _
        ValidarUpload = ValidarUpload & "  Aba " & ABA_UP_RAZAO & " sem dados (linha 2 vazia)" & vbCrLf
End Function

' Cada item da lista pode ter alternativas "A|B"; compara sem acento/"_"
Private Function FaltamCabecalhos(ws As Worksheet, req As Variant) As String
    Dim lastC As Long, c As Long, d As Object, i As Long, alts() As String, k As Long, ok As Boolean
    Set d = CreateObject("Scripting.Dictionary")
    lastC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastC
        If Not IsError(ws.Cells(1, c).Value) Then d(NormTxt(CStr(ws.Cells(1, c).Value))) = 1
    Next c
    For i = 0 To UBound(req)
        alts = Split(req(i), "|")
        ok = False
        For k = 0 To UBound(alts)
            If d.Exists(NormTxt(alts(k))) Then ok = True: Exit For
        Next k
        If Not ok Then FaltamCabecalhos = FaltamCabecalhos & IIf(FaltamCabecalhos <> "", ", ", "") & alts(0)
    Next i
End Function


'==============================================================================
'  LEITURA DAS ABAS DE RESULTADO
'==============================================================================
' Acha a linha de cabecalho (a que tem a coluna "chave") nas 40 primeiras
' linhas e devolve os dados abaixo dela; mapa = NOME NORMALIZADO -> coluna.
Private Function LerTabelaAba(ws As Worksheet, ByVal chave As String, _
                              ByRef mapa As Object) As Variant
    Set mapa = CreateObject("Scripting.Dictionary")
    Dim lastC As Long, r As Long, c As Long, hr As Long
    lastC = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If lastC > 60 Then lastC = 60

    Dim topo As Variant
    topo = ws.Range(ws.Cells(1, 1), ws.Cells(40, lastC)).Value
    For r = 1 To 40
        For c = 1 To lastC
            If Not IsError(topo(r, c)) Then
                If NormTxt(CStr(topo(r, c))) = chave Then hr = r: Exit For
            End If
        Next c
        If hr > 0 Then Exit For
    Next r
    If hr = 0 Then Err.Raise vbObjectError + 1, , _
        "Cabecalho '" & chave & "' nao encontrado na aba " & ws.Name

    Dim h As String
    For c = 1 To lastC
        h = ""
        If Not IsError(topo(hr, c)) Then h = NormTxt(CStr(topo(hr, c)))
        If h <> "" And Not mapa.Exists(h) Then mapa.Add h, c
    Next c

    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.Count, mapa(chave)).End(xlUp).Row
    If lastR <= hr Then
        LerTabelaAba = Empty
    Else
        LerTabelaAba = ws.Range(ws.Cells(hr + 1, 1), ws.Cells(lastR, lastC)).Value
    End If
End Function

Private Function Col(mapa As Object, ByVal nome As String) As Long
    If mapa.Exists(nome) Then Col = mapa(nome)
End Function

Private Function Cel(a As Variant, ByVal i As Long, ByVal c As Long) As String
    If c <= 0 Then Exit Function
    If IsError(a(i, c)) Then Exit Function
    Cel = Trim$(CStr(a(i, c)))
End Function

Private Function Num(a As Variant, ByVal i As Long, ByVal c As Long) As Double
    If c <= 0 Then Exit Function
    If IsError(a(i, c)) Then Exit Function
    If IsNumeric(a(i, c)) Then Num = CDbl(a(i, c))
End Function

' % de PEP4 em comum, medido sobre a base MENOR (a RAZAO costuma cobrir so
' parte das obras do inventario: o que importa e se os PEP4 dela batem)
Private Function PercPep4EmComum(aInv As Variant, ci As Object, _
                                 aCk As Variant, cc As Object) As Double
    If IsEmpty(aInv) Or IsEmpty(aCk) Then Exit Function
    Dim dCk As Object: Set dCk = CreateObject("Scripting.Dictionary")
    Dim dInv As Object: Set dInv = CreateObject("Scripting.Dictionary")
    Dim i As Long, k As String, n As Long
    For i = 1 To UBound(aCk, 1)
        k = KeyPep(Cel(aCk, i, Col(cc, "PEP4NIVEL")))
        If k <> "" Then dCk(k) = 1
    Next i
    For i = 1 To UBound(aInv, 1)
        k = KeyPep(Cel(aInv, i, Col(ci, "PEP4NIVEL")))
        If k <> "" Then dInv(k) = 1
    Next i
    If dInv.Count = 0 Or dCk.Count = 0 Then Exit Function
    Dim kv As Variant
    For Each kv In dInv.Keys
        If dCk.Exists(kv) Then n = n + 1
    Next kv
    If dInv.Count < dCk.Count Then
        PercPep4EmComum = n / dInv.Count
    Else
        PercPep4EmComum = n / dCk.Count
    End If
End Function


'==============================================================================
'  AGREGACAO DAS PERNAS
'==============================================================================
' Perna 1 + 2: projetado x baixado (inventario)
Private Sub AgregarInventario(a As Variant, mp As Object, ByVal porPep4 As Boolean)
    If IsEmpty(a) Then Exit Sub
    Dim cP3 As Long, cP4 As Long, cFam As Long, cTipo As Long, cLib As Long
    Dim cPrj As Long, cVal As Long, cSit As Long, cApr As Long, cMot As Long
    cP3 = Col(mp, "PEP3NIVEL"): cP4 = Col(mp, "PEP4NIVEL")
    cFam = Col(mp, "FAMILIA"): cTipo = Col(mp, "TIPO")
    cLib = Col(mp, "MAT LIB SAP"): cPrj = Col(mp, "MAT PRJ CAD")
    cVal = Col(mp, "VALOR"): cSit = Col(mp, "SIT MAT")
    cApr = Col(mp, "APROVACAO"): cMot = Col(mp, "MOTIVO DEVOLUCAO PEP3")
    Dim cCod As Long, cDesc As Long
    cCod = Col(mp, "COD MAT"): cDesc = Col(mp, "DESC MAT")

    Dim i As Long, p3 As String, p4 As String, p3k As String, ix As Long
    Dim apr As String, sit As String
    For i = 1 To UBound(a, 1)
        p4 = Cel(a, i, cP4): If p4 = "" Then GoTo Prox
        p3 = Cel(a, i, cP3): If p3 = "" Then p3 = Pep3DePep4(p4)
        p3k = KeyPep(p3)
        If Not dP3Nome.Exists(p3k) Then dP3Nome(p3k) = p3

        ' veredito da obra no SAP x PRJ (pior caso entre as linhas)
        apr = NormTxt(Cel(a, i, cApr))
        If InStr(apr, "REPROVADO") > 0 Then
            dInvVer(p3k) = "REPROVADO"
        ElseIf InStr(apr, "APROVADO") > 0 Then
            If dInvVer(p3k) <> "REPROVADO" Then dInvVer(p3k) = "APROVADO"
        ElseIf InStr(apr, "SEM UC") > 0 Then
            If dInvVer(p3k) = "" Then dInvVer(p3k) = "SEM UC"
        End If
        If Cel(a, i, cMot) <> "" And Not dInvMot.Exists(p3k) Then dInvMot(p3k) = Cel(a, i, cMot)

        ix = PegarGrupo(p3k, p3, IIf(porPep4, p4, ""), Cel(a, i, cFam))
        With g(ix)
            .TemInv = True
            .QPrj = .QPrj + Num(a, i, cPrj)
            .QSap = .QSap + Num(a, i, cLib)
            .VSap = .VSap + Num(a, i, cVal)
            If .Tipo = "" Or .Tipo = "COM" Then
                If NormTxt(Cel(a, i, cTipo)) <> "" Then .Tipo = NormTxt(Cel(a, i, cTipo))
            End If
            sit = NormTxt(Cel(a, i, cSit))
            If InStr(sit, "NAO ADER") > 0 Or InStr(sit, "DIVERG") > 0 Then .InvNaoAder = True
        End With
        AddMaterial ix, Cel(a, i, cCod), Cel(a, i, cDesc)
Prox:
    Next i
End Sub

' Perna 2 + 3: material lancado x servico pago (CKCP)
Private Sub AgregarCKCP(a As Variant, mp As Object, ByVal porPep4 As Boolean)
    If IsEmpty(a) Then Exit Sub
    Dim cSt As Long, cP3 As Long, cP4 As Long, cCls As Long, cTipo As Long
    Dim cMat As Long, cSrv As Long, cSit As Long, cObs1 As Long, cObs2 As Long
    cSt = Col(mp, "STATUS PEP4"): cP3 = Col(mp, "PEP3NIVEL"): cP4 = Col(mp, "PEP4NIVEL")
    cCls = Col(mp, "CLS2"): cTipo = Col(mp, "TIPO")
    cMat = Col(mp, "MAT"): cSrv = Col(mp, "SRV"): cSit = Col(mp, "SITUACAO")
    cObs1 = Col(mp, "OBS1"): cObs2 = Col(mp, "OBS2")

    Dim i As Long, p3 As String, p4 As String, p3k As String, ix As Long
    Dim st As String, sit As String, tp As String, mot As String
    For i = 1 To UBound(a, 1)
        p4 = Cel(a, i, cP4): If p4 = "" Then GoTo Prox
        p3 = Cel(a, i, cP3): If p3 = "" Then p3 = Pep3DePep4(p4)
        p3k = KeyPep(p3)
        If Not dP3Nome.Exists(p3k) Then dP3Nome(p3k) = p3

        st = NormTxt(Cel(a, i, cSt))
        If st = "REPROVADO" Then
            dCkVer(p3k) = "REPROVADO"
        ElseIf st = "APROVADO" Then
            If dCkVer(p3k) <> "REPROVADO" Then dCkVer(p3k) = "APROVADO"
        ElseIf st <> "" Then
            If dCkVer(p3k) = "" Then dCkVer(p3k) = st      ' SEM UC
        End If
        mot = Cel(a, i, cObs2)
        If InStr(NormTxt(mot), "REPROVADO") = 0 Then mot = ""
        If mot <> "" And Not dCkMot.Exists(p3k) Then dCkMot(p3k) = mot

        ix = PegarGrupo(p3k, p3, IIf(porPep4, p4, ""), Cel(a, i, cCls))
        With g(ix)
            .TemCk = True
            .QMatCk = .QMatCk + Num(a, i, cMat)
            .QSrv = .QSrv + Num(a, i, cSrv)
            tp = NormTxt(Cel(a, i, cTipo))
            If .Tipo = "" Or (.Tipo = "SERV" And tp <> "SERV") Then .Tipo = tp
            sit = NormTxt(Cel(a, i, cSit))
            If InStr(sit, "NAO ADER") > 0 Then .CkNaoAder = True
        End With
Prox:
    Next i
End Sub

' Acrescenta "COD - DESCRICAO" na lista do grupo (sem repetir; ate MAX_MAT_DESC)
Private Sub AddMaterial(ByVal ix As Long, ByVal cod As String, ByVal desc As String)
    If cod = "" And desc = "" Then Exit Sub
    Dim k As String: k = ix & "|" & IIf(cod <> "", cod, desc)
    If dMatSeen.Exists(k) Then Exit Sub
    dMatSeen.Add k, 1
    Dim item As String
    If cod <> "" And desc <> "" Then
        item = cod & " - " & desc
    Else
        item = cod & desc
    End If
    g(ix).NMat = g(ix).NMat + 1
    If g(ix).NMat <= MAX_MAT_DESC Then
        g(ix).Materiais = g(ix).Materiais & IIf(g(ix).Materiais <> "", " | ", "") & item
    End If
End Sub

Private Function PegarGrupo(ByVal p3k As String, ByVal p3 As String, _
                            ByVal p4 As String, ByVal famRaw As String) As Long
    Dim fam As String: fam = NormFam(famRaw)
    If fam = "" Then fam = "(SEM FAMILIA)"
    Dim k As String
    If p4 = "" Then k = p3k & "|" & fam Else k = KeyPep(p4) & "|" & fam
    If dG.Exists(k) Then PegarGrupo = dG(k): Exit Function
    nG = nG + 1
    If nG > UBound(g) Then ReDim Preserve g(1 To UBound(g) * 2)
    g(nG).Pep3Key = p3k
    g(nG).Pep3 = p3
    g(nG).Pep4 = IIf(p4 = "", "(todos)", p4)
    g(nG).Familia = fam
    dG.Add k, nG
    PegarGrupo = nG
End Function


'==============================================================================
'  REGRA DO TRIPE (situacao de cada PEP4 + FAMILIA)
'==============================================================================
' A aderencia de cada perna vem da propria macro (com todas as regras
' especiais: equivalencias, cabo, ramal, religador, trafo ...). O TRIPE so
' junta as duas e diz QUAL perna quebrou.
Private Sub ClassificarGrupos()
    Dim i As Long, sap As Double, prj As Double, srv As Double
    For i = 1 To nG
        With g(i)
            sap = Abs(.QSap): prj = Abs(.QPrj): srv = Abs(.QSrv)
            If .TemCk And Not .TemInv And .Tipo = "SERV" Then
                .Situacao = "SERVICO PURO"
                .Perna = "-"
            ElseIf .TemCk And Not .TemInv Then
                If .CkNaoAder Then
                    .Situacao = "SO NO CKCP - MAT x SRV DIVERGENTE"
                    .Perna = "SERVICO"
                Else
                    .Situacao = "SO NO CKCP"
                    .Perna = "-"
                End If
            ElseIf prj > 0 And sap = 0 Then
                .Situacao = "PROJETADO NAO BAIXADO"
                .Perna = "BAIXA SAP"
            ElseIf .TemInv And Not .TemCk And sap > 0 And .Tipo = "UC" Then
                .Situacao = "MATERIAL BAIXADO SEM SERVICO PAGO"
                .Perna = "SERVICO"
            ElseIf .InvNaoAder And .CkNaoAder Then
                .Situacao = "DIVERGENTE NAS TRES PERNAS"
                .Perna = "PROJETO + SERVICO"
            ElseIf .InvNaoAder Then
                .Situacao = "DIVERGENTE PROJETO x BAIXA"
                .Perna = "PROJETO"
            ElseIf .CkNaoAder Then
                If sap > 0 And srv = 0 Then
                    .Situacao = "MATERIAL BAIXADO SEM SERVICO PAGO"
                ElseIf sap = 0 And srv > 0 Then
                    .Situacao = "SERVICO PAGO SEM MATERIAL"
                ElseIf srv > sap Then
                    .Situacao = "SERVICO PAGO ACIMA DO MATERIAL"
                Else
                    .Situacao = "SERVICO PAGO ABAIXO DO MATERIAL"
                End If
                .Perna = "SERVICO"
            ElseIf .TemInv And Not .TemCk Then
                .Situacao = "OK (COM - SERVICO NAO EXIGIDO)"
                If .Tipo <> "COM" Then .Situacao = "OK (SEM LANCAMENTO NO CKCP)"
                .Perna = "-"
            Else
                .Situacao = "ADERENTE (PRJ = SAP = SRV)"
                .Perna = "-"
            End If
        End With
    Next i
End Sub

' Veredito final da obra: APROVADO so quando as duas analises aprovam.
' Obra que esta so numa das abas de upload fica SEM RAZAO / SEM INVENTARIO
' (nao da para fechar o tripe), mas uma reprovacao ja conhecida prevalece.
Private Function VereditoFinal(ByVal vInv As String, ByVal vCk As String) As String
    If vInv = "SEM DADOS" Then vInv = ""
    If vCk = "SEM DADOS" Then vCk = ""
    If vInv = "REPROVADO" Or vCk = "REPROVADO" Then
        VereditoFinal = "REPROVADO"
    ElseIf vCk = "" Then
        VereditoFinal = "SEM RAZAO"
    ElseIf vInv = "" Then
        VereditoFinal = "SEM INVENTARIO"
    ElseIf vInv = "APROVADO" And vCk = "APROVADO" Then
        VereditoFinal = "APROVADO"
    Else
        VereditoFinal = "VERIFICAR"
    End If
End Function


'==============================================================================
'  SAIDA: TRIPE DETALHE
'==============================================================================
Private Sub EscreverDetalhe(wb As Workbook, ByVal porPep4 As Boolean, ByVal pctMatch As Double)
    Dim ws As Worksheet: Set ws = NovaAba(wb, ABA_DET)
    Const HR As Long = 4
    Dim hdr As Variant
    hdr = Array("VEREDITO FINAL PEP3", "PEP3NIVEL", "PEP4NIVEL", "FAMILIA", _
                "DESCRICAO DO MATERIAL", "TIPO", _
                "QTD PROJETO (PRJ)", "QTD BAIXADA SAP", "QTD MAT CKCP", "QTD SERVICO PAGO", _
                "DIF SAP - PRJ", "DIF SRV - SAP", "VALOR SAP (R$)", _
                "SIT PRJ x SAP", "SIT MAT x SRV", "SITUACAO TRIPE", "PERNA QUE QUEBROU")
    Dim nC As Long: nC = UBound(hdr) + 1

    Titulo ws, "TRIPE  |  Detalhe por " & IIf(porPep4, "PEP4", "PEP3") & " + Familia", _
           "Projeto x Baixa SAP x Servico pago  -  " & Format$(pctMatch, "0%") & _
           " dos PEP4 da base menor existem na outra" & _
           IIf(porPep4, "", " (abaixo do minimo: cruzamento feito por PEP3)"), nC

    Dim c As Long
    For c = 1 To nC: ws.Cells(HR, c).Value = hdr(c - 1): Next c
    If nG = 0 Then GoTo Formata

    ' ordena por PEP3 > PEP4 > FAMILIA
    Dim ord() As Long, chv() As String, i As Long
    ReDim ord(1 To nG): ReDim chv(1 To nG)
    For i = 1 To nG
        ord(i) = i
        chv(i) = g(i).Pep3Key & "|" & g(i).Pep4 & "|" & g(i).Familia
    Next i
    QuickSortIdx chv, ord, 1, nG

    Dim out() As Variant: ReDim out(1 To nG, 1 To nC)
    Dim r As Long, v As String
    For r = 1 To nG
        i = ord(r)
        With g(i)
            out(r, 1) = VereditoFinal(CStr(dInvVer(.Pep3Key)), CStr(dCkVer(.Pep3Key)))
            out(r, 2) = .Pep3
            out(r, 3) = .Pep4
            out(r, 4) = .Familia
            If .NMat = 0 Then
                out(r, 5) = "-"
            ElseIf .NMat > MAX_MAT_DESC Then
                out(r, 5) = .Materiais & " | (+" & (.NMat - MAX_MAT_DESC) & " outros)"
            Else
                out(r, 5) = .Materiais
            End If
            out(r, 6) = .Tipo
            out(r, 7) = Round(.QPrj, 2)
            out(r, 8) = Round(.QSap, 2)
            out(r, 9) = Round(.QMatCk, 2)
            out(r, 10) = Round(.QSrv, 2)
            out(r, 11) = Round(Abs(.QSap) - Abs(.QPrj), 2)
            out(r, 12) = Round(Abs(.QSrv) - Abs(.QSap), 2)
            out(r, 13) = Round(.VSap, 2)
            out(r, 14) = IIf(.TemInv, IIf(.InvNaoAder, "NAO ADERENTE", "ADERENTE"), "-")
            out(r, 15) = IIf(.TemCk, IIf(.CkNaoAder, "NAO ADERENTE", "ADERENTE"), "-")
            out(r, 16) = .Situacao
            out(r, 17) = .Perna
        End With
    Next r
    ws.Range(ws.Cells(HR + 1, 1), ws.Cells(HR + nG, nC)).Value = out

Formata:
    FormatarTabela ws, HR, nG, nC
    ws.Range(ws.Cells(HR + 1, 7), ws.Cells(HR + nG + 1, 12)).NumberFormat = "#,##0.##;[Red]-#,##0.##;-"
    ws.Range(ws.Cells(HR + 1, 13), ws.Cells(HR + nG + 1, 13)).NumberFormat = "#,##0.00;[Red]-#,##0.00;-"
    For r = 1 To nG
        PintarVeredito ws.Cells(HR + r, 1)
        PintarSituacao ws.Cells(HR + r, 14)
        PintarSituacao ws.Cells(HR + r, 15)
        PintarSituacao ws.Cells(HR + r, 16)
    Next r
    ws.Columns(1).ColumnWidth = 14
    ws.Columns(2).ColumnWidth = 22: ws.Columns(3).ColumnWidth = 24
    ws.Columns(4).ColumnWidth = 20
    ws.Columns(5).ColumnWidth = 60          ' DESCRICAO DO MATERIAL (texto quebra na celula)
    ws.Range(ws.Cells(HR + 1, 5), ws.Cells(HR + nG + 1, 5)).WrapText = True
    ws.Columns(6).ColumnWidth = 7
    ws.Range("G:M").ColumnWidth = 13
    ws.Columns(14).ColumnWidth = 14: ws.Columns(15).ColumnWidth = 14
    ws.Columns(16).ColumnWidth = 38: ws.Columns(17).ColumnWidth = 18
End Sub


'==============================================================================
'  SAIDA: TRIPE CONCLUSAO (1 linha por obra)
'==============================================================================
Private Sub EscreverConclusao(wb As Workbook, ByVal porPep4 As Boolean, ByVal pctMatch As Double, _
                              ByRef nAprov As Long, ByRef nReprov As Long, ByRef nVerif As Long, _
                              ByRef nSem As Long)
    Dim ws As Worksheet: Set ws = NovaAba(wb, ABA_CONC)
    Const HR As Long = 8
    Dim hdr As Variant
    hdr = Array("VEREDITO FINAL", "PEP3NIVEL", "PROJETO x BAIXA (inventario)", _
                "MATERIAL x SERVICO (CKCP)", "FAMILIAS DIVERGENTES", _
                "QUEBROU NO PROJETO", "QUEBROU NO SERVICO", "VALOR SAP (R$)", "MOTIVO")
    Dim nC As Long: nC = UBound(hdr) + 1

    Titulo ws, "TRIPE  |  Conclusao por obra", _
           "Aprovada so quando as tres pernas batem: Projeto = Baixa SAP = Servico pago", nC

    ' totais por obra a partir do detalhe
    Dim dDiv As Object, dQP As Object, dQS As Object, dVal As Object
    Set dDiv = CreateObject("Scripting.Dictionary"): Set dQP = CreateObject("Scripting.Dictionary")
    Set dQS = CreateObject("Scripting.Dictionary"): Set dVal = CreateObject("Scripting.Dictionary")
    Dim i As Long, k As String
    For i = 1 To nG
        k = g(i).Pep3Key
        dVal(k) = dVal(k) + g(i).VSap
        If g(i).Perna <> "-" Then
            dDiv(k) = dDiv(k) + 1
            If InStr(g(i).Perna, "PROJETO") > 0 Or g(i).Perna = "BAIXA SAP" Then dQP(k) = dQP(k) + 1
            If InStr(g(i).Perna, "SERVICO") > 0 Then dQS(k) = dQS(k) + 1
        End If
    Next i

    Dim ks As Variant, n As Long: n = dP3Nome.Count
    Dim c As Long
    For c = 1 To nC: ws.Cells(HR, c).Value = hdr(c - 1): Next c
    If n = 0 Then GoTo Formata

    ks = dP3Nome.Keys
    Dim ord() As Long, chv() As String
    ReDim ord(1 To n): ReDim chv(1 To n)
    For i = 1 To n: ord(i) = i: chv(i) = CStr(ks(i - 1)): Next i
    QuickSortIdx chv, ord, 1, n

    Dim out() As Variant: ReDim out(1 To n, 1 To nC)
    Dim r As Long, vI As String, vC As String, vF As String, mot As String
    For r = 1 To n
        k = CStr(ks(ord(r) - 1))
        vI = CStr(dInvVer(k)): If vI = "" Then vI = "SEM DADOS"
        vC = CStr(dCkVer(k)): If vC = "" Then vC = "SEM DADOS"
        vF = VereditoFinal(vI, vC)
        Select Case vF
            Case "APROVADO": nAprov = nAprov + 1
            Case "REPROVADO": nReprov = nReprov + 1
            Case "SEM RAZAO", "SEM INVENTARIO": nSem = nSem + 1
            Case Else: nVerif = nVerif + 1
        End Select
        mot = ""
        If dInvMot.Exists(k) Then mot = "PRJ x SAP: " & dInvMot(k)
        If dCkMot.Exists(k) Then mot = mot & IIf(mot <> "", "  |  ", "") & "MAT x SRV: " & dCkMot(k)
        If mot = "" Then
            If vF = "SEM RAZAO" Then
                mot = "Obra sem lancamento na aba RAZAO: servico pago nao avaliado"
            ElseIf vF = "SEM INVENTARIO" Then
                mot = "Obra sem linha na aba INVENTARIO: projeto x baixa nao avaliado"
            ElseIf vF = "VERIFICAR" Then
                mot = "Sem familia UC para avaliar"
            End If
        End If
        out(r, 1) = vF
        out(r, 2) = dP3Nome(k)
        out(r, 3) = vI
        out(r, 4) = vC
        out(r, 5) = CLng(dDiv(k))
        out(r, 6) = CLng(dQP(k))
        out(r, 7) = CLng(dQS(k))
        out(r, 8) = Round(dVal(k), 2)
        out(r, 9) = mot
    Next r
    ws.Range(ws.Cells(HR + 1, 1), ws.Cells(HR + n, nC)).Value = out

Formata:
    ' cartoes de KPI
    Cartao ws, 4, 1, "OBRAS", n, COR_NEU_BG, COR_TIT
    Cartao ws, 4, 2, "APROVADAS", nAprov, COR_OK_BG, COR_OK_FG
    Cartao ws, 4, 3, "REPROVADAS", nReprov, COR_BAD_BG, COR_BAD_FG
    Cartao ws, 4, 4, "VERIFICAR", nVerif, COR_WARN_BG, COR_WARN_FG
    Cartao ws, 4, 5, "SEM CRUZAMENTO", nSem, COR_NEU_BG, COR_NEU_FG
    ws.Cells(6, 6).Value = "Cruzamento por " & IIf(porPep4, "PEP4", "PEP3") & _
                           "  |  PEP4 em comum (base menor): " & Format$(pctMatch, "0%")
    ws.Cells(6, 6).Font.Color = COR_NEU_FG: ws.Cells(6, 6).Font.Size = 9

    FormatarTabela ws, HR, n, nC
    ws.Range(ws.Cells(HR + 1, 8), ws.Cells(HR + n + 1, 8)).NumberFormat = "#,##0.00;[Red]-#,##0.00;-"
    ws.Range(ws.Cells(HR + 1, 5), ws.Cells(HR + n + 1, 7)).HorizontalAlignment = xlCenter
    For r = 1 To n
        PintarVeredito ws.Cells(HR + r, 1)
        PintarVeredito ws.Cells(HR + r, 3)
        PintarVeredito ws.Cells(HR + r, 4)
    Next r
    ws.Columns(1).ColumnWidth = 16: ws.Columns(2).ColumnWidth = 24
    ws.Columns(3).ColumnWidth = 18: ws.Columns(4).ColumnWidth = 18
    ws.Range("E:G").ColumnWidth = 14: ws.Columns(8).ColumnWidth = 16
    ws.Columns(9).ColumnWidth = 90
    ws.Range(ws.Cells(HR, 1), ws.Cells(HR, nC)).WrapText = True
    ws.Rows(HR).RowHeight = 32
End Sub


'==============================================================================
'  FORMATACAO
'==============================================================================
Private Function NovaAba(wb As Workbook, ByVal nome As String) As Worksheet
    ApagarAba wb, nome
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    ws.Name = nome
    ws.Tab.Color = COR_HDR
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 9.5
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 90
    On Error GoTo 0
    Set NovaAba = ws
End Function

Private Sub Titulo(ws As Worksheet, ByVal t As String, ByVal subT As String, ByVal nC As Long)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, nC))
        .Interior.Color = COR_TIT
    End With
    ws.Cells(1, 1).Value = t
    ws.Cells(1, 1).Font.Size = 15: ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Color = vbWhite
    ws.Rows(1).RowHeight = 30
    ws.Cells(2, 1).Value = subT
    ws.Cells(2, 1).Font.Color = COR_NEU_FG
    ws.Cells(2, 1).Font.Italic = True
End Sub

Private Sub Cartao(ws As Worksheet, ByVal r As Long, ByVal c As Long, _
                   ByVal rot As String, ByVal v As Long, ByVal bg As Long, ByVal fg As Long)
    With ws.Cells(r, c)
        .Value = rot
        .Interior.Color = bg: .Font.Color = fg
        .Font.Size = 8: .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    With ws.Cells(r + 1, c)
        .Value = v
        .Interior.Color = bg: .Font.Color = fg
        .Font.Size = 18: .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(r + 1).RowHeight = 30
End Sub

Private Sub FormatarTabela(ws As Worksheet, ByVal hr As Long, ByVal n As Long, ByVal nC As Long)
    With ws.Range(ws.Cells(hr, 1), ws.Cells(hr, nC))
        .Interior.Color = COR_HDR
        .Font.Color = vbWhite: .Font.Bold = True
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Rows(hr).RowHeight = 30
    If n > 0 Then
        With ws.Range(ws.Cells(hr + 1, 1), ws.Cells(hr + n, nC)).Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous: .Color = RGB(230, 230, 230)
        End With
    End If
    ws.Range(ws.Cells(hr, 1), ws.Cells(hr + IIf(n > 0, n, 1), nC)).AutoFilter
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(hr + 1, 3).Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
End Sub

Private Sub PintarVeredito(c As Range)
    Dim v As String: v = UCase$(CStr(c.Value))
    If v = "APROVADO" Then
        c.Interior.Color = COR_OK_BG: c.Font.Color = COR_OK_FG
    ElseIf v = "REPROVADO" Then
        c.Interior.Color = COR_BAD_BG: c.Font.Color = COR_BAD_FG
    ElseIf v = "VERIFICAR" Or v = "SEM UC" Then
        c.Interior.Color = COR_WARN_BG: c.Font.Color = COR_WARN_FG
    Else
        c.Interior.Color = COR_NEU_BG: c.Font.Color = COR_NEU_FG
    End If
    c.Font.Bold = True
    c.HorizontalAlignment = xlCenter
End Sub

Private Sub PintarSituacao(c As Range)
    Dim v As String: v = UCase$(CStr(c.Value))
    If v = "-" Or v = "" Or Left$(v, 7) = "SO NO C" Or v = "SERVICO PURO" Then
        c.Font.Color = COR_NEU_FG
    ElseIf Left$(v, 2) = "OK" Or Left$(v, 8) = "ADERENTE" Then
        c.Interior.Color = COR_OK_BG: c.Font.Color = COR_OK_FG
    Else
        c.Interior.Color = COR_BAD_BG: c.Font.Color = COR_BAD_FG
    End If
End Sub


'==============================================================================
'  NORMALIZACAO / APOIO
'==============================================================================
' Maiusculas, sem acento, "_" vira espaco, espacos colapsados.
Private Function NormTxt(ByVal s As String) As String
    Dim cods As Variant, subs As Variant, k As Long
    s = UCase$(Trim$(s))
    cods = Array(192, 193, 194, 195, 196, 200, 201, 202, 205, 206, 211, 212, 213, 214, 218, 219, 220, 199)
    subs = Array("A", "A", "A", "A", "A", "E", "E", "E", "I", "I", "O", "O", "O", "O", "U", "U", "U", "C")
    For k = 0 To UBound(cods)
        s = Replace(s, ChrW(CLng(cods(k))), subs(k))
    Next k
    s = Replace(s, "_", " ")
    s = Replace(s, ChrW(&H25CF), "")         ' bolinha usada na coluna SIT MAT
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormTxt = Trim$(s)
End Function

' Chave de PEP: so letras e numeros (ignora ".", "-", espacos)
Private Function KeyPep(ByVal s As String) As String
    Dim i As Long, ch As String, r As String
    s = UCase$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then r = r & ch
    Next i
    KeyPep = r
End Function

Private Function Pep3DePep4(ByVal pep As String) As String
    Dim s As String: s = UCase$(Right$(pep, 2))
    If s = ".I" Or s = ".D" Or s = ".M" Or s = ".S" Then
        Pep3DePep4 = Left$(pep, Len(pep) - 2)
    Else
        Pep3DePep4 = pep
    End If
End Function

' Familia normalizada + de-para (o inventario usa FAMILIA, o CKCP usa CLS2)
Private Function NormFam(ByVal s As String) As String
    s = NormTxt(s)
    If dAlias.Exists(s) Then s = dAlias(s)
    NormFam = s
End Function

Private Sub CarregarAlias(wb As Workbook)
    Set dAlias = CreateObject("Scripting.Dictionary")
    ' de-para padrao (mesmo do CKCP)
    dAlias("COND ISOLADO/PROT") = "COND PROT"
    dAlias("COND ISOLADO") = "COND PROT"
    dAlias("TRAFO OU REGULADOR") = "TRAFO_OU_REGULADOR"
    ' de-para extra da CONFIG: "DE=PARA;DE2=PARA2"
    Dim lista As String, p() As String, q() As String, i As Long
    lista = CfgTxtTripe(wb, "TRIPE_ALIAS_FAMILIA", "")
    If lista = "" Then Exit Sub
    p = Split(lista, ";")
    For i = 0 To UBound(p)
        q = Split(p(i), "=")
        If UBound(q) = 1 Then
            If NormTxt(q(0)) <> "" Then dAlias(NormTxt(q(0))) = NormTxt(q(1))
        End If
    Next i
End Sub

Private Function CfgTxtTripe(wb As Workbook, ByVal chave As String, ByVal padrao As String) As String
    CfgTxtTripe = padrao
    Dim ws As Worksheet: Set ws = PegarAba(wb, "CONFIG")
    If ws Is Nothing Then Exit Function
    GarantirChaveTripe ws, "TRIPE_ALIAS_FAMILIA", "", _
        "TRIPE: de-para extra de familia (inventario x CKCP), ex.: CABO MULTIPLEX=COND PROT;X=Y"
    GarantirChaveTripe ws, "TRIPE_MIN_MATCH_PEP4", "20", _
        "TRIPE: % minimo de PEP4 em comum para cruzar por PEP4 (abaixo: por PEP3)"
    Dim f As Range
    Set f = ws.Columns(1).Find(chave, LookAt:=xlWhole, MatchCase:=False)
    If Not f Is Nothing Then
        If Trim$(CStr(f.Offset(0, 1).Value)) <> "" Then CfgTxtTripe = Trim$(CStr(f.Offset(0, 1).Value))
    End If
End Function

Private Function CfgNumTripe(wb As Workbook, ByVal chave As String, ByVal padrao As Double) As Double
    Dim s As String: s = Replace(CfgTxtTripe(wb, chave, ""), ",", ".")
    If s = "" Then CfgNumTripe = padrao Else CfgNumTripe = Val(s)
End Function

Private Sub GarantirChaveTripe(ws As Worksheet, ByVal chave As String, _
                               ByVal valor As String, ByVal desc As String)
    If Not ws.Columns(1).Find(chave, LookAt:=xlWhole, MatchCase:=False) Is Nothing Then Exit Sub
    Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = chave
    ws.Cells(r, 2).Value = valor
    ws.Cells(r, 3).Value = desc
End Sub

Private Function PegarAba(wb As Workbook, ByVal nome As String) As Worksheet
    On Error Resume Next
    Set PegarAba = wb.Worksheets(nome)
    On Error GoTo 0
End Function

Private Sub ApagarAba(wb As Workbook, ByVal nome As String)
    Dim ws As Worksheet: Set ws = PegarAba(wb, nome)
    If ws Is Nothing Then Exit Sub
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
End Sub

Private Sub RenomearAba(wb As Workbook, ByVal de As String, ByVal para As String)
    Dim ws As Worksheet: Set ws = PegarAba(wb, de)
    If Not ws Is Nothing Then ws.Name = para
End Sub

' Deixa visiveis so as abas do TRIPE, nesta ordem. REGRAS fica oculta.
' Nada e apagado: as demais abas (INVENTARIO, CONFIG, PAINEIS, ...) apenas
' ficam ocultas (botao direito numa guia > Reexibir).
Private Sub OrganizarAbasFinal(wb As Workbook)
    Dim ordem As Variant, ocultas As Variant
    ordem = Array("TRIPE CONCLUSAO", "TRIPE DETALHE", "ALERTAS CRITICOS", "MATERIAL vs SERVICO", _
                  "MATERIAL", "SERVICO", "SERVICO SEM MATERIAL", "ANALISE DE CA", _
                  "NAO CLASSIFICADOS", "RAZAO", "REGRAS", "ANALISE SAP x PRJ", _
                  "RESUMO SAP x PRJ", "ALERTA CRITICO")
    Dim dOrd As Object: Set dOrd = CreateObject("Scripting.Dictionary")
    Dim i As Long, ws As Worksheet, pos As Long
    For i = 0 To UBound(ordem): dOrd(UCase$(ordem(i))) = 1: Next i

    ' 1) deixa visivel o que fica (garante pelo menos uma aba visivel)
    For i = 0 To UBound(ordem)
        Set ws = PegarAba(wb, CStr(ordem(i)))
        If Not ws Is Nothing Then ws.Visible = xlSheetVisible
    Next i
    ' 2) ordem das abas
    For i = 0 To UBound(ordem)
        Set ws = PegarAba(wb, CStr(ordem(i)))
        If Not ws Is Nothing Then
            pos = pos + 1
            If ws.Index <> pos Then ws.Move Before:=wb.Worksheets(pos)
        End If
    Next i
    ' 3) oculta tudo que nao esta na lista (inclusive REGRAS)
    On Error Resume Next
    wb.Worksheets(ABA_CONC).Activate
    For Each ws In wb.Worksheets
        If Not dOrd.Exists(UCase$(ws.Name)) Or UCase$(ws.Name) = "REGRAS" Then ws.Visible = xlSheetHidden
    Next ws
    On Error GoTo 0
End Sub

' Roda a macro procurando-a no arquivo do TRIPE e depois em todas as pastas
' abertas (inclusive a Pasta de Macros Pessoal). Se nao achar em lugar
' nenhum, explica o que importar e devolve False (a execucao para).
Private Function RodarOuAvisar(ByVal macro As String, ByVal modulo As String, _
                               ByRef falhas As String) As Boolean
    Dim alvos As Collection: Set alvos = New Collection
    alvos.Add ThisWorkbook.Name
    Dim w As Workbook
    For Each w In Application.Workbooks
        If w.Name <> ThisWorkbook.Name Then alvos.Add w.Name
    Next w

    Dim a As Variant, nErr As Long, dErr As String
    For Each a In alvos
        On Error Resume Next
        Err.Clear
        Application.Run "'" & a & "'!" & macro
        nErr = Err.Number: dErr = Err.Description
        On Error GoTo 0
        If nErr = 0 Then
            If macro = "GerarRelatorio" Then mOrigemCKCP = CStr(a)
            RodarOuAvisar = True
            Exit Function
        End If
        ' 1004 = macro inexistente neste arquivo: tenta o proximo
        If nErr <> 1004 Then
            falhas = falhas & vbCrLf & "  - " & macro & ": " & dErr
            RodarOuAvisar = True
            Exit Function
        End If
    Next a

    Application.ScreenUpdating = True
    MsgBox "Nao encontrei a macro " & macro & " em nenhum arquivo aberto." & vbCrLf & vbCrLf & _
           "O TRIPE usa 3 modulos que precisam estar JUNTOS no mesmo arquivo:" & vbCrLf & _
           "   - Tripe.bas" & vbCrLf & _
           "   - inventario.bas    (macro GerarInventario)" & vbCrLf & _
           "   - AnaliseCKCP.bas   (macro GerarRelatorio)" & vbCrLf & vbCrLf & _
           "Falta importar o " & modulo & " no arquivo '" & ThisWorkbook.Name & "':" & vbCrLf & _
           "   Alt+F11 > Arquivo > Importar arquivo..." & vbCrLf & vbCrLf & _
           "Dica: guarde os 3 modulos num arquivo de macros .xlsm (ex.: TRIPE.xlsm)." & vbCrLf & _
           "Numa planilha .xlsx os modulos sao perdidos ao salvar.", _
           vbExclamation, "TRIPE - modulo faltando"
End Function

' Fecha a tela de resultado (HUD) do CKCP, se ela ficou aberta
Private Sub LimparSplashCKCP(wb As Workbook)
    On Error Resume Next
    Application.Run "'" & mOrigemCKCP & "'!FecharSplashCKCP"
    On Error GoTo 0
End Sub

Private Sub QuickSortIdx(keys() As String, idx() As Long, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As String, tk As String, ti As Long
    i = lo: j = hi
    p = keys((lo + hi) \ 2)
    Do While i <= j
        Do While keys(i) < p: i = i + 1: Loop
        Do While keys(j) > p: j = j - 1: Loop
        If i <= j Then
            tk = keys(i): keys(i) = keys(j): keys(j) = tk
            ti = idx(i): idx(i) = idx(j): idx(j) = ti
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSortIdx keys, idx, lo, j
    If i < hi Then QuickSortIdx keys, idx, i, hi
End Sub

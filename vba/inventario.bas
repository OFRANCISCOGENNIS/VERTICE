Attribute VB_Name = "inventario"
Option Explicit

' Cache da faixa de precos (evita reabrir o arquivo externo de 200k+ linhas
' a cada chamada). Vale por execucao; e limpo no inicio de GerarInventario.
Private mFaixaCache As Object

' Etapa atual (p/ diagnostico em caso de erro)
Private mStep As String

' Infra de melhorias: CONFIG editavel, log de execucao e erro granular.
Private mCfg As Object       ' CONFIG: CHAVE -> valor (string)
Private mComCrit As Object   ' tokens normalizados de familia COM critica
Private mLog As Object       ' Collection de linhas do log de execucao
Private mFalhas As String    ' etapas que falharam (erro granular, nao aborta tudo)
Private mT0 As Double        ' Timer no inicio da execucao
Private mTolAder As Double   ' tolerancia de aderencia (fracao) - hot path
Private mCaboMax As Double   ' cabo isolado isento de UC abaixo deste comprimento
Private mDifCondNu As Double ' COND NU: dif ABSOLUTA tolerada (CONFIG DIF_ABS_COND_NU)
Private mDifAbsCabo As Double ' familias DIF_ABS_FAMILIAS: dif ABSOLUTA tolerada
Private mFamDifAbs As Object ' tokens normalizados das familias com dif absoluta
Private mOddValMin As Variant ' CONFIG ODD_VALOR_MIN (Empty = padrao -100)
Private mNormCache As Object ' NormStr: texto original -> normalizado (perf, ver NormStr)

' Materiais equivalentes (aba EQUIVALENCIAS): COD normalizado -> GRUPO, e
' GRUPO -> lista legivel dos codigos ("105300003 = 105300005").
Private mEquiv As Object
Private mEquivLista As Object

' Buffer de escrita da aba ALERTA CRITICO (perf: as ~10 escritas de celula por
' alerta viram 1 unica atribuicao de Range no final). Transposto (col, linha)
' porque ReDim Preserve so cresce a ULTIMA dimensao.
Private mAlBuf() As Variant
Private mAlCap As Long

' Aba oculta com um registro de KPIs por execucao (base do "vs. execucao
' anterior" nos cards do PAINEL DO GESTOR). Nao entra na lista de abas
' recriadas a cada rodada, senao o historico se perderia.
Private Const ABA_KPI As String = "HISTORICO KPI"

' Versao da lista padrao de materiais equivalentes. A aba EQUIVALENCIAS guarda
' (no nome oculto EQUIV_VERSAO) a versao que ja recebeu; so os grupos de
' versoes mais novas sao acrescentados, sem reinserir o que o usuario apagou.
Private Const EQUIV_VERSAO_ATUAL As Long = 5

' Modo lote de CarregarBases (uma analise por base): GerarInventario nao abre
' MsgBox (o texto fica em mAvisoLote para o resumo) e nao grava HISTORICO nem
' HISTORICO KPI - compararia uma base com outra. mUltimoOk: a ultima execucao
' de GerarInventario chegou ao fim.
Private mModoLote As Boolean
Private mAvisoLote As String
Private mUltimoOk As Boolean

' Saude da execucao (bloco no PAINEL): uma linha por etapa e os avisos nao
' fatais. Fonte de precos: qual elo da cadeia (aba interna / arquivo externo)
' foi usado nesta execucao - antes uma troca de fonte passava despercebida.
Private mEtapas As Collection       ' Array(etapa, ok, segundos, detalhe)
Private mAvisosExec As String       ' avisos nao fatais (vbCrLf & "  - " & ...)
Private mFontePrecos As String      ' faixa MIN/MAX (PRECO UNITARIO e referencias)
Private mFontePrecosData As Variant ' data do arquivo externo; Empty = sem data
Private mFontePrecosRef As String   ' preco de referencia do UC SUBVALORIZADO

' Resultado da avaliacao de uma linha SAP x PRJ. Fonte UNICA da regra de
' aderencia: alimenta tanto o calculo do veredito (loop 1) quanto a
' exibicao de SIT MAT/APROVACAO/MOTIVO (loop 2) em ProcessarSAPxPRJ.
Private Type tAvalLinha
    EhAder       As Boolean  ' aderente (regras especiais ja aplicadas)
    UcZerado     As Boolean  ' material UC com LIB e PRJ zerados (devolve)
    Avaliavel    As Boolean  ' entra no veredito (UC, ou COM critico)
    CaboIsento   As Boolean  ' CABO ISOLADO < CABO_ISOLADO_MAX_M (isento)
    ReprovaLinha As Boolean  ' esta linha reprova o PEP4/PEP3
    RecalculaSit As Boolean  ' SIT MAT recalculada (cabo/condutor/ramal numerico)
    ViolaOD      As Boolean  ' viola premissa de sinal/OD (sinal por OD; valor da ODD; SAP x PRJ opostos)
    MotivoOD     As String   ' descricao da(s) violacao(oes) ("" quando ok)
End Type

' Estrutura de dados para NT.006
Private Type tMaterial
    Familia     As String   ' familia NT.006 (CRUZETA, ISOLADOR PILAR, ...)
    CodNT006    As String   ' codigo NT.006 (R-02, I-05, ...)
    DescrNT006  As String   ' descricao resumida
    EhAncora    As Boolean  ' se True, e referencia para calcular os demais
    AncoraDep   As String   ' familia ancora de que depende (se nao for ancora)
    RazaoMin    As Double   ' razao minima em relacao a ancora
    RazaoMax    As Double   ' razao maxima em relacao a ancora
    DescrRegra  As String   ' texto da regra NT.006
End Type

' Mapa NT.006: codigo SAP -> tMaterial
Private Function CriarMapaNT006() As Object

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    ' Helper para adicionar entradas
    ' CriarMapaNT006 preenche o dicionario e retorna

    ' === CRUZETAS (ancora) ===
    AddMat d, "133100007", "CRUZETA",        "R-02", "Cruzeta concreto T 1900mm", True, "", 0, 0, ""
    AddMat d, "133100001", "CRUZETA",        "R-02", "Cruzeta concreto L 1700mm", True, "", 0, 0, ""
    AddMat d, "133100006", "CRUZETA",        "R-02", "Cruzeta concreto T 2200mm", True, "", 0, 0, ""
    AddMat d, "133400012", "CRUZETA",        "R-02", "Cruzeta PRFV 90x112,5 2,4m", True, "", 0, 0, ""
    AddMat d, "133400003", "CRUZETA",        "R-02", "Cruzeta PRFV", True, "", 0, 0, ""
    AddMat d, "133400004", "CRUZETA",        "R-02", "Cruzeta PRFV", True, "", 0, 0, ""

    ' === ISOLADOR PILAR: 2-3 por cruzeta (N1 trifasico=3; monofasico=1) ===
    AddMat d, "123140003", "ISOLADOR PILAR", "I-05", "Isolador pilar 15kV M16",     False, "CRUZETA", 2, 3.5, "2-3 iso. pilar por cruzeta (N1=3)"
    AddMat d, "123140016", "ISOLADOR PILAR", "I-05", "Isolador pilar 24,2kV M16",   False, "CRUZETA", 2, 3.5, "2-3 iso. pilar por cruzeta"
    AddMat d, "123140015", "ISOLADOR PILAR", "I-05", "Isolador pilar polim. 25kV",  False, "CRUZETA", 2, 3.5, "2-3 iso. pilar por cruzeta"
    AddMat d, "123140014", "ISOLADOR PILAR", "I-05", "Isolador pilar",              False, "CRUZETA", 2, 3.5, "2-3 iso. pilar por cruzeta"

    ' === ISOLADOR SUSPENSAO (ancora para estruturas U/B) ===
    AddMat d, "123230001", "ISOL SUSPENSAO", "I-06", "Isolador suspensao polim. 15kV", True, "", 0, 0, ""
    AddMat d, "123230002", "ISOL SUSPENSAO", "I-06", "Isolador suspensao",             True, "", 0, 0, ""

    ' === ARRUELA: 2-8 por cruzeta (N1=4; N2/T2/T3/T4=8) ===
    AddMat d, "134830013", "ARRUELA", "A-02", "Arruela quad. 38x38x3mm F18", False, "CRUZETA", 2, 9, "2-8 arruelas por cruzeta (N1=4; N2=8)"
    AddMat d, "134830014", "ARRUELA", "A-02", "Arruela quad. lis 18x50x3mm", False, "CRUZETA", 2, 9, "2-8 arruelas por cruzeta"
    AddMat d, "134830051", "ARRUELA", "A-02", "Arruela red pres M18",        False, "CRUZETA", 2, 9, "2-8 arruelas por cruzeta"

    ' === PARAFUSO: 1-8 por cruzeta (varia por estrutura) ===
    AddMat d, "134700040", "PARAFUSO", "F-30", "Parafuso cab qd 125mm M16x2", False, "CRUZETA", 1, 8, "1-8 parafusos por cruzeta"
    AddMat d, "134700043", "PARAFUSO", "F-30", "Parafuso cab qd 200mm M16x2", False, "CRUZETA", 1, 8, "1-8 parafusos por cruzeta"
    AddMat d, "134700046", "PARAFUSO", "F-30", "Parafuso cab qd 250mm M16x2", False, "CRUZETA", 1, 8, "1-8 parafusos por cruzeta"
    AddMat d, "134700047", "PARAFUSO", "F-30", "Parafuso cab qd 300mm M16x2", False, "CRUZETA", 1, 8, "1-8 parafusos por cruzeta"
    AddMat d, "134700049", "PARAFUSO", "F-30", "Parafuso cab qd 400mm M16x2", False, "CRUZETA", 1, 8, "1-8 parafusos por cruzeta"
    AddMat d, "134700028", "PARAFUSO", "F-30", "Parafuso cab abaul 16x45mm",  False, "CRUZETA", 1, 8, "1-8 parafusos por cruzeta"
    AddMat d, "134700030", "PARAFUSO", "F-30", "Parafuso cab abaul 16x150mm", False, "CRUZETA", 1, 8, "1-8 parafusos por cruzeta"
    AddMat d, "134700082", "PARAFUSO", "F-30", "Parafuso rosca dupla 16x500", False, "CRUZETA", 1, 8, "1-8 parafusos por cruzeta"

    ' === PINO: 2-3 por cruzeta (1 por isolador pilar) ===
    AddMat d, "134280005", "PINO", "F-36", "Pino iso pilar autotrav M16x2", False, "CRUZETA", 2, 3.5, "~3 pinos por cruzeta (1 por isolador pilar)"
    AddMat d, "134280002", "PINO", "F-37", "Pino curto suporte topo",       False, "CRUZETA", 1, 2.5, "1-2 pinos curtos por cruzeta (est. terminal)"

    ' === PORCA: 2-6 por cruzeta ===
    AddMat d, "134800002", "PORCA", "A-21", "Porca quad. M16x2", False, "CRUZETA", 2, 6.5, "2-6 porcas por cruzeta (N2/T2/T3/T4=4)"

    ' === SELA DE CRUZETA: 2-4 por cruzeta ===
    AddMat d, "134380004", "SELA CRUZETA", "-", "Sela cruzeta 110x116mm", False, "CRUZETA", 2, 4, "2-3 selas por cruzeta trifasica"
    AddMat d, "134380005", "SELA CRUZETA", "-", "Sela cruzeta",           False, "CRUZETA", 2, 4, "2-3 selas por cruzeta"

    ' === MAO FRANCESA: 1-2 por cruzeta ===
    AddMat d, "134100006", "MAO FRANCESA", "-", "Mao francesa plana 726x38x5mm", False, "CRUZETA", 0.5, 2.5, "1-2 maos-francesas por cruzeta"

    ' === GANCHO OLHAL (ancora suspensao U/B) ===
    AddMat d, "134250015", "GANCHO OLHAL",   "F-13", "Gancho olhal 5000daN", True, "", 0, 0, ""

    ' === MANILHA / OLHAL: 1:1 com gancho ===
    AddMat d, "134200006", "MANILHA",        "F-22", "Manilha sapatilha 5000daN",   False, "GANCHO OLHAL", 0.8, 1.2, "1 manilha por ponto de suspensao"
    AddMat d, "134250023", "OLHAL PARAFUSO", "-",    "Olhal parafuso M16 5000daN",  False, "GANCHO OLHAL", 0.8, 1.2, "1 olhal por ponto de suspensao"
    AddMat d, "134740023", "PARAFUSO OLHAL", "F-34", "Parafuso olhal M16x250mm",    False, "GANCHO OLHAL", 0.8, 1.2, "1 parafuso olhal por ponto de suspensao"

    ' === HASTE DE ATERRAMENTO (ancora AT) ===
    AddMat d, "134600010", "HASTE TERRA", "F-17", "Haste aco-cobreado 14,3mm 2,4m", True, "", 0, 0, ""
    AddMat d, "134600004", "HASTE TERRA", "F-17", "Haste aco-cobreado 12,7mm 2,4m", True, "", 0, 0, ""

    ' === CONECTOR HASTE: 1:1 com haste ===
    AddMat d, "124140026", "CONEC HASTE", "M-10", "Conector cunha haste 6-16mm",     False, "HASTE TERRA", 0.8, 1.2, "1 conector por haste de aterramento"
    AddMat d, "124140078", "CONEC HASTE", "M-10", "Conector aterramento p/haste",    False, "HASTE TERRA", 0.8, 1.2, "1 conector por haste de aterramento"
    AddMat d, "124140011", "CONEC HASTE", "M-10", "Conector cunha haste",            False, "HASTE TERRA", 0.8, 1.2, "1 conector por haste de aterramento"

    ' === SUPORTE PARA-RAIOS (ancora PR) ===
    AddMat d, "134190064", "SUP PARA-RAIO", "F-47", "Suporte L para-raios 38x205", True, "", 0, 0, ""

    ' === PARA-RAIOS: 1:1 com suporte ===
    AddMat d, "104010001", "PARA-RAIO", "E-29", "Para-raios ZnO 12kV 10kA", False, "SUP PARA-RAIO", 0.8, 1.2, "1 para-raios por suporte (1:1)"
    AddMat d, "104010004", "PARA-RAIO", "E-29", "Para-raios ZnO 15kV",      False, "SUP PARA-RAIO", 0.8, 1.2, "1 para-raios por suporte (1:1)"

    ' === CHAVE FUSIVEL (ancora) ===
    AddMat d, "105300003", "CHAVE FUSIVEL", "E-09", "Chave fusivel 15kV 100A base C", True, "", 0, 0, ""

    ' === TRANSFORMADOR (ancora) ===
    AddMat d, "102100035", "TRAFO", "E-45", "Trafo trifasico 13,8kV 500kVA", True, "", 0, 0, ""
    AddMat d, "102100036", "TRAFO", "E-45", "Trafo trifasico 13,8kV",        True, "", 0, 0, ""
    AddMat d, "102100030", "TRAFO", "E-45", "Trafo monofasico",               True, "", 0, 0, ""

    ' === CONECTOR RAMAL (sem ancora definida - informativo) ===
    AddMat d, "124010010", "CONEC RAMAL", "O-02", "Conector cunha CuEst tipo II",  True, "", 0, 0, ""
    AddMat d, "124010012", "CONEC RAMAL", "O-02", "Conector cunha CuEst tipo III", True, "", 0, 0, ""

    Set CriarMapaNT006 = d

End Function

Private Sub AddMat(d As Object, cod As String, familia As String, nt006 As String, _
                   descr As String, ehAnc As Boolean, ancDep As String, _
                   rMin As Double, rMax As Double, regra As String)
    Dim tm As tMaterial
    tm.Familia    = familia
    tm.CodNT006   = nt006
    tm.DescrNT006 = descr
    tm.EhAncora   = ehAnc
    tm.AncoraDep  = ancDep
    tm.RazaoMin   = rMin
    tm.RazaoMax   = rMax
    tm.DescrRegra = regra
    d.Add cod, tm.Familia & "|" & tm.CodNT006 & "|" & tm.DescrNT006 & "|" & IIf(tm.EhAncora, "1", "0") & "|" & tm.AncoraDep & "|" & Trim$(Str$(tm.RazaoMin)) & "|" & Trim$(Str$(tm.RazaoMax)) & "|" & tm.DescrRegra
End Sub

' Desempacota tMaterial da string no Dictionary
Private Function GetMat(d As Object, cod As String) As tMaterial
    Dim p() As String, tm As tMaterial
    p = Split(CStr(d(cod)), "|")
    tm.Familia    = p(0)
    tm.CodNT006   = p(1)
    tm.DescrNT006 = p(2)
    tm.EhAncora   = (p(3) = "1")
    tm.AncoraDep  = p(4)
    tm.RazaoMin   = Val(p(5))
    tm.RazaoMax   = Val(p(6))
    tm.DescrRegra = p(7)
    GetMat = tm
End Function

' ---------------------------------------------------------------------------
'  HELPERS
' ---------------------------------------------------------------------------
' Upper-case sem acentos. E chamada centenas de milhares de vezes por execucao
' (varias por linha da base, em quase todas as etapas) e os textos se repetem
' muito (PEP, familia, tipo, situacao): o resultado fica guardado e as 21
' substituicoes rodam uma vez por texto distinto.
Private Function NormStr(s As String) As String
    If mNormCache Is Nothing Then
        Set mNormCache = CreateObject("Scripting.Dictionary")
    ElseIf mNormCache.Exists(s) Then
        NormStr = mNormCache(s) : Exit Function
    End If
    Dim r As String : r = NormStrCalc(s)
    If mNormCache.Count >= 300000 Then mNormCache.RemoveAll   ' teto de memoria
    mNormCache.Add s, r
    NormStr = r
End Function

Private Function NormStrCalc(ByVal s As String) As String
    ' Upper-case sem acentos (simplificado para os campos relevantes)
    Dim r As String : r = UCase(Trim(s))
    r = Replace(r, ".", "")
    r = Replace(r, "/", "")
    r = Replace(r, "-", "")
    r = Replace(r, "_", "")
    r = Replace(r, Chr(231), "C")   ' c cedilha
    r = Replace(r, Chr(199), "C")   ' C cedilha
    r = Replace(r, Chr(227), "A")   ' a til
    r = Replace(r, Chr(195), "A")   ' A til
    r = Replace(r, Chr(245), "O")   ' o til
    r = Replace(r, Chr(213), "O")   ' O til
    r = Replace(r, Chr(225), "A")   ' a agudo
    r = Replace(r, Chr(193), "A")   ' A agudo
    r = Replace(r, Chr(233), "E")   ' e agudo
    r = Replace(r, Chr(201), "E")   ' E agudo
    r = Replace(r, Chr(237), "I")   ' i agudo
    r = Replace(r, Chr(205), "I")   ' I agudo
    r = Replace(r, Chr(243), "O")   ' o agudo
    r = Replace(r, Chr(211), "O")   ' O agudo
    r = Replace(r, Chr(250), "U")   ' u agudo
    r = Replace(r, Chr(218), "U")   ' U agudo
    r = Replace(r, "  ", " ")
    NormStrCalc = Trim(r)
End Function

Private Function NormCod(v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then NormCod = "" : Exit Function
    Dim s As String : s = Trim(CStr(v))
    ' Remove ".0" de numeros que o Excel converte em double
    If Right(s, 2) = ".0" Then s = Left(s, Len(s) - 2)
    ' Converte notacao cientifica se necessario
    If InStr(s, "E+") > 0 Or InStr(s, "e+") > 0 Then
        On Error Resume Next
        s = Format$(CDbl(v), "0")
        On Error GoTo 0
    End If
    NormCod = s
End Function

Private Function ToNum(v As Variant) As Double
    If IsEmpty(v) Or IsNull(v) Then ToNum = 0 : Exit Function
    On Error Resume Next
    ToNum = CDbl(v)
    If Err.Number <> 0 Then ToNum = 0
    On Error GoTo 0
End Function
'=============================================================
' INVENTARIO: bloco principal
'=============================================================

Public Sub GerarInventario()

    Set mFaixaCache = Nothing   ' limpa cache de precos para recarregar a base atual
    Set mNormCache = Nothing    ' libera a memoria do cache de NormStr da execucao anterior
    mUltimoOk = False : mAvisoLote = ""

    Dim wb      As Workbook
    Dim wsBase  As Worksheet
    Dim wsDet   As Worksheet
    Dim wsCom   As Worksheet
    Dim wsAlertaC As Worksheet

    Set wb = ActiveWorkbook
    Set wsBase = AcharBaseInventario(wb)

    If wsBase Is Nothing Then
        Aviso "Nenhuma aba com as colunas 'MAT LIB SAP' (ou 'ATENDIDO SAP') e " & _
              "'MAT PRJ CAD' (ou 'TOMBADO PROJ') foi encontrada.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo ErrHandler

    LogInit
    Call GarantirConfig(wb)
    Call GarantirEquivalencias(wb)
    Dim faltam As String : faltam = ValidarBase(wsBase)
    If faltam <> "" Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Aviso "A base '" & wsBase.Name & "' esta sem colunas obrigatorias:" & vbCrLf & vbCrLf & _
              "  " & faltam & vbCrLf & vbCrLf & _
              "Ajuste os cabecalhos da linha 1 e rode novamente.", vbCritical
        Exit Sub
    End If
    LogAdd "CONFIG carregada e base validada: " & wsBase.Name

    Application.DisplayAlerts = False
    ' Cada Delete pode falhar (aba inexistente) - esperado.
    On Error Resume Next
    wb.Worksheets("PAINEL DO GESTOR").Delete
    wb.Worksheets("ANALISE SAP x PRJ").Delete
    wb.Worksheets("RESUMO SAP x PRJ").Delete
    wb.Worksheets("RACIONALIZACAO COM").Delete
    wb.Worksheets("ALERTA PEPS SEM UC").Delete
    wb.Worksheets("ALERTA CRITICO").Delete
    wb.Worksheets("SUBSTITUICOES").Delete
    wb.Worksheets("PRECO UNITARIO").Delete
    wb.Worksheets("RANKING DE RISCO").Delete
    wb.Worksheets("DIAGNOSTICO BASE").Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    Set wsDet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsDet.Name = "ANALISE SAP x PRJ"
    Set wsCom = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsCom.Name = "RACIONALIZACAO COM"
    Dim wsPU As Worksheet
    Set wsPU = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsPU.Name = "PRECO UNITARIO"
    Set wsAlertaC = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsAlertaC.Name = "ALERTA CRITICO"

    ' Cada etapa e isolada: se falhar, registra no log e segue para a proxima
    ' (as demais abas ainda sao geradas), em vez de abortar tudo.
    Dim tt As Double
    Dim wsGestor As Worksheet

    mStep = "ANALISE SAP x PRJ" : tt = Timer : On Error Resume Next : Err.Clear
    Call ProcessarSAPxPRJ(wsBase, wsDet)
    FimEtapa Err.Number, Err.Description, mStep, tt : On Error GoTo ErrHandler

    mStep = "RESUMO SAP x PRJ" : tt = Timer : On Error Resume Next : Err.Clear
    Call ProcessarResumoPEP3(wb, wsDet)
    FimEtapa Err.Number, Err.Description, mStep, tt : On Error GoTo ErrHandler

    mStep = "RACIONALIZACAO COM" : tt = Timer : On Error Resume Next : Err.Clear
    Call ProcessarCOMInventario(wsBase, wsCom)
    FimEtapa Err.Number, Err.Description, mStep, tt : On Error GoTo ErrHandler

    mStep = "PRECO UNITARIO" : tt = Timer : On Error Resume Next : Err.Clear
    Call ProcessarPrecoUnitario(wsBase, wsPU)
    FimEtapa Err.Number, Err.Description, mStep, tt : On Error GoTo ErrHandler

    mStep = "ALERTA CRITICO" : tt = Timer : On Error Resume Next : Err.Clear
    Call ProcessarAlertaCritico(wsBase, wsAlertaC)
    FimEtapa Err.Number, Err.Description, mStep, tt : On Error GoTo ErrHandler

    mStep = "DIAGNOSTICO BASE" : tt = Timer : On Error Resume Next : Err.Clear
    Call ProcessarDiagnosticoBase(wsBase, wb)
    FimEtapa Err.Number, Err.Description, mStep, tt : On Error GoTo ErrHandler

    mStep = "RANKING DE RISCO" : tt = Timer : On Error Resume Next : Err.Clear
    Call ProcessarRankingRisco(wb, wsDet, wsCom, wsPU, wsAlertaC)
    FimEtapa Err.Number, Err.Description, mStep, tt : On Error GoTo ErrHandler

    mStep = "PAINEL DO GESTOR" : tt = Timer : On Error Resume Next : Err.Clear
    Set wsGestor = ProcessarPainelGestor(wb, wsBase, wsDet, wsPU, wsAlertaC)
    FimEtapa Err.Number, Err.Description, mStep, tt : On Error GoTo ErrHandler

    mStep = "DESIGN GLOBAL" : tt = Timer : On Error Resume Next : Err.Clear
    Call AplicarDesignGlobal(wb)
    FimEtapa Err.Number, Err.Description, mStep, tt : On Error GoTo ErrHandler

    mStep = "HISTORICO" : tt = Timer : On Error Resume Next : Err.Clear
    If Not mModoLote Then Call AtualizarHistorico(wb)    ' lote: bases diferentes
    FimEtapa Err.Number, Err.Description, mStep, tt : On Error GoTo ErrHandler

    mStep = "NOMES DEFINIDOS" : tt = Timer : On Error Resume Next : Err.Clear
    Call CriarNomesDefinidos(wb)
    FimEtapa Err.Number, Err.Description, mStep, tt : On Error GoTo ErrHandler

    ' Saude da execucao no PAINEL, depois de TODAS as etapas (vai junto no
    ' arquivo exportado pelo modo lote)
    On Error Resume Next
    If Not wsGestor Is Nothing Then EscreverSaudeExecucao wsGestor, wsBase
    On Error GoTo ErrHandler

    LogAdd "Fim da geracao"
    Call GravarLog(wb)

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    On Error Resume Next
    wsGestor.Activate
    wsGestor.Range("A1").Select
    On Error GoTo ErrHandler

    mUltimoOk = True
    Dim msgFim As String
    msgFim = "Abas geradas:" & vbCrLf & _
           "  - PAINEL DO GESTOR (resumo executivo)" & vbCrLf & _
           "  - ANALISE SAP x PRJ" & vbCrLf & _
           "  - RESUMO SAP x PRJ (1 linha por obra)" & vbCrLf & _
           "  - RACIONALIZACAO COM (NT.006)" & vbCrLf & _
           "  - PRECO UNITARIO (faixa MIN/MAX)" & vbCrLf & _
           "  - ALERTA CRITICO" & vbCrLf & _
           "  - RANKING DE RISCO (score por obra)" & vbCrLf & _
           "  - DIAGNOSTICO BASE (saude dos dados 0-100)" & vbCrLf & _
           "  - CONFIG / HISTORICO (parametros e evolucao)" & vbCrLf & vbCrLf & _
           "Base: " & wsBase.Name
    If mFalhas <> "" Then
        Aviso msgFim & vbCrLf & vbCrLf & _
              "ATENCAO - etapas com falha (as demais foram geradas):" & mFalhas, vbExclamation
    Else
        Aviso msgFim, vbInformation
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Aviso "Erro " & Err.Number & " na etapa [" & mStep & "]:" & vbCrLf & _
          Err.Description, vbCritical
End Sub

' Mensagem de GerarInventario. No modo lote (CarregarBases, uma analise por
' base) nao interrompe: guarda o texto para o resumo final.
Private Sub Aviso(ByVal msg As String, ByVal estilo As VbMsgBoxStyle)
    If mModoLote Then
        mAvisoLote = msg
    Else
        MsgBox msg, estilo, "Analise Inventario"
    End If
End Sub


Private Sub ProcessarSAPxPRJ(wsBase As Worksheet, wsDet As Worksheet)

    Dim azul As Long, verde As Long
    azul = RGB(10, 37, 64)
    verde = RGB(0, 179, 136)

    Dim HR As Long, DS As Long
    HR = 9
    DS = 10

    Dim outHdr As Variant
    outHdr = Array("PEP3NIVEL", "PEP4NIVEL", "NOTA", "CLASSE", "COD MAT", "VALOR", _
                   "DESC MAT", "UND", "MAT LIB SAP", "MAT PRJ CAD", "TIPO", "FAMILIA", "SIT MAT", "APROVACAO", "MOTIVO", "MOTIVO DEVOLUCAO PEP3")
    Dim nCols As Long : nCols = 16

    Dim lastC As Long
    lastC = wsBase.Cells(1, wsBase.Columns.Count).End(xlToLeft).Column

    Dim srcIdx(1 To 16) As Long
    Dim oc As Long, c As Long
    Dim mapa As Object : Set mapa = MapaColunas(wsBase)
    For oc = 1 To nCols
        srcIdx(oc) = ColIdx(mapa, NormStr(CStr(outHdr(oc - 1))))
    Next oc
    ' TIPO PEP nao vai para a aba de saida, mas classifica o tipo de OD quando o
    ' PEP4 nao tem sufixo (.I/.S/.D) - usado na premissa de sinal
    Dim iTpPep As Long : iTpPep = ColIdx(mapa, "TIPO PEP")

    Dim refCol As Long : refCol = srcIdx(5)
    If refCol = 0 Then refCol = 1
    Dim lastRow As Long
    lastRow = wsBase.Cells(wsBase.Rows.Count, refCol).End(xlUp).Row

    Dim nRows As Long : nRows = 0
    Dim nPep4 As Long, nAprov As Long, nReprov As Long
    Dim sumValor As Double, sumNaoAder As Double
    Dim dot As String, okIco As String, noIco As String
    dot = ChrW(&H25CF)
    okIco = ChrW(&H2705)
    noIco = ChrW(&H274C)

    If lastRow >= 2 Then
        Dim baseData As Variant
        baseData = wsBase.Range(wsBase.Cells(2, 1), wsBase.Cells(lastRow, lastC)).Value
        nRows = lastRow - 1

        Dim aprova As Object : Set aprova = CreateObject("Scripting.Dictionary")
        Dim temUC As Object : Set temUC = CreateObject("Scripting.Dictionary")
        Dim allP4 As Object : Set allP4 = CreateObject("Scripting.Dictionary")
        Dim p4pep3 As Object : Set p4pep3 = CreateObject("Scripting.Dictionary")
        Dim pep3RepFam As Object : Set pep3RepFam = CreateObject("Scripting.Dictionary")  ' PEP3 -> familias que reprovaram
        Dim p3RepLinha As Object : Set p3RepLinha = CreateObject("Scripting.Dictionary")  ' PEP3 reprovado por regra de linha (qualquer item)
        Dim i As Long, pep4 As String, tipoR As String, sitR As String, pep3R As String
        Dim famR As String, famRaw As String, libV As Variant, prjV As Variant

        ' ---- Pre-passe: regra do UC zerado por OBRA (PEP3) ----
        ' A devolucao por valores zerados so vale quando TODOS os materiais UC
        ' do PEP3 estao zerados. Se apenas parte esta zerada, essas linhas nao
        ' reprovam e o PEP3 fica APROVADO quando as demais regras passam.
        Dim p3nUC As Object : Set p3nUC = CreateObject("Scripting.Dictionary")
        Dim p3nUC0 As Object : Set p3nUC0 = CreateObject("Scripting.Dictionary")
        For i = 1 To nRows
            tipoR = "" : pep3R = ""
            If srcIdx(11) > 0 Then tipoR = NormStr(CStr(baseData(i, srcIdx(11))))
            If tipoR = "UC" Then
                If srcIdx(1) > 0 Then pep3R = NormStr(CStr(baseData(i, srcIdx(1))))
                libV = "" : prjV = ""
                If srcIdx(9) > 0 Then libV = baseData(i, srcIdx(9))
                If srcIdx(10) > 0 Then prjV = baseData(i, srcIdx(10))
                p3nUC(pep3R) = p3nUC(pep3R) + 1
                If MaterialUCZerado(tipoR, libV, prjV) Then p3nUC0(pep3R) = p3nUC0(pep3R) + 1
            End If
        Next i

        ' ---- Pre-passe: grupos de materiais EQUIVALENTES em cada PEP4 ----
        ' Quando 2+ codigos do mesmo grupo (aba EQUIVALENCIAS) aparecem no mesmo
        ' PEP4, houve troca de material: todos sao aderentes so se a quantidade
        ' TOTAL do grupo e a mesma no SAP e no PRJ, cada linha com o proprio
        ' codigo e descricao. Guarda os totais e os codigos do grupo (motivo).
        Dim eqLib As Object : Set eqLib = CreateObject("Scripting.Dictionary")
        Dim eqPrj As Object : Set eqPrj = CreateObject("Scripting.Dictionary")
        Dim eqCods As Object : Set eqCods = CreateObject("Scripting.Dictionary")
        Dim eqK As String, eqL As Double, eqP As Double, eqCod As String, eqDsc As String
        For i = 1 To nRows
            eqK = ChaveEquiv(baseData, i, srcIdx(2), srcIdx(5))
            If eqK <> "" Then
                eqL = 0 : eqP = 0
                If srcIdx(9) > 0 Then eqL = Val0(baseData(i, srcIdx(9)))
                If srcIdx(10) > 0 Then eqP = Val0(baseData(i, srcIdx(10)))
                eqLib(eqK) = eqLib(eqK) + eqL
                eqPrj(eqK) = eqPrj(eqK) + eqP
                If Not eqCods.Exists(eqK) Then eqCods.Add eqK, CreateObject("Scripting.Dictionary")
                eqCod = NormCod(baseData(i, srcIdx(5)))      ' ChaveEquiv ja descartou erro
                eqDsc = ""
                If srcIdx(7) > 0 Then
                    If Not IsError(baseData(i, srcIdx(7))) Then eqDsc = Trim(CStr(baseData(i, srcIdx(7))))
                End If
                If Not eqCods(eqK).Exists(eqCod) Then eqCods(eqK).Add eqCod, eqDsc
            End If
        Next i

        For i = 1 To nRows
            pep4 = "" : tipoR = "" : sitR = "" : famR = "" : famRaw = "" : pep3R = ""
            If srcIdx(2) > 0 Then pep4 = NormStr(CStr(baseData(i, srcIdx(2))))
            If srcIdx(1) > 0 Then pep3R = NormStr(CStr(baseData(i, srcIdx(1))))
            If srcIdx(11) > 0 Then tipoR = NormStr(CStr(baseData(i, srcIdx(11))))
            If srcIdx(13) > 0 Then sitR = NormStr(CStr(baseData(i, srcIdx(13))))
            If srcIdx(12) > 0 Then famR = NormStr(CStr(baseData(i, srcIdx(12))))
            If srcIdx(12) > 0 Then famRaw = Trim(CStr(baseData(i, srcIdx(12))))
            libV = "" : prjV = ""
            If srcIdx(9) > 0 Then libV = baseData(i, srcIdx(9))
            If srcIdx(10) > 0 Then prjV = baseData(i, srcIdx(10))
            ' Avaliacao unificada da linha (mesma funcao usada no loop de exibicao).
            ' UC zerado so reprova quando a obra (PEP3) esta TODA zerada.
            ' tipo de OD pelo PEP4 CRU (o pep4 acima passou por NormStr, sem o ponto)
            Dim p4Raw As String, tpPepR As String, valR As Variant
            p4Raw = "" : tpPepR = "" : valR = 0
            If srcIdx(2) > 0 Then p4Raw = CStr(baseData(i, srcIdx(2)))
            If iTpPep > 0 Then tpPepR = CStr(baseData(i, iTpPep))
            If srcIdx(6) > 0 Then valR = baseData(i, srcIdx(6))
            ' troca por material EQUIVALENTE no PEP4: aderente so se a QUANTIDADE TOTAL
            ' do grupo e a mesma no SAP e no PRJ (sem tolerancia)
            Dim eqK1 As String : eqK1 = ChaveEquiv(baseData, i, srcIdx(2), srcIdx(5))
            Dim eqAt1 As Boolean : eqAt1 = False
            If eqK1 <> "" Then eqAt1 = (eqCods(eqK1).Count >= 2)
            Dim libAv As Variant, prjAv As Variant, sitAv As String
            libAv = libV : prjAv = prjV : sitAv = sitR
            If eqAt1 Then
                libAv = eqLib(eqK1) : prjAv = eqPrj(eqK1) : sitAv = SitPorQuantidade(libAv, prjAv)
            End If
            Dim av As tAvalLinha
            av = AvaliarLinha(tipoR, famR, libAv, prjAv, sitAv, TodosUCZerados(p3nUC, p3nUC0, pep3R), _
                              ClassificarOD(p4Raw, tpPepR), valR, libV, prjV)
            If eqAt1 Then AjustarAderEquiv av, libAv, prjAv
            ' Regras de LINHA que reprovam o PEP3 inteiro, seja qual for o tipo do
            ' item (o veredito normal so conta UC e COM critico: um COM comum assim
            ' ficava NAO ADERENTE com a obra APROVADA):
            '  - ODI com qualquer VALOR negativo
            '  - MAT LIB SAP e MAT PRJ CAD da propria linha com sinais opostos
            If pep3R <> "" Then
                Dim fBase As String : fBase = IIf(famRaw <> "", famRaw, "(sem familia)")
                If NegativoEmODI(ClassificarOD(p4Raw, tpPepR), valR) Then
                    p3RepLinha(pep3R) = True
                    RegistrarFamRep pep3RepFam, pep3R, fBase & " (valor negativo em ODI)"
                End If
                If SinaisOpostos(libV, prjV) <> "" Then
                    p3RepLinha(pep3R) = True
                    RegistrarFamRep pep3RepFam, pep3R, fBase & " (SAP x PRJ com sinais opostos)"
                End If
            End If
            If pep4 <> "" And Not allP4.Exists(pep4) Then allP4.Add pep4, True
            If pep4 <> "" And Not p4pep3.Exists(pep4) Then p4pep3.Add pep4, pep3R
            If pep4 <> "" And av.Avaliavel And Not av.CaboIsento Then
                If Not temUC.Exists(pep4) Then temUC.Add pep4, True
                If Not aprova.Exists(pep4) Then aprova.Add pep4, True
                If av.ReprovaLinha Then
                    aprova(pep4) = False
                    ' registra a familia culpada no PEP3 (sem duplicar)
                    Dim fLbl As String : fLbl = IIf(famRaw <> "", famRaw, "(sem familia)")
                    If av.UcZerado Then fLbl = fLbl & " (valores zerados)"
                    If av.ViolaOD Then fLbl = fLbl & " (regra de OD)"
                    If Not pep3RepFam.Exists(pep3R) Then
                        pep3RepFam.Add pep3R, fLbl
                    ElseIf InStr(pep3RepFam(pep3R), fLbl) = 0 Then
                        pep3RepFam(pep3R) = pep3RepFam(pep3R) & ", " & fLbl
                    End If
                End If
            End If
            If srcIdx(6) > 0 Then
                Dim vv As Double : vv = 0
                If IsNumeric(baseData(i, srcIdx(6))) Then vv = CDbl(baseData(i, srcIdx(6)))
                sumValor = sumValor + vv
                If av.ReprovaLinha Then sumNaoAder = sumNaoAder + Abs(vv)
            End If
        Next i

        nPep4 = allP4.Count

        ' Rollup PEP3: se qualquer PEP4 do PEP3 reprova, todo o PEP3 reprova
        Dim pep3Rep As Object : Set pep3Rep = CreateObject("Scripting.Dictionary")
        Dim pep3UC As Object : Set pep3UC = CreateObject("Scripting.Dictionary")
        Dim k4 As Variant, p3 As String
        For Each k4 In temUC.Keys
            p3 = CStr(p4pep3(k4))
            If Not pep3UC.Exists(p3) Then pep3UC.Add p3, True
            If Not pep3Rep.Exists(p3) Then pep3Rep.Add p3, False
            If Not aprova(k4) Then pep3Rep(p3) = True
        Next k4
        ' Regras de linha (ODI negativa; SAP x PRJ opostos): PEP3 reprovado mesmo
        ' sem item UC/COM critico reprovado (e se nao tinha item avaliavel, deixa
        ' de ser "SEM UC")
        Dim kN As Variant
        For Each kN In p3RepLinha.Keys
            If Not pep3UC.Exists(kN) Then pep3UC.Add kN, True
            pep3Rep(kN) = True
        Next kN

        Dim kk As Variant
        For Each kk In pep3Rep.Keys
            If pep3Rep(kk) Then nReprov = nReprov + 1 Else nAprov = nAprov + 1
        Next kk

        Dim outArr() As Variant : ReDim outArr(1 To nRows, 1 To nCols)
        ' Helpers p/ KPIs DINAMICOS (escritos nas colunas ocultas AE:AH):
        ' 1=1a linha do PEP4 | 2=1a linha do PEP3 avaliavel | 3=PEP3 reprovado | 4=valor nao aderente
        Dim hlp() As Double : ReDim hlp(1 To nRows, 1 To 4)
        Dim seen4 As Object : Set seen4 = CreateObject("Scripting.Dictionary")
        Dim seen3 As Object : Set seen3 = CreateObject("Scripting.Dictionary")
        Dim pk As String, rawSit As String
        For i = 1 To nRows
            For oc = 1 To nCols
                If oc <> 13 And oc <> 14 And oc <> 15 And srcIdx(oc) > 0 Then outArr(i, oc) = baseData(i, srcIdx(oc))
            Next oc
            rawSit = "" : sitR = "" : famR = ""
            If srcIdx(13) > 0 Then rawSit = Trim(CStr(baseData(i, srcIdx(13))))
            sitR = NormStr(rawSit)
            If srcIdx(12) > 0 Then famR = NormStr(CStr(baseData(i, srcIdx(12))))
            libV = "" : prjV = ""
            If srcIdx(9) > 0 Then libV = baseData(i, srcIdx(9))
            If srcIdx(10) > 0 Then prjV = baseData(i, srcIdx(10))
            ' TIPO deste item + avaliacao unificada (mesma funcao do loop 1).
            ' UC zerado so reprova quando a obra (PEP3) esta TODA zerada.
            Dim tipoRow As String : tipoRow = ""
            If srcIdx(11) > 0 Then tipoRow = NormStr(CStr(baseData(i, srcIdx(11))))
            Dim p3row As String : p3row = ""
            If srcIdx(1) > 0 Then p3row = NormStr(CStr(baseData(i, srcIdx(1))))
            Dim p4RawRow As String, tpPepRow As String, valRow As Variant
            p4RawRow = "" : tpPepRow = "" : valRow = 0
            If srcIdx(2) > 0 Then p4RawRow = CStr(baseData(i, srcIdx(2)))
            If iTpPep > 0 Then tpPepRow = CStr(baseData(i, iTpPep))
            If srcIdx(6) > 0 Then valRow = baseData(i, srcIdx(6))
            ' grupo EQUIVALENTE: a partir daqui libV/prjV/sitR sao do TOTAL do grupo
            ' no PEP4 (avaliacao e motivo); as colunas de quantidade da linha ja
            ' foram copiadas acima com os valores proprios do item
            Dim libRow As Variant : libRow = libV
            Dim prjRow As Variant : prjRow = prjV
            Dim eqK2 As String : eqK2 = ChaveEquiv(baseData, i, srcIdx(2), srcIdx(5))
            Dim eqAt2 As Boolean : eqAt2 = False
            If eqK2 <> "" Then eqAt2 = (eqCods(eqK2).Count >= 2)
            If eqAt2 Then
                libV = eqLib(eqK2) : prjV = eqPrj(eqK2) : sitR = SitPorQuantidade(libV, prjV)
            End If
            Dim avRow As tAvalLinha
            avRow = AvaliarLinha(tipoRow, famR, libV, prjV, sitR, TodosUCZerados(p3nUC, p3nUC0, p3row), _
                                 ClassificarOD(p4RawRow, tpPepRow), valRow, libRow, prjRow)
            If eqAt2 Then AjustarAderEquiv avRow, libV, prjV

            Dim sitText As String
            If avRow.UcZerado Or avRow.ViolaOD Then
                sitText = "NAO ADERENTE"
            ElseIf eqAt2 Then
                sitText = IIf(avRow.EhAder, "ADERENTE", "NAO ADERENTE")
            ElseIf avRow.RecalculaSit Then
                sitText = IIf(avRow.EhAder, "ADERENTE", "NAO ADERENTE")
            Else
                sitText = rawSit
            End If
            outArr(i, 13) = dot & " " & sitText

            ' PRIORIDADE 1: se o PEP3 esta reprovado, TODA linha vira REPROVADO
            ' (inclusive CABO ISOLADO que seria isento) - arrastado pela reprovacao do PEP3.
            ' Nota: And do VBA NAO curto-circuita; acessar chave inexistente de
            ' Scripting.Dictionary CRIA a chave. Por isso o teste e aninhado.
            Dim motivo As String
            Dim p3Reprovado As Boolean : p3Reprovado = False
            If pep3UC.Exists(p3row) Then p3Reprovado = CBool(pep3Rep(p3row))
            If p3Reprovado Then
                outArr(i, 14) = noIco & " REPROVADO"
                If avRow.UcZerado Then
                    motivo = "Obra com TODOS os materiais UC zerados (MAT LIB SAP=0 e MAT PRJ CAD=0) - devolvida"
                ElseIf avRow.ViolaOD Then
                    motivo = "Este item nao aderente: " & avRow.MotivoOD
                ElseIf tipoRow = "UC" And Not avRow.EhAder And sitR <> "NULO" Then
                    motivo = "Este item (UC) nao aderente: SAP=" & CStr(libV) & " / PRJ=" & CStr(prjV) & FmtDif(libV, prjV)
                ElseIf tipoRow = "COM" And EhComCritico(famR) And Not avRow.EhAder And sitR <> "NULO" Then
                    motivo = "Este item (COM critico - " & famR & ") nao aderente: SAP=" & CStr(libV) & " / PRJ=" & CStr(prjV) & FmtDif(libV, prjV)
                ElseIf avRow.CaboIsento Then
                    motivo = "Cabo isolado arrastado | Familia reprovada: " & pep3RepFam(p3row)
                Else
                    motivo = "Arrastado pela reprovacao do PEP3 | Familia reprovada: " & pep3RepFam(p3row)
                End If
            ElseIf avRow.CaboIsento Then
                outArr(i, 14) = okIco & " APROVADO"
                motivo = "Cabo isolado < " & Format(IIf(mCaboMax > 0, mCaboMax, 15), "0.#") & "m - isento de UC"
            ElseIf Not pep3UC.Exists(p3row) Then
                outArr(i, 14) = "SEM UC"
                motivo = "PEP3 sem UC nem COM critico (CH FUS/PARA RAIO) p/ avaliar"
            Else
                outArr(i, 14) = okIco & " APROVADO"
                If MaterialUCZerado(tipoRow, libV, prjV) Then
                    motivo = "Item UC zerado ignorado (obra nao esta toda zerada) | Demais itens aderentes"
                Else
                    motivo = "Todos os itens avaliados do PEP3 aderentes"
                End If
            End If
            ' item com sinal invertido em PEP3 nao reprovado (ex.: COM nao critico,
            ' que nao entra no veredito): a linha e NAO ADERENTE e o motivo diz por que
            If avRow.ViolaOD And Not p3Reprovado Then
                motivo = motivo & " | Item NAO ADERENTE: " & avRow.MotivoOD
            End If
            If eqAt2 Then
                motivo = motivo & " | Material equivalente: " & TextoEquiv(eqCods(eqK2)) & _
                         " | total do grupo no PEP4 SAP=" & CStr(libV) & " / PRJ=" & CStr(prjV) & _
                         IIf(avRow.EhAder, "", FmtDif(libV, prjV))
            End If
            outArr(i, 15) = motivo
            outArr(i, 16) = IIf(p3Reprovado, "Devolvido por divergencia de " & pep3RepFam(p3row), "")

            ' --- helpers dos KPIs dinamicos ---
            Dim p4key As String : p4key = ""
            If srcIdx(2) > 0 Then p4key = NormStr(CStr(baseData(i, srcIdx(2))))
            If p4key <> "" Then
                If Not seen4.Exists(p4key) Then seen4.Add p4key, True : hlp(i, 1) = 1
            End If
            If p3row <> "" Then
                If pep3UC.Exists(p3row) Then
                    If Not seen3.Exists(p3row) Then
                        seen3.Add p3row, True
                        hlp(i, 2) = 1
                        If p3Reprovado Then hlp(i, 3) = 1
                    End If
                End If
            End If
            If srcIdx(6) > 0 Then
                If InStr(NormStr(sitText), "NAO ADER") > 0 Then hlp(i, 4) = Abs(Val0(baseData(i, srcIdx(6))))
            End If
        Next i
        wsDet.Range(wsDet.Cells(DS, 1), wsDet.Cells(DS + nRows - 1, nCols)).Value = outArr

        ' Colunas auxiliares OCULTAS (fora do AutoFilter A:P):
        ' AD=VIS (1 se a linha esta visivel no filtro), AE..AH = helpers acima.
        ' VIS usa SUBTOTAL(102) sobre AE, que e sempre numerica (robusto a vazios).
        wsDet.Range(wsDet.Cells(DS, 31), wsDet.Cells(DS + nRows - 1, 34)).Value = hlp
        wsDet.Range(wsDet.Cells(DS, 30), wsDet.Cells(DS + nRows - 1, 30)).Formula = _
            "=SUBTOTAL(102,$AE" & DS & ")"
        wsDet.Cells(HR, 30).Value = "vis" : wsDet.Cells(HR, 31).Value = "p4"
        wsDet.Cells(HR, 32).Value = "p3" : wsDet.Cells(HR, 33).Value = "rep"
        wsDet.Cells(HR, 34).Value = "vna"
        wsDet.Range(wsDet.Cells(HR, 30), wsDet.Cells(HR, 34)).Font.Color = RGB(200, 200, 200)
        wsDet.Range(wsDet.Columns(30), wsDet.Columns(34)).Hidden = True
    End If

    ' Best-effort (so formatacao visual): falha aqui nao altera a analise.
    On Error Resume Next

    With wsDet.Range("A1:P2")
        .Merge
        .Value = ChrW(&H26A1) & " ANALISE SAP x PROJETO"
        .Font.Name = "Segoe UI" : .Font.Size = 20 : .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = azul
        .HorizontalAlignment = xlLeft : .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    With wsDet.Range("A3:P3")
        .Merge
        .Value = "Inventario Inteligente  |  Atualizacao: " & Format(Now, "dd" & Chr(47) & "mm" & Chr(47) & "yyyy hh:nn") & "  |  Os cards reagem aos filtros"
        .Font.Name = "Segoe UI" : .Font.Size = 9 : .Font.Italic = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = verde
        .HorizontalAlignment = xlLeft : .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    wsDet.Rows(1).RowHeight = 20 : wsDet.Rows(2).RowHeight = 20
    wsDet.Rows(3).RowHeight = 16 : wsDet.Rows(4).RowHeight = 6

    Dim cardCol As Variant, cardLbl As Variant, cardVal As Variant, cardClr As Variant
    cardCol = Array(1, 4, 7, 10, 13)
    cardLbl = Array("ODs (PEP4)", "PEP3 APROVADOS", "PEP3 REPROVADOS", "VALOR SAP", "VALOR NAO ADER.")
    Dim pctA As Double, pctR As Double, denom As Long
    denom = nAprov + nReprov
    If denom > 0 Then
        pctA = nAprov/denom
        pctR = nReprov/denom
    End If
    cardVal = Array(Format(nPep4, "#,##0"), Format(pctA, "0%"), Format(pctR, "0%"), _
                    FmtKPI(sumValor), FmtKPI(sumNaoAder))
    cardClr = Array(azul, verde, RGB(192, 0, 0), azul, RGB(192, 87, 0))
    Dim cidx As Long, cc As Long
    For cidx = 0 To 4
        cc = cardCol(cidx)
        With wsDet.Range(wsDet.Cells(5, cc), wsDet.Cells(5, cc + 1))
            .Merge : .Value = cardLbl(cidx)
            .Font.Name = "Segoe UI" : .Font.Size = 9 : .Font.Bold = True
            .Font.Color = RGB(255, 255, 255) : .Interior.Color = cardClr(cidx)
            .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
        End With
        With wsDet.Range(wsDet.Cells(6, cc), wsDet.Cells(6, cc + 1))
            .Merge : .Value = cardVal(cidx)
            .Font.Name = "Segoe UI" : .Font.Size = 18 : .Font.Bold = True
            .Font.Color = cardClr(cidx) : .Interior.Color = RGB(248, 249, 250)
            .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous : .Borders.Color = RGB(220, 220, 220)
        End With
    Next cidx
    wsDet.Rows(5).RowHeight = 16 : wsDet.Rows(6).RowHeight = 34
    wsDet.Rows(7).RowHeight = 6

    ' KPIs DINAMICOS: os cards reagem ao AutoFilter. As formulas usam as
    ' colunas auxiliares ocultas (AD:AH) escritas junto com os dados.
    If nRows > 0 Then
        Dim ultL As String : ultL = CStr(DS + nRows - 1)
        Dim rVis As String : rVis = "$AD$" & DS & ":$AD$" & ultL
        Dim rP4 As String  : rP4 = "$AE$" & DS & ":$AE$" & ultL
        Dim rP3 As String  : rP3 = "$AF$" & DS & ":$AF$" & ultL
        Dim rRep As String : rRep = "$AG$" & DS & ":$AG$" & ultL
        Dim rVna As String : rVna = "$AH$" & DS & ":$AH$" & ultL
        Dim fmtRS As String
        fmtRS = "[>=1000000]""R$"" #,##0.0,,"" MM"";[>=1000]""R$"" #,##0.0,"" mil"";""R$"" #,##0.00"
        With wsDet.Cells(6, 1)   ' ODs (PEP4) visiveis
            .Formula = "=SUMPRODUCT(" & rVis & "," & rP4 & ")"
            .NumberFormat = "#,##0"
        End With
        With wsDet.Cells(6, 4)   ' % PEP3 aprovados (entre os visiveis)
            .Formula = "=IFERROR((SUMPRODUCT(" & rVis & "," & rP3 & ")-SUMPRODUCT(" & _
                       rVis & "," & rRep & "))/SUMPRODUCT(" & rVis & "," & rP3 & "),0)"
            .NumberFormat = "0%"
        End With
        With wsDet.Cells(6, 7)   ' % PEP3 reprovados (entre os visiveis)
            .Formula = "=IFERROR(SUMPRODUCT(" & rVis & "," & rRep & ")/SUMPRODUCT(" & _
                       rVis & "," & rP3 & "),0)"
            .NumberFormat = "0%"
        End With
        With wsDet.Cells(6, 10)  ' Valor SAP visivel
            .Formula = "=SUBTOTAL(109,$F$" & DS & ":$F$" & ultL & ")"
            .NumberFormat = fmtRS
        End With
        With wsDet.Cells(6, 13)  ' Valor nao aderente visivel
            .Formula = "=SUMPRODUCT(" & rVis & "," & rVna & ")"
            .NumberFormat = fmtRS
        End With
    End If

    Dim navCol As Variant, navTxt As Variant, navSheet As Variant
    navCol = Array(1, 5, 9)
    navTxt = Array("SAP x PROJETO", "RACIONALIZACAO COM", "ALERTAS CRITICOS")
    navSheet = Array("ANALISE SAP x PRJ", "RACIONALIZACAO COM", "ALERTA CRITICO")
    Dim ni As Long
    For ni = 0 To 2
        With wsDet.Range(wsDet.Cells(8, navCol(ni)), wsDet.Cells(8, navCol(ni) + 3))
            .Merge
            .Interior.Color = IIf(ni = 0, verde, azul)
            .Borders.LineStyle = xlContinuous : .Borders.Color = RGB(255, 255, 255)
        End With
        wsDet.Hyperlinks.Add Anchor:=wsDet.Cells(8, navCol(ni)), Address:="", _
            SubAddress:="'" & navSheet(ni) & "'!A1", TextToDisplay:=navTxt(ni)
        With wsDet.Cells(8, navCol(ni))
            .Font.Name = "Segoe UI" : .Font.Size = 10 : .Font.Bold = True
            .Font.Color = RGB(255, 255, 255) : .Font.Underline = xlUnderlineStyleNone
            .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
        End With
    Next ni
    wsDet.Rows(8).RowHeight = 24
    On Error GoTo 0

    For oc = 1 To nCols
        With wsDet.Cells(HR, oc)
            .Value = outHdr(oc - 1)
            .Font.Name = "Segoe UI" : .Font.Bold = True : .Font.Size = 10
            .Font.Color = RGB(255, 255, 255)
            Select Case oc
                Case 9:  .Interior.Color = RGB(0, 128, 0)
                Case 10: .Interior.Color = RGB(255, 25, 25)
                Case 13: .Interior.Color = RGB(89, 89, 89)
                Case 15: .Interior.Color = RGB(112, 48, 160)
                Case 16: .Interior.Color = RGB(192, 0, 0)
                Case Else: .Interior.Color = azul
            End Select
            .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous : .Borders.Color = RGB(170, 170, 170)
        End With
    Next oc
    wsDet.Rows(HR).RowHeight = 26

    Dim widths As Variant
    widths = Array(24.6, 26.6, 10.6, 11.7, 10.6, 11, 50.8, 5.2, 13.4, 14.5, 5.5, 24.3, 18, 15, 48, 48)
    For oc = 1 To nCols
        wsDet.Columns(oc).ColumnWidth = widths(oc - 1)
    Next oc

    If nRows > 0 Then
        Dim lastD As Long : lastD = DS + nRows - 1
        With wsDet.Range(wsDet.Cells(DS, 1), wsDet.Cells(lastD, nCols))
            .Font.Name = "Segoe UI" : .Font.Size = 9
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous : .Borders.Color = RGB(225, 225, 225)
        End With

        ' Best-effort (so formatacao visual): falha aqui nao altera a analise.
        On Error Resume Next
        Dim bRng As Range : Set bRng = wsDet.Range(wsDet.Cells(DS, 1), wsDet.Cells(lastD, 11))
        bRng.FormatConditions.Delete
        With bRng.FormatConditions.Add(xlExpression, , "MOD(ROW(),2)=0")
            .Interior.Color = RGB(244, 247, 250)
        End With

        Dim fL As Range : Set fL = wsDet.Range(wsDet.Cells(DS, 12), wsDet.Cells(lastD, 12))
        fL.FormatConditions.Delete
        Dim r0 As String : r0 = "$L" & DS
        With fL.FormatConditions.Add(xlExpression, , "ISNUMBER(SEARCH(""ARMA""," & r0 & "))")
            .Interior.Color = RGB(189, 215, 238)
        End With
        With fL.FormatConditions.Add(xlExpression, , "ISNUMBER(SEARCH(""CRUZETA""," & r0 & "))")
            .Interior.Color = RGB(198, 239, 206)
        End With
        With fL.FormatConditions.Add(xlExpression, , "ISNUMBER(SEARCH(""ISOLADOR""," & r0 & "))")
            .Interior.Color = RGB(225, 213, 231)
        End With
        With fL.FormatConditions.Add(xlExpression, , "ISNUMBER(SEARCH(""TRAFO""," & r0 & "))")
            .Interior.Color = RGB(252, 228, 214)
        End With
        With fL.FormatConditions.Add(xlExpression, , "ISNUMBER(SEARCH(""POSTE""," & r0 & "))")
            .Interior.Color = RGB(217, 217, 217)
        End With

        Dim sR As Range : Set sR = wsDet.Range(wsDet.Cells(DS, 13), wsDet.Cells(lastD, 13))
        sR.FormatConditions.Delete
        Dim rm As String : rm = "$M" & DS
        With sR.FormatConditions.Add(xlExpression, , "ISNUMBER(SEARCH(""NAO ADER""," & rm & "))")
            .Interior.Color = RGB(255, 199, 206) : .Font.Color = RGB(156, 0, 6) : .StopIfTrue = True
        End With
        With sR.FormatConditions.Add(xlExpression, , "ISNUMBER(SEARCH(""NULO""," & rm & "))")
            .Interior.Color = RGB(255, 235, 156) : .Font.Color = RGB(156, 87, 0) : .StopIfTrue = True
        End With
        With sR.FormatConditions.Add(xlExpression, , "ISNUMBER(SEARCH(""ADERENTE""," & rm & "))")
            .Interior.Color = RGB(198, 239, 206) : .Font.Color = RGB(0, 97, 0)
        End With
        sR.HorizontalAlignment = xlCenter

        ' APROVACAO: coloracao DIRETA (FC com formula em ingles falha no Excel PT-BR).
        ' Le o veredito do array em memoria (sem reler a planilha) e pinta BLOCOS
        ' de linhas consecutivas com o mesmo status - 1 chamada por bloco.
        Dim rrA As Long, rIni As Long, avA As String
        wsDet.Range(wsDet.Cells(DS, 14), wsDet.Cells(lastD, 14)).HorizontalAlignment = xlCenter
        rrA = DS
        Do While rrA <= lastD
            avA = CStr(outArr(rrA - DS + 1, 14))
            rIni = rrA
            Do While rrA < lastD
                If CStr(outArr(rrA - DS + 2, 14)) <> avA Then Exit Do
                rrA = rrA + 1
            Loop
            With wsDet.Range(wsDet.Cells(rIni, 14), wsDet.Cells(rrA, 14))
                If InStr(avA, "REPROVADO") > 0 Then
                    .Interior.Color = RGB(192, 0, 0) : .Font.Color = RGB(255, 255, 255) : .Font.Bold = True
                ElseIf InStr(avA, "APROVADO") > 0 Then
                    .Interior.Color = RGB(0, 176, 80) : .Font.Color = RGB(255, 255, 255) : .Font.Bold = True
                Else
                    .Interior.Color = RGB(217, 217, 217) : .Font.Color = RGB(64, 64, 64)
                End If
            End With
            rrA = rrA + 1
        Loop

        Dim vR As Range : Set vR = wsDet.Range(wsDet.Cells(DS, 6), wsDet.Cells(lastD, 6))
        vR.FormatConditions.Delete
        With vR.FormatConditions.AddDatabar
            .BarColor.Color = RGB(0, 179, 136)
            .BarFillType = xlDataBarFillGradient
        End With

        Dim cRng As Range : Set cRng = wsDet.Range(wsDet.Cells(DS, 1), wsDet.Cells(lastD, nCols))
        Dim corte As Double : corte = CfgD("CORTE_DIVERG_QTD", 20)
        With cRng.FormatConditions.Add(xlExpression, , "OR($J" & DS & ">" & corte & ",$I" & DS & "<-" & corte & ")")
            .Interior.Color = RGB(192, 0, 0) : .Font.Color = RGB(255, 255, 255) : .Font.Bold = True
            .SetFirstPriority
        End With
        On Error GoTo 0
    End If

    wsDet.Activate
    On Error Resume Next
    ActiveWindow.FreezePanes = False
    wsDet.Cells(DS, 1).Select
    ActiveWindow.FreezePanes = True
    wsDet.Range(wsDet.Cells(HR, 1), wsDet.Cells(HR, nCols)).AutoFilter
    On Error GoTo 0
    wsDet.Cells(HR, 1).Select

End Sub

' CABO ISOLADO tratado como COM (isento) quando MAT LIB SAP < 15 metros
Private Function CaboComoCOM(fam As String, libV As Variant) As Boolean
    If fam = "CABO ISOLADO" And IsNumeric(libV) Then
        Dim mx As Double : mx = mCaboMax : If mx <= 0 Then mx = 15
        CaboComoCOM = (Abs(CDbl(libV)) < mx)
    Else
        CaboComoCOM = False
    End If
End Function

' Materiais TIPO=COM que, se NAO aderentes, reprovam o PEP3NIVEL inteiro.
' Sao apenas: CH FUS, PARA RAIO BT, PARA RAIO MT.
' famNorm ja vem normalizado (UCase, sem espacos/pontos/acentos).
Private Function EhComCritico(famNorm As String) As Boolean
    ' Remove espacos p/ casar "CH FUS", "PARA RAIO MT", "PARA RAIO BT", "PARA-RAIO" etc.
    Dim f As String : f = Replace(famNorm, " ", "")
    ' Lista configuravel (aba CONFIG, chave COM_CRITICO). Sem config -> fallback fixo.
    If mComCrit Is Nothing Then
        EhComCritico = (InStr(f, "CHFUS") > 0 Or InStr(f, "CHAVEFUS") > 0 Or InStr(f, "PARARAIO") > 0)
        Exit Function
    End If
    Dim k As Variant
    For Each k In mComCrit.Keys
        If InStr(f, CStr(k)) > 0 Then EhComCritico = True : Exit Function
    Next k
End Function

' Match por PALAVRA inteira (evita falso positivo de substring:
' "ELO" em PARALELO, "ANEL" em PAINEL etc.). Virgulas/parenteses viram
' separadores. descNorm ja vem de NormStr.
Private Function TemPalavra(descNorm As String, termo As String) As Boolean
    Dim s As String, i As Long, ch As String
    s = " "
    For i = 1 To Len(descNorm)
        ch = Mid$(descNorm, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
            s = s & ch
        Else
            s = s & " "
        End If
    Next i
    TemPalavra = (InStr(s & " ", " " & termo & " ") > 0)
End Function

' Fallback de classificacao por DESCRICAO quando o codigo SAP nao esta no mapa NT.006.
' Baseado na padronizacao de estruturas de distribuicao 15kV (NT.006 / NT.018 Equatorial
' e normas equivalentes ABNT/CELESC/ENEL). Preenche tm ByRef; retorna True se classificou.
' descNorm = descricao ja normalizada (UCase, sem acentos).
Private Function ClassificarDesc(descNorm As String, ByRef tm As tMaterial) As Boolean
    tm.Familia = "" : tm.CodNT006 = "~WEB" : tm.DescrNT006 = "Classif. por descricao (NT.006/018)"
    tm.EhAncora = False : tm.AncoraDep = "CRUZETA"
    tm.RazaoMin = 0 : tm.RazaoMax = 0 : tm.DescrRegra = ""
    ClassificarDesc = True

    ' ---- PINO ISOLADOR: checado ANTES de CRUZETA/CONECTOR pois a descricao do
    ' pino costuma citar a cruzeta ("PINO,ISOL,CRUZETA MAD...") e a palavra
    ' "PINO" tambem aparece em conectores tipo pino de compressao (excluidos aqui). ----
    If TemPalavra(descNorm, "PINO") And InStr(descNorm, "CONECTOR") = 0 And InStr(descNorm, "TERM") = 0 Then
        tm.Familia = "PINO" : tm.AncoraDep = "CRUZETA" : tm.RazaoMin = 1 : tm.RazaoMax = 3.5
        tm.CodNT006 = "F-36" : tm.DescrRegra = "1-3 pinos por cruzeta (1 por isolador pilar)"

    ' ---- PLACA (identificacao/sinalizacao): checado ANTES de POSTE pois a
    ' descricao da placa costuma citar "POSTE" (onde ela e fixada). ----
    ElseIf InStr(descNorm, "PLACA") > 0 Then
        tm.Familia = "PLACA IDENTIFICACAO" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "PL"
        tm.DescrRegra = "Placa de identificacao/sinalizacao (neutro)"

    ' ---- ANCORAS ----
    ElseIf InStr(descNorm, "CRUZETA") > 0 Then
        tm.Familia = "CRUZETA" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "R-02"
    ElseIf InStr(descNorm, "ISOLADOR") > 0 And (InStr(descNorm, "DISCO") > 0 Or InStr(descNorm, "SUSPENS") > 0) Then
        tm.Familia = "ISOL SUSPENSAO" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "I-06"
    ElseIf (InStr(descNorm, "HASTE") > 0 And (InStr(descNorm, "TERRA") > 0 Or InStr(descNorm, "ATERR") > 0 Or InStr(descNorm, "COBRE") > 0)) Then
        tm.Familia = "HASTE TERRA" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "F-17"
    ElseIf InStr(descNorm, "CHAVE") > 0 And InStr(descNorm, "FUS") > 0 Then
        tm.Familia = "CHAVE FUSIVEL" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "E-09"
    ElseIf InStr(descNorm, "CHAVE") > 0 Then
        tm.Familia = "CHAVE FACA/SECC" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "E-10"
        tm.DescrRegra = "Chave faca/seccionadora (ancora)"
    ElseIf InStr(descNorm, "TRAFO") > 0 Or InStr(descNorm, "TRANSFORMADOR") > 0 Then
        tm.Familia = "TRAFO" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "E-45"
    ElseIf InStr(descNorm, "GANCHO") > 0 And InStr(descNorm, "OLHAL") > 0 Then
        tm.Familia = "GANCHO OLHAL" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "F-13"
    ElseIf InStr(descNorm, "SUPORTE") > 0 And (InStr(descNorm, "PARA RAIO") > 0 Or InStr(descNorm, "PARARAIO") > 0) Then
        tm.Familia = "SUP PARA-RAIO" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "F-47"

    ' ---- DEPENDENTES DA CRUZETA (qtd por cruzeta) ----
    ElseIf InStr(descNorm, "ISOLADOR") > 0 And InStr(descNorm, "PILAR") > 0 Then
        tm.Familia = "ISOLADOR PILAR" : tm.AncoraDep = "CRUZETA" : tm.RazaoMin = 2 : tm.RazaoMax = 3.5
        tm.CodNT006 = "I-05" : tm.DescrRegra = "2-3 isolador pilar por cruzeta (1/fase, N1=3)"
    ElseIf InStr(descNorm, "MAO FRANCESA") > 0 Then
        tm.Familia = "MAO FRANCESA" : tm.AncoraDep = "CRUZETA" : tm.RazaoMin = 0.5 : tm.RazaoMax = 2.5
        tm.DescrRegra = "1-2 maos-francesas por cruzeta"
    ElseIf TemPalavra(descNorm, "SELA") Then
        tm.Familia = "SELA CRUZETA" : tm.AncoraDep = "CRUZETA" : tm.RazaoMin = 2 : tm.RazaoMax = 4
        tm.DescrRegra = "2-4 selas por cruzeta"
    ElseIf InStr(descNorm, "ARRUELA") > 0 Then
        tm.Familia = "ARRUELA" : tm.AncoraDep = "CRUZETA" : tm.RazaoMin = 2 : tm.RazaoMax = 9
        tm.CodNT006 = "A-02" : tm.DescrRegra = "2-8 arruelas por cruzeta (N1=4; N2=8)"
    ElseIf InStr(descNorm, "PORCA") > 0 Then
        tm.Familia = "PORCA" : tm.AncoraDep = "CRUZETA" : tm.RazaoMin = 2 : tm.RazaoMax = 6.5
        tm.CodNT006 = "A-21" : tm.DescrRegra = "2-6 porcas por cruzeta"
    ElseIf InStr(descNorm, "PARAFUSO") > 0 And InStr(descNorm, "OLHAL") > 0 Then
        tm.Familia = "PARAFUSO OLHAL" : tm.AncoraDep = "GANCHO OLHAL" : tm.RazaoMin = 0.8 : tm.RazaoMax = 1.2
        tm.CodNT006 = "F-34" : tm.DescrRegra = "1 parafuso olhal por ponto de suspensao"
    ElseIf InStr(descNorm, "PARAFUSO") > 0 Then
        tm.Familia = "PARAFUSO" : tm.AncoraDep = "CRUZETA" : tm.RazaoMin = 1 : tm.RazaoMax = 8
        tm.CodNT006 = "F-30" : tm.DescrRegra = "1-8 parafusos por cruzeta"

    ' ---- DEPENDENTES DE OUTRAS ANCORAS ----
    ElseIf InStr(descNorm, "PARA RAIO") > 0 Or InStr(descNorm, "PARARAIO") > 0 Then
        tm.Familia = "PARA-RAIO" : tm.AncoraDep = "SUP PARA-RAIO" : tm.RazaoMin = 0.8 : tm.RazaoMax = 1.2
        tm.CodNT006 = "E-29" : tm.DescrRegra = "1 para-raios por suporte (1:1)"
    ElseIf InStr(descNorm, "MANILHA") > 0 Then
        tm.Familia = "MANILHA" : tm.AncoraDep = "GANCHO OLHAL" : tm.RazaoMin = 0.8 : tm.RazaoMax = 1.2
        tm.CodNT006 = "F-22" : tm.DescrRegra = "1 manilha por ponto de suspensao"
    ElseIf InStr(descNorm, "OLHAL") > 0 Then
        tm.Familia = "OLHAL" : tm.AncoraDep = "GANCHO OLHAL" : tm.RazaoMin = 0.8 : tm.RazaoMax = 1.2
        tm.DescrRegra = "1 olhal por ponto de suspensao"
    ElseIf InStr(descNorm, "CONECTOR") > 0 And (InStr(descNorm, "HASTE") > 0 Or InStr(descNorm, "ATERR") > 0 Or InStr(descNorm, "CUNHA") > 0) Then
        tm.Familia = "CONEC HASTE" : tm.AncoraDep = "HASTE TERRA" : tm.RazaoMin = 0.8 : tm.RazaoMax = 1.2
        tm.CodNT006 = "M-10" : tm.DescrRegra = "1 conector por haste de aterramento"

    ' ---- ESTAI (cordoalha / esticador / sapata / haste ancora) ----
    ElseIf InStr(descNorm, "CORDOALHA") > 0 Or InStr(descNorm, "ESTICADOR") > 0 Or _
           InStr(descNorm, "SAPATA") > 0 Or (InStr(descNorm, "HASTE") > 0 And InStr(descNorm, "ANCORA") > 0) Then
        tm.Familia = "ESTAI" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "ES"
        tm.DescrRegra = "Conjunto de estai (ancoragem de poste)"

    ' ---- DPS / PROTETOR DE SURTO BT ----
    ElseIf TemPalavra(descNorm, "DPS") Or (InStr(descNorm, "PROTETOR") > 0 And InStr(descNorm, "SURTO") > 0) Then
        tm.Familia = "DPS" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "DPS"
        tm.DescrRegra = "Protetor de surto BT (neutro)"

    ' ---- CONDUTORES / CABOS / FIOS (medidos em metros - referencia por vao;
    ' exclui ALCA, que e acessorio de amarracao de cabo, nao o cabo em si) ----
    ElseIf (InStr(descNorm, "CABO") > 0 Or InStr(descNorm, "CONDUTOR") > 0 Or TemPalavra(descNorm, "FIO") Or TemPalavra(descNorm, "CAA") Or TemPalavra(descNorm, "CAZ")) _
        And InStr(descNorm, "ALCA") = 0 Then
        tm.Familia = "CABO/CONDUTOR" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "COND"
        tm.DescrRegra = "Condutor - medido em metros (referencia por vao)"

    ' ---- POSTE (ancora estrutural) ----
    ElseIf InStr(descNorm, "POSTE") > 0 Then
        tm.Familia = "POSTE" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "POSTE"
        tm.DescrRegra = "Poste - estrutura de suporte"

    ' ---- CAIXA DE MEDICAO: 1 por medidor ----
    ElseIf InStr(descNorm, "CAIXA") > 0 And (InStr(descNorm, "MEDI") > 0 Or _
           InStr(descNorm, "POLICARB") > 0 Or InStr(descNorm, "MONOF") > 0 Or InStr(descNorm, "POLIF") > 0) Then
        tm.Familia = "CAIXA MEDICAO" : tm.AncoraDep = "MEDIDOR" : tm.RazaoMin = 0.8 : tm.RazaoMax = 1.2
        tm.CodNT006 = "CX-M" : tm.DescrRegra = "1 caixa de medicao por medidor"
    ElseIf InStr(descNorm, "CAIXA") > 0 Then
        tm.Familia = "CAIXA" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "CX"
        tm.DescrRegra = "Caixa (neutro)"

    ' ---- LACRE: 2 por medidor instalado (ODM) ----
    ' Antes de MEDIDOR: "LACRE P/ MEDIDOR" e lacre, nao medidor.
    ElseIf InStr(descNorm, "LACRE") > 0 Then
        tm.Familia = "LACRE" : tm.AncoraDep = "MEDIDOR" : tm.RazaoMin = 1.8 : tm.RazaoMax = 2.2
        tm.CodNT006 = "LACRE" : tm.DescrRegra = "2 lacres por medidor instalado"

    ' ---- MEDIDOR (ancora p/ lacre e caixa) ----
    ElseIf InStr(descNorm, "MEDIDOR") > 0 Then
        tm.Familia = "MEDIDOR" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "MED"
        tm.DescrRegra = "Medidor de energia"

    ' ---- ELO / FUSIVEL: depende da chave fusivel (1/fase) ----
    ElseIf TemPalavra(descNorm, "ELO") Or (InStr(descNorm, "FUSIVEL") > 0 And InStr(descNorm, "CHAVE") = 0) Then
        tm.Familia = "ELO FUSIVEL" : tm.AncoraDep = "CHAVE FUSIVEL" : tm.RazaoMin = 0.8 : tm.RazaoMax = 3.5
        tm.CodNT006 = "E-12" : tm.DescrRegra = "1-3 elos por chave fusivel (1/fase)"

    ' ---- CINTA: depende do poste ----
    ElseIf TemPalavra(descNorm, "CINTA") Then
        tm.Familia = "CINTA POSTE" : tm.AncoraDep = "POSTE" : tm.RazaoMin = 1 : tm.RazaoMax = 4
        tm.DescrRegra = "1-4 cintas por poste"

    ' ---- ARMACAO SECUNDARIA (ancora BT) ----
    ElseIf InStr(descNorm, "ARMACAO") > 0 Then
        tm.Familia = "ARMACAO SEC" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "F-01"
        tm.DescrRegra = "Armacao secundaria BT (ancora)"

    ' ---- ROLDANA: depende da armacao ----
    ElseIf InStr(descNorm, "ROLDANA") > 0 Then
        tm.Familia = "ROLDANA" : tm.AncoraDep = "ARMACAO SEC" : tm.RazaoMin = 0.8 : tm.RazaoMax = 1.2
        tm.DescrRegra = "1 roldana por armacao"

    ' ---- GRAMPO (linha morta / ancoragem / suspensao - neutro) ----
    ElseIf InStr(descNorm, "GRAMPO") > 0 Then
        tm.Familia = "GRAMPO" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "F-20"
        tm.DescrRegra = "Grampo de ancoragem/suspensao"

    ' ---- CONECTOR (perfurante / cunha / derivacao - neutro; cobre tambem as
    ' abreviacoes de catalogo "CONEC" e "CON PER" p/ conector perfurante) ----
    ElseIf InStr(descNorm, "CONECTOR") > 0 Or InStr(descNorm, "CONEX") > 0 _
        Or InStr(descNorm, "CONEC") > 0 Or InStr(descNorm, "CON PER") > 0 Then
        tm.Familia = "CONECTOR" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "M-01"
        tm.DescrRegra = "Conector eletrico (neutro)"

    ' ---- LUVA DE EMENDA (neutro) ----
    ElseIf TemPalavra(descNorm, "LUVA") Then
        tm.Familia = "LUVA EMENDA" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "M-05"
        tm.DescrRegra = "Luva de emenda de condutor"

    ' ---- FITA / FECHO / FIVELA (rede compacta - neutro) ----
    ElseIf TemPalavra(descNorm, "FITA") Or InStr(descNorm, "FECHO") > 0 Or InStr(descNorm, "FIVELA") > 0 Then
        tm.Familia = "FITA/FECHO" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "F-50"
        tm.DescrRegra = "Fita/fecho de aco inox (fixacao)"

    ' ---- TERMINAL / CABECOTE / MUFLA (neutro) ----
    ElseIf InStr(descNorm, "TERMINAL") > 0 Or InStr(descNorm, "CABECOTE") > 0 Or InStr(descNorm, "MUFLA") > 0 Then
        tm.Familia = "TERMINAL/MUFLA" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "M-30"
        tm.DescrRegra = "Terminacao de cabo (neutro)"

    ' ---- ESTRIBO (conexao de derivacao - neutro) ----
    ElseIf InStr(descNorm, "ESTRIBO") > 0 Then
        tm.Familia = "ESTRIBO" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "M-40"
        tm.DescrRegra = "Estribo (conexao de derivacao)"

    ' ---- EQUIPAMENTO DE MANOBRA/PROTECAO (religador/seccionador/disjuntor/regulador - ancora) ----
    ElseIf InStr(descNorm, "RELIGADOR") > 0 Or InStr(descNorm, "SECCIONA") > 0 Or InStr(descNorm, "DISJUNTOR") > 0 Or InStr(descNorm, "REGULADOR") > 0 Then
        tm.Familia = "EQUIP MANOBRA" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "EQ"
        tm.DescrRegra = "Equipamento de manobra/protecao (ancora)"

    ' ---- BUCHA / fixacao generica por cruzeta ----
    ElseIf TemPalavra(descNorm, "BUCHA") Then
        tm.Familia = "BUCHA" : tm.AncoraDep = "CRUZETA" : tm.RazaoMin = 1 : tm.RazaoMax = 8
        tm.DescrRegra = "Bucha/fixacao por cruzeta"

    ' ---- REDE COMPACTA (NT.018) - referencia por vao; tratado como neutro ----
    ElseIf InStr(descNorm, "ESPACADOR") > 0 Or InStr(descNorm, "BRACO") > 0 Or InStr(descNorm, "LOSANGULAR") > 0 Then
        tm.Familia = "REDE COMPACTA" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "NT018"
        tm.DescrRegra = "Rede compacta - referencia por vao (3-5 espacadores/vao)"
    ElseIf InStr(descNorm, "ALCA") > 0 And (InStr(descNorm, "PREFORM") > 0 Or InStr(descNorm, "PREF") > 0) Then
        tm.Familia = "ALCA PREFORMADA" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "F-04"
        tm.DescrRegra = "Alca preformada (amarracao de condutor; PREF = abreviacao de preformada)"
    ElseIf TemPalavra(descNorm, "ANEL") Or TemPalavra(descNorm, "LACO") Or TemPalavra(descNorm, "ALCA") Then
        tm.Familia = "REDE COMPACTA" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "NT018"
        tm.DescrRegra = "Acessorio rede compacta (anel/laco/alca)"

    ' ---- ELETRODUTO / FERRAGENS / ARAME / SUPORTE (identificacao, neutros) ----
    ElseIf InStr(descNorm, "ELETRODUTO") > 0 Then
        tm.Familia = "ELETRODUTO" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "ED"
        tm.DescrRegra = "Eletroduto/duto de descida"
    ElseIf InStr(descNorm, "ABRACADEIRA") > 0 Or InStr(descNorm, "CANTONEIRA") > 0 Or _
           InStr(descNorm, "PERFIL") > 0 Or InStr(descNorm, "CHAPA") > 0 Then
        tm.Familia = "FERRAGEM" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "FE"
        tm.DescrRegra = "Ferragem de fixacao (neutro)"
    ElseIf TemPalavra(descNorm, "ARAME") Then
        tm.Familia = "ARAME" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "AR"
        tm.DescrRegra = "Arame de amarracao"
    ElseIf InStr(descNorm, "SUPORTE") > 0 Then
        tm.Familia = "SUPORTE" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "SUP"
        tm.DescrRegra = "Suporte generico (neutro)"
    ElseIf TemPalavra(descNorm, "SUP") And (InStr(descNorm, "CH FACA") > 0 Or InStr(descNorm, "CHAVE FACA") > 0 _
        Or InStr(descNorm, "BYPASS") > 0 Or InStr(descNorm, "BY PASS") > 0) Then
        tm.Familia = "SUPORTE" : tm.EhAncora = True : tm.AncoraDep = "" : tm.CodNT006 = "SUP"
        tm.DescrRegra = "Suporte p/ equipamento (chave faca/by-pass) - abreviacao SUP"

    Else
        ClassificarDesc = False
    End If
End Function

' Formata a diferenca SAP-PRJ com sinal e percentual (p/ o MOTIVO de reprovados)
' Ex: " | Dif=+3 (42,9% acima)"  /  " | Dif=-2 (-25,0% abaixo)"
Private Function FmtDif(libV As Variant, prjV As Variant) As String
    If IsNumeric(libV) And IsNumeric(prjV) Then
        Dim l As Double, p As Double, d As Double
        l = CDbl(libV) : p = CDbl(prjV) : d = l - p
        Dim s As String
        s = " | Dif=" & IIf(d >= 0, "+", "") & Format(d, "#,##0.##")
        If p <> 0 Then
            s = s & " (" & Format(Abs(d / p), "0.0%") & IIf(d >= 0, " acima", " abaixo") & ")"
        End If
        FmtDif = s
    Else
        FmtDif = ""
    End If
End Function

' Avalia uma linha SAP x PRJ (fonte UNICA de regra, usada pelos 2 loops de
' ProcessarSAPxPRJ). Recebe tipo/familia ja normalizados via NormStr.
' ucZeradoReprova: a regra do UC zerado so devolve a obra quando TODOS os
' materiais UC do PEP3 estao zerados - o chamador informa via este flag
' (ver TodosUCZerados). Itens zerados isolados sao ignorados e o PEP3
' fica APROVADO quando as demais regras passam.
' Tipo de OD do PEP: sufixo do PEP4 (.I/.M/.S/.D) ou, sem sufixo, a coluna
' TIPO PEP (I/M/S/D ou ODI/ODM/ODS/ODD). Recebe o PEP4 CRU: NormStr remove o
' ponto e apagaria o sufixo. Retorna "ODI", "ODM", "ODS", "ODD" ou "".
Private Function ClassificarOD(ByVal pep4Raw As String, ByVal tipoPepRaw As String) As String
    Select Case UCase(Right(Trim(pep4Raw), 2))
        Case ".I": ClassificarOD = "ODI"
        Case ".M": ClassificarOD = "ODM"
        Case ".S": ClassificarOD = "ODS"
        Case ".D": ClassificarOD = "ODD"
        Case Else
            Select Case UCase(Trim(tipoPepRaw))
                Case "I", "ODI": ClassificarOD = "ODI"
                Case "M", "ODM": ClassificarOD = "ODM"
                Case "S", "ODS": ClassificarOD = "ODS"
                Case "D", "ODD": ClassificarOD = "ODD"
            End Select
    End Select
End Function

' Premissa de sinal por tipo de OD, sobre VALOR e QUANTIDADE (MAT LIB SAP):
'   ODD (desmonte)         -> ambos NEGATIVOS; positivo = NAO ADERENTE
'   ODI / ODS (instalacao) -> ambos POSITIVOS; negativo = NAO ADERENTE
' Zero nao viola (nao e positivo nem negativo). ODM fica fora da premissa.
' Devolve a descricao da violacao, ou "" quando o sinal esta correto.
Private Function SinalInvalidoOD(ByVal tipoOD As String, ByVal qtdV As Variant, _
                                 ByVal valV As Variant) As String
    Dim q As Double : q = Val0(qtdV)
    Dim v As Double : v = Val0(valV)
    Dim errado As String, esperado As String
    Select Case tipoOD
        Case "ODD"
            esperado = "negativo"
            If q > 0 And v > 0 Then
                errado = "valor e quantidade POSITIVOS"
            ElseIf q > 0 Then
                errado = "quantidade POSITIVA"
            ElseIf v > 0 Then
                errado = "valor POSITIVO"
            End If
        Case "ODI", "ODS"
            esperado = "positivo"
            If q < 0 And v < 0 Then
                errado = "valor e quantidade NEGATIVOS"
            ElseIf q < 0 Then
                errado = "quantidade NEGATIVA"
            ElseIf v < 0 Then
                errado = "valor NEGATIVO"
            End If
    End Select
    If errado <> "" Then
        SinalInvalidoOD = "PEP " & tipoOD & " com " & errado & " (esperado " & esperado & ")"
    End If
End Function

' Alerta PU BAIXO EM ODI: material em PEP ODI com preco unitario (VALOR / QTD)
' dentro de [puMin, puMax], limites inclusive. Condutores e cabos (familia
' COND* / CABO*) ficam de fora: o PU deles e por metro e naturalmente baixo.
' QTD zero nao tem preco unitario e nao alerta.
Private Function PuBaixoODI(ByVal odTp As String, ByVal famN As String, _
        ByVal valor As Double, ByVal qtd As Double, _
        ByVal puMin As Double, ByVal puMax As Double) As Boolean
    If odTp <> "ODI" Or qtd = 0 Then Exit Function
    If Left(famN, 4) = "COND" Or Left(famN, 4) = "CABO" Then Exit Function
    Dim pu As Double : pu = valor / qtd
    ' folga de 1e-6: 0,01 calculado em ponto flutuante pode sair 0,00999999...
    PuBaixoODI = (pu >= puMin - 0.000001 And pu <= puMax + 0.000001)
End Function

' Alerta PU 0,01 EM CABO/COND: cabo ou condutor (familia CABO* / COND*) com
' preco unitario (VALOR / QTD) igual ao valor simbolico puAlvo (R$ 0,01) -
' vale "o que aparece como 0,01": a menos de meio centavo dele (0,0051 a 0,0149).
' Qualquer tipo de PEP. QTD zero nao tem preco unitario e nao alerta.
Private Function PuSimbolicoCabo(ByVal famN As String, ByVal valor As Double, _
        ByVal qtd As Double, ByVal puAlvo As Double) As Boolean
    If qtd = 0 Then Exit Function
    If Left(famN, 4) <> "COND" And Left(famN, 4) <> "CABO" Then Exit Function
    PuSimbolicoCabo = (Abs(valor / qtd - puAlvo) < 0.005)
End Function

' Premissa de valor em ODD: material com VALOR abaixo do limite (padrao
' R$ -100) e NAO ADERENTE. "Abaixo" = mais negativo: -150 viola, -100 nao.
' Devolve a descricao da violacao, ou "" quando ok.
Private Function ValorBaixoODD(ByVal tipoOD As String, ByVal valV As Variant, _
                               ByVal limite As Double) As String
    If tipoOD <> "ODD" Then Exit Function
    Dim v As Double : v = Val0(valV)
    If v < limite Then
        ValorBaixoODD = "PEP ODD com valor R$ " & Format(v, "#,##0.00") & _
                        " abaixo do limite de R$ " & Format(limite, "#,##0.00")
    End If
End Function

' Limite de ODD_VALOR_MIN (CONFIG); -100 quando a CONFIG nao foi carregada.
Private Function LimiteValorODD() As Double
    If IsEmpty(mOddValMin) Then LimiteValorODD = -100 Else LimiteValorODD = mOddValMin
End Function

' Alerta POSTE ODI PU BAIXO: poste (familia contendo POSTE) em PEP ODI com
' preco unitario (VALOR / QTD) de 0 ate abaixo do limite (padrao R$ 500).
' PU negativo fica de fora (ja e alerta de sinal); QTD zero nao alerta.
Private Function PosteBaratoODI(ByVal odTp As String, ByVal famN As String, _
        ByVal valor As Double, ByVal qtd As Double, ByVal limite As Double) As Boolean
    If odTp <> "ODI" Or qtd = 0 Then Exit Function
    If InStr(famN, "POSTE") = 0 Then Exit Function
    Dim pu As Double : pu = valor / qtd
    PosteBaratoODI = (pu >= 0 And pu < limite)
End Function

' Alerta MEDIDOR EM PEP ODI: medidor (descricao ou familia com MEDIDOR) em PEP
' ODI. Mesmo criterio de medidor da regra de ODM (sem CAIXA, para nao pegar
' "CAIXA PARA MEDIDOR"), e tambem sem LACRE ("LACRE P/ MEDIDOR" nao e medidor).
Private Function MedidorEmODI(ByVal odTp As String, ByVal descNorm As String, _
                              ByVal famN As String) As Boolean
    If odTp <> "ODI" Then Exit Function
    Dim dN As String : dN = descNorm & " " & famN
    MedidorEmODI = EhMedidorODM(dN)
End Function

' Medidor na regra LACRE x MEDIDOR (ODM) e no alerta MEDIDOR EM PEP ODI:
' descricao/familia com MEDIDOR, sem CAIXA ("CAIXA PARA MEDIDOR") e sem LACRE
' ("LACRE P/ MEDIDOR"). dN = descricao normalizada & " " & familia.
Private Function EhMedidorODM(ByVal dN As String) As Boolean
    EhMedidorODM = (InStr(dN, "MEDIDOR") > 0 And InStr(dN, "CAIXA") = 0 _
                    And InStr(dN, "LACRE") = 0)
End Function

' PEP3 NAO ADERENTE por valor negativo em ODI: linha de PEP ODI com VALOR
' negativo (qualquer tipo de item). Zero nao conta.
Private Function NegativoEmODI(ByVal odTp As String, ByVal valV As Variant) As Boolean
    NegativoEmODI = (odTp = "ODI" And Val0(valV) < 0)
End Function

' SAP e PRJ com sinais opostos na mesma linha (SAP +3 / PRJ -3, ou SAP -3 /
' PRJ +3): incoerente em qualquer tipo de OD. Zero nao e positivo nem
' negativo. Devolve a descricao, ou "" quando ok.
Private Function SinaisOpostos(ByVal libV As Variant, ByVal prjV As Variant) As String
    Dim l As Double : l = Val0(libV)
    Dim p As Double : p = Val0(prjV)
    If (l > 0 And p < 0) Or (l < 0 And p > 0) Then
        SinaisOpostos = "MAT LIB SAP (" & CStr(l) & ") e MAT PRJ CAD (" & CStr(p) & _
                        ") com sinais opostos"
    End If
End Function

' Acrescenta a familia culpada a lista de devolucao do PEP3 (sem repetir).
Private Sub RegistrarFamRep(d As Object, ByVal p3 As String, ByVal lbl As String)
    If Not d.Exists(p3) Then
        d.Add p3, lbl
    ElseIf InStr(d(p3), lbl) = 0 Then
        d(p3) = d(p3) & ", " & lbl
    End If
End Sub

Private Function AvaliarLinha(tipoR As String, famR As String, _
                              libV As Variant, prjV As Variant, _
                              sitRNorm As String, _
                              Optional ucZeradoReprova As Boolean = True, _
                              Optional ByVal tipoOD As String = "", _
                              Optional ByVal valV As Variant = 0, _
                              Optional ByVal qtdSinal As Variant, _
                              Optional ByVal prjSinal As Variant) As tAvalLinha
    Dim av As tAvalLinha
    av.UcZerado = MaterialUCZerado(tipoR, libV, prjV) And ucZeradoReprova
    av.EhAder = EhAderente(famR, libV, prjV, sitRNorm)
    If av.UcZerado Then av.EhAder = False   ' UC zerado nunca e aderente
    ' Premissa de sinal por tipo de OD: valor/quantidade com sinal oposto ao
    ' esperado torna o item NAO ADERENTE, independente da comparacao SAP x PRJ
    ' qtdSinal: quantidade da propria linha quando libV e o TOTAL de um grupo
    ' de equivalentes - o sinal vale por item, nao pelo total do grupo
    Dim qS As Variant
    If IsMissing(qtdSinal) Then qS = libV Else qS = qtdSinal
    av.MotivoOD = SinalInvalidoOD(tipoOD, qS, valV)
    ' ODD: valor abaixo do limite tambem torna o item NAO ADERENTE
    Dim mVal As String : mVal = ValorBaixoODD(tipoOD, valV, LimiteValorODD())
    If mVal <> "" Then av.MotivoOD = av.MotivoOD & IIf(av.MotivoOD <> "", "; ", "") & mVal
    ' SAP e PRJ com sinais opostos: confere as quantidades da PROPRIA linha
    ' (prjSinal/qtdSinal), nao o total de um grupo de equivalentes
    Dim pS As Variant
    If IsMissing(prjSinal) Then pS = prjV Else pS = prjSinal
    Dim mOp As String : mOp = SinaisOpostos(qS, pS)
    If mOp <> "" Then av.MotivoOD = av.MotivoOD & IIf(av.MotivoOD <> "", "; ", "") & mOp
    av.ViolaOD = (av.MotivoOD <> "")
    If av.ViolaOD Then av.EhAder = False
    av.CaboIsento = CaboComoCOM(famR, libV)
    ' UC sempre avalia; COM avalia apenas se for familia critica (COM_CRITICO)
    av.Avaliavel = (tipoR = "UC") Or (tipoR = "COM" And EhComCritico(famR))
    ' reprova: nao aderente com situacao valida, OU material UC zerado, OU sinal
    ' invertido (fato do dado: nao depende do texto da SIT MAT, nem de "NULO")
    av.ReprovaLinha = (Not av.EhAder And sitRNorm <> "NULO") Or av.UcZerado _
                      Or av.ViolaOD
    ' SIT MAT recalculada p/ cabo/condutor/ramal com valores numericos
    av.RecalculaSit = (Left(famR, 4) = "COND" Or Left(famR, 4) = "CABO" Or famR = "RAMAL") _
                      And IsNumeric(libV) And IsNumeric(prjV)
    AvaliarLinha = av
End Function

' Aderencia com margem percentual p/ condutores, cabos e ramal (kg ou m).
' Limiares configuraveis na aba CONFIG: TOL_ADERENCIA_PCT (margem percentual),
' DIF_ABS_COND_NU (dif absoluta COND NU) e DIF_ABS_CABO (dif absoluta das
' familias de DIF_ABS_FAMILIAS). Fallbacks fixos valem quando a CONFIG ainda
' nao foi carregada (ex.: TestarLogicaInventario standalone).
Private Function EhAderente(fam As String, libV As Variant, prjV As Variant, rawSitNorm As String) As Boolean
    Dim isMarg As Boolean
    isMarg = (Left(fam, 4) = "COND" Or Left(fam, 4) = "CABO" Or fam = "RAMAL")
    If isMarg And IsNumeric(libV) And IsNumeric(prjV) Then
        Dim l As Double, pp As Double
        l = CDbl(libV) : pp = CDbl(prjV)
        Dim dNu As Double : dNu = mDifCondNu : If dNu <= 0 Then dNu = 5
        Dim dAbs As Double : dAbs = mDifAbsCabo : If dAbs <= 0 Then dAbs = 15
        ' COND NU: diferenca ABSOLUTA ate DIF_ABS_COND_NU e ADERENTE, independente
        ' da margem percentual. Match tolerante a espaco/pontuacao ("COND NU",
        ' "COND-NU", "COND. NU" -> "CONDNU").
        If EhCondNu(fam) And Abs(l - pp) <= dNu Then
            EhAderente = True
        ' Familias de DIF_ABS_FAMILIAS (CABO ISOLADO/CABO COBRE/COND COBRE): dif
        ' ABSOLUTA ate DIF_ABS_CABO (nao e %) e ADERENTE, mesmo com projetado 0.
        ElseIf FamiliaDifAbs(fam) And Abs(l - pp) <= dAbs Then
            EhAderente = True
        ElseIf pp = 0 Then
            EhAderente = (l = 0)
        Else
            Dim tol As Double : tol = mTolAder : If tol <= 0 Then tol = 0.1
            EhAderente = (Abs(l - pp) <= tol*Abs(pp))
        End If
    Else
        EhAderente = (rawSitNorm = "ADERENTE")
    End If
End Function

' True quando a familia (ja NormStr) e COND NU, tolerante a espaco/pontuacao.
Private Function EhCondNu(fam As String) As Boolean
    EhCondNu = (Replace(fam, " ", "") = "CONDNU")
End Function

' Familias em que a divergencia ABSOLUTA (nao percentual) ate DIF_ABS_CABO e
' considerada ADERENTE. Lista configuravel na CONFIG (chave DIF_ABS_FAMILIAS,
' separada por ;). Fallback fixo quando a CONFIG nao foi carregada.
Private Function FamiliaDifAbs(fam As String) As Boolean
    Dim f As String : f = Replace(fam, " ", "")
    If Not mFamDifAbs Is Nothing Then
        If mFamDifAbs.Count > 0 Then
            FamiliaDifAbs = mFamDifAbs.Exists(f)
            Exit Function
        End If
    End If
    FamiliaDifAbs = (f = "CABOISOLADO" Or f = "CABOCOBRE" Or f = "CONDCOBRE")
End Function

' True quando um material UC vem com MAT LIB SAP e MAT PRJ CAD zerados/vazios.
' A devolucao da obra por esta regra so ocorre quando TODOS os UC do PEP3
' estao zerados (ver TodosUCZerados); itens zerados isolados sao ignorados.
Private Function MaterialUCZerado(tipoR As String, libV As Variant, prjV As Variant) As Boolean
    MaterialUCZerado = (tipoR = "UC" And Val0(libV) = 0 And Val0(prjV) = 0)
End Function

' True quando TODOS os materiais UC do PEP3 estao com valores zerados (ai sim
' a obra e devolvida). PEP3 sem nenhum UC, ou com qualquer UC valorado -> False.
' p3nUC: PEP3 -> qtd de linhas UC | p3nUC0: PEP3 -> qtd de linhas UC zeradas.
Private Function TodosUCZerados(p3nUC As Object, p3nUC0 As Object, pep3 As String) As Boolean
    Dim t As Double, z As Double
    t = 0 : If p3nUC.Exists(pep3) Then t = CDbl(p3nUC(pep3))
    z = 0 : If p3nUC0.Exists(pep3) Then z = CDbl(p3nUC0(pep3))
    TodosUCZerados = (t > 0 And z = t)
End Function

Private Function FmtKPI(v As Double) As String
    Dim a As Double : a = Abs(v)
    If a >= 1000000 Then
        FmtKPI = "R$ " & Format(v*0.000001, "0.0") & " MM"
    ElseIf a >= 1000 Then
        FmtKPI = "R$ " & Format(v*0.001, "#,##0") & " mil"
    Else
        FmtKPI = "R$ " & Format(v, "#,##0")
    End If
End Function

' True quando a unidade do material e CONTAVEL (UN, PC, CJ...) e a faixa
' prevista deve ser exibida/avaliada em numeros INTEIROS. Unidades continuas
' (metro, quilo, litro...) mantem decimais.
Private Function EhUnidadeInteira(uml As String) As Boolean
    Dim u As String : u = NormStr(uml)
    Select Case u
        Case "M", "MT", "MTS", "KM", "KG", "G", "GR", "T", "TON", "L", "LT", "M2", "M3"
            EhUnidadeInteira = False
        Case Else
            EhUnidadeInteira = True   ' UN, PC, CJ, JG, vazio etc.
    End Select
End Function

Private Sub ProcessarCOMInventario(wsBase As Worksheet, wsCom As Worksheet)

    Dim lastC As Long
    lastC = wsBase.Cells(1, wsBase.Columns.Count).End(xlToLeft).Column

    Dim colPEP3 As Long, colPEP4 As Long, colMat As Long
    Dim colDescM As Long, colQtd As Long, colUML As Long
    Dim colTipo As Long, colPEP3n As Long
    Dim c As Long
    Dim mapa As Object : Set mapa = MapaColunas(wsBase)
    colPEP3  = ColIdx(mapa, "PEP3NIVEL")
    colPEP4  = ColIdx(mapa, "PEP4NIVEL")
    colMat   = ColIdx(mapa, "COD MAT")
    colDescM = ColIdx(mapa, "DESC MAT")
    colQtd   = ColIdx(mapa, "MAT LIB SAP")
    colUML   = ColIdx(mapa, "UND")
    colTipo  = ColIdx(mapa, "TIPO")
    If colPEP4 = 0 Then colPEP4 = 2
    If colMat  = 0 Then colMat  = 5
    If colDescM= 0 Then colDescM= 7
    If colQtd  = 0 Then colQtd  = 9
    If colUML  = 0 Then colUML  = 8
    If colTipo = 0 Then colTipo = 11
    If colPEP3 = 0 Then colPEP3 = 1

    Dim nt006 As Object
    Set nt006 = CriarMapaNT006()

    Dim liquidQ    As Object : Set liquidQ    = CreateObject("Scripting.Dictionary")
    Dim liquidInfo As Object : Set liquidInfo = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    lastRow = wsBase.Cells(wsBase.Rows.Count, colPEP4).End(xlUp).Row

    Dim d As Variant
    d = wsBase.Range(wsBase.Cells(2, 1), wsBase.Cells(lastRow, lastC)).Value
    Dim nRows As Long : nRows = lastRow - 1

    Dim r As Long
    For r = 1 To nRows
        Dim tipoR As String : tipoR = NormStr(CStr(d(r, colTipo)))
        If tipoR = "COM" Then
            Dim pep4   As String : pep4   = Trim(CStr(d(r, colPEP4)))
            Dim pep3   As String : pep3   = Trim(CStr(d(r, colPEP3)))
            Dim matCod As String : matCod = NormCod(d(r, colMat))
            Dim qtd    As Double : qtd    = ToNum(d(r, colQtd))
            Dim key    As String : key    = pep4 & "|" & matCod
            If liquidQ.Exists(key) Then
                liquidQ(key) = liquidQ(key) + qtd
            Else
                liquidQ.Add key, qtd
                Dim info(5) As String
                info(0) = pep4
                info(1) = pep3
                info(2) = matCod
                info(3) = Trim(CStr(d(r, colDescM)))
                info(4) = Trim(CStr(d(r, colUML)))
                info(5) = pep3
                liquidInfo.Add key, info
            End If
        End If
    Next r

    Dim ancoras As Object : Set ancoras = CreateObject("Scripting.Dictionary")
    Dim ancorasP3 As Object : Set ancorasP3 = CreateObject("Scripting.Dictionary")  ' pep3|familia (fallback)
    Dim ancoraInfo As Object : Set ancoraInfo = CreateObject("Scripting.Dictionary")  ' pep|familia -> "cod - desc"
    Dim k As Variant
    For Each k In liquidQ.Keys
        Dim matCodK As String : matCodK = Split(CStr(k), "|")(1)
        Dim pepK    As String : pepK    = Split(CStr(k), "|")(0)
        Dim tm As tMaterial
        Dim achouTm As Boolean : achouTm = False
        If nt006.Exists(matCodK) Then
            tm = GetMat(nt006, matCodK) : achouTm = True
        Else
            ' Fallback por descricao (web NT.006/018)
            Dim descK As String : descK = NormStr(CStr(liquidInfo(k)(3)))
            achouTm = ClassificarDesc(descK, tm)
        End If
        If achouTm Then
            If tm.EhAncora Then
                Dim aKey As String : aKey = pepK & "|" & tm.Familia
                If ancoras.Exists(aKey) Then
                    ancoras(aKey) = ancoras(aKey) + liquidQ(k)
                Else
                    ancoras.Add aKey, liquidQ(k)
                    ' guarda o material ancora representativo (cod - desc)
                    ancoraInfo.Add aKey, liquidInfo(k)(2) & " - " & liquidInfo(k)(3)
                End If
                ' agregado no PEP3 (fallback p/ PEP4 irmaos que dividem materiais)
                Dim aKey3 As String : aKey3 = liquidInfo(k)(1) & "|" & tm.Familia
                If ancorasP3.Exists(aKey3) Then
                    ancorasP3(aKey3) = ancorasP3(aKey3) + liquidQ(k)
                Else
                    ancorasP3.Add aKey3, liquidQ(k)
                End If
            End If
        End If
    Next k

    CabecalhoRacioNT wsCom

    Dim keys() As String
    Dim ki As Long : ki = 0
    If liquidQ.Count > 0 Then
    ReDim keys(0 To liquidQ.Count - 1)
    For Each k In liquidQ.Keys
        keys(ki) = CStr(k) : ki = ki + 1
    Next k
    QuickSortStr keys, 0, UBound(keys)

    ' Saida bufferizada: monta as linhas em array e escreve na planilha UMA vez
    Dim outA() As Variant : ReDim outA(1 To liquidQ.Count, 1 To 10)
    Dim stArr() As String : ReDim stArr(1 To liquidQ.Count)
    Dim i As Long

    For i = 0 To UBound(keys)
        k = keys(i)
        Dim pepV  As String : pepV = Split(CStr(k), "|")(0)
        Dim matV  As String : matV = Split(CStr(k), "|")(1)
        Dim qtdL  As Double : qtdL = liquidQ(k)
        Dim inf() As String : inf  = liquidInfo(k)

        Dim familia   As String  : familia   = "SEM REFERENCIA"
        Dim codNT     As String  : codNT     = "-"
        Dim descrNT   As String  : descrNT   = ""
        Dim ancoraFam As String  : ancoraFam = ""
        Dim qtdAnc    As Double  : qtdAnc    = 0
        Dim faixaMin  As Double  : faixaMin  = 0
        Dim faixaMax  As Double  : faixaMax  = 0
        Dim regra     As String  : regra     = ""
        Dim status    As String  : status    = "SEM REFERENCIA"
        Dim ehAnc     As Boolean : ehAnc     = False
        Dim ancNivel3 As Boolean : ancNivel3 = False

        ' Tenta mapa por codigo SAP; se nao houver, fallback por descricao (web NT.006/018)
        Dim tmV As tMaterial
        Dim temRef As Boolean : temRef = False
        If nt006.Exists(matV) Then
            tmV = GetMat(nt006, matV) : temRef = True
        Else
            Dim descV As String : descV = NormStr(CStr(inf(3)))
            temRef = ClassificarDesc(descV, tmV)
        End If

        If temRef Then
            familia   = tmV.Familia
            codNT     = tmV.CodNT006
            descrNT   = tmV.DescrNT006
            ehAnc     = tmV.EhAncora
            regra     = tmV.DescrRegra
            ancoraFam = tmV.AncoraDep
            If ehAnc Then
                If qtdL < 0 Then
                    status = "ESTORNO SEM ENTRADA"
                ElseIf qtdL = 0 Then
                    status = "QTD ZERO"
                Else
                    status = "ANCORA"
                End If
            Else
                Dim ancKey As String : ancKey = pepV & "|" & ancoraFam
                If ancoras.Exists(ancKey) Then qtdAnc = ancoras(ancKey)
                ' FALLBACK: ancora ausente/zerada no PEP4 -> usa o agregado do
                ' PEP3 (estruturas costumam dividir materiais entre PEP4 irmaos)
                If qtdAnc <= 0 Then
                    Dim ancKey3 As String : ancKey3 = inf(1) & "|" & ancoraFam
                    If ancorasP3.Exists(ancKey3) Then
                        If ancorasP3(ancKey3) > 0 Then
                            qtdAnc = ancorasP3(ancKey3) : ancNivel3 = True
                        End If
                    End If
                End If
                If qtdAnc > 0 Then
                    ' FOLGA UNICA: ja embutida nas razoes min/max do mapa NT.006.
                    ' (removido o multiplicador extra 0.9/1.1, que somava
                    '  tolerancia sobre tolerancia e afrouxava a regra)
                    faixaMin = qtdAnc * tmV.RazaoMin
                    faixaMax = qtdAnc * tmV.RazaoMax
                    ' Material contavel (UN/PC/CJ...): faixa em INTEIROS
                    ' (min arredonda p/ baixo, max p/ cima - nao gera falso alerta)
                    If EhUnidadeInteira(inf(4)) Then
                        faixaMin = Int(faixaMin)
                        faixaMax = -Int(-faixaMax)
                    End If
                    If qtdL < 0 Then
                        status = "ESTORNO SEM ENTRADA"
                    ElseIf qtdL < faixaMin Then
                        status = "INSUFICIENTE"
                    ElseIf faixaMax > 0 And qtdL >= faixaMax * 2# Then
                        status = "EXCESSO EXAGERADO"
                    ElseIf qtdL > faixaMax Then
                        status = "EXCESSO"
                    Else
                        status = "OK"
                    End If
                Else
                    status = "SEM ANCORA"
                End If
            End If
        End If

        ' Texto do previsto (faixa) p/ leitura humana
        Dim previstoTxt As String
        If faixaMax > 0 Then
            previstoTxt = Format(faixaMin, "#,##0.##") & " a " & Format(faixaMax, "#,##0.##")
        Else
            previstoTxt = "-"
        End If

        ' Observacao em linguagem clara: quanto veio x o previsto
        Dim obsTxt As String
        Dim qVeio As String : qVeio = Format(qtdL, "#,##0.##")
        Select Case status
            Case "ANCORA"
                obsTxt = "Material de referencia (ancora)"
            Case "OK"
                obsTxt = "Dentro do previsto"
            Case "INSUFICIENTE"
                obsTxt = "Veio MENOS que o previsto (faltaram " & Format(faixaMin - qtdL, "#,##0.##") & ")"
            Case "EXCESSO"
                obsTxt = "Veio MAIS que o previsto (excedeu " & Format(qtdL - faixaMax, "#,##0.##") & ")"
            Case "EXCESSO EXAGERADO"
                obsTxt = "Veio MUITO acima do previsto (" & Format(qtdL / faixaMax, "0.0") & "x o maximo)"
            Case "SEM ANCORA"
                obsTxt = "Sem " & ancoraFam & " no PEP para comparar"
            Case "QTD ZERO"
                obsTxt = "Quantidade zero na obra"
            Case "ESTORNO SEM ENTRADA"
                obsTxt = "Estorno sem entrada correspondente"
            Case "SEM REFERENCIA"
                obsTxt = "Material fora da NT.006"
            Case Else
                obsTxt = ""
        End Select
        If ancNivel3 Then obsTxt = obsTxt & " [ancora agregada no PEP3]"
        If temRef Then
            If tmV.CodNT006 = "~WEB" Then obsTxt = obsTxt & " [classif. por descricao]"
        End If

        ' LIGADO A: a qual material/ancora este COM deveria corresponder
        Dim ligadoTxt As String : ligadoTxt = "-"
        If temRef Then
            If ehAnc Then
                ligadoTxt = "(e a propria referencia)"
            ElseIf ancoraFam <> "" Then
                Dim aK As String : aK = pepV & "|" & ancoraFam
                If ancoras.Exists(aK) Then
                    Dim aInfo As String : aInfo = ancoraFam
                    If ancoraInfo.Exists(aK) Then aInfo = ancoraInfo(aK)
                    ligadoTxt = aInfo & " (" & ancoraFam & ", qtd " & Format(ancoras(aK), "#,##0.##") & ")"
                ElseIf ancNivel3 Then
                    ligadoTxt = ancoraFam & " (agregada no PEP3, qtd " & Format(qtdAnc, "#,##0.##") & ")"
                Else
                    ligadoTxt = ancoraFam & " (AUSENTE no PEP)"
                End If
            End If
        End If

        Dim bgS As Long, fgS As Long, icoS As String
        CoresStatusRacio status, bgS, fgS, icoS
        outA(i + 1, 1) = pepV : outA(i + 1, 2) = inf(1) : outA(i + 1, 3) = matV
        outA(i + 1, 4) = inf(3) : outA(i + 1, 5) = familia : outA(i + 1, 6) = ligadoTxt
        outA(i + 1, 7) = qtdL : outA(i + 1, 8) = previstoTxt : outA(i + 1, 9) = icoS
        outA(i + 1, 10) = obsTxt
        stArr(i + 1) = status
    Next i

    wsCom.Range(wsCom.Cells(3, 1), wsCom.Cells(2 + liquidQ.Count, 10)).Value = outA
    FormatarLinhasRacioNT wsCom, stArr, liquidQ.Count
    End If   ' liquidQ.Count > 0

    FormatarColunasRacioNT wsCom

    On Error Resume Next
    wsCom.Activate
    wsCom.Range("A3").Select
    ActiveWindow.FreezePanes = True
    wsCom.Range("A2:J2").AutoFilter
    wsCom.Range("A1").Select
    On Error GoTo 0

End Sub

Private Sub CabecalhoRacioNT(ws As Worksheet)
    With ws.Range("A1:J1")
        .Merge
        .Value = "RACIONALIZACAO COM  -  o que veio na obra x o previsto (NT.006)"
        .Font.Name = "Segoe UI Semibold" : .Font.Bold = True : .Font.Size = 13
        .Font.Color = RGB(235, 240, 248)
        .Interior.Color = RGB(17, 24, 39)
        .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 30
    Dim hdrs As Variant
    hdrs = Array("PEP4NIVEL", "PEP3NIVEL", "COD MATERIAL", "MATERIAL", "FAMILIA", _
                 "LIGADO A (referencia)", "VEIO", "PREVISTO", "STATUS", "OBSERVACAO")
    Dim c As Long
    For c = 1 To 10
        With ws.Cells(2, c)
            .Value = hdrs(c - 1)
            .Font.Name = "Segoe UI Semibold" : .Font.Bold = True : .Font.Size = 10
            .Font.Color = RGB(235, 240, 248)
            .Interior.Color = RGB(31, 41, 59)
            .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous : .Borders.Color = RGB(60, 70, 90)
        End With
    Next c
    ws.Rows(2).RowHeight = 22
End Sub

' Tabela de cores semanticas por STATUS da racionalizacao (fundo, texto e rotulo)
Private Sub CoresStatusRacio(st As String, ByRef bgS As Long, ByRef fgS As Long, ByRef ico As String)
    Select Case st
        Case "OK":                bgS = RGB(198, 239, 206) : fgS = RGB(0, 97, 0)   : ico = "OK"
        Case "ANCORA":            bgS = RGB(221, 235, 247) : fgS = RGB(31, 78, 121): ico = "REFERENCIA"
        Case "EXCESSO":           bgS = RGB(255, 235, 156) : fgS = RGB(156, 87, 0) : ico = "EXCESSO"
        Case "EXCESSO EXAGERADO": bgS = RGB(192, 0, 0)     : fgS = RGB(255, 255, 255): ico = "EXAGERADO"
        Case "INSUFICIENTE":      bgS = RGB(255, 199, 206) : fgS = RGB(156, 0, 6)  : ico = "FALTANDO"
        Case "SEM ANCORA":        bgS = RGB(255, 242, 204) : fgS = RGB(119, 107, 0): ico = "SEM ANCORA"
        Case "QTD ZERO":          bgS = RGB(230, 230, 230) : fgS = RGB(64, 64, 64) : ico = "ZERO"
        Case "ESTORNO SEM ENTRADA": bgS = RGB(230, 230, 230): fgS = RGB(64, 64, 64): ico = "ESTORNO"
        Case Else:                bgS = RGB(242, 242, 242) : fgS = RGB(89, 89, 89) : ico = st
    End Select
End Sub

' Formata as linhas de dados da racionalizacao em LOTE: base (fonte/borda/
' alinhamento) numa chamada so, zebra via formatacao condicional e cores
' semanticas aplicadas por BLOCOS de linhas consecutivas com o mesmo status.
Private Sub FormatarLinhasRacioNT(ws As Worksheet, stArr() As String, nRows As Long)
    If nRows < 1 Then Exit Sub
    Dim r1 As Long : r1 = 3
    Dim r2 As Long : r2 = r1 + nRows - 1

    With ws.Range(ws.Cells(r1, 1), ws.Cells(r2, 10))
        .Font.Name = "Segoe UI" : .Font.Size = 9
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous : .Borders.Color = RGB(225, 225, 225)
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
    End With
    ws.Range(ws.Cells(r1, 1), ws.Cells(r2, 3)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(r1, 7), ws.Cells(r2, 8)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(r1, 7), ws.Cells(r2, 7)).NumberFormat = "#,##0.##"
    ws.Rows(r1 & ":" & r2).RowHeight = 16

    ' zebra nas colunas descritivas (mesma tecnica MOD(ROW(),2) da ANALISE SAP x PRJ)
    On Error Resume Next
    Dim zb As Range : Set zb = ws.Range(ws.Cells(r1, 1), ws.Cells(r2, 8))
    zb.FormatConditions.Delete
    With zb.FormatConditions.Add(xlExpression, , "MOD(ROW(),2)=0")
        .Interior.Color = RGB(244, 247, 250)
    End With
    On Error GoTo 0

    Dim i As Long, iIni As Long, st As String
    Dim bgS As Long, fgS As Long, ico As String
    i = 1
    Do While i <= nRows
        st = stArr(i) : iIni = i
        Do While i < nRows
            If stArr(i + 1) <> st Then Exit Do
            i = i + 1
        Loop
        CoresStatusRacio st, bgS, fgS, ico
        With ws.Range(ws.Cells(r1 + iIni - 1, 9), ws.Cells(r1 + i - 1, 9))
            .Interior.Color = bgS : .Font.Color = fgS : .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        With ws.Range(ws.Cells(r1 + iIni - 1, 10), ws.Cells(r1 + i - 1, 10))
            .Interior.Color = bgS : .Font.Color = fgS
            .Font.Bold = (st = "EXCESSO EXAGERADO" Or st = "INSUFICIENTE" Or st = "EXCESSO")
        End With
        i = i + 1
    Loop
End Sub

' Ordena um vetor de strings in-place (quicksort; substitui o bubble sort O(n^2))
Private Sub QuickSortStr(a() As String, ByVal lo As Long, ByVal hi As Long)
    If lo >= hi Then Exit Sub
    Dim i As Long, j As Long, pv As String, tmp As String
    i = lo : j = hi : pv = a((lo + hi) \ 2)
    Do While i <= j
        Do While a(i) < pv : i = i + 1 : Loop
        Do While a(j) > pv : j = j - 1 : Loop
        If i <= j Then
            tmp = a(i) : a(i) = a(j) : a(j) = tmp
            i = i + 1 : j = j - 1
        End If
    Loop
    QuickSortStr a, lo, j
    QuickSortStr a, i, hi
End Sub

Private Sub FormatarColunasRacioNT(ws As Worksheet)
    Dim wds As Variant
    wds = Array(26, 24, 14, 42, 18, 40, 9, 12, 16, 46)
    Dim c As Long
    For c = 1 To 10
        ws.Columns(c).ColumnWidth = wds(c - 1)
    Next c
End Sub

Private Sub ProcessarAlertaCritico(wsBase As Worksheet, ws As Worksheet)

    ' Cabecalho PRIMEIRO: se algo abaixo falhar, a aba ainda tem estrutura e o
    ' anexo do diagnostico (AnexarAtencoesAlerta) tem onde se apoiar. Antes, uma
    ' falha ao carregar precos abortava a etapa com a aba totalmente vazia.
    EscreverCabecalhoAlerta ws, 1

    ' Precos de referencia so alimentam o alerta UC SUBVALORIZADO: se a base de
    ' precos falhar, a etapa segue sem esse alerta em vez de abortar inteira.
    mStep = "ALERTA CRITICO (precos de referencia)"
    Dim precos As Object
    On Error Resume Next
    Set precos = CarregarPrecos(wsBase.Parent)
    If Err.Number <> 0 Then
        AvisoExec "ALERTA CRITICO: precos de referencia indisponiveis, alerta UC " & _
                  "SUBVALORIZADO desativado nesta execucao - " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    If precos Is Nothing Then Set precos = CreateObject("Scripting.Dictionary")
    mStep = "ALERTA CRITICO (leitura da base)"

    Dim lastC As Long
    lastC = wsBase.Cells(1, wsBase.Columns.Count).End(xlToLeft).Column
    Dim idxPep3 As Long, idxPep4 As Long, idxCod As Long, idxDesc As Long
    Dim idxFam As Long, idxTipo As Long, idxValor As Long, idxQtd As Long, idxTipoPep As Long
    Dim c As Long
    Dim mapa As Object : Set mapa = MapaColunas(wsBase)
    idxPep3    = ColIdx(mapa, "PEP3NIVEL")
    idxPep4    = ColIdx(mapa, "PEP4NIVEL")
    idxCod     = ColIdx(mapa, "COD MAT")
    idxDesc    = ColIdx(mapa, "DESC MAT")
    idxFam     = ColIdx(mapa, "FAMILIA")
    idxTipo    = ColIdx(mapa, "TIPO")
    idxValor   = ColIdx(mapa, "VALOR")
    idxQtd     = ColIdx(mapa, "MAT LIB SAP")
    idxTipoPep = ColIdx(mapa, "TIPO PEP")

    ' outRow e declarado ANTES do early-exit abaixo: o GoTo Finaliza2 pula a
    ' secao de declaracoes seguinte, e Finaliza2 usa outRow para saber quantos
    ' alertas foram gerados (totAl = outRow - 2).
    Dim outRow As Long : outRow = 2
    Dim refCol As Long : refCol = idxPep4 : If refCol = 0 Then refCol = 1
    Dim lastRow As Long
    lastRow = wsBase.Cells(wsBase.Rows.Count, refCol).End(xlUp).Row
    If lastRow < 2 Then GoTo Finaliza2

    Dim d As Variant
    d = wsBase.Range(wsBase.Cells(2, 1), wsBase.Cells(lastRow, lastC)).Value
    Dim n As Long : n = lastRow - 1
    mStep = "ALERTA CRITICO (regras por linha)"

    ' Saida bufferizada (perf): alertas acumulados em memoria via AlertaCel e
    ' escritos na planilha numa unica atribuicao em Finaliza2
    mAlCap = 4 * n + 256
    ReDim mAlBuf(1 To 11, 1 To mAlCap)

    Dim p4pep3b  As Object : Set p4pep3b = CreateObject("Scripting.Dictionary")
    Dim p4uc     As Object : Set p4uc    = CreateObject("Scripting.Dictionary")
    Dim p4odi    As Object : Set p4odi   = CreateObject("Scripting.Dictionary")
    Dim p4com    As Object : Set p4com   = CreateObject("Scripting.Dictionary")  ' pep4 -> tem COM?
    Dim ucNzP4   As Object : Set ucNzP4  = CreateObject("Scripting.Dictionary")  ' pep4 -> tem UC com MAT LIB SAP <> 0?
    ' ODM: medidor x lacre (1 medidor -> 2 lacres)
    Dim medP4   As Object : Set medP4   = CreateObject("Scripting.Dictionary")  ' pep4 -> qtd medidores
    Dim lacreP4 As Object : Set lacreP4 = CreateObject("Scripting.Dictionary")  ' pep4 -> qtd lacres
    Dim medPep3 As Object : Set medPep3 = CreateObject("Scripting.Dictionary")  ' pep4 -> pep3
    Dim pep3All As Object : Set pep3All = CreateObject("Scripting.Dictionary")  ' pep3 -> pep4 representativo
    Dim pep3UCb As Object : Set pep3UCb = CreateObject("Scripting.Dictionary")  ' pep3 -> tem UC?
    Dim i As Long, pep3 As String, pep4 As String, tipo As String, tipoPep As String
    For i = 1 To n
        pep3 = "" : pep4 = "" : tipo = "" : tipoPep = ""
        If idxPep3 > 0 Then pep3 = Trim(CStr(d(i, idxPep3)))
        If idxPep4 > 0 Then pep4 = Trim(CStr(d(i, idxPep4)))
        If idxTipo > 0 Then tipo = NormStr(CStr(d(i, idxTipo)))
        If idxTipoPep > 0 Then tipoPep = NormStr(CStr(d(i, idxTipoPep)))
        If pep4 <> "" Then
            If Not p4pep3b.Exists(pep4) Then p4pep3b.Add pep4, pep3
            If Not p4odi.Exists(pep4) Then p4odi.Add pep4, False
            If tipoPep = "I" Or Right(pep4, 2) = ".I" Then p4odi(pep4) = True
            If tipo = "UC" And Not p4uc.Exists(pep4) Then p4uc.Add pep4, True
            ' UC com quantidade efetiva (MAT LIB SAP <> 0)
            If tipo = "UC" And idxQtd > 0 Then
                If Val0(d(i, idxQtd)) <> 0 And Not ucNzP4.Exists(pep4) Then ucNzP4.Add pep4, True
            End If
            If tipo = "COM" And Not p4com.Exists(pep4) Then p4com.Add pep4, True
        End If
        ' agrega por PEP3: existe e se tem alguma UC
        If pep3 <> "" Then
            If Not pep3All.Exists(pep3) Then pep3All.Add pep3, pep4
            If tipo = "UC" And Not pep3UCb.Exists(pep3) Then pep3UCb.Add pep3, True
        End If
    Next i

    Dim codMat As String, descMat As String, familia As String, famN As String
    Dim valor As Double, qtd As Double, unit As Double, refP As Double
    Dim puOdiMin As Double : puOdiMin = CfgD("PU_ODI_MIN", 0.01)
    Dim puOdiMax As Double : puOdiMax = CfgD("PU_ODI_MAX", 10)
    Dim puCabo As Double : puCabo = CfgD("PU_CABO_ALERTA", 0.01)
    Dim oddLim As Double : oddLim = LimiteValorODD()
    Dim posteMin As Double : posteMin = CfgD("POSTE_ODI_PU_MIN", 500)

    Dim kk As Variant
    For Each kk In p4pep3b.Keys
        If p4odi(kk) And Not ucNzP4.Exists(kk) Then
            AlertaCel outRow, 1, "ODI SEM UC"
            AlertaCel outRow, 2, p4pep3b(kk)
            AlertaCel outRow, 3, kk
            If p4uc.Exists(kk) Then
                AlertaCel outRow, 11, "ODI (PEP .I) com TODAS as UC zeradas (MAT LIB SAP = 0)"
            Else
                AlertaCel outRow, 11, "ODI (PEP .I) sem nenhum item TIPO=UC"
            End If
            outRow = outRow + 1
        End If
        ' ODI sem nenhum material COM
        If p4odi(kk) And Not p4com.Exists(kk) Then
            AlertaCel outRow, 1, "ODI SEM COM"
            AlertaCel outRow, 2, p4pep3b(kk)
            AlertaCel outRow, 3, kk
            AlertaCel outRow, 11, "ODI (PEP .I) sem nenhum material TIPO=COM"
            outRow = outRow + 1
        End If
    Next kk

    ' --- PEP3 sem nenhuma UC ---
    Dim kP3 As Variant
    For Each kP3 In pep3All.Keys
        If Not pep3UCb.Exists(kP3) Then
            AlertaCel outRow, 1, "PEP SEM UC"
            AlertaCel outRow, 2, kP3
            AlertaCel outRow, 3, pep3All(kP3)
            AlertaCel outRow, 11, "PEP3 sem nenhum item TIPO=UC cadastrado"
            outRow = outRow + 1
        End If
    Next kP3

    For i = 1 To n
        pep3 = "" : pep4 = "" : tipo = "" : codMat = "" : descMat = "" : familia = ""
        If idxPep3 > 0 Then pep3 = Trim(CStr(d(i, idxPep3)))
        If idxPep4 > 0 Then pep4 = Trim(CStr(d(i, idxPep4)))
        If idxTipo > 0 Then tipo = NormStr(CStr(d(i, idxTipo)))
        If idxCod  > 0 Then codMat  = Trim(CStr(d(i, idxCod)))
        If idxDesc > 0 Then descMat = Trim(CStr(d(i, idxDesc)))
        If idxFam  > 0 Then familia = Trim(CStr(d(i, idxFam)))
        famN = NormStr(familia)
        valor = 0 : qtd = 0
        If idxValor > 0 Then valor = Val0(d(i, idxValor))
        If idxQtd   > 0 Then qtd   = Val0(d(i, idxQtd))
        If pep4 = "" Then GoTo ProxLinha2

        If (Right(pep4, 2) = ".M" Or InStr(pep4, ".M.") > 0) And InStr(famN, "POSTE") > 0 Then
            AlertaCel outRow, 1, "POSTE EM PEP .M"
            AlertaCel outRow, 2, pep3
            AlertaCel outRow, 3, pep4
            AlertaCel outRow, 4, codMat
            AlertaCel outRow, 5, descMat
            AlertaCel outRow, 6, familia
            AlertaCel outRow, 7, valor
            AlertaCel outRow, 8, qtd
            AlertaCel outRow, 11, "Poste em PEP4 com sufixo .M (ODM)"
            outRow = outRow + 1
        End If

        If (Right(pep4, 2) = ".S" Or InStr(pep4, ".S.") > 0) And InStr(famN, "POSTE") > 0 Then
            AlertaCel outRow, 1, "POSTE EM PEP .S"
            AlertaCel outRow, 2, pep3
            AlertaCel outRow, 3, pep4
            AlertaCel outRow, 4, codMat
            AlertaCel outRow, 5, descMat
            AlertaCel outRow, 6, familia
            AlertaCel outRow, 7, valor
            AlertaCel outRow, 8, qtd
            AlertaCel outRow, 11, "Poste em PEP4 com sufixo .S (ODS)"
            outRow = outRow + 1
        End If

        ' Tipo de OD: sufixo do PEP4 ou, sem sufixo, a coluna TIPO PEP
        Dim tpPepRow As String : tpPepRow = ""
        If idxTipoPep > 0 Then tpPepRow = CStr(d(i, idxTipoPep))
        Dim odTp As String : odTp = ClassificarOD(pep4, tpPepRow)

        ' ODM: acumula medidores e lacres por PEP4 (regra 1 medidor = 2 lacres)
        If odTp = "ODM" Then
            Dim dN As String : dN = NormStr(descMat) & " " & famN
            If Not medPep3.Exists(pep4) Then medPep3.Add pep4, pep3
            If EhMedidorODM(dN) Then
                If medP4.Exists(pep4) Then medP4(pep4) = medP4(pep4) + qtd Else medP4.Add pep4, qtd
            ElseIf InStr(dN, "LACRE") > 0 Then
                If lacreP4.Exists(pep4) Then lacreP4(pep4) = lacreP4(pep4) + qtd Else lacreP4.Add pep4, qtd
            End If
        End If
        ' Premissa de sinal (valor E quantidade) - a mesma de AvaliarLinha, que
        ' marca estes itens como NAO ADERENTE na ANALISE SAP x PRJ
        Dim motSinal As String : motSinal = SinalInvalidoOD(odTp, qtd, valor)
        If motSinal <> "" And codMat <> "" Then
            AlertaCel outRow, 1, IIf(odTp = "ODD", "MATERIAL POSITIVO EM ODD", "MATERIAL NEGATIVO")
            AlertaCel outRow, 2, pep3
            AlertaCel outRow, 3, pep4
            AlertaCel outRow, 4, codMat
            AlertaCel outRow, 5, descMat
            AlertaCel outRow, 6, familia
            AlertaCel outRow, 7, valor
            AlertaCel outRow, 8, qtd
            AlertaCel outRow, 11, motSinal & " - item NAO ADERENTE"
            outRow = outRow + 1
        ElseIf odTp = "ODM" And qtd < 0 And codMat <> "" Then
            ' ODM fica fora da premissa de aderencia; segue apenas como alerta
            AlertaCel outRow, 1, "MATERIAL NEGATIVO"
            AlertaCel outRow, 2, pep3
            AlertaCel outRow, 3, pep4
            AlertaCel outRow, 4, codMat
            AlertaCel outRow, 5, descMat
            AlertaCel outRow, 6, familia
            AlertaCel outRow, 7, valor
            AlertaCel outRow, 8, qtd
            AlertaCel outRow, 11, "Material NEGATIVO em PEP ODM (esperado positivo)"
            outRow = outRow + 1
        End If

        If PuBaixoODI(odTp, famN, valor, qtd, puOdiMin, puOdiMax) And codMat <> "" Then
            AlertaCel outRow, 1, "PU BAIXO EM ODI"
            AlertaCel outRow, 2, pep3
            AlertaCel outRow, 3, pep4
            AlertaCel outRow, 4, codMat
            AlertaCel outRow, 5, descMat
            AlertaCel outRow, 6, familia
            AlertaCel outRow, 7, valor
            AlertaCel outRow, 8, qtd
            AlertaCel outRow, 9, valor / qtd
            AlertaCel outRow, 10, Format(puOdiMin, "#,##0.00") & " a " & Format(puOdiMax, "#,##0.00")
            AlertaCel outRow, 11, "Preco unitario R$ " & Format(valor / qtd, "#,##0.00") & _
                " em PEP ODI (faixa de alerta R$ " & Format(puOdiMin, "#,##0.00") & _
                " a R$ " & Format(puOdiMax, "#,##0.00") & ")"
            outRow = outRow + 1
        End If

        ' ODD com valor abaixo do limite - a mesma regra de AvaliarLinha, que marca
        ' o item como NAO ADERENTE na ANALISE SAP x PRJ
        Dim motOdd As String : motOdd = ValorBaixoODD(odTp, valor, oddLim)
        If motOdd <> "" And codMat <> "" Then
            AlertaCel outRow, 1, "ODD VALOR ABAIXO DO LIMITE"
            AlertaCel outRow, 2, pep3
            AlertaCel outRow, 3, pep4
            AlertaCel outRow, 4, codMat
            AlertaCel outRow, 5, descMat
            AlertaCel outRow, 6, familia
            AlertaCel outRow, 7, valor
            AlertaCel outRow, 8, qtd
            AlertaCel outRow, 10, Format(oddLim, "#,##0.00")
            AlertaCel outRow, 11, motOdd & " - item NAO ADERENTE"
            outRow = outRow + 1
        End If

        ' medidor e material de ODM: lancado em ODI esta no tipo de obra errado
        If MedidorEmODI(odTp, NormStr(descMat), famN) And codMat <> "" Then
            AlertaCel outRow, 1, "MEDIDOR EM PEP ODI"
            AlertaCel outRow, 2, pep3
            AlertaCel outRow, 3, pep4
            AlertaCel outRow, 4, codMat
            AlertaCel outRow, 5, descMat
            AlertaCel outRow, 6, familia
            AlertaCel outRow, 7, valor
            AlertaCel outRow, 8, qtd
            AlertaCel outRow, 11, "Medidor em PEP ODI (medidor e material de ODM)"
            outRow = outRow + 1
        End If

        If PosteBaratoODI(odTp, famN, valor, qtd, posteMin) And codMat <> "" Then
            AlertaCel outRow, 1, "POSTE ODI PU BAIXO"
            AlertaCel outRow, 2, pep3
            AlertaCel outRow, 3, pep4
            AlertaCel outRow, 4, codMat
            AlertaCel outRow, 5, descMat
            AlertaCel outRow, 6, familia
            AlertaCel outRow, 7, valor
            AlertaCel outRow, 8, qtd
            AlertaCel outRow, 9, valor / qtd
            AlertaCel outRow, 10, Format(posteMin, "#,##0.00")
            AlertaCel outRow, 11, "Poste em PEP ODI com preco unitario R$ " & Format(valor / qtd, "#,##0.00") & _
                " abaixo de R$ " & Format(posteMin, "#,##0.00")
            outRow = outRow + 1
        End If

        If PuSimbolicoCabo(famN, valor, qtd, puCabo) And codMat <> "" Then
            AlertaCel outRow, 1, "PU 0,01 EM CABO/COND"
            AlertaCel outRow, 2, pep3
            AlertaCel outRow, 3, pep4
            AlertaCel outRow, 4, codMat
            AlertaCel outRow, 5, descMat
            AlertaCel outRow, 6, familia
            AlertaCel outRow, 7, valor
            AlertaCel outRow, 8, qtd
            AlertaCel outRow, 9, valor / qtd
            AlertaCel outRow, 10, Format(puCabo, "#,##0.00")
            AlertaCel outRow, 11, "Cabo/condutor com preco unitario R$ " & Format(valor / qtd, "#,##0.00") & _
                " (valor simbolico) em PEP " & IIf(odTp <> "", odTp, pep4)
            outRow = outRow + 1
        End If

        If tipo = "UC" And Left(famN, 4) <> "COND" And Left(famN, 4) <> "CABO" Then
            If codMat <> "" Then
                Dim codMatNorm As String : codMatNorm = NormCod(codMat)
                If precos.Exists(codMatNorm) And qtd > 0 And valor > 0 Then
                    refP = precos(codMatNorm) : unit = valor / qtd
                    Dim tolSub As Double : tolSub = CfgD("TOL_SUBVAL", 0.9)
                    Dim minDiv As Double : minDiv = CfgD("MIN_DIVERG_RS", 100)
                    If refP > 0 And unit < refP * tolSub And (refP - unit) * qtd >= minDiv Then
                        AlertaCel outRow, 1, "UC SUBVALORIZADO"
                        AlertaCel outRow, 2, pep3
                        AlertaCel outRow, 3, pep4
                        AlertaCel outRow, 4, codMat
                        AlertaCel outRow, 5, descMat
                        AlertaCel outRow, 6, familia
                        AlertaCel outRow, 7, valor
                        AlertaCel outRow, 8, qtd
                        AlertaCel outRow, 9, unit
                        AlertaCel outRow, 10, refP
                        AlertaCel outRow, 11, "PU " & Format(unit, "#,##0.00") & " abaixo de " & Format(tolSub, "0%") & " da referencia (" & Format(refP, "#,##0.00") & ")"
                        outRow = outRow + 1
                    End If
                ElseIf codMatNorm <> "" And qtd > 0 And valor > 0 Then
                    AlertaCel outRow, 1, "UC - PRECO NAO ENCONTRADO"
                    AlertaCel outRow, 2, pep3
                    AlertaCel outRow, 3, pep4
                    AlertaCel outRow, 4, codMat
                    AlertaCel outRow, 5, descMat
                    AlertaCel outRow, 6, familia
                    AlertaCel outRow, 7, valor
                    AlertaCel outRow, 8, qtd
                    AlertaCel outRow, 11, "Material nao tem preco referencia cadastrado na aba PRECOS"
                    outRow = outRow + 1
                End If
            Else
                AlertaCel outRow, 1, "UC - COD MATERIAL VAZIO"
                AlertaCel outRow, 2, pep3
                AlertaCel outRow, 3, pep4
                AlertaCel outRow, 5, descMat
                AlertaCel outRow, 6, familia
                AlertaCel outRow, 7, valor
                AlertaCel outRow, 11, "Codigo de material nao preenchido"
                outRow = outRow + 1
            End If
        End If
ProxLinha2:
    Next i

    ' --- ODM: valida 1 medidor = 2 lacres (por PEP4) ---
    Dim kMed As Variant
    For Each kMed In medP4.Keys
        Dim qMed As Double : qMed = medP4(kMed)
        Dim qLac As Double : qLac = 0
        If lacreP4.Exists(kMed) Then qLac = lacreP4(kMed)
        Dim espLac As Double : espLac = 2 * qMed
        If qMed > 0 And qLac <> espLac Then
            AlertaCel outRow, 1, "LACRE x MEDIDOR"
            AlertaCel outRow, 2, medPep3(kMed)
            AlertaCel outRow, 3, kMed
            AlertaCel outRow, 6, "MEDIDOR / LACRE"
            AlertaCel outRow, 8, qLac
            AlertaCel outRow, 10, espLac
            AlertaCel outRow, 11, Format(qMed, "#,##0") & " medidor(es) -> esperado " & _
                Format(espLac, "#,##0") & " lacres, veio " & Format(qLac, "#,##0") & _
                IIf(qLac < espLac, " (FALTAM " & Format(espLac - qLac, "#,##0") & ")", _
                                   " (SOBRAM " & Format(qLac - espLac, "#,##0") & ")")
            outRow = outRow + 1
        End If
    Next kMed
    ' lacre sem medidor no ODM
    Dim kLac As Variant
    For Each kLac In lacreP4.Keys
        If Not medP4.Exists(kLac) And lacreP4(kLac) <> 0 Then
            AlertaCel outRow, 1, "LACRE x MEDIDOR"
            AlertaCel outRow, 2, medPep3(kLac)
            AlertaCel outRow, 3, kLac
            AlertaCel outRow, 6, "MEDIDOR / LACRE"
            AlertaCel outRow, 8, lacreP4(kLac)
            AlertaCel outRow, 10, 0
            AlertaCel outRow, 11, "Lacre sem medidor no PEP ODM (veio " & Format(lacreP4(kLac), "#,##0") & ")"
            outRow = outRow + 1
        End If
    Next kLac

    ' --- Importa discrepancias de PRECO UNITARIO ---
    Dim wsPU As Worksheet
    On Error Resume Next
    Set wsPU = ws.Parent.Worksheets("PRECO UNITARIO")
    On Error GoTo 0
    If Not wsPU Is Nothing Then
        Dim lastRowP As Long : lastRowP = wsPU.Cells(wsPU.Rows.Count, 1).End(xlUp).Row
        If lastRowP >= 2 Then
            ' 1 PEP3|2 PEP4|3 TIPO|4 COD|5 DESC|6 UND|7 QTD|8 VALOR|9 PU|10 MIN|11 MAX|12 STATUS|13 OBS
            Dim pD As Variant : pD = wsPU.Range(wsPU.Cells(2, 1), wsPU.Cells(lastRowP, 13)).Value
            Dim rp As Long, stP As String
            For rp = 1 To UBound(pD, 1)
                stP = Trim(CStr(pD(rp, 12)))
                If stP = "ABAIXO DO MINIMO" Or stP = "ACIMA DO MAXIMO" Then
                    AlertaCel outRow, 1, IIf(stP = "ABAIXO DO MINIMO", "PU ABAIXO MIN", "PU ACIMA MAX")
                    AlertaCel outRow, 2, pD(rp, 1)   ' PEP3
                    AlertaCel outRow, 3, pD(rp, 2)   ' PEP4
                    AlertaCel outRow, 4, pD(rp, 4)   ' COD
                    AlertaCel outRow, 5, pD(rp, 5)   ' DESC
                    AlertaCel outRow, 6, pD(rp, 3)   ' TIPO (na col familia)
                    AlertaCel outRow, 7, pD(rp, 8)   ' VALOR
                    AlertaCel outRow, 8, pD(rp, 7)   ' QTD
                    AlertaCel outRow, 9, pD(rp, 9)   ' PU
                    AlertaCel outRow, 10, IIf(stP = "ABAIXO DO MINIMO", pD(rp, 10), pD(rp, 11)) ' ref
                    AlertaCel outRow, 11, pD(rp, 13) ' MOTIVO (obs)
                    outRow = outRow + 1
                End If
            Next rp
        End If
    End If

Finaliza2:
    mStep = "ALERTA CRITICO (montagem da aba)"
    Dim totAl As Long : totAl = outRow - 2

    ' Despeja o buffer de alertas na planilha numa UNICA atribuicao de Range
    ' (o restante da formatacao le os dados de volta a partir da planilha)
    If totAl > 0 Then
        Dim dumpA() As Variant : ReDim dumpA(1 To totAl, 1 To 11)
        Dim rDp As Long, cDp As Long
        For rDp = 1 To totAl
            For cDp = 1 To 11
                dumpA(rDp, cDp) = mAlBuf(cDp, rDp)
            Next cDp
        Next rDp
        ws.Range(ws.Cells(2, 1), ws.Cells(outRow - 1, 11)).Value = dumpA
    End If
    Erase mAlBuf : mAlCap = 0
    Dim wds As Variant
    wds = Array(22, 24, 26, 12, 46, 18, 12, 9, 12, 12, 42)
    For c = 1 To 11
        ws.Columns(c).ColumnWidth = wds(c - 1)
    Next c

    ' Formatacao em LOTE: le os dados de volta em array (1 chamada), calcula o
    ' PU em memoria, aplica a base numa chamada so e pinta BLOCOS de linhas
    ' consecutivas com o mesmo tipo de alerta.
    If totAl > 0 Then
        Dim dA As Variant
        dA = ws.Range(ws.Cells(2, 1), ws.Cells(outRow - 1, 11)).Value

        ' PRECO UNIT = VALOR / QTD (col 9 = col 7 / col 8), calculado em memoria
        Dim puCol() As Variant : ReDim puCol(1 To totAl, 1 To 1)
        Dim rAlert As Long
        For rAlert = 1 To totAl
            Dim vAl As Double : vAl = Val0(dA(rAlert, 7))
            Dim qAl As Double : qAl = Val0(dA(rAlert, 8))
            If qAl <> 0 Then puCol(rAlert, 1) = vAl / qAl Else puCol(rAlert, 1) = dA(rAlert, 9)
        Next rAlert
        ws.Range(ws.Cells(2, 9), ws.Cells(outRow - 1, 9)).Value = puCol

        With ws.Range(ws.Cells(2, 1), ws.Cells(outRow - 1, 11))
            .Font.Name = "Segoe UI" : .Font.Size = 10 : .Font.Color = RGB(45, 52, 64)
            .VerticalAlignment = xlCenter
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Color = RGB(255, 255, 255)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(255, 255, 255)
        End With
        ws.Range(ws.Cells(2, 7), ws.Cells(outRow - 1, 10)).NumberFormat = "#,##0.00"

        Dim iA As Long, iIni As Long, tA As String
        iA = 1
        Do While iA <= totAl
            tA = Trim(CStr(dA(iA, 1))) : iIni = iA
            Do While iA < totAl
                If Trim(CStr(dA(iA + 1, 1))) <> tA Then Exit Do
                iA = iA + 1
            Loop
            Dim bgA As Long, chip As Long
            CoresAlerta tA, bgA, chip
            ws.Range(ws.Cells(1 + iIni, 1), ws.Cells(1 + iA, 11)).Interior.Color = bgA
            ' "chip" colorido do tipo de alerta (col 1)
            With ws.Range(ws.Cells(1 + iIni, 1), ws.Cells(1 + iA, 1))
                .Interior.Color = chip
                .Font.Bold = True : .Font.Size = 9 : .Font.Color = RGB(255, 255, 255)
                .HorizontalAlignment = xlCenter
            End With
            iA = iA + 1
        Loop
    End If

    ' --- Resumo por tipo: contagem e primeira linha de cada tipo ---
    ' Os tipos se intercalam na lista (o loop principal emite varios tipos por
    ' linha da base), entao a contagem vai por dicionario, e nao pelos blocos
    ' contiguos usados na pintura acima.
    Dim tpNome() As String, tpQtd() As Long, tpLin() As Long, nTp As Long
    If totAl > 0 Then
        Dim dCnt As Object : Set dCnt = CreateObject("Scripting.Dictionary")
        Dim dFst As Object : Set dFst = CreateObject("Scripting.Dictionary")
        Dim rS As Long, tS As String
        For rS = 1 To totAl
            tS = Trim(CStr(dA(rS, 1)))
            If tS <> "" Then
                If dCnt.Exists(tS) Then
                    dCnt(tS) = dCnt(tS) + 1
                Else
                    dCnt.Add tS, 1
                    dFst.Add tS, 1 + rS      ' linha na planilha ANTES da insercao do topo
                End If
            End If
        Next rS
        nTp = dCnt.Count
        If nTp > 0 Then
            ReDim tpNome(1 To nTp) : ReDim tpQtd(1 To nTp) : ReDim tpLin(1 To nTp)
            Dim kS As Variant, ixS As Long : ixS = 0
            For Each kS In dCnt.Keys
                ixS = ixS + 1
                tpNome(ixS) = CStr(kS) : tpQtd(ixS) = dCnt(kS) : tpLin(ixS) = dFst(kS)
            Next kS
            ' ordena por quantidade desc (selection sort: no maximo ~15 tipos)
            Dim iSel As Long, jSel As Long, mSel As Long
            For iSel = 1 To nTp - 1
                mSel = iSel
                For jSel = iSel + 1 To nTp
                    If tpQtd(jSel) > tpQtd(mSel) Then mSel = jSel
                Next jSel
                If mSel <> iSel Then
                    Dim swN As String, swQ As Long, swL As Long
                    swN = tpNome(iSel) : tpNome(iSel) = tpNome(mSel) : tpNome(mSel) = swN
                    swQ = tpQtd(iSel) : tpQtd(iSel) = tpQtd(mSel) : tpQtd(mSel) = swQ
                    swL = tpLin(iSel) : tpLin(iSel) = tpLin(mSel) : tpLin(mSel) = swL
                End If
            Next iSel
        End If
    End If

    ' --- Banda de titulo + resumo no topo ---
    ' Linhas inseridas: 2 de titulo + N de chips + 1 de respiro. Como a altura
    ' varia, quem le esta aba usa PrimeiraLinhaAlerta em vez de linha fixa.
    Dim nLinSum As Long, nIns As Long
    If nTp > 0 Then nLinSum = -Int(-nTp / 3)          ' teto da divisao por 3
    If nTp > 0 Then nIns = 2 + nLinSum + 1 Else nIns = 2
    ws.Range("1:" & nIns).Insert Shift:=xlDown

    EscreverTituloAlerta ws

    If nTp > 0 Then
        Dim cIni As Variant : cIni = Array(1, 4, 7)
        Dim cFim As Variant : cFim = Array(3, 6, 11)
        Dim t As Long
        For t = 1 To nTp
            Dim rowCh As Long : rowCh = 3 + Int((t - 1) / 3)
            Dim slot As Long : slot = (t - 1) Mod 3
            Dim bgS As Long, chS As Long
            CoresAlerta tpNome(t), bgS, chS
            With ws.Range(ws.Cells(rowCh, cIni(slot)), ws.Cells(rowCh, cFim(slot)))
                .Merge
                .Value = tpNome(t) & "   " & ChrW(&H2022) & "   " & Format(tpQtd(t), "#,##0")
                .Interior.Color = chS
                .Font.Name = "Segoe UI Semibold" : .Font.Size = 10 : .Font.Bold = True
                .Font.Color = RGB(255, 255, 255)
                .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
            End With
            ' atalho para o primeiro apontamento do tipo (os blocos sao pintados
            ' por tipo, entao cair na primeira linha ja mostra o grupo)
            On Error Resume Next
            ws.Hyperlinks.Add Anchor:=ws.Cells(rowCh, cIni(slot)), Address:="", _
                SubAddress:="'" & ws.Name & "'!A" & (tpLin(t) + nIns), _
                ScreenTip:="Ir para o primeiro: " & tpNome(t)
            On Error GoTo 0
            ' Hyperlinks.Add repinta a celula com o estilo de link
            With ws.Cells(rowCh, cIni(slot)).Font
                .Name = "Segoe UI Semibold" : .Size = 10 : .Bold = True
                .Color = RGB(255, 255, 255) : .Underline = xlUnderlineStyleNone
            End With
            ws.Rows(rowCh).RowHeight = 20
        Next t
        ws.Rows(3 + nLinSum).RowHeight = 6      ' respiro antes do cabecalho
    End If

    Dim rHdr As Long : rHdr = nIns + 1
    If totAl = 0 Then
        ws.Cells(rHdr + 1, 1).Value = "Nenhum alerta critico encontrado nesta base."
        ws.Cells(rHdr + 1, 1).Font.Italic = True
        ws.Cells(rHdr + 1, 1).Font.Color = RGB(0, 128, 0)
    Else
        ws.Range(ws.Cells(rHdr, 1), ws.Cells(rHdr, 11)).AutoFilter
    End If

    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range("A" & (rHdr + 1)).Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
    ws.Range("A1").Select
    mStep = "ALERTA CRITICO"     ' concluiu: o log registra a etapa pelo nome normal
End Sub

' Paleta por tipo de alerta: cor de fundo da linha e cor do "chip" da coluna 1.
' Fonte unica, usada tanto na pintura das linhas quanto nos chips do resumo.
Private Sub CoresAlerta(ByVal tipo As String, ByRef bgA As Long, ByRef chip As Long)
    Select Case tipo
        Case "ODI SEM UC":                bgA = RGB(255, 244, 214) : chip = RGB(214, 158, 0)
        Case "PEP SEM UC":                bgA = RGB(255, 238, 230) : chip = RGB(196, 89, 17)
        Case "ODI SEM COM":               bgA = RGB(230, 240, 248) : chip = RGB(46, 117, 182)
        Case "PU ABAIXO MIN":             bgA = RGB(255, 244, 214) : chip = RGB(176, 124, 0)
        Case "PU ACIMA MAX":              bgA = RGB(252, 226, 228) : chip = RGB(192, 0, 0)
        Case "MATERIAL NEGATIVO":         bgA = RGB(252, 226, 228) : chip = RGB(192, 0, 0)
        Case "MATERIAL POSITIVO EM ODD":  bgA = RGB(252, 233, 219) : chip = RGB(214, 99, 0)
        Case "LACRE x MEDIDOR":           bgA = RGB(225, 223, 245) : chip = RGB(96, 64, 196)
        Case "POSTE EM PEP .M":           bgA = RGB(222, 236, 250) : chip = RGB(31, 107, 183)
        Case "POSTE EM PEP .S":           bgA = RGB(220, 245, 240) : chip = RGB(15, 130, 115)
        Case "UC SUBVALORIZADO":          bgA = RGB(252, 226, 228) : chip = RGB(192, 0, 0)
        Case "UC - PRECO NAO ENCONTRADO": bgA = RGB(255, 244, 214) : chip = RGB(176, 124, 0)
        Case "UC - COD MATERIAL VAZIO":   bgA = RGB(238, 238, 238) : chip = RGB(110, 110, 110)
        Case "PU BAIXO EM ODI":           bgA = RGB(252, 231, 243) : chip = RGB(190, 24, 93)
        Case "PU 0,01 EM CABO/COND":      bgA = RGB(255, 237, 213) : chip = RGB(194, 65, 12)
        Case "ODD VALOR ABAIXO DO LIMITE": bgA = RGB(252, 226, 228) : chip = RGB(153, 27, 27)
        Case "POSTE ODI PU BAIXO":        bgA = RGB(224, 231, 255) : chip = RGB(55, 48, 163)
        Case "MEDIDOR EM PEP ODI":        bgA = RGB(237, 233, 254) : chip = RGB(109, 40, 217)
        Case Else:                        bgA = RGB(245, 245, 245) : chip = RGB(120, 120, 120)
    End Select
End Sub

' Primeira linha de DADOS da aba ALERTA CRITICO. O topo tem banda de titulo e,
' quando ha apontamentos, um resumo por tipo de altura variavel - por isso a
' linha do cabecalho nao pode ser fixa. Procura a linha cujo A = "TIPO ALERTA".
Private Function PrimeiraLinhaAlerta(wsA As Worksheet) As Long
    Dim r As Long : r = LinhaCabecalhoAlerta(wsA)
    If r > 0 Then PrimeiraLinhaAlerta = r + 1 Else PrimeiraLinhaAlerta = 4
End Function

' Linha do cabecalho (A = "TIPO ALERTA") da aba ALERTA CRITICO; 0 se nao houver.
Private Function LinhaCabecalhoAlerta(wsA As Worksheet) As Long
    Dim r As Long
    For r = 1 To 40
        If NormStr(CStr(wsA.Cells(r, 1).Value)) = "TIPO ALERTA" Then
            LinhaCabecalhoAlerta = r
            Exit Function
        End If
    Next r
End Function

' Cabecalho da tabela da aba ALERTA CRITICO, na linha indicada.
Private Sub EscreverCabecalhoAlerta(ws As Worksheet, ByVal linha As Long)
    Dim hdrs As Variant
    hdrs = Array("TIPO ALERTA", "PEP3NIVEL", "PEP4NIVEL", "COD MAT", "DESC MAT", "FAMILIA", _
                 "VALOR", "QTD", "PRECO UNIT", "REFERENCIA", "MOTIVO")
    Dim c As Long
    For c = 1 To 11
        With ws.Cells(linha, c)
            .Value = hdrs(c - 1)
            .Font.Name = "Segoe UI Semibold" : .Font.Bold = True : .Font.Size = 10
            .Font.Color = RGB(235, 240, 248)
            .Interior.Color = RGB(17, 24, 39)        ' navy escuro
            .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous : .Borders.Color = RGB(60, 70, 90)
        End With
    Next c
    ws.Rows(linha).RowHeight = 30
End Sub

' Banda de titulo (A1:K2) da aba ALERTA CRITICO.
Private Sub EscreverTituloAlerta(ws As Worksheet)
    With ws.Range("A1:K2")
        .Merge
        .Value = ChrW(&H26A0) & "  PAINEL DE ALERTAS CRITICOS"
        .Font.Name = "Segoe UI Semibold" : .Font.Size = 18 : .Font.Bold = True
        .Font.Color = RGB(255, 255, 255) : .Interior.Color = RGB(176, 0, 0)
        .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 22 : ws.Rows(2).RowHeight = 22
End Sub

' Escreve uma celula de alerta no buffer (substitui ws.Cells(outRow, c).Value =).
' Cresce o buffer sob demanda; outRow comeca em 2 (linha 1 = cabecalho).
Private Sub AlertaCel(outRow As Long, c As Long, v As Variant)
    ' buffer ainda nao alocado (chamada fora da ordem esperada): aloca em vez
    ' de estourar no ReDim Preserve com mAlCap = 0, que faria laco infinito
    If mAlCap < 1 Then
        mAlCap = 256
        ReDim mAlBuf(1 To 11, 1 To mAlCap)
    End If
    Do While outRow - 1 > mAlCap
        mAlCap = mAlCap * 2
        ReDim Preserve mAlBuf(1 To 11, 1 To mAlCap)
    Loop
    mAlBuf(c, outRow - 1) = v
End Sub

Private Function Val0(v As Variant) As Double
    On Error Resume Next
    If IsNumeric(v) Then Val0 = CDbl(v) Else Val0 = 0
    On Error GoTo 0
End Function

' Carrega precos de uma aba "PRECOS" (codigo na col 1, preco na col 2)
' Retorna Dictionary vazio se aba nao existir
Private Function CarregarPrecos(wb As Workbook) As Object
    Dim precos As Object : Set precos = CreateObject("Scripting.Dictionary")

    Dim wsPrecos As Worksheet
    On Error Resume Next
    Set wsPrecos = wb.Worksheets("PRECOS")
    On Error GoTo 0

    If Not wsPrecos Is Nothing Then
        Dim lastRow As Long
        lastRow = wsPrecos.Cells(wsPrecos.Rows.Count, 1).End(xlUp).Row
        If lastRow >= 2 Then
            ' leitura em array (1 chamada) em vez de celula a celula
            Dim dPr As Variant : dPr = wsPrecos.Range(wsPrecos.Cells(2, 1), wsPrecos.Cells(lastRow, 2)).Value
            Dim r As Long
            For r = 1 To UBound(dPr, 1)
                Dim cod As String : cod = ""
                If Not IsError(dPr(r, 1)) Then cod = NormCod(dPr(r, 1))
                Dim preco As Double : preco = Val0(dPr(r, 2))
                If cod <> "" And preco > 0 Then
                    If Not precos.Exists(cod) Then precos.Add cod, preco
                End If
            Next r
        End If
    End If

    If precos.Count > 0 Then mFontePrecosRef = "aba PRECOS (" & Format(precos.Count, "#,##0") & " materiais)"
    ' Fallback: se nao houver aba "PRECOS", deriva preco = ponto medio (MIN+MAX)/2
    ' da BASE DE PRECOS (aba interna ou arquivo externo).
    If precos.Count = 0 Then
        Dim fx As Object : Set fx = CarregarFaixaPrecos(wb)
        Dim k As Variant, p() As String, mn As Double, mx As Double, md As Double
        For Each k In fx.Keys
            p = Split(CStr(fx(k)), "|")
            mn = Val0(p(0)) : mx = Val0(p(1))
            If mx <= 0 Then mx = mn
            md = (mn + mx) / 2
            If md > 0 Then precos.Add CStr(k), md
        Next k
        mFontePrecosRef = "ponto medio MIN/MAX da faixa de precos (" & Format(precos.Count, "#,##0") & " materiais)"
    End If

    Set CarregarPrecos = precos
End Function

' Localiza o arquivo externo da base de precos. NAO usa literais acentuados
' (que corrompem na importacao do .bas no VBA). Usa ChrW(193)=A-acento e
' curinga '?' no lugar do C-cedilha. Retorna o caminho completo ou "".
Private Function CaminhoBasePrecos(wb As Workbook) As String
    ' '?' casa PRECOS e PRE<cedilha>OS; nao pega 'BASE DE PRECOS A CLASSIFICAR.xlsx'
    Dim padrao As String : padrao = "BASE DE PRE?OS.xlsx"

    ' 'Area de Trabalho' montado com ChrW p/ evitar corrupcao de encoding
    Dim areaTrab As String : areaTrab = ChrW(193) & "rea de Trabalho"
    Dim up As String : up = Environ$("USERPROFILE")

    Dim pastas() As String
    pastas = Split( _
        wb.Path & "|" & _
        up & "\OneDrive - GRUPO EQUATORIAL ENERGIA\" & areaTrab & "\claude" & "|" & _
        up & "\OneDrive\" & areaTrab & "\claude" & "|" & _
        up & "\" & areaTrab & "\claude" & "|" & _
        up & "\Desktop\claude", "|")

    Dim i As Long, ach As String
    For i = LBound(pastas) To UBound(pastas)
        If Len(pastas(i)) > 0 Then
            ach = Dir(pastas(i) & "\" & padrao)
            If Len(ach) > 0 Then
                CaminhoBasePrecos = pastas(i) & "\" & ach
                Exit Function
            End If
        End If
    Next i
    CaminhoBasePrecos = ""
End Function

' Carrega faixa MIN/MAX de preco unitario.
' 1) tenta aba interna "BASE PRECOS"/"BASE DE PRECOS"/"BASE DE PRECOS" (com cedilha);
' 2) se nao houver, abre o arquivo externo BASE DE PRECOS.xlsx (read-only).
' Colunas: MATERIAL | TEXTO MATERIAL | UML | MIN PU | MAX PU | CLS2 | CLASSIFICACAO
' Retorna Dictionary: codNorm -> "minPU|maxPU|texto"
Private Function CarregarFaixaPrecos(wb As Workbook) As Object
    ' usa cache de execucao se disponivel
    If Not mFaixaCache Is Nothing Then
        Set CarregarFaixaPrecos = mFaixaCache : Exit Function
    End If
    Dim fx As Object : Set fx = CreateObject("Scripting.Dictionary")
    Dim wsP As Worksheet
    Dim wbExt As Workbook       ' workbook externo aberto por nos (p/ fechar depois)
    Set wbExt = Nothing
    ' sem fonte o cache nao e gravado e a funcao roda de novo na etapa seguinte:
    ' avisa so na primeira vez
    Dim jaAvisou As Boolean : jaAvisou = (mFontePrecos <> "")

    ' --- 1) aba interna ---
    On Error Resume Next
    Set wsP = wb.Worksheets("BASE PRECOS")
    If wsP Is Nothing Then Set wsP = wb.Worksheets("BASE DE PRECOS")
    If wsP Is Nothing Then Set wsP = wb.Worksheets("BASE DE PRE" & ChrW(199) & "OS")
    On Error GoTo 0
    If Not wsP Is Nothing Then mFontePrecos = "aba interna '" & wsP.Name & "'" : mFontePrecosData = Empty

    ' --- 2) arquivo externo ---
    If wsP Is Nothing Then
        Dim cam As String : cam = CaminhoBasePrecos(wb)
        If Len(cam) = 0 Then
            mFontePrecos = "NENHUMA (sem aba BASE PRECOS e sem arquivo BASE DE PRECOS.xlsx nas pastas conhecidas)"
            mFontePrecosData = Empty
            If Not jaAvisou Then AvisoExec "PRECOS: nenhuma base de precos encontrada - PRECO UNITARIO " & _
                                          "e alertas de preco ficam sem referencia"
            Set CarregarFaixaPrecos = fx : Exit Function
        End If
        mFontePrecos = "arquivo externo '" & cam & "'"
        mFontePrecosData = Empty
        On Error Resume Next
        mFontePrecosData = FileDateTime(cam)
        On Error GoTo 0
        Dim nomeArq As String : nomeArq = Dir(cam)
        On Error Resume Next
        Dim wbJa As Workbook : Set wbJa = Workbooks(nomeArq)   ' ja aberto?
        On Error GoTo 0
        If wbJa Is Nothing Then
            On Error Resume Next
            Set wbExt = Workbooks.Open(cam, ReadOnly:=True, UpdateLinks:=0)
            On Error GoTo 0
            If wbExt Is Nothing Then
                mFontePrecos = "arquivo externo '" & cam & "' NAO ABRIU"
                If Not jaAvisou Then AvisoExec "PRECOS: o arquivo " & cam & " nao abriu - PRECO UNITARIO " & _
                                              "e alertas de preco ficam sem referencia"
                Set CarregarFaixaPrecos = fx : Exit Function
            End If
            Set wsP = wbExt.Worksheets(1)
        Else
            Set wsP = wbJa.Worksheets(1)
        End If
    End If
    If wsP Is Nothing Then Set CarregarFaixaPrecos = fx : Exit Function

    Dim lastC As Long : lastC = wsP.Cells(1, wsP.Columns.Count).End(xlToLeft).Column
    Dim cMat As Long, cTxt As Long, cMin As Long, cMax As Long
    Dim c As Long, h As String
    For c = 1 To lastC
        h = NormStr(CStr(wsP.Cells(1, c).Value))
        Select Case h
            Case "MATERIAL":        cMat = c
            Case "TEXTO MATERIAL":  cTxt = c
            Case "MIN PU":          cMin = c
            Case "MAX PU":          cMax = c
        End Select
    Next c
    If cMat = 0 Then cMat = 1
    If cTxt = 0 Then cTxt = 2
    If cMin = 0 Then cMin = 4
    If cMax = 0 Then cMax = 5
    ' base com menos colunas que as posicoes-padrao: ler ao menos ate a 5a evita
    ' estourar o array (erro 9) - a coluna ausente e lida como vazia
    If lastC < 5 Then lastC = 5

    Dim lastR As Long : lastR = wsP.Cells(wsP.Rows.Count, cMat).End(xlUp).Row
    If lastR >= 2 Then
        Dim dd As Variant : dd = wsP.Range(wsP.Cells(2, 1), wsP.Cells(lastR, lastC)).Value
        Dim r As Long
        For r = 1 To UBound(dd, 1)
            ' IsError: celula com #N/D derrubava a carga inteira (CStr falha nela)
            Dim cod As String : cod = ""
            If Not IsError(dd(r, cMat)) Then cod = NormCod(dd(r, cMat))
            If cod <> "" And Not fx.Exists(cod) Then
                Dim txtM As String : txtM = ""
                If Not IsError(dd(r, cTxt)) Then txtM = Trim(CStr(dd(r, cTxt)))
                fx.Add cod, Val0(dd(r, cMin)) & "|" & Val0(dd(r, cMax)) & "|" & txtM
            End If
        Next r
    End If

    If Not wbExt Is Nothing Then wbExt.Close SaveChanges:=False
    If fx.Count = 0 Then
        mFontePrecos = mFontePrecos & " (vazia: nenhum material lido)"
        If Not jaAvisou Then AvisoExec "PRECOS: a base de precos usada nao tem materiais legiveis"
    Else
        mFontePrecos = mFontePrecos & " - " & Format(fx.Count, "#,##0") & " materiais"
    End If
    Set mFaixaCache = fx
    Set CarregarFaixaPrecos = fx
End Function

' ===========================================================================
'  ABA PRECO UNITARIO: PU = VALOR/QTD x faixa MIN/MAX (BASE PRECOS)
' ===========================================================================
Private Sub ProcessarPrecoUnitario(wsBase As Worksheet, ws As Worksheet)

    Dim fx As Object : Set fx = CarregarFaixaPrecos(wsBase.Parent)

    Dim lastC As Long : lastC = wsBase.Cells(1, wsBase.Columns.Count).End(xlToLeft).Column
    Dim iPep3 As Long, iPep4 As Long, iCod As Long, iDesc As Long
    Dim iFam As Long, iTipo As Long, iVal As Long, iQtd As Long, iUml As Long
    Dim iTipoPep As Long
    Dim c As Long
    Dim mapa As Object : Set mapa = MapaColunas(wsBase)
    iPep3    = ColIdx(mapa, "PEP3NIVEL")
    iPep4    = ColIdx(mapa, "PEP4NIVEL")
    iCod     = ColIdx(mapa, "COD MAT")
    iDesc    = ColIdx(mapa, "DESC MAT")
    iFam     = ColIdx(mapa, "FAMILIA")
    iTipo    = ColIdx(mapa, "TIPO")
    iVal     = ColIdx(mapa, "VALOR")
    iQtd     = ColIdx(mapa, "MAT LIB SAP")
    iUml     = ColIdx(mapa, "UND")
    iTipoPep = ColIdx(mapa, "TIPO PEP")

    ' Cabecalho
    Dim hdrs As Variant
    hdrs = Array("PEP3NIVEL", "PEP4NIVEL", "TIPO", "COD MAT", "DESC MAT", "UND", _
                 "QTD", "VALOR", "PU (VALOR/QTD)", "MIN PU", "MAX PU", "STATUS", "OBSERVACAO", "TIPO OD")
    For c = 1 To 14
        With ws.Cells(1, c)
            .Value = hdrs(c - 1)
            .Font.Name = "Segoe UI Semibold" : .Font.Bold = True : .Font.Size = 10
            .Font.Color = RGB(235, 240, 248) : .Interior.Color = RGB(17, 24, 39)
            .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous : .Borders.Color = RGB(60, 70, 90)
        End With
    Next c
    ws.Rows(1).RowHeight = 28

    Dim refCol As Long : refCol = iCod : If refCol = 0 Then refCol = 1
    Dim lastR As Long : lastR = wsBase.Cells(wsBase.Rows.Count, refCol).End(xlUp).Row
    Dim outRow As Long : outRow = 2
    Dim nOut As Long : nOut = 0
    Dim stArr() As String
    If lastR >= 2 Then
        Dim d As Variant : d = wsBase.Range(wsBase.Cells(2, 1), wsBase.Cells(lastR, lastC)).Value
        ' Saida bufferizada: monta as linhas em array e escreve UMA vez no final
        Dim outA() As Variant : ReDim outA(1 To UBound(d, 1), 1 To 14)
        ReDim stArr(1 To UBound(d, 1))
        Dim i As Long
        For i = 1 To UBound(d, 1)
            Dim cod As String : cod = "" : If iCod > 0 Then cod = NormCod(d(i, iCod))
            Dim qtd As Double : qtd = 0 : If iQtd > 0 Then qtd = Val0(d(i, iQtd))
            Dim vlrPU As Double : vlrPU = 0 : If iVal > 0 Then vlrPU = Val0(d(i, iVal))
            If cod = "" Or qtd = 0 Or vlrPU = 0 Then GoTo Prox
            Dim pu As Double : pu = vlrPU / qtd
            Dim mn As Double : mn = 0 : Dim mx As Double : mx = 0
            Dim temFx As Boolean : temFx = fx.Exists(cod)
            Dim st As String, obs As String
            If temFx Then
                Dim p() As String : p = Split(CStr(fx(cod)), "|")
                mn = Val0(p(0)) : mx = Val0(p(1))
                If pu < mn Then
                    st = "ABAIXO DO MINIMO"
                    obs = "PU " & Format(pu, "#,##0.00") & " < min " & Format(mn, "#,##0.00")
                ElseIf pu > mx Then
                    st = "ACIMA DO MAXIMO"
                    obs = "PU " & Format(pu, "#,##0.00") & " > max " & Format(mx, "#,##0.00")
                Else
                    st = "DENTRO"
                    obs = "Dentro da faixa (" & Format(mn, "#,##0.00") & " a " & Format(mx, "#,##0.00") & ")"
                End If
            Else
                st = "SEM REFERENCIA"
                obs = "Material sem faixa na BASE PRECOS"
            End If

            nOut = nOut + 1
            outA(nOut, 1) = IIf(iPep3 > 0, d(i, iPep3), "")
            outA(nOut, 2) = IIf(iPep4 > 0, d(i, iPep4), "")
            outA(nOut, 3) = IIf(iTipo > 0, d(i, iTipo), "")
            outA(nOut, 4) = cod
            outA(nOut, 5) = IIf(iDesc > 0, d(i, iDesc), "")
            outA(nOut, 6) = IIf(iUml > 0, d(i, iUml), "")
            outA(nOut, 7) = qtd
            outA(nOut, 8) = vlrPU
            outA(nOut, 9) = pu
            outA(nOut, 10) = IIf(temFx, mn, "")
            outA(nOut, 11) = IIf(temFx, mx, "")
            ' CORRECAO: STATUS vai na coluna 12 (antes era escrito na 11 por cima
            ' do MAX PU, deixando a coluna STATUS vazia - o que quebrava as cores
            ' da aba, a importacao de alertas de PU, o ranking e o painel)
            outA(nOut, 12) = st
            outA(nOut, 13) = obs
            stArr(nOut) = st

            ' TIPO OD: classifica o PEP pelo sufixo do PEP4 (.I/.M/.S/.D) ou pela coluna TIPO PEP
            Dim pep4OD As String : pep4OD = "" : If iPep4 > 0 Then pep4OD = UCase(Trim(CStr(d(i, iPep4))))
            Dim tpRaw As String : tpRaw = "" : If iTipoPep > 0 Then tpRaw = UCase(Trim(CStr(d(i, iTipoPep))))
            Dim tipoOD As String : tipoOD = ClassificarOD(pep4OD, tpRaw)
            If tipoOD = "" Then tipoOD = "-"
            outA(nOut, 14) = tipoOD
Prox:
        Next i
        If nOut > 0 Then ws.Range(ws.Cells(2, 1), ws.Cells(1 + nOut, 14)).Value = outA
    End If
    outRow = 2 + nOut

    ' Formatacao
    Dim wds As Variant : wds = Array(24, 26, 8, 13, 46, 6, 10, 12, 14, 11, 11, 18, 40, 9)
    For c = 1 To 14 : ws.Columns(c).ColumnWidth = wds(c - 1) : Next c

    ' Formatacao das linhas em LOTE: base numa chamada so; cores por BLOCOS de
    ' linhas consecutivas com o mesmo status, lendo do array (sem reler celulas)
    If nOut > 0 Then
        With ws.Range(ws.Cells(2, 1), ws.Cells(1 + nOut, 14))
            .Font.Name = "Segoe UI" : .Font.Size = 9 : .Font.Color = RGB(45, 52, 64)
            .VerticalAlignment = xlCenter
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Color = RGB(255, 255, 255)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(255, 255, 255)
        End With
        ws.Range(ws.Cells(2, 7), ws.Cells(1 + nOut, 11)).NumberFormat = "#,##0.00"
        With ws.Range(ws.Cells(2, 14), ws.Cells(1 + nOut, 14))
            .Font.Bold = True : .HorizontalAlignment = xlCenter : .Font.Color = RGB(31, 41, 59)
        End With

        Dim rr As Long, rIni As Long, stB As String, bg As Long, chip As Long
        rr = 1
        Do While rr <= nOut
            stB = stArr(rr) : rIni = rr
            Do While rr < nOut
                If stArr(rr + 1) <> stB Then Exit Do
                rr = rr + 1
            Loop
            Select Case stB
                Case "DENTRO":           bg = RGB(231, 244, 234) : chip = RGB(33, 130, 70)
                Case "ABAIXO DO MINIMO": bg = RGB(255, 244, 214) : chip = RGB(176, 124, 0)
                Case "ACIMA DO MAXIMO":  bg = RGB(252, 226, 228) : chip = RGB(192, 0, 0)
                Case Else:               bg = RGB(238, 238, 238) : chip = RGB(110, 110, 110)
            End Select
            ws.Range(ws.Cells(1 + rIni, 1), ws.Cells(1 + rr, 14)).Interior.Color = bg
            With ws.Range(ws.Cells(1 + rIni, 12), ws.Cells(1 + rr, 12))
                .Interior.Color = chip : .Font.Color = RGB(255, 255, 255) : .Font.Bold = True
                .HorizontalAlignment = xlCenter
            End With
            rr = rr + 1
        Loop
    End If

    If outRow > 2 Then ws.Range("A1:N1").AutoFilter
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
    ws.Range("A1").Select
End Sub


' ===========================================================================
'  RANKING DE RISCO POR OBRA (PEP3)
'  Consolida TODOS os sinais de auditoria ja gerados (reprovacao SAP x PRJ,
'  alertas criticos, divergencias de preco unitario e racionalizacao NT.006)
'  num SCORE 0-100 por obra, ordenado do maior risco para o menor.
'  Responde a pergunta do auditor: "qual obra eu olho primeiro?"
' ===========================================================================
Private Sub ProcessarRankingRisco(wb As Workbook, wsDet As Worksheet, _
        wsCom As Worksheet, wsPU As Worksheet, wsAlertaC As Worksheet)

    ' Pesos do score (ajustaveis na aba CONFIG; soma maxima = 100)
    Dim PESO_REPROV As Double : PESO_REPROV = CfgD("PESO_REPROV", 40)  ' obra reprovada
    Dim PESO_ALERTA As Double : PESO_ALERTA = CfgD("PESO_ALERTA", 4)   ' por alerta critico
    Dim CAP_ALERTA As Double : CAP_ALERTA = CfgD("CAP_ALERTA", 24)     ' teto alertas
    Dim PESO_PU As Double : PESO_PU = CfgD("PESO_PU", 3)               ' por diverg. de PU
    Dim CAP_PU As Double : CAP_PU = CfgD("CAP_PU", 18)                 ' teto PU
    Dim PESO_COM As Double : PESO_COM = CfgD("PESO_COM", 2)            ' por COM fora NT.006
    Dim CAP_COM As Double : CAP_COM = CfgD("CAP_COM", 18)             ' teto COM
    Dim PESO_DISCREP As Double : PESO_DISCREP = CfgD("PESO_DISCREP", 2) ' por discrepancia de valor (ODI)
    Dim CAP_DISCREP As Double : CAP_DISCREP = CfgD("CAP_DISCREP", 12)   ' teto discrepancia de valor

    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wsAlertaC)
    ws.Name = "RANKING DE RISCO"

    ' ---------------- coleta por PEP3 ----------------
    Dim valOb As Object : Set valOb = CreateObject("Scripting.Dictionary")  ' valor total
    Dim repOb As Object : Set repOb = CreateObject("Scripting.Dictionary")  ' reprovado?
    Dim alOb  As Object : Set alOb  = CreateObject("Scripting.Dictionary")  ' n alertas
    Dim puOb  As Object : Set puOb  = CreateObject("Scripting.Dictionary")  ' n diverg PU
    Dim sobOb As Object : Set sobOb = CreateObject("Scripting.Dictionary")  ' sobrepreco R$
    Dim comOb As Object : Set comOb = CreateObject("Scripting.Dictionary")  ' n COM fora

    Dim r As Long, p3 As String

    ' PEPs sem material baixado no projeto: item "baixado" = MAT PRJ CAD diferente
    ' de zero; PEP sem baixa = TODOS os seus itens com MAT PRJ CAD zero ou vazio
    Dim p3Itn As Object : Set p3Itn = CreateObject("Scripting.Dictionary")  ' n itens
    Dim p3Bx  As Object : Set p3Bx  = CreateObject("Scripting.Dictionary")  ' n itens baixados
    Dim p3Lib As Object : Set p3Lib = CreateObject("Scripting.Dictionary")  ' soma |MAT LIB SAP|
    Dim p4Itn As Object : Set p4Itn = CreateObject("Scripting.Dictionary")
    Dim p4Bx  As Object : Set p4Bx  = CreateObject("Scripting.Dictionary")
    Dim p4Lib As Object : Set p4Lib = CreateObject("Scripting.Dictionary")
    Dim p4Val As Object : Set p4Val = CreateObject("Scripting.Dictionary")  ' soma |VALOR|
    Dim p4P3  As Object : Set p4P3  = CreateObject("Scripting.Dictionary")  ' PEP4 -> PEP3

    ' Discrepancia de valor de material (ODI): PU de cada linha x media do
    ' codigo. Vira secao propria no fim da aba e entra no SCORE e no DIAGNOSTICO.
    Dim dvTol As Double : dvTol = CfgD("DISCREP_VALOR_PCT", 0.3)
    Dim dvItens As Collection : Set dvItens = DiscrepValorODI(AcharBaseInventario(wb), dvTol)
    Dim dvOb As Object : Set dvOb = CreateObject("Scripting.Dictionary")      ' PEP3 -> n discrepancias
    Dim dvIt As Variant, dvP3 As String
    For Each dvIt In dvItens
        dvP3 = CStr(dvIt(0))
        If dvP3 <> "" Then
            If dvOb.Exists(dvP3) Then dvOb(dvP3) = dvOb(dvP3) + 1 Else dvOb.Add dvP3, 1
        End If
    Next dvIt

    ' 1) ANALISE SAP x PRJ (dados a partir da linha 10): valor + aprovacao
    '    (col 2 = PEP4, 6 = VALOR, 9 = MAT LIB SAP, 10 = MAT PRJ CAD, 14 = APROVACAO)
    Dim lastD As Long : lastD = wsDet.Cells(wsDet.Rows.Count, 1).End(xlUp).Row
    If lastD >= 10 Then
        Dim dd As Variant : dd = wsDet.Range(wsDet.Cells(10, 1), wsDet.Cells(lastD, 15)).Value
        For r = 1 To UBound(dd, 1)
            p3 = Trim(CStr(dd(r, 1)))
            If p3 <> "" Then
                If Not valOb.Exists(p3) Then valOb.Add p3, 0#
                valOb(p3) = valOb(p3) + Abs(Val0(dd(r, 6)))
                If Not repOb.Exists(p3) Then repOb.Add p3, False
                If InStr(UCase(CStr(dd(r, 14))), "REPROVADO") > 0 Then repOb(p3) = True

                Dim bxR As Boolean : bxR = (Val0(dd(r, 10)) <> 0)
                Dim lbR As Double : lbR = Abs(Val0(dd(r, 9)))
                p3Itn(p3) = p3Itn(p3) + 1
                p3Lib(p3) = p3Lib(p3) + lbR
                If bxR Then p3Bx(p3) = p3Bx(p3) + 1
                Dim p4R As String : p4R = Trim(CStr(dd(r, 2)))
                If p4R <> "" Then
                    p4Itn(p4R) = p4Itn(p4R) + 1
                    p4Lib(p4R) = p4Lib(p4R) + lbR
                    p4Val(p4R) = p4Val(p4R) + Abs(Val0(dd(r, 6)))
                    If bxR Then p4Bx(p4R) = p4Bx(p4R) + 1
                    If Not p4P3.Exists(p4R) Then p4P3.Add p4R, p3
                End If
            End If
        Next r
    End If

    ' 2) ALERTA CRITICO (col 2 = PEP3) - leitura em array
    Dim r0A As Long : r0A = PrimeiraLinhaAlerta(wsAlertaC)
    Dim lastA As Long : lastA = wsAlertaC.Cells(wsAlertaC.Rows.Count, 1).End(xlUp).Row
    If lastA >= r0A Then
        Dim aa As Variant : aa = wsAlertaC.Range(wsAlertaC.Cells(r0A, 1), wsAlertaC.Cells(lastA, 2)).Value
        For r = 1 To UBound(aa, 1)
            p3 = Trim(CStr(aa(r, 2)))
            If p3 <> "" Then
                If alOb.Exists(p3) Then alOb(p3) = alOb(p3) + 1 Else alOb.Add p3, 1
                If Not valOb.Exists(p3) Then valOb.Add p3, 0#
            End If
        Next r
    End If

    ' 3) PRECO UNITARIO (linha 2+): diverg + sobrepreco potencial
    Dim lastP As Long : lastP = wsPU.Cells(wsPU.Rows.Count, 1).End(xlUp).Row
    If lastP >= 2 Then
        Dim pp As Variant : pp = wsPU.Range(wsPU.Cells(2, 1), wsPU.Cells(lastP, 13)).Value
        For r = 1 To UBound(pp, 1)
            p3 = Trim(CStr(pp(r, 1)))
            Dim stP As String : stP = UCase(Trim(CStr(pp(r, 12))))
            If p3 <> "" And (stP = "ABAIXO DO MINIMO" Or stP = "ACIMA DO MAXIMO") Then
                If puOb.Exists(p3) Then puOb(p3) = puOb(p3) + 1 Else puOb.Add p3, 1
                If Not valOb.Exists(p3) Then valOb.Add p3, 0#
                If stP = "ACIMA DO MAXIMO" Then
                    Dim qq As Double : qq = Val0(pp(r, 7))
                    Dim puu As Double : puu = Val0(pp(r, 9))
                    Dim mxx As Double : mxx = Val0(pp(r, 11))
                    If puu > mxx And qq > 0 Then
                        If Not sobOb.Exists(p3) Then sobOb.Add p3, 0#
                        sobOb(p3) = sobOb(p3) + (puu - mxx) * qq
                    End If
                End If
            End If
        Next r
    End If

    ' 4) RACIONALIZACAO COM (linha 3+, col 2 = PEP3, col 9 = STATUS rotulo) - array
    Dim lastM As Long : lastM = wsCom.Cells(wsCom.Rows.Count, 1).End(xlUp).Row
    If lastM >= 3 Then
        Dim mm As Variant : mm = wsCom.Range(wsCom.Cells(3, 1), wsCom.Cells(lastM, 9)).Value
        For r = 1 To UBound(mm, 1)
            p3 = Trim(CStr(mm(r, 2)))
            Dim stC As String : stC = UCase(Trim(CStr(mm(r, 9))))
            If p3 <> "" Then
                If stC = "FALTANDO" Or stC = "EXCESSO" Or stC = "EXAGERADO" Or stC = "ZERO" Then
                    If comOb.Exists(p3) Then comOb(p3) = comOb(p3) + 1 Else comOb.Add p3, 1
                    If Not valOb.Exists(p3) Then valOb.Add p3, 0#
                End If
            End If
        Next r
    End If

    ' 5) obra so com discrepancia de valor tambem entra no ranking
    Dim kDv As Variant
    For Each kDv In dvOb.Keys
        If Not valOb.Exists(kDv) Then valOb.Add kDv, 0#
    Next kDv

    ' ---------------- score + ordenacao ----------------
    Dim nOb As Long : nOb = valOb.Count
    Dim ks() As String, sc() As Double
    Dim k As Variant, ix As Long

    If nOb > 0 Then
        ReDim ks(0 To nOb - 1) : ReDim sc(0 To nOb - 1)
        ix = 0
        For Each k In valOb.Keys
            Dim s As Double : s = 0
            If repOb.Exists(k) Then If repOb(k) Then s = s + PESO_REPROV
            Dim pAl As Double : pAl = 0
            If alOb.Exists(k) Then pAl = alOb(k) * PESO_ALERTA
            If pAl > CAP_ALERTA Then pAl = CAP_ALERTA
            Dim pPu As Double : pPu = 0
            If puOb.Exists(k) Then pPu = puOb(k) * PESO_PU
            If pPu > CAP_PU Then pPu = CAP_PU
            Dim pCo As Double : pCo = 0
            If comOb.Exists(k) Then pCo = comOb(k) * PESO_COM
            If pCo > CAP_COM Then pCo = CAP_COM
            Dim pDv As Double : pDv = 0
            If dvOb.Exists(k) Then pDv = ParcelaRisco(dvOb(k), PESO_DISCREP, CAP_DISCREP)
            s = s + pAl + pPu + pCo + pDv
            If s > 100 Then s = 100
            ks(ix) = CStr(k) : sc(ix) = s : ix = ix + 1
        Next k

        ' ordena por score desc; empate -> maior valor de obra primeiro (estavel:
        ' mesma ordem do bubble sort de antes, em O(n log n))
        Dim oIdx() As Long, oVal() As Double, ksO() As String, scO() As Double
        ReDim oIdx(0 To nOb - 1) : ReDim oVal(0 To nOb - 1)
        ReDim ksO(0 To nOb - 1) : ReDim scO(0 To nOb - 1)
        For ix = 0 To nOb - 1
            oIdx(ix) = ix : oVal(ix) = valOb(ks(ix))
        Next ix
        OrdenarIdxDesc oIdx, sc, oVal, 0, nOb - 1
        For ix = 0 To nOb - 1
            ksO(ix) = ks(oIdx(ix)) : scO(ix) = sc(oIdx(ix))
        Next ix
        ks = ksO : sc = scO
    End If

    ' ---------------- saida ----------------
    With ws.Range("A1:K2")
        .Merge
        .Value = "RANKING DE RISCO POR OBRA  -  por onde comecar a auditoria"
        .Font.Name = "Segoe UI Semibold" : .Font.Size = 16 : .Font.Bold = True
        .Font.Color = RGB(255, 255, 255) : .Interior.Color = RGB(17, 24, 39)
        .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 20 : ws.Rows(2).RowHeight = 20

    Dim hdrs As Variant
    hdrs = Array("#", "PEP3NIVEL", "VALOR OBRA", "SITUACAO", "ALERTAS", _
                 "DIVERG PRECO", "SOBREPRECO R$", "COM FORA NT.006", "SCORE", "RISCO", "DIAGNOSTICO")
    Dim c As Long
    For c = 1 To 11
        With ws.Cells(3, c)
            .Value = hdrs(c - 1)
            .Font.Name = "Segoe UI Semibold" : .Font.Bold = True : .Font.Size = 10
            .Font.Color = RGB(235, 240, 248) : .Interior.Color = RGB(31, 41, 59)
            .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous : .Borders.Color = RGB(60, 70, 90)
        End With
    Next c
    ws.Rows(3).RowHeight = 26

    Dim outRow As Long : outRow = 4
    If nOb = 0 Then
        ws.Cells(4, 1).Value = "Nenhuma obra encontrada na base."
        ws.Cells(4, 1).Font.Italic = True
    Else
        ' Saida bufferizada: monta as linhas em array e escreve UMA vez
        Dim outA() As Variant : ReDim outA(1 To nOb, 1 To 11)
        Dim rkArr() As String : ReDim rkArr(1 To nOb)
        For ix = 0 To nOb - 1
            p3 = ks(ix)
            Dim nAl As Long : nAl = 0 : If alOb.Exists(p3) Then nAl = alOb(p3)
            Dim nPu As Long : nPu = 0 : If puOb.Exists(p3) Then nPu = puOb(p3)
            Dim nCo As Long : nCo = 0 : If comOb.Exists(p3) Then nCo = comOb(p3)
            Dim vSo As Double : vSo = 0 : If sobOb.Exists(p3) Then vSo = sobOb(p3)
            Dim ehRep As Boolean : ehRep = False
            If repOb.Exists(p3) Then ehRep = CBool(repOb(p3))

            Dim risco As String
            If sc(ix) >= 60 Then
                risco = "ALTO"
            ElseIf sc(ix) >= 30 Then
                risco = "MEDIO"
            ElseIf sc(ix) > 0 Then
                risco = "BAIXO"
            Else
                risco = "OK"
            End If

            ' diagnostico em linguagem clara
            Dim diag As String : diag = ""
            If ehRep Then diag = "REPROVADO na analise SAP x PRJ"
            If nAl > 0 Then diag = diag & IIf(diag <> "", " | ", "") & nAl & " alerta(s) critico(s)"
            If nPu > 0 Then
                diag = diag & IIf(diag <> "", " | ", "") & nPu & " diverg. de preco"
                If vSo > 0 Then diag = diag & " (sobrepreco " & FmtKPI(vSo) & ")"
            End If
            If nCo > 0 Then diag = diag & IIf(diag <> "", " | ", "") & nCo & " COM fora do previsto"
            If dvOb.Exists(p3) Then
                diag = diag & IIf(diag <> "", " | ", "") & dvOb(p3) & " discrepancia(s) de valor de material"
            End If
            If p3Itn.Exists(p3) And Not p3Bx.Exists(p3) Then
                diag = diag & IIf(diag <> "", " | ", "") & "nenhum material baixado no PRJ"
            End If
            If diag = "" Then diag = "Sem apontamentos"

            outA(ix + 1, 1) = ix + 1
            outA(ix + 1, 2) = p3
            outA(ix + 1, 3) = valOb(p3)
            outA(ix + 1, 4) = IIf(ehRep, "REPROVADO", "APROVADO")
            outA(ix + 1, 5) = nAl
            outA(ix + 1, 6) = nPu
            outA(ix + 1, 7) = vSo
            outA(ix + 1, 8) = nCo
            outA(ix + 1, 9) = sc(ix)
            outA(ix + 1, 10) = risco
            outA(ix + 1, 11) = diag
            rkArr(ix + 1) = risco
            outRow = outRow + 1
        Next ix
        ' escreve tudo de uma vez
        ws.Range(ws.Cells(4, 1), ws.Cells(3 + nOb, 11)).Value = outA
    End If

    ' ---------------- formatacao ----------------
    Dim wds As Variant : wds = Array(7, 26, 14, 13, 9, 13, 14, 15, 8, 9, 70)
    For c = 1 To 11 : ws.Columns(c).ColumnWidth = wds(c - 1) : Next c

    ' Formatacao em LOTE: base numa chamada so; cores por BLOCOS de risco
    ' (a lista ja vem ordenada por score, entao os niveis formam blocos contiguos)
    If nOb > 0 Then
        With ws.Range(ws.Cells(4, 1), ws.Cells(3 + nOb, 11))
            .Font.Name = "Segoe UI" : .Font.Size = 9 : .Font.Color = RGB(45, 52, 64)
            .VerticalAlignment = xlCenter
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Color = RGB(255, 255, 255)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(255, 255, 255)
        End With
        ws.Range(ws.Cells(4, 1), ws.Cells(3 + nOb, 1)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(4, 4), ws.Cells(3 + nOb, 6)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(4, 8), ws.Cells(3 + nOb, 10)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(4, 9), ws.Cells(3 + nOb, 9)).Font.Bold = True
        ws.Range(ws.Cells(4, 3), ws.Cells(3 + nOb, 3)).NumberFormat = "#,##0.00"
        ws.Range(ws.Cells(4, 7), ws.Cells(3 + nOb, 7)).NumberFormat = "#,##0.00"

        Dim rr As Long, rIni As Long, rsk As String, bg As Long, chip As Long
        rr = 1
        Do While rr <= nOb
            rsk = rkArr(rr) : rIni = rr
            Do While rr < nOb
                If rkArr(rr + 1) <> rsk Then Exit Do
                rr = rr + 1
            Loop
            Select Case rsk
                Case "ALTO":  bg = RGB(252, 226, 228) : chip = RGB(192, 0, 0)
                Case "MEDIO": bg = RGB(255, 244, 214) : chip = RGB(176, 124, 0)
                Case "BAIXO": bg = RGB(234, 242, 250) : chip = RGB(31, 78, 121)
                Case Else:    bg = RGB(231, 244, 234) : chip = RGB(33, 130, 70)
            End Select
            ws.Range(ws.Cells(3 + rIni, 1), ws.Cells(3 + rr, 11)).Interior.Color = bg
            With ws.Range(ws.Cells(3 + rIni, 10), ws.Cells(3 + rr, 10))
                .Interior.Color = chip : .Font.Color = RGB(255, 255, 255) : .Font.Bold = True
            End With
            ws.Range(ws.Cells(3 + rIni, 9), ws.Cells(3 + rr, 9)).Font.Color = chip
            rr = rr + 1
        Loop

        ' destaque de SITUACAO reprovada (poucas obras - pode ser por linha)
        For rr = 1 To nOb
            If CStr(outA(rr, 4)) = "REPROVADO" Then
                With ws.Cells(3 + rr, 4)
                    .Font.Color = RGB(192, 0, 0) : .Font.Bold = True
                End With
            End If
        Next rr
    End If

    If outRow > 4 Then ws.Range(ws.Cells(3, 1), ws.Cells(3, 11)).AutoFilter

    ' Secao de PEPs sem material baixado no PRJ, 3 linhas abaixo do ranking. As
    ' linhas em branco isolam os blocos: autofiltro, contador do titulo e o nome
    ' dadosRanking param no fim do ranking (UltimaLinhaContigua).
    Dim rSec As Long
    If nOb > 0 Then rSec = 3 + nOb + 3 Else rSec = 7
    SecaoSemBaixaPRJ ws, rSec, p3Itn, p3Bx, p3Lib, valOb, repOb, p4Itn, p4Bx, p4Lib, p4Val, p4P3

    ' Secao de discrepancia de valor de material, 3 linhas abaixo da anterior
    Dim uSec As Range : Set uSec = ws.Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious)
    Dim rDv As Long : rDv = rSec + 6
    If Not uSec Is Nothing Then rDv = uSec.Row + 4
    SecaoDiscrepValor ws, rDv, dvItens, dvTol

    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range("A4").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
    ws.Range("A1").Select
End Sub


' ===========================================================================
'  DESIGN GLOBAL - identidade visual unificada, aplicada APOS a geracao
'  (por isso nao interfere nas leituras entre abas, que usam o layout bruto):
'    1. Cor da guia (Tab) por aba - paleta propria de cada relatorio
'    2. Barra de navegacao clicavel no topo de TODAS as abas
'    3. Contadores dinamicos nos titulos (apontamentos, obras, divergencias)
'    4. PRECO UNITARIO ganha banda de titulo (nao tinha)
'    5. Drill-down: cards do PAINEL viram atalhos para a aba de origem
' ===========================================================================
' Secao "PEPs sem material baixado no projeto" do RANKING DE RISCO.
'   PEP3: todos os itens do PEP3 com MAT PRJ CAD zero ou vazio.
'   PEP4: idem no PEP4 - listado so quando o PEP3 dele TEM baixa em outro PEP4
'         (se o PEP3 inteiro esta sem baixa, a linha do PEP3 ja o cobre).
' Cada grupo vem ordenado por valor decrescente.
Private Sub SecaoSemBaixaPRJ(ws As Worksheet, ByVal r0 As Long, _
        p3Itn As Object, p3Bx As Object, p3Lib As Object, p3Val As Object, p3Rep As Object, _
        p4Itn As Object, p4Bx As Object, p4Lib As Object, p4Val As Object, p4P3 As Object)

    ' quantos PEP4 cada PEP3 tem (para a observacao da linha do PEP3)
    Dim p3NP4 As Object : Set p3NP4 = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In p4P3.Keys
        p3NP4(p4P3(k)) = p3NP4(p4P3(k)) + 1
    Next k

    ' lista: PEP3 sem baixa primeiro (1..n3), depois PEP4 sem baixa (n3+1..n)
    Dim pp() As String, vl() As Double, n As Long, n3 As Long
    Dim nMax As Long : nMax = p3Itn.Count + p4Itn.Count
    If nMax > 0 Then ReDim pp(1 To nMax) : ReDim vl(1 To nMax)
    For Each k In p3Itn.Keys
        If Not p3Bx.Exists(k) Then
            n = n + 1 : pp(n) = CStr(k)
            If p3Val.Exists(k) Then vl(n) = p3Val(k)
        End If
    Next k
    n3 = n
    For Each k In p4Itn.Keys
        If Not p4Bx.Exists(k) Then
            If p3Bx.Exists(p4P3(k)) Then
                n = n + 1 : pp(n) = CStr(k) : vl(n) = p4Val(k)
            End If
        End If
    Next k
    OrdenarPorValor pp, vl, 1, n3
    OrdenarPorValor pp, vl, n3 + 1, n

    ' -- titulo da secao --
    With ws.Range(ws.Cells(r0, 1), ws.Cells(r0, 11))
        .Merge
        .Value = "PEPs SEM MATERIAL BAIXADO NO PROJETO  (MAT PRJ CAD zerado ou vazio em todos os itens)" & _
                 "   -   " & n3 & " PEP3  |  " & (n - n3) & " PEP4"
        .Font.Name = "Segoe UI Semibold" : .Font.Size = 12 : .Font.Bold = True
        .Font.Color = RGB(255, 255, 255) : .Interior.Color = RGB(176, 0, 0)
        .HorizontalAlignment = xlLeft : .VerticalAlignment = xlCenter : .IndentLevel = 1
    End With
    ws.Rows(r0).RowHeight = 24

    ' -- cabecalho (H:J mesclado = PEP3 do PEP4) --
    Dim h As Variant
    h = Array("NIVEL", "PEP", "VALOR", "SITUACAO", "ITENS", "MAT LIB SAP", "TIPO OD", _
              "PEP3 DO PEP4", "", "", "OBSERVACAO")
    Dim c As Long
    For c = 1 To 11
        With ws.Cells(r0 + 1, c)
            .Value = h(c - 1)
            .Font.Name = "Segoe UI Semibold" : .Font.Bold = True : .Font.Size = 10
            .Font.Color = RGB(235, 240, 248) : .Interior.Color = RGB(31, 41, 59)
            .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous : .Borders.Color = RGB(60, 70, 90)
        End With
    Next c
    ws.Range(ws.Cells(r0 + 1, 8), ws.Cells(r0 + 1, 10)).Merge
    ws.Rows(r0 + 1).RowHeight = 26

    If n = 0 Then
        With ws.Cells(r0 + 2, 1)
            .Value = "Todos os PEP3 e PEP4 tem ao menos um material baixado em MAT PRJ CAD."
            .Font.Italic = True : .Font.Color = RGB(0, 128, 0)
        End With
        Exit Sub
    End If

    ' -- linhas (escrita unica em array) --
    Dim outS() As Variant : ReDim outS(1 To n, 1 To 11)
    Dim i As Long, pep As String, p3d As String, libS As Double, obs As String, nP4 As Long
    For i = 1 To n
        pep = pp(i)
        If i <= n3 Then
            p3d = pep
            libS = p3Lib(pep)
            nP4 = 0 : If p3NP4.Exists(pep) Then nP4 = p3NP4(pep)
            outS(i, 1) = "PEP3"
            outS(i, 5) = p3Itn(pep)
            obs = "Nenhum dos " & p3Itn(pep) & " itens (em " & nP4 & _
                  " PEP4) tem MAT PRJ CAD baixado"
        Else
            p3d = CStr(p4P3(pep))
            libS = p4Lib(pep)
            outS(i, 1) = "PEP4"
            outS(i, 5) = p4Itn(pep)
            outS(i, 7) = ClassificarOD(pep, "")
            outS(i, 8) = p3d
            obs = "O PEP3 tem baixa em outros PEP4; este PEP4 nao tem nenhuma (" & _
                  p4Itn(pep) & " itens)"
        End If
        ' SAP com material liberado e projeto sem nada baixado = o sinal mais forte
        ' (Format "#,##0.##" deixaria virgula sobrando em inteiro: "1.234,")
        If libS > 0 Then obs = obs & " | SAP liberou " & _
            IIf(libS = Int(libS), Format(libS, "#,##0"), Format(libS, "#,##0.00")) & _
            " - material sem projeto"
        outS(i, 2) = pep
        outS(i, 3) = vl(i)
        If p3Rep.Exists(p3d) Then outS(i, 4) = IIf(p3Rep(p3d), "REPROVADO", "APROVADO")
        outS(i, 6) = libS
        outS(i, 11) = obs
    Next i
    Dim rI As Long : rI = r0 + 2
    Dim rF As Long : rF = r0 + 1 + n
    ws.Range(ws.Cells(rI, 1), ws.Cells(rF, 11)).Value = outS

    ' -- formatacao: PEP3 em vermelho, PEP4 em ambar --
    With ws.Range(ws.Cells(rI, 1), ws.Cells(rF, 11))
        .Font.Name = "Segoe UI" : .Font.Size = 9 : .Font.Color = RGB(45, 52, 64)
        .VerticalAlignment = xlCenter
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Color = RGB(255, 255, 255)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(255, 255, 255)
    End With
    ws.Range(ws.Cells(rI, 8), ws.Cells(rF, 10)).Merge True     ' H:J por linha
    ws.Range(ws.Cells(rI, 3), ws.Cells(rF, 3)).NumberFormat = "#,##0.00"
    ws.Range(ws.Cells(rI, 6), ws.Cells(rF, 6)).NumberFormat = "#,##0.00"
    ws.Range(ws.Cells(rI, 4), ws.Cells(rF, 5)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(rI, 7), ws.Cells(rF, 7)).HorizontalAlignment = xlCenter
    If n3 > 0 Then
        ws.Range(ws.Cells(rI, 1), ws.Cells(rI + n3 - 1, 11)).Interior.Color = RGB(252, 226, 228)
        With ws.Range(ws.Cells(rI, 1), ws.Cells(rI + n3 - 1, 1))
            .Interior.Color = RGB(192, 0, 0) : .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True : .HorizontalAlignment = xlCenter
        End With
    End If
    If n > n3 Then
        ws.Range(ws.Cells(rI + n3, 1), ws.Cells(rF, 11)).Interior.Color = RGB(255, 244, 214)
        With ws.Range(ws.Cells(rI + n3, 1), ws.Cells(rF, 1))
            .Interior.Color = RGB(176, 124, 0) : .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True : .HorizontalAlignment = xlCenter
        End With
    End If
    For i = 1 To n
        If CStr(outS(i, 4)) = "REPROVADO" Then
            With ws.Cells(rI + i - 1, 4)
                .Font.Color = RGB(192, 0, 0) : .Font.Bold = True
            End With
        End If
    Next i
End Sub

' Ordena pp/vl por vl decrescente no intervalo [lo, hi] (selection sort:
' algumas centenas de PEPs no maximo).
Private Sub OrdenarPorValor(pp() As String, vl() As Double, ByVal lo As Long, ByVal hi As Long)
    Dim a As Long, b As Long, m As Long, tS As String, tV As Double
    For a = lo To hi - 1
        m = a
        For b = a + 1 To hi
            If vl(b) > vl(m) Then m = b
        Next b
        If m <> a Then
            tS = pp(a) : pp(a) = pp(m) : pp(m) = tS
            tV = vl(a) : vl(a) = vl(m) : vl(m) = tV
        End If
    Next a
End Sub

' ===========================================================================
'  DISCREPANCIA DE VALOR DE MATERIAL (ODI)
'  Preco unitario (VALOR / MAT LIB SAP) de cada linha de PEP ODI comparado com
'  a MEDIA do mesmo codigo de material em todas as linhas ODI da base. Fora de
'  +/- DISCREP_VALOR_PCT (padrao 30%) = discrepancia. Codigo com um unico
'  lancamento ODI nao tem com o que comparar e fica de fora.
' ===========================================================================

' Parcela de um tipo de apontamento no SCORE do ranking: n x peso, limitada ao teto.
Private Function ParcelaRisco(ByVal n As Double, ByVal peso As Double, ByVal teto As Double) As Double
    ParcelaRisco = n * peso
    If ParcelaRisco > teto Then ParcelaRisco = teto
End Function

' Desvio do PU em relacao a media, em fracao (+0,45 = 45% acima).
Private Function DesvioMedia(ByVal pu As Double, ByVal media As Double) As Double
    If media = 0 Then Exit Function
    DesvioMedia = (pu - media) / Abs(media)
End Function

' True quando o PU passa de tol (fracao) acima ou abaixo da media. O limite
' exato (30,0%) nao e discrepancia; folga de 1e-9 contra arredondamento.
Private Function EhDiscrepValor(ByVal pu As Double, ByVal media As Double, _
                                ByVal tol As Double) As Boolean
    If media = 0 Then Exit Function
    EhDiscrepValor = (Abs(DesvioMedia(pu, media)) > tol + 0.000000001)
End Function

' Linhas ODI com PU discrepante, da maior para a menor diferenca. Cada item:
' Array(PEP3, PEP4, COD, DESC, QTD, VALOR, PU, MEDIA, DESVIO, N LANCAMENTOS).
Private Function DiscrepValorODI(ByVal wsBase As Worksheet, ByVal tol As Double) As Collection
    Dim out As New Collection
    Set DiscrepValorODI = out
    If wsBase Is Nothing Then Exit Function

    Dim mapa As Object : Set mapa = MapaColunas(wsBase)
    Dim iP3 As Long : iP3 = ColIdx(mapa, "PEP3NIVEL")
    Dim iP4 As Long : iP4 = ColIdx(mapa, "PEP4NIVEL")
    Dim iCod As Long : iCod = ColIdx(mapa, "COD MAT")
    Dim iDsc As Long : iDsc = ColIdx(mapa, "DESC MAT")
    Dim iVal As Long : iVal = ColIdx(mapa, "VALOR")
    Dim iQtd As Long : iQtd = ColIdx(mapa, "MAT LIB SAP")
    Dim iTpp As Long : iTpp = ColIdx(mapa, "TIPO PEP")
    If iP4 = 0 Or iCod = 0 Or iVal = 0 Or iQtd = 0 Then Exit Function

    Dim lastC As Long : lastC = wsBase.Cells(1, wsBase.Columns.Count).End(xlToLeft).Column
    Dim lastR As Long : lastR = wsBase.Cells(wsBase.Rows.Count, iP4).End(xlUp).Row
    If lastR < 2 Then Exit Function
    Dim d As Variant : d = wsBase.Range(wsBase.Cells(2, 1), wsBase.Cells(lastR, lastC)).Value
    Dim n As Long : n = UBound(d, 1)

    ' 1a passada: PU de cada linha ODI valida e soma/contagem por codigo
    Dim puL() As Double, okL() As Boolean, codL() As String
    ReDim puL(1 To n) : ReDim okL(1 To n) : ReDim codL(1 To n)
    Dim soma As Object : Set soma = CreateObject("Scripting.Dictionary")
    Dim qtdC As Object : Set qtdC = CreateObject("Scripting.Dictionary")
    Dim i As Long, p4 As String, tpp As String, cod As String, q As Double
    For i = 1 To n
        p4 = "" : tpp = "" : cod = ""
        If Not IsError(d(i, iP4)) Then p4 = Trim(CStr(d(i, iP4)))
        If iTpp > 0 Then
            If Not IsError(d(i, iTpp)) Then tpp = CStr(d(i, iTpp))
        End If
        If Not IsError(d(i, iCod)) Then cod = NormCod(d(i, iCod))
        q = Val0(d(i, iQtd))
        If p4 <> "" And cod <> "" And q <> 0 Then
            If ClassificarOD(p4, tpp) = "ODI" Then
                puL(i) = Val0(d(i, iVal)) / q : okL(i) = True : codL(i) = cod
                If soma.Exists(cod) Then
                    soma(cod) = soma(cod) + puL(i) : qtdC(cod) = qtdC(cod) + 1
                Else
                    soma.Add cod, puL(i) : qtdC.Add cod, 1
                End If
            End If
        End If
    Next i

    ' 2a passada: compara cada linha com a media do codigo
    Dim itens() As Variant, dv() As Double, k As Long
    ReDim itens(1 To n) : ReDim dv(1 To n)
    Dim media As Double, nC As Long, sP3 As String, sDsc As String
    For i = 1 To n
        If okL(i) Then
            nC = qtdC(codL(i))
            media = soma(codL(i)) / nC
            If nC >= 2 And EhDiscrepValor(puL(i), media, tol) Then
                k = k + 1
                dv(k) = DesvioMedia(puL(i), media)
                ' sem IIf: ele avalia os dois ramos e leria a coluna 0 quando a coluna nao existe
                sP3 = "" : sDsc = ""
                If iP3 > 0 Then If Not IsError(d(i, iP3)) Then sP3 = Trim(CStr(d(i, iP3)))
                If iDsc > 0 Then If Not IsError(d(i, iDsc)) Then sDsc = Trim(CStr(d(i, iDsc)))
                itens(k) = Array(sP3, Trim(CStr(d(i, iP4))), codL(i), sDsc, Val0(d(i, iQtd)), _
                                 Val0(d(i, iVal)), puL(i), media, dv(k), nC)
            End If
        End If
    Next i

    ' ordena pela maior diferenca (em modulo); empate fica na ordem da base
    If k = 0 Then Exit Function
    Dim oIdx() As Long, aDv() As Double, zero() As Double, a As Long
    ReDim oIdx(1 To k) : ReDim aDv(1 To k) : ReDim zero(1 To k)
    For a = 1 To k
        oIdx(a) = a : aDv(a) = Abs(dv(a))
    Next a
    OrdenarIdxDesc oIdx, aDv, zero, 1, k
    For a = 1 To k
        out.Add itens(oIdx(a))
    Next a
End Function

' Ordena os indices idx(lo..hi) por k1 decrescente e, no empate, k2
' decrescente (k1/k2 indexados pelo valor do indice). Merge sort ESTAVEL:
' empate total mantem a ordem de entrada. O(n log n) no lugar dos laces
' O(n^2) que travavam com milhares de itens.
Private Sub OrdenarIdxDesc(idx() As Long, k1() As Double, k2() As Double, _
                           ByVal lo As Long, ByVal hi As Long)
    If hi <= lo Then Exit Sub
    Dim tmp() As Long : ReDim tmp(lo To hi)
    MergeIdxDesc idx, tmp, k1, k2, lo, hi
End Sub

Private Sub MergeIdxDesc(idx() As Long, tmp() As Long, k1() As Double, k2() As Double, _
                         ByVal lo As Long, ByVal hi As Long)
    If hi <= lo Then Exit Sub
    Dim md As Long : md = (lo + hi) \ 2
    MergeIdxDesc idx, tmp, k1, k2, lo, md
    MergeIdxDesc idx, tmp, k1, k2, md + 1, hi
    Dim i As Long, j As Long, t As Long, xi As Long, xj As Long, daDireita As Boolean
    i = lo : j = md + 1 : t = lo
    Do While i <= md And j <= hi
        xi = idx(i) : xj = idx(j)
        ' da direita so quando ESTRITAMENTE maior (estabilidade)
        daDireita = False
        If k1(xj) > k1(xi) Then
            daDireita = True
        ElseIf k1(xj) = k1(xi) Then
            If k2(xj) > k2(xi) Then daDireita = True
        End If
        If daDireita Then
            tmp(t) = xj : j = j + 1
        Else
            tmp(t) = xi : i = i + 1
        End If
        t = t + 1
    Loop
    Do While i <= md
        tmp(t) = idx(i) : i = i + 1 : t = t + 1
    Loop
    Do While j <= hi
        tmp(t) = idx(j) : j = j + 1 : t = t + 1
    Loop
    For t = lo To hi
        idx(t) = tmp(t)
    Next t
End Sub

' Secao "DISCREPANCIA DE VALOR DE MATERIAL (ODI)" do RANKING DE RISCO, a partir
' da linha r0: titulo, cabecalho e uma linha por item (acima da media em
' vermelho, abaixo em ambar).
Private Sub SecaoDiscrepValor(ws As Worksheet, ByVal r0 As Long, itens As Collection, _
                              ByVal tol As Double)
    Dim n As Long : n = itens.Count
    With ws.Range(ws.Cells(r0, 1), ws.Cells(r0, 11))
        .Merge
        .Value = "DISCREPANCIA DE VALOR DE MATERIAL (ODI)  -  PU fora de +/-" & Format(tol, "0%") & _
                 " da media do codigo   -   " & n & " item(ns)"
        .Font.Name = "Segoe UI Semibold" : .Font.Size = 12 : .Font.Bold = True
        .Font.Color = RGB(255, 255, 255) : .Interior.Color = RGB(190, 24, 93)
        .HorizontalAlignment = xlLeft : .VerticalAlignment = xlCenter : .IndentLevel = 1
    End With
    ws.Rows(r0).RowHeight = 24

    Dim h As Variant
    h = Array("PEP3NIVEL", "PEP4NIVEL", "COD MAT", "DESC MAT", "QTD", "VALOR", "PU (VALOR/QTD)", _
              "MEDIA PU DO COD", "DESVIO", "SITUACAO", "OBSERVACAO")
    Dim c As Long
    For c = 1 To 11
        With ws.Cells(r0 + 1, c)
            .Value = h(c - 1)
            .Font.Name = "Segoe UI Semibold" : .Font.Bold = True : .Font.Size = 10
            .Font.Color = RGB(235, 240, 248) : .Interior.Color = RGB(31, 41, 59)
            .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous : .Borders.Color = RGB(60, 70, 90)
        End With
    Next c
    ws.Rows(r0 + 1).RowHeight = 26

    If n = 0 Then
        With ws.Cells(r0 + 2, 1)
            .Value = "Nenhum material de ODI com valor fora de +/-" & Format(tol, "0%") & " da media do codigo."
            .Font.Italic = True : .Font.Color = RGB(0, 128, 0)
        End With
        Exit Sub
    End If

    Dim outD() As Variant : ReDim outD(1 To n, 1 To 11)
    Dim i As Long, it As Variant, acima As Boolean
    For i = 1 To n
        it = itens(i)
        acima = (CDbl(it(8)) > 0)
        outD(i, 1) = it(0) : outD(i, 2) = it(1) : outD(i, 3) = it(2) : outD(i, 4) = it(3)
        outD(i, 5) = it(4) : outD(i, 6) = it(5) : outD(i, 7) = it(6) : outD(i, 8) = it(7)
        outD(i, 9) = it(8)
        outD(i, 10) = IIf(acima, "ACIMA DA MEDIA", "ABAIXO DA MEDIA")
        outD(i, 11) = "Discrepancia de valor de material: PU R$ " & Format(it(6), "#,##0.00") & _
                      " esta " & Format(Abs(CDbl(it(8))), "0.0%") & IIf(acima, " acima", " abaixo") & _
                      " da media do codigo (R$ " & Format(it(7), "#,##0.00") & " em " & it(9) & _
                      " lancamentos ODI)"
    Next i
    Dim rI As Long : rI = r0 + 2
    Dim rF As Long : rF = r0 + 1 + n
    ws.Range(ws.Cells(rI, 3), ws.Cells(rF, 3)).NumberFormat = "@"
    ws.Range(ws.Cells(rI, 1), ws.Cells(rF, 11)).Value = outD
    ws.Range(ws.Cells(rI, 5), ws.Cells(rF, 8)).NumberFormat = "#,##0.00"
    ws.Range(ws.Cells(rI, 9), ws.Cells(rF, 9)).NumberFormat = "+0.0%;-0.0%"
    ws.Range(ws.Cells(rI, 9), ws.Cells(rF, 10)).HorizontalAlignment = xlCenter
    With ws.Range(ws.Cells(rI, 1), ws.Cells(rF, 11))
        .Font.Name = "Segoe UI" : .Font.Size = 9 : .Font.Color = RGB(45, 52, 64)
        .VerticalAlignment = xlCenter
    End With
    ' Cor: ambar (abaixo da media) no bloco inteiro e vermelho (acima) por
    ' formatacao condicional - 5 chamadas em vez de 5 por linha (sao milhares
    ' de itens em base grande). Formula sem funcao: independe do idioma do Excel.
    Dim bloco As Range : Set bloco = ws.Range(ws.Cells(rI, 1), ws.Cells(rF, 11))
    bloco.Interior.Color = RGB(255, 244, 214)
    With ws.Range(ws.Cells(rI, 10), ws.Cells(rF, 10)).Font
        .Color = RGB(176, 124, 0) : .Bold = True
    End With
    On Error Resume Next
    With bloco.FormatConditions.Add(xlExpression, , "=$J" & rI & "=""ACIMA DA MEDIA""")
        .Interior.Color = RGB(252, 226, 228)
    End With
    With ws.Range(ws.Cells(rI, 10), ws.Cells(rF, 10)).FormatConditions.Add(xlExpression, , _
                  "=$J" & rI & "=""ACIMA DA MEDIA""")
        .Font.Color = RGB(192, 0, 0)
    End With
    On Error GoTo 0
End Sub

' Ultima linha do bloco CONTIGUO de dados na coluna, a partir de linhaIni.
' Diferente de End(xlUp) a partir do fim da planilha, nao engole blocos
' colocados abaixo (ex.: a secao de PEPs sem baixa no RANKING DE RISCO).
' Devolve linhaIni - 1 quando o bloco esta vazio.
Private Function UltimaLinhaContigua(ws As Worksheet, ByVal linhaIni As Long, _
                                     ByVal col As Long) As Long
    If IsEmpty(ws.Cells(linhaIni, col).Value) Then
        UltimaLinhaContigua = linhaIni - 1
    ElseIf IsEmpty(ws.Cells(linhaIni + 1, col).Value) Then
        UltimaLinhaContigua = linhaIni
    Else
        UltimaLinhaContigua = ws.Cells(linhaIni, col).End(xlDown).Row
    End If
End Function

Private Function CorAcento(nome As String) As Long
    Select Case nome
        Case "PAINEL DO GESTOR":   CorAcento = RGB(17, 24, 39)     ' navy
        Case "ANALISE SAP x PRJ":  CorAcento = RGB(31, 78, 121)    ' azul
        Case "RACIONALIZACAO COM": CorAcento = RGB(94, 84, 158)    ' indigo
        Case "PRECO UNITARIO":     CorAcento = RGB(21, 96, 130)    ' petroleo
        Case "ALERTA CRITICO":     CorAcento = RGB(176, 0, 0)      ' vermelho
        Case "RANKING DE RISCO":   CorAcento = RGB(176, 124, 0)    ' ambar
        Case Else:                 CorAcento = RGB(90, 98, 110)
    End Select
End Function

' Barra de navegacao com BOTOES (Shapes) inserida na LINHA 1 da aba.
' Shapes tem largura uniforme, independente da largura das colunas de dados.
Private Sub InserirBarraNav(ws As Worksheet, nomes As Variant, rots As Variant, _
                            atual As String, congRow As Long)
    ws.Activate
    On Error Resume Next
    ActiveWindow.FreezePanes = False
    On Error GoTo 0
    ws.Rows(1).Insert
    ws.Rows(1).Clear
    ws.Rows(1).RowHeight = 32
    ws.Rows(1).Interior.Color = RGB(246, 247, 249)
    Dim j As Long, x As Double
    x = 6
    For j = 0 To UBound(nomes)
        BotaoNav ws, x, CStr(nomes(j)), CStr(rots(j)), (CStr(nomes(j)) = atual)
        x = x + 116
    Next j
    ' refaz o congelamento, que desceu 1 linha junto com o conteudo
    If congRow > 0 Then
        On Error Resume Next
        ws.Cells(congRow, 1).Select
        ActiveWindow.FreezePanes = True
        On Error GoTo 0
    End If
    ws.Range("A1").Select
End Sub

' Um botao da barra: retangulo arredondado clicavel na cor da aba destino.
' O botao da propria aba aparece cinza com a marca de "voce esta aqui".
Private Sub BotaoNav(ws As Worksheet, x As Double, destino As String, _
                     rotulo As String, ehAtual As Boolean)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, 4, 110, 24)
    With shp
        .Line.Visible = msoFalse
        If ehAtual Then
            .Fill.ForeColor.RGB = RGB(225, 228, 233)
        Else
            .Fill.ForeColor.RGB = CorAcento(destino)
        End If
        With .TextFrame2
            .WordWrap = msoFalse
            .MarginLeft = 0 : .MarginRight = 0 : .MarginTop = 0 : .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.Text = IIf(ehAtual, ChrW(9679) & " " & rotulo, rotulo)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            With .TextRange.Font
                .Size = 9 : .Bold = msoTrue : .Name = "Segoe UI"
                .Fill.ForeColor.RGB = IIf(ehAtual, RGB(80, 88, 100), RGB(255, 255, 255))
            End With
        End With
        .Placement = xlFreeFloating
    End With
    If Not ehAtual Then
        ws.Hyperlinks.Add Anchor:=shp, Address:="", _
            SubAddress:="'" & destino & "'!A1", ScreenTip:="Abrir " & destino
    End If
End Sub

' Acrescenta " - N <sufixo>" ao titulo (A1) de uma aba, ignorando a linha
' informativa "Nenhum..." das abas vazias.
Private Sub AppendContagem(wb As Workbook, aba As String, _
                           primeiraLinhaDados As Long, sufixo As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(aba)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    ' bloco contiguo: nao conta secoes colocadas abaixo dos dados
    Dim lastR As Long : lastR = UltimaLinhaContigua(ws, primeiraLinhaDados, 1)
    Dim n As Long : n = lastR - primeiraLinhaDados + 1
    If n < 0 Then n = 0
    If n = 1 Then
        If InStr(UCase(CStr(ws.Cells(primeiraLinhaDados, 1).Value)), "NENHUM") > 0 Then n = 0
    End If
    ws.Range("A1").Value = ws.Range("A1").Value & "  -  " & Format(n, "#,##0") & sufixo
End Sub

' Transforma o rotulo de um card do PAINEL em atalho para a aba de origem.
Private Sub LinkCard(ws As Worksheet, topRow As Long, leftCol As Long, destino As String)
    Dim cel As Range : Set cel = ws.Cells(topRow, leftCol)
    Dim corLbl As Long : corLbl = cel.Font.Color
    Dim txt As String : txt = CStr(cel.Value)
    On Error Resume Next
    ws.Hyperlinks.Add Anchor:=cel, Address:="", _
        SubAddress:="'" & destino & "'!A1", ScreenTip:="Abrir " & destino
    On Error GoTo 0
    With cel
        .Value = txt & "  " & ChrW(8594)   ' seta indica que o card e clicavel
        .Font.Name = "Segoe UI" : .Font.Size = 9 : .Font.Bold = True
        .Font.Color = corLbl : .Font.Underline = False
        .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub AplicarDesignGlobal(wb As Workbook)
    Dim nomes As Variant, rots As Variant
    nomes = Array("PAINEL DO GESTOR", "ANALISE SAP x PRJ", "RACIONALIZACAO COM", _
                  "PRECO UNITARIO", "ALERTA CRITICO", "RANKING DE RISCO")
    rots = Array("PAINEL", "ANALISE", "RACIONALIZACAO", "PRECOS", "ALERTAS", "RANKING")

    Dim ws As Worksheet, r As Long, i As Long, j As Long

    ' ---- 1) PRECO UNITARIO ganha banda de titulo com contadores ----
    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets("PRECO UNITARIO")
    On Error GoTo 0
    If Not ws Is Nothing Then
        ws.Activate
        On Error Resume Next
        ActiveWindow.FreezePanes = False
        On Error GoTo 0
        ws.Rows("1:2").Insert
        Dim lastP As Long : lastP = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim nIt As Long, nDv As Long
        If lastP >= 4 Then
            Dim dP As Variant : dP = ws.Range(ws.Cells(4, 1), ws.Cells(lastP, 12)).Value
            For r = 1 To UBound(dP, 1)
                nIt = nIt + 1
                Dim stP As String : stP = UCase(Trim(CStr(dP(r, 12))))
                If stP = "ABAIXO DO MINIMO" Or stP = "ACIMA DO MAXIMO" Then nDv = nDv + 1
            Next r
        End If
        With ws.Range("A1:N2")
            .Merge
            .Value = "PRECO UNITARIO  -  " & Format(nIt, "#,##0") & " itens  |  " & _
                     Format(nDv, "#,##0") & " divergencia(s) de faixa"
            .Font.Name = "Segoe UI Semibold" : .Font.Size = 13 : .Font.Bold = True
            .Font.Color = RGB(235, 240, 248) : .Interior.Color = RGB(17, 24, 39)
            .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
        End With
        ws.Rows(1).RowHeight = 20 : ws.Rows(2).RowHeight = 20
        On Error Resume Next
        ws.Range("A4").Select
        ActiveWindow.FreezePanes = True
        On Error GoTo 0
        ws.Range("A1").Select
    End If

    ' ---- 2) contadores dinamicos nos titulos ----
    ' ALERTA CRITICO tem topo de altura variavel (resumo por tipo): pergunta
    ' onde os dados comecam em vez de assumir a linha 4.
    Dim wsAC As Worksheet
    On Error Resume Next
    Set wsAC = wb.Worksheets("ALERTA CRITICO")
    On Error GoTo 0
    If Not wsAC Is Nothing Then
        AppendContagem wb, "ALERTA CRITICO", PrimeiraLinhaAlerta(wsAC), " apontamento(s)"
    End If
    AppendContagem wb, "RANKING DE RISCO", 4, " obra(s)"

    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets("RACIONALIZACAO COM")
    On Error GoTo 0
    If Not ws Is Nothing Then
        Dim lastC As Long : lastC = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim nFora As Long : nFora = 0
        If lastC >= 3 Then
            Dim dC As Variant : dC = ws.Range(ws.Cells(3, 1), ws.Cells(lastC, 9)).Value
            For r = 1 To UBound(dC, 1)
                Select Case UCase(Trim(CStr(dC(r, 9))))
                    Case "FALTANDO", "EXCESSO", "EXAGERADO", "ZERO" : nFora = nFora + 1
                End Select
            Next r
        End If
        ws.Range("A1").Value = ws.Range("A1").Value & "  -  " & _
                               Format(nFora, "#,##0") & " fora do previsto"
    End If

    ' ---- 3) cor da guia de cada aba ----
    For i = 0 To UBound(nomes)
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Worksheets(CStr(nomes(i)))
        On Error GoTo 0
        If Not ws Is Nothing Then ws.Tab.Color = CorAcento(CStr(nomes(i)))
    Next i

    ' ---- 4) drill-down: cards do PAINEL viram atalhos ----
    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets("PAINEL DO GESTOR")
    On Error GoTo 0
    If Not ws Is Nothing Then
        LinkCard ws, 4, 1, "RANKING DE RISCO"     ' Valor total -> visao por obra
        LinkCard ws, 4, 3, "ANALISE SAP x PRJ"    ' PEPs analisados
        LinkCard ws, 4, 5, "ANALISE SAP x PRJ"    ' PEPs aprovados
        LinkCard ws, 4, 7, "ANALISE SAP x PRJ"    ' PEPs reprovados
        LinkCard ws, 8, 1, "ALERTA CRITICO"       ' Alertas criticos
        LinkCard ws, 8, 3, "PRECO UNITARIO"       ' Divergencias de preco
        LinkCard ws, 8, 5, "PRECO UNITARIO"       ' Sobrepreco potencial
        LinkCard ws, 8, 7, "RANKING DE RISCO"     ' Valor em risco
    End If

    ' ---- 5) barra de navegacao no TOPO (linha 1) de cada aba ----
    ' Roda por ULTIMO: o conteudo desce 1 linha e o congelamento e refeito.
    For i = 0 To UBound(nomes)
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Worksheets(CStr(nomes(i)))
        On Error GoTo 0
        If Not ws Is Nothing Then
            Dim congRow As Long
            Select Case CStr(nomes(i))
                Case "PAINEL DO GESTOR":   congRow = 0    ' sem congelamento
                Case "ANALISE SAP x PRJ":  congRow = 11   ' dados comecam na 11
                Case "RACIONALIZACAO COM": congRow = 4
                ' +1: a barra de navegacao empurra tudo uma linha para baixo
                Case "ALERTA CRITICO":     congRow = PrimeiraLinhaAlerta(ws) + 1
                Case Else:                 congRow = 5    ' PRECOS, RANKING
            End Select
            InserirBarraNav ws, nomes, rots, CStr(nomes(i)), congRow
        End If
    Next i
End Sub

' Localiza a aba base pelas colunas MAT LIB SAP / MAT PRJ CAD
Private Function AcharBaseInventario(wb As Workbook) As Worksheet

    Dim ws      As Worksheet
    Dim mp      As Object

    For Each ws In wb.Worksheets
        Select Case ws.Name
            Case "PAINEL DO GESTOR", "ANALISE SAP x PRJ", "RESUMO SAP x PRJ", "RACIONALIZACAO COM", "ALERTA CRITICO", "PRECO UNITARIO", "RANKING DE RISCO", "CONFIG", "HISTORICO", "LOG EXECUCAO", "TESTES", "DIAGNOSTICO BASE", "DEVOLUCOES", "EQUIVALENCIAS", "HISTORICO KPI"
                ' aba de saida / apoio - ignora
            Case Else
                If ws.UsedRange.Rows.Count > 1 Then
                    ' pelo mapa de colunas: aceita MAT LIB SAP / ATENDIDO SAP e
                    ' MAT PRJ CAD / TOMBADO PROJ (layout antigo ou novo)
                    Set mp = MapaColunas(ws)
                    If mp.Exists("MAT LIB SAP") And mp.Exists("MAT PRJ CAD") Then
                        Set AcharBaseInventario = ws
                        Exit Function
                    End If
                End If
        End Select
    Next ws

    Set AcharBaseInventario = Nothing
End Function

' ===========================================================================
'  PAINEL DO GESTOR: visao executiva com KPIs em cartoes + ranking de alertas
'  Le os resultados ja gerados (ANALISE SAP x PRJ, PRECO UNITARIO, ALERTA CRITICO)
' ===========================================================================
Private Function ProcessarPainelGestor(wb As Workbook, wsBase As Worksheet, _
        wsDet As Worksheet, wsPU As Worksheet, wsAlertaC As Worksheet) As Worksheet

    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    ws.Name = "PAINEL DO GESTOR"

    ' ---------------- COLETA DE KPIs ----------------
    Dim valTot As Double, valRisk As Double
    Dim nPep3 As Long, nReprov As Long, nAprov As Long
    Dim nAlertas As Long, nDiverg As Long, sobrepreco As Double

    ' PEPs + valor + reprovacao (le ANALISE SAP x PRJ, dados a partir da linha 10)
    Dim pep3St As Object : Set pep3St = CreateObject("Scripting.Dictionary")  ' pep3 -> reprovado?
    Dim lastD As Long : lastD = wsDet.Cells(wsDet.Rows.Count, 1).End(xlUp).Row
    If lastD >= 10 Then
        Dim dd As Variant : dd = wsDet.Range(wsDet.Cells(10, 1), wsDet.Cells(lastD, 15)).Value
        Dim r As Long
        For r = 1 To UBound(dd, 1)
            Dim p3 As String : p3 = Trim(CStr(dd(r, 1)))
            Dim vlr As Double : vlr = Val0(dd(r, 6))
            Dim apv As String : apv = UCase(CStr(dd(r, 14)))
            Dim isRep As Boolean : isRep = (InStr(apv, "REPROVADO") > 0)
            valTot = valTot + vlr
            If isRep Then valRisk = valRisk + Abs(vlr)
            If p3 <> "" Then
                If Not pep3St.Exists(p3) Then
                    pep3St.Add p3, isRep
                ElseIf isRep Then
                    pep3St(p3) = True
                End If
            End If
        Next r
    End If
    nPep3 = pep3St.Count
    Dim k As Variant
    For Each k In pep3St.Keys
        If pep3St(k) Then nReprov = nReprov + 1
    Next k
    nAprov = nPep3 - nReprov

    ' Alertas criticos (le ALERTA CRITICO) + ranking por tipo
    Dim tipos As Object : Set tipos = CreateObject("Scripting.Dictionary")
    Dim r0A As Long : r0A = PrimeiraLinhaAlerta(wsAlertaC)
    Dim lastA As Long : lastA = wsAlertaC.Cells(wsAlertaC.Rows.Count, 1).End(xlUp).Row
    If lastA >= r0A Then
        Dim aa As Variant : aa = wsAlertaC.Range(wsAlertaC.Cells(r0A, 1), wsAlertaC.Cells(lastA, 2)).Value
        For r = 1 To UBound(aa, 1)
            Dim tA As String : tA = Trim(CStr(aa(r, 1)))
            ' ignora a linha informativa "Nenhum alerta..." da aba sem apontamentos
            ' e os apontamentos QUALIDADE DA BASE, que sao replicas do
            ' DIAGNOSTICO DA BASE e ja tem indicador proprio (nota de saude):
            ' conta-los aqui inflaria o KPI de alertas por obra.
            If tA <> "" And InStr(UCase(tA), "NENHUM ALERTA") = 0 _
               And InStr(UCase(tA), "QUALIDADE DA BASE") = 0 Then
                nAlertas = nAlertas + 1
                If tipos.Exists(tA) Then tipos(tA) = tipos(tA) + 1 Else tipos.Add tA, 1
            End If
        Next r
    End If

    ' Divergencias de preco + sobrepreco potencial (le PRECO UNITARIO, dados linha 2+)
    Dim lastP As Long : lastP = wsPU.Cells(wsPU.Rows.Count, 1).End(xlUp).Row
    If lastP >= 2 Then
        Dim pp As Variant : pp = wsPU.Range(wsPU.Cells(2, 1), wsPU.Cells(lastP, 13)).Value
        For r = 1 To UBound(pp, 1)
            Dim stP As String : stP = UCase(Trim(CStr(pp(r, 12))))
            If stP = "ABAIXO DO MINIMO" Or stP = "ACIMA DO MAXIMO" Then nDiverg = nDiverg + 1
            If stP = "ACIMA DO MAXIMO" Then
                Dim qq As Double : qq = Val0(pp(r, 7))
                Dim puu As Double : puu = Val0(pp(r, 9))
                Dim mxx As Double : mxx = Val0(pp(r, 11))
                If puu > mxx And qq > 0 Then sobrepreco = sobrepreco + (puu - mxx) * qq
            End If
        Next r
    End If

    Dim pctAprov As Double
    If nPep3 > 0 Then pctAprov = nAprov / nPep3 * 100

    ' ---------------- LAYOUT ----------------
    ws.Cells.Interior.Color = RGB(248, 249, 251)
    Dim cc As Long
    For cc = 1 To 8 : ws.Columns(cc).ColumnWidth = 16.5 : Next cc

    ' Titulo
    With ws.Range("A1:H1")
        .Merge : .Value = "PAINEL DO GESTOR"
        .Font.Name = "Segoe UI" : .Font.Size = 24 : .Font.Bold = True
        .Font.Color = RGB(255, 255, 255) : .Interior.Color = RGB(17, 24, 39)
        .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 42
    With ws.Range("A2:H2")
        .Merge
        .Value = "Analise de Inventario  -  Base: " & wsBase.Name & "   |   Gerado em " & Format(Now, "dd/mm/yyyy hh:mm")
        .Font.Name = "Segoe UI" : .Font.Size = 10 : .Font.Color = RGB(120, 128, 140)
        .Interior.Color = RGB(31, 41, 59) : .Font.Color = RGB(200, 208, 220)
        .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
    End With
    ws.Rows(2).RowHeight = 20

    ' --- variacao vs. execucao anterior (aba HISTORICO KPI) ---
    ' Lido ANTES de registrar esta rodada, senao a comparacao seria consigo mesma.
    Dim kAnt As Variant
    If Not mModoLote Then kAnt = LerUltimoKpiRun(wb)    ' lote: bases diferentes
    Dim dTxt(1 To 4) As String, dCor(1 To 4) As Long
    Dim temAnt As Boolean : temAnt = Not IsEmpty(kAnt)
    If temAnt Then
        dTxt(1) = TxtDelta(nPep3, kAnt(1), False, dCor(1))
        dTxt(2) = TxtDelta(nReprov, kAnt(2), True, dCor(2))
        dTxt(3) = TxtDelta(nAlertas, kAnt(3), True, dCor(3))
        dTxt(4) = TxtDelta(nDiverg, kAnt(4), True, dCor(4))
    End If

    ' --- KPI cards (linha 1 de cards: 4) ---
    CardGestor ws, 4, 1, "Valor total da obra", FmtKPI(valTot), RGB(255, 255, 255), RGB(17, 24, 39)
    CardGestor ws, 4, 3, "PEPs analisados", Format(nPep3, "#,##0"), RGB(255, 255, 255), RGB(17, 24, 39), dTxt(1), dCor(1)
    CardGestor ws, 4, 5, "PEPs aprovados", Format(nAprov, "#,##0") & "  (" & Format(pctAprov, "0") & "%)", RGB(231, 247, 237), RGB(21, 115, 71)
    CardGestor ws, 4, 7, "PEPs reprovados", Format(nReprov, "#,##0"), RGB(252, 232, 233), RGB(192, 0, 0), dTxt(2), dCor(2)
    ' --- KPI cards (linha 2 de cards: 4) ---
    CardGestor ws, 8, 1, "Alertas criticos", Format(nAlertas, "#,##0"), RGB(255, 248, 230), RGB(176, 124, 0), dTxt(3), dCor(3)
    CardGestor ws, 8, 3, "Divergencias de preco", Format(nDiverg, "#,##0"), RGB(234, 242, 250), RGB(31, 78, 121), dTxt(4), dCor(4)
    CardGestor ws, 8, 5, "Sobrepreco potencial", FmtKPI(sobrepreco), RGB(252, 232, 233), RGB(192, 0, 0)
    CardGestor ws, 8, 7, "Valor em risco (reprov.)", FmtKPI(valRisk), RGB(252, 232, 233), RGB(192, 0, 0)

    ' --- Ranking de alertas por tipo (com mini-barras) ---
    Dim secRow As Long : secRow = 13
    With ws.Range(ws.Cells(secRow, 1), ws.Cells(secRow, 8))
        .Merge : .Value = "ALERTAS CRITICOS POR TIPO"
        .Font.Name = "Segoe UI" : .Font.Size = 12 : .Font.Bold = True
        .Font.Color = RGB(255, 255, 255) : .Interior.Color = RGB(31, 41, 59)
        .HorizontalAlignment = xlLeft : .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    ws.Rows(secRow).RowHeight = 26

    ' ordena tipos por contagem desc
    Dim nT As Long : nT = tipos.Count
    Dim hr As Long : hr = secRow + 1
    If nT = 0 Then
        With ws.Range(ws.Cells(hr, 1), ws.Cells(hr, 8))
            .Merge : .Value = "Nenhum alerta critico encontrado. Parabens!"
            .Font.Name = "Segoe UI" : .Font.Size = 11 : .Font.Italic = True
            .Font.Color = RGB(21, 115, 71)
            .HorizontalAlignment = xlCenter
        End With
        ws.Rows(hr).RowHeight = 24
    Else
        Dim ks() As String, vs() As Long
        ReDim ks(0 To nT - 1) : ReDim vs(0 To nT - 1)
        Dim ix As Long : ix = 0
        For Each k In tipos.Keys
            ks(ix) = CStr(k) : vs(ix) = tipos(k) : ix = ix + 1
        Next k
        Dim a As Long, b As Long, tS As String, tV As Long
        For a = 0 To nT - 2
            For b = 0 To nT - a - 2
                If vs(b) < vs(b + 1) Then
                    tV = vs(b) : vs(b) = vs(b + 1) : vs(b + 1) = tV
                    tS = ks(b) : ks(b) = ks(b + 1) : ks(b + 1) = tS
                End If
            Next b
        Next a
        Dim maxV As Long : maxV = vs(0) : If maxV < 1 Then maxV = 1

        ' cabecalho da mini-tabela
        ws.Cells(hr, 1).Value = "TIPO DE ALERTA"
        ws.Cells(hr, 4).Value = "QTD"
        ws.Cells(hr, 5).Value = "PARTICIPACAO"
        Dim hc As Long
        For hc = 1 To 8
            With ws.Cells(hr, hc)
                .Font.Name = "Segoe UI" : .Font.Bold = True : .Font.Size = 9
                .Font.Color = RGB(120, 128, 140)
            End With
        Next hc
        ws.Range(ws.Cells(hr, 1), ws.Cells(hr, 3)).Merge
        ws.Range(ws.Cells(hr, 5), ws.Cells(hr, 8)).Merge

        Dim rr As Long : rr = hr + 1
        For a = 0 To nT - 1
            With ws.Range(ws.Cells(rr, 1), ws.Cells(rr, 3))
                .Merge : .Value = ks(a)
                .Font.Name = "Segoe UI" : .Font.Size = 10 : .Font.Color = RGB(40, 46, 56)
                .HorizontalAlignment = xlLeft : .IndentLevel = 1
            End With
            With ws.Cells(rr, 4)
                .Value = vs(a)
                .Font.Name = "Segoe UI" : .Font.Size = 11 : .Font.Bold = True
                .Font.Color = RGB(17, 24, 39) : .HorizontalAlignment = xlCenter
            End With
            Dim nB As Long : nB = Int(vs(a) / maxV * 24) : If nB < 1 Then nB = 1
            With ws.Range(ws.Cells(rr, 5), ws.Cells(rr, 8))
                .Merge : .Value = String(nB, ChrW(9608)) & "  " & Format(vs(a) / nAlertas, "0%")
                .Font.Name = "Consolas" : .Font.Size = 10 : .Font.Color = RGB(214, 89, 17)
                .HorizontalAlignment = xlLeft
            End With
            If a Mod 2 = 1 Then
                ws.Range(ws.Cells(rr, 1), ws.Cells(rr, 8)).Interior.Color = RGB(255, 255, 255)
            End If
            ws.Rows(rr).RowHeight = 20
            rr = rr + 1
        Next a
    End If

    ' Nota de saude da base (DIAGNOSTICO BASE ja rodou nesta altura)
    Dim notaB As Double
    Dim wsDg As Worksheet
    On Error Resume Next
    Set wsDg = wb.Worksheets("DIAGNOSTICO BASE")
    If Not wsDg Is Nothing Then notaB = Val0(wsDg.Range("A5").Value)
    On Error GoTo 0

    ' Fecha a rodada no historico de KPIs (a proxima execucao compara com esta)
    If Not mModoLote Then RegistrarKpiRun wb, nPep3, nReprov, nAlertas, nDiverg, valRisk, notaB

    ' Aparencia geral
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range("A1").Select
    On Error GoTo 0
    ws.Tab.Color = RGB(17, 24, 39)

    Set ProcessarPainelGestor = ws
End Function

' Desenha um cartao de KPI ocupando 2 colunas x 3 linhas (label + valor grande)
' Card de KPI. Com 'delta' preenchido, a 3a linha do card deixa de fazer parte
' do valor e vira o rodape de variacao vs. execucao anterior - a altura total
' do card nao muda, entao o layout do painel continua o mesmo.
Private Sub CardGestor(ws As Worksheet, ByVal topRow As Long, ByVal leftCol As Long, _
        titulo As String, valor As String, corFundo As Long, corValor As Long, _
        Optional ByVal delta As String = "", Optional ByVal corDelta As Long = 0)
    With ws.Range(ws.Cells(topRow, leftCol), ws.Cells(topRow + 2, leftCol + 1))
        .Interior.Color = corFundo
        .Borders.LineStyle = xlContinuous : .Borders.Color = RGB(222, 226, 232)
    End With
    With ws.Range(ws.Cells(topRow, leftCol), ws.Cells(topRow, leftCol + 1))
        .Merge : .Value = UCase(titulo)
        .Font.Name = "Segoe UI" : .Font.Size = 9 : .Font.Bold = True
        .Font.Color = RGB(110, 118, 130)
        .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
    End With
    If delta = "" Then
        With ws.Range(ws.Cells(topRow + 1, leftCol), ws.Cells(topRow + 2, leftCol + 1))
            .Merge : .Value = valor
            .Font.Name = "Segoe UI" : .Font.Size = 19 : .Font.Bold = True
            .Font.Color = corValor
            .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
        End With
    Else
        With ws.Range(ws.Cells(topRow + 1, leftCol), ws.Cells(topRow + 1, leftCol + 1))
            .Merge : .Value = valor
            .Font.Name = "Segoe UI" : .Font.Size = 17 : .Font.Bold = True
            .Font.Color = corValor
            .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
        End With
        With ws.Range(ws.Cells(topRow + 2, leftCol), ws.Cells(topRow + 2, leftCol + 1))
            .Merge : .Value = delta
            .Font.Name = "Segoe UI" : .Font.Size = 9 : .Font.Bold = True
            .Font.Color = corDelta
            .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
        End With
    End If
    ws.Rows(topRow).RowHeight = 18
    ws.Rows(topRow + 1).RowHeight = 22
    ws.Rows(topRow + 2).RowHeight = 18
End Sub

' ===========================================================================
'  MELHORIAS: CONFIG, LOG, VALIDACAO, RESUMO PEP3, DEVOLUCOES, HISTORICO,
'  TESTES DE LOGICA e NOMES/TABELA ESTRUTURADA
' ===========================================================================

' Registra o resultado de uma etapa no log; erro nao aborta o processo (erro granular).
Private Sub FimEtapa(erroNum As Long, erroDesc As String, etapa As String, t0 As Double)
    Dim seg As Double : seg = Timer - t0
    If seg < 0 Then seg = seg + 86400           ' virada da meia-noite
    If erroNum <> 0 Then
        mFalhas = mFalhas & vbCrLf & "  - " & etapa & ": " & erroDesc
        LogAdd "FALHA [" & etapa & "]: (" & erroNum & ") " & erroDesc
    Else
        LogAdd etapa & " OK (" & Format(seg, "0.0") & "s)"
    End If
    If mEtapas Is Nothing Then Set mEtapas = New Collection
    mEtapas.Add Array(etapa, (erroNum = 0), seg, IIf(erroNum <> 0, erroDesc, ""))
End Sub

' Aviso nao fatal da execucao (a analise seguiu, mas sem algo): vai para a
' mensagem final, para o LOG e para o bloco SAUDE DA EXECUCAO do painel.
Private Sub AvisoExec(ByVal msg As String)
    mFalhas = mFalhas & vbCrLf & "  - " & msg
    mAvisosExec = mAvisosExec & vbCrLf & "  - " & msg
    LogAdd "AVISO: " & msg
End Sub

' Bloco SAUDE DA EXECUCAO no fim do PAINEL DO GESTOR: cada etapa com status e
' tempo, a fonte de precos usada e os avisos. Com falha ou aviso, o subtitulo
' do painel ganha o selo da execucao. Escrito depois de TODAS as etapas, entao
' vai junto no arquivo exportado pelo modo lote (antes o aviso so existia na
' MsgBox final e no LOG oculto, e o resultado exportado parecia completo).
Private Sub EscreverSaudeExecucao(ws As Worksheet, wsBase As Worksheet)
    Dim it As Variant, nFal As Long
    If Not mEtapas Is Nothing Then
        For Each it In mEtapas
            If Not CBool(it(1)) Then nFal = nFal + 1
        Next it
    End If
    Dim precoRuim As Boolean
    precoRuim = (Left(mFontePrecos, 7) = "NENHUMA" Or InStr(mFontePrecos, "NAO ABRIU") > 0 _
                 Or InStr(mFontePrecos, "vazia") > 0)
    Dim selo As String, corSelo As Long
    If nFal > 0 Then
        selo = "ANALISE INCOMPLETA: " & nFal & " etapa(s) com falha" : corSelo = RGB(192, 0, 0)
    ElseIf mAvisosExec <> "" Or precoRuim Then
        selo = "ANALISE COM AVISOS" : corSelo = RGB(176, 124, 0)
    Else
        selo = "EXECUCAO OK" : corSelo = RGB(21, 115, 71)
    End If

    ' 3 linhas abaixo do ultimo conteudo do painel
    Dim uR As Range : Set uR = ws.Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious)
    Dim r As Long : r = 3 : If Not uR Is Nothing Then r = uR.Row + 3

    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 8))
        .Merge : .Value = "SAUDE DA EXECUCAO  -  " & selo
        .Font.Name = "Segoe UI" : .Font.Size = 12 : .Font.Bold = True
        .Font.Color = RGB(255, 255, 255) : .Interior.Color = corSelo
        .HorizontalAlignment = xlLeft : .VerticalAlignment = xlCenter : .IndentLevel = 1
    End With
    ws.Rows(r).RowHeight = 26
    r = r + 1
    SaudeLinha ws, r, "ETAPA", "STATUS", "TEMPO (s)", "DETALHE", RGB(31, 41, 59), RGB(235, 240, 248), True
    r = r + 1
    If Not mEtapas Is Nothing Then
        For Each it In mEtapas
            SaudeLinha ws, r, CStr(it(0)), IIf(CBool(it(1)), "OK", "FALHA"), Format(it(2), "0.0"), _
                       CStr(it(3)), IIf(CBool(it(1)), RGB(231, 247, 237), RGB(252, 226, 228)), _
                       IIf(CBool(it(1)), RGB(21, 115, 71), RGB(192, 0, 0)), False
            r = r + 1
        Next it
    End If
    ' fonte de precos (cadeia: aba interna -> arquivo externo)
    Dim txtF As String : txtF = IIf(mFontePrecos <> "", mFontePrecos, "nao consultada")
    If Not IsEmpty(mFontePrecosData) Then txtF = txtF & " - alterado em " & Format(mFontePrecosData, "dd/mm/yyyy")
    SaudeLinha ws, r, "FONTE DE PRECOS (faixa MIN/MAX)", IIf(precoRuim, "ATENCAO", "OK"), "", txtF, _
               IIf(precoRuim, RGB(255, 244, 214), RGB(244, 247, 250)), _
               IIf(precoRuim, RGB(176, 124, 0), RGB(45, 52, 64)), False
    r = r + 1
    If mFontePrecosRef <> "" Then
        SaudeLinha ws, r, "PRECO DE REFERENCIA (UC SUBVAL.)", "OK", "", mFontePrecosRef, _
                   RGB(244, 247, 250), RGB(45, 52, 64), False
        r = r + 1
    End If
    ' avisos nao fatais, um por linha
    If mAvisosExec <> "" Then
        Dim av As Variant
        For Each av In Split(Mid(mAvisosExec, 3), vbCrLf)
            If Trim(CStr(av)) <> "" Then
                SaudeLinha ws, r, "AVISO", "ATENCAO", "", Trim(Replace(CStr(av), "  - ", "", 1, 1)), _
                           RGB(255, 244, 214), RGB(176, 124, 0), False
                r = r + 1
            End If
        Next av
    End If

    ' selo no subtitulo do painel ("Analise de Inventario - Base: ...")
    If selo <> "EXECUCAO OK" Then
        Dim rs As Long
        For rs = 1 To 6
            If Left(CStr(ws.Cells(rs, 1).Value), 21) = "Analise de Inventario" Then
                ws.Cells(rs, 1).Value = CStr(ws.Cells(rs, 1).Value) & "   |   " & ChrW(&H26A0) & " " & selo
                ws.Cells(rs, 1).Font.Color = RGB(255, 196, 0)
                ws.Cells(rs, 1).Font.Bold = True
                Exit For
            End If
        Next rs
    End If
End Sub

' Uma linha do bloco SAUDE DA EXECUCAO: A:C etapa | D status | E tempo | F:H detalhe.
Private Sub SaudeLinha(ws As Worksheet, ByVal r As Long, ByVal etapa As String, _
        ByVal status As String, ByVal tempo As String, ByVal detalhe As String, _
        ByVal corFundo As Long, ByVal corStatus As Long, ByVal cabecalho As Boolean)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Merge
    ws.Range(ws.Cells(r, 6), ws.Cells(r, 8)).Merge
    ws.Cells(r, 1).Value = etapa : ws.Cells(r, 4).Value = status
    ws.Cells(r, 5).Value = tempo : ws.Cells(r, 6).Value = detalhe
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 8))
        .Interior.Color = corFundo
        .Font.Name = "Segoe UI" : .Font.Size = 9
        .Font.Color = IIf(cabecalho, corStatus, RGB(45, 52, 64))
        .Font.Bold = cabecalho
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(222, 226, 232)
    End With
    ws.Range(ws.Cells(r, 4), ws.Cells(r, 5)).HorizontalAlignment = xlCenter
    If Not cabecalho Then
        ws.Cells(r, 4).Font.Bold = True : ws.Cells(r, 4).Font.Color = corStatus
    End If
    ws.Cells(r, 1).IndentLevel = 1
    ws.Rows(r).RowHeight = 18
End Sub

' --------- CONFIG editavel (parametros centralizados na aba CONFIG) ---------
' Parametros padrao da aba CONFIG: CHAVE | VALOR | DESCRICAO.
Private Function CfgPadrao() As Collection
    Dim c As Collection : Set c = New Collection
    c.Add Array("TOL_ADERENCIA_PCT", 0.1, "Tolerancia de aderencia SAP x PRJ (fracao; 0,1 = 10%)")
    c.Add Array("TOL_SUBVAL", 0.9, "UC subvalorizado: alerta se PU menor que fator x referencia")
    c.Add Array("MIN_DIVERG_RS", 100, "Materialidade minima da divergencia em R$ p/ alertar")
    c.Add Array("CABO_ISOLADO_MAX_M", 15, "Cabo isolado abaixo deste comprimento (m) fica isento de UC")
    c.Add Array("CORTE_DIVERG_QTD", 20, "Destaca linhas com diferenca de qtd acima deste valor")
    c.Add Array("PESO_REPROV", 40, "Ranking: peso de obra reprovada na SAP x PRJ")
    c.Add Array("PESO_ALERTA", 4, "Ranking: peso por alerta critico")
    c.Add Array("CAP_ALERTA", 24, "Ranking: teto da parcela de alertas")
    c.Add Array("PESO_PU", 3, "Ranking: peso por divergencia de preco unitario")
    c.Add Array("CAP_PU", 18, "Ranking: teto da parcela de PU")
    c.Add Array("PESO_COM", 2, "Ranking: peso por COM fora do previsto NT.006")
    c.Add Array("CAP_COM", 18, "Ranking: teto da parcela de COM")
    c.Add Array("PESO_DISCREP", 2, "Ranking: peso por discrepancia de valor de material (ODI)")
    c.Add Array("CAP_DISCREP", 12, "Ranking: teto da parcela de discrepancia de valor")
    c.Add Array("COM_CRITICO", "CH FUS; PARA RAIO", "Familias COM criticas (separadas por ;) que reprovam o PEP3")
    c.Add Array("DIF_ABS_COND_NU", 5, "COND NU: diferenca ABSOLUTA (nao %) ate este valor = ADERENTE")
    c.Add Array("DIF_ABS_CABO", 15, "Familias de DIF_ABS_FAMILIAS: dif ABSOLUTA ate este valor = ADERENTE")
    c.Add Array("DIF_ABS_FAMILIAS", "CABO ISOLADO; CABO COBRE; COND COBRE", "Familias (separadas por ;) que usam a regra DIF_ABS_CABO")
    c.Add Array("PU_ODI_MIN", 0.01, "Alerta PU BAIXO EM ODI: preco unitario minimo da faixa (R$)")
    c.Add Array("PU_ODI_MAX", 10, "Alerta PU BAIXO EM ODI: preco unitario maximo da faixa (R$)")
    c.Add Array("PU_CABO_ALERTA", 0.01, "Alerta PU 0,01 EM CABO/COND: preco unitario simbolico (R$)")
    c.Add Array("ODD_VALOR_MIN", -100, "ODD: material com VALOR abaixo deste limite (R$) e NAO ADERENTE")
    c.Add Array("POSTE_ODI_PU_MIN", 500, "Alerta POSTE ODI PU BAIXO: poste em ODI com preco unitario abaixo de (R$)")
    c.Add Array("PRECOS_MAX_DIAS", 90, "Diagnostico: arquivo da base de precos mais antigo que isto (dias) gera ATENCAO")
    c.Add Array("DISCREP_VALOR_PCT", 0.3, "Discrepancia de valor de material (ODI): PU fora de +/- esta fracao da media do codigo")
    Set CfgPadrao = c
End Function

Private Sub GarantirConfig(wb As Workbook)
    Set mCfg = CreateObject("Scripting.Dictionary")

    ' Defaults: usados na criacao da aba E para acrescentar chaves novas
    ' em planilhas que ja tem a aba CONFIG de versao anterior do modulo.
    ' (montado por CfgPadrao, uma linha por parametro: o Array com "_" ja estava
    ' em 22 das 24 continuacoes de linha que o VBA aceita numa instrucao)
    Dim cl As Collection : Set cl = CfgPadrao()
    Dim df() As Variant : ReDim df(0 To cl.Count - 1)
    Dim jD As Long
    For jD = 1 To cl.Count
        df(jD - 1) = cl(jD)
    Next jD

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("CONFIG")
    On Error GoTo 0
    Dim j As Long
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "CONFIG"
        ws.Cells(1, 1).Value = "CHAVE"
        ws.Cells(1, 2).Value = "VALOR"
        ws.Cells(1, 3).Value = "DESCRICAO"
        For j = 0 To UBound(df)
            ws.Cells(2 + j, 1).Value = df(j)(0)
            ws.Cells(2 + j, 2).Value = df(j)(1)
            ws.Cells(2 + j, 3).Value = df(j)(2)
        Next j
        With ws.Range("A1:C1")
            .Font.Bold = True : .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(17, 24, 39) : .HorizontalAlignment = xlCenter
        End With
        ws.Columns(1).ColumnWidth = 22 : ws.Columns(2).ColumnWidth = 20 : ws.Columns(3).ColumnWidth = 62
        ws.Range(ws.Cells(2, 2), ws.Cells(2 + UBound(df), 2)).HorizontalAlignment = xlCenter
        ws.Tab.Color = RGB(90, 90, 90)
    Else
        ' Upgrade: acrescenta no fim da aba as chaves default que ainda nao existem
        Dim lastX As Long : lastX = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim jaTem As Object : Set jaTem = CreateObject("Scripting.Dictionary")
        Dim rX As Long, ch As String
        For rX = 2 To lastX
            ch = Trim(CStr(ws.Cells(rX, 1).Value))
            If ch <> "" And Not jaTem.Exists(ch) Then jaTem.Add ch, True
        Next rX
        For j = 0 To UBound(df)
            If Not jaTem.Exists(CStr(df(j)(0))) Then
                lastX = lastX + 1
                ws.Cells(lastX, 1).Value = df(j)(0)
                ws.Cells(lastX, 2).Value = df(j)(1)
                ws.Cells(lastX, 3).Value = df(j)(2)
                ws.Cells(lastX, 2).HorizontalAlignment = xlCenter
            End If
        Next j
    End If

    Dim lastR As Long : lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long, chave As String
    For r = 2 To lastR
        chave = Trim(CStr(ws.Cells(r, 1).Value))
        If chave <> "" And Not mCfg.Exists(chave) Then mCfg.Add chave, CStr(ws.Cells(r, 2).Value)
    Next r

    ' valores usados em hot path (evita reparse por chamada)
    mTolAder = CfgD("TOL_ADERENCIA_PCT", 0.1)
    mOddValMin = CfgD("ODD_VALOR_MIN", -100)
    mCaboMax = CfgD("CABO_ISOLADO_MAX_M", 15)
    mDifCondNu = CfgD("DIF_ABS_COND_NU", 5)
    mDifAbsCabo = CfgD("DIF_ABS_CABO", 15)

    ' familias COM criticas -> tokens normalizados (UCase, sem espaco)
    Set mComCrit = CreateObject("Scripting.Dictionary")
    Dim parts() As String : parts = Split(CfgS("COM_CRITICO", "CH FUS; PARA RAIO"), ";")
    Dim tk As String
    For r = 0 To UBound(parts)
        tk = Replace(NormStr(parts(r)), " ", "")
        If tk <> "" And Not mComCrit.Exists(tk) Then mComCrit.Add tk, True
    Next r

    ' familias com dif ABSOLUTA tolerada -> tokens normalizados (UCase, sem espaco)
    Set mFamDifAbs = CreateObject("Scripting.Dictionary")
    parts = Split(CfgS("DIF_ABS_FAMILIAS", "CABO ISOLADO; CABO COBRE; COND COBRE"), ";")
    For r = 0 To UBound(parts)
        tk = Replace(NormStr(parts(r)), " ", "")
        If tk <> "" And Not mFamDifAbs.Exists(tk) Then mFamDifAbs.Add tk, True
    Next r

    OcultarAba ws      ' aba de apoio: some das guias (Reexibir para editar)
End Sub

' Oculta uma aba de apoio (CONFIG, EQUIVALENCIAS). xlSheetHidden, e nao
' VeryHidden: o usuario reexibe pelo clique direito numa guia > Reexibir para
' editar, e a proxima execucao oculta de novo. A aba ativa so pode ser oculta
' se outra estiver visivel: antes ativa a primeira aba visivel diferente dela.
Private Sub OcultarAba(ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    If ws.Visible <> xlSheetVisible Then Exit Sub
    On Error Resume Next
    If ws.Parent.ActiveSheet Is ws Then
        Dim w As Worksheet
        For Each w In ws.Parent.Worksheets
            If w.Visible = xlSheetVisible And Not w Is ws Then w.Activate : Exit For
        Next w
    End If
    ws.Visible = xlSheetHidden
    On Error GoTo 0
End Sub

Private Function CfgS(chave As String, padrao As String) As String
    CfgS = padrao
    If mCfg Is Nothing Then Exit Function
    If mCfg.Exists(chave) Then
        Dim v As String : v = Trim(CStr(mCfg(chave)))
        If v <> "" Then CfgS = v
    End If
End Function

Private Function CfgD(chave As String, padrao As Double) As Double
    Dim s As String : s = CfgS(chave, "")
    If s = "" Then CfgD = padrao : Exit Function
    ' Val interpreta "." como decimal independente do locale
    CfgD = Val(Replace(s, ",", "."))
End Function

' --------- Log de execucao (aba oculta, tempo por etapa) ---------
Private Sub LogInit()
    Set mLog = New Collection
    mFalhas = "" : mT0 = Timer
    Set mEtapas = New Collection : mAvisosExec = ""
    mFontePrecos = "" : mFontePrecosData = Empty : mFontePrecosRef = ""
    LogAdd "Inicio da execucao"
End Sub

Private Sub LogAdd(msg As String)
    If mLog Is Nothing Then Set mLog = New Collection
    mLog.Add Format(Now, "hh:nn:ss") & "  (+" & Format(Timer - mT0, "0.0") & "s)  " & msg
End Sub

Private Sub GravarLog(wb As Workbook)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("LOG EXECUCAO")   ' Nothing se a aba ainda nao existe
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "LOG EXECUCAO"
    Else
        ws.Visible = xlSheetVisible : ws.Cells.Clear
    End If
    ws.Cells(1, 1).Value = "LOG DE EXECUCAO - " & Format(Now, "dd/mm/yyyy hh:nn")
    ws.Cells(1, 1).Font.Bold = True
    Dim i As Long
    If Not mLog Is Nothing Then
        For i = 1 To mLog.Count
            ws.Cells(1 + i, 1).Value = mLog(i)
        Next i
    End If
    ws.Columns(1).ColumnWidth = 95
    ws.Tab.Color = RGB(120, 120, 120)
    ' sai da aba LOG (recem-criada fica ativa) para poder oculta-la
    On Error Resume Next   ' Worksheets(1) pode estar oculta
    wb.Worksheets(1).Activate
    ws.Visible = xlSheetHidden
    On Error GoTo 0
End Sub

' --------- Materiais equivalentes (aba EQUIVALENCIAS) ---------
' Codigos com o mesmo numero em EQUIVALENCIA sao intercambiaveis: na ANALISE
' SAP x PRJ, quando 2+ codigos do grupo aparecem no mesmo PEP4, todos ficam
' ADERENTE. Um codigo em dois grupos une os grupos (equivalencia transitiva).
' A aba e mantida pelo usuario; a macro so acrescenta grupos de versoes novas
' da lista padrao (EquivPadrao).
' Layout (o da tabela da area, colunas localizadas pelo cabecalho):
'   COD MAT | DESC MAT | UND | TIPO | FAMILIA | EQUIVALENCIA | PAR DE EQUIVALENCIA
' So COD MAT e EQUIVALENCIA sao lidas; as demais sao informativas.

' Lista padrao, uma linha por codigo:
'   COD | DESC | UND | TIPO | FAMILIA | EQUIVALENCIA | VERSAO em que entrou
' (Collection, e nao Array com "_": o VBA limita a 24 continuacoes de linha.)
Private Function EquivPadrao() As Collection
    Dim c As Collection : Set c = New Collection
    ' versao 1
    c.Add Split("105000058|CH SEC 1P VHI 15KV 630A 12,5KA RD PDE|UN|UC|CH SU FACA|1|1", "|")
    c.Add Split("105000029|CH SEC BYPAS 1F 13,8KV 630A 16KA|UN|UC|CH SU FACA|1|1", "|")
    c.Add Split("102210080|TP IND 15KV 1,2P75 500VA G1 EXT PDE|UN|UC|TP_TC|2|1", "|")
    c.Add Split("102310194|TP,13,8-11,4kV,EXT.GSCT004/21,EXC GO|UN|UC|TP_TC|2|1", "|")
    c.Add Split("105010022|CH SEC 1P VHI 36,2KV 630A 12,5KA RD PDE|UN|UC|CH SU FACA|3|1", "|")
    c.Add Split("105010003|CH SEC BYPAS 1P 34,5KV 630A 16KA|UN|UC|CH SU FACA|3|1", "|")
    c.Add Split("102100066|TD 3F 13,8KV 112,5KVA OV 3T 380/220V PDE|UN|UC|TRAFO|4|1", "|")
    c.Add Split("102100387|TD,3P,13.8-0.38/0.22kV,112.5KVA,BLQ|UN|UC|TRAFO|4|1", "|")
    c.Add Split("133000038|POSTE CONC DT B-1,5 12M 1000DAN PDE|UN|UC|POSTE RD|5|1", "|")
    c.Add Split("133000150|POSTE DT CONC B AUTO ATER 12M 1000DAN|UN|UC|POSTE RD|5|1", "|")
    ' versao 2 (o grupo 10 da tabela da area repete o 8 e ficou de fora; o 11
    ' compartilha 102310194 com o 2 e e unido a ele na leitura)
    c.Add Split("105310001|CHAVE FUS C POR/POL 34,5KV 100A 5KA PDE|UN|UC|CH FUS|6|2", "|")
    c.Add Split("105310008|CHAVE FUS DT C POL 34,5KV 100A 5KA PDE|UN|UC|CH FUS|6|2", "|")
    c.Add Split("102110399|TD,3P,34.5-0.38/0.22kV,112.5KVA,BLQ|UN|UC|TRAFO|7|2", "|")
    c.Add Split("102110065|TRAFO DT 3F 34,5KV 112,5KVA 380/220V OV|UN|UC|TRAFO|7|2", "|")
    c.Add Split("102100149|TRAFO DT 1F 13,8 FN 37,5KVA 440/220(M)OV|UN|UC|TRAFO|8|2", "|")
    c.Add Split("102100175|TRAFO DT 1F 13,8KV FN 37,5KVA 220V OV|UN|UC|TRAFO|8|2", "|")
    c.Add Split("102100160|TD 1F FN 13,8KV 10KVA 220V 3T OV PDE|UN|UC|TRAFO|9|2", "|")
    c.Add Split("102100163|TRAFO DT 1F 13,8KV FN 10KVA 440/220(M)OV|UN|UC|TRAFO|9|2", "|")
    c.Add Split("102310194|TP,13,8-11,4kV,EXT.GSCT004/21,EXC GO|UN|UC|TP_TC|11|2", "|")
    c.Add Split("102210093|TP,17,5KV,USO EXTERNO,TIPO GSCT004/1,BLQ|UN|UC|TP_TC|11|2", "|")
    c.Add Split("133000036|POSTE CONC DT B 12M 600DAN PDE|UN|UC|POSTE RD|12|2", "|")
    c.Add Split("133000149|POSTE DT CONC B AUTO ATER 12M 600DAN|UN|UC|POSTE RD|12|2", "|")
    ' versao 3 (13, 16, 17, 18 e 21 compartilham codigos com grupos anteriores
    ' e sao unidos a eles na leitura: 9+16+18, 8+17, 6+21)
    c.Add Split("102110402|TD,3P,34.5-0.38/0.22kV,75KVA,BLQ|UN|UC|TRAFO|13|3", "|")
    c.Add Split("102110064|TRAFO DT 3F 34,5KV 75KVA 380/220V OV|UN|UC|TRAFO|13|3", "|")
    c.Add Split("102110018|TRAFO DT 1F 34,5KV FN 10KVA 440/22(BLQ)|UN|UC|TRAFO|14|3", "|")
    c.Add Split("102110181|TRAFO DT 1F 34,5KV FN 10KVA 440/220(M)OV|UN|UC|TRAFO|14|3", "|")
    c.Add Split("104020017|PARA-RAIO DT POL 21KV 10KA 23,1KV PDE|UN|COM|PARA RAIO MT|15|3", "|")
    c.Add Split("104010001|PARA-RAIOS DT POLIM 12KV 10KA 13,8KV|UN|COM|PARA RAIO MT|15|3", "|")
    c.Add Split("102100021|TRAFO DT 1F 13,8KV FN 10KVA 440/22(BLQ)|UN|UC|TRAFO|16|3", "|")
    c.Add Split("102100163|TRAFO DT 1F 13,8KV FN 10KVA 440/220(M)OV|UN|UC|TRAFO|16|3", "|")
    c.Add Split("102100102|TRAFO DT 1F 13,8KV FN 37,5KVA 440/220(BQ|UN|UC|TRAFO|17|3", "|")
    c.Add Split("102100149|TRAFO DT 1F 13,8 FN 37,5KVA 440/220(M)OV|UN|UC|TRAFO|17|3", "|")
    c.Add Split("102100309|TD,1F,7,97-0,44/0,22kV,10KVA,4TAPs(DAN)|UN|UC|TRAFO|18|3", "|")
    c.Add Split("102100163|TRAFO DT 1F 13,8KV FN 10KVA 440/220(M)OV|UN|UC|TRAFO|18|3", "|")
    c.Add Split("102110020|TRAFO DT 1F 34,5KV FN 25KVA 440/220V(BQ)|UN|UC|TRAFO|19|3", "|")
    c.Add Split("102110191|TRAFO DT 1F 34,5KV FN 25KVA 440/220(M)OV|UN|UC|TRAFO|19|3", "|")
    c.Add Split("102100388|TD,3P,13.8-0.38/0.22kV,150KVA,GST01/606|UN|UC|TRAFO|20|3", "|")
    c.Add Split("102100067|TRAFO DT 3F 13,8KV 150KVA 380/220V OV|UN|UC|TRAFO|20|3", "|")
    c.Add Split("105310029|CH,FUS,36.2KV,100A,5KA,PORC,S.L,BLQ|UN|UC|CH FUS|21|3", "|")
    c.Add Split("105310001|CHAVE FUS DT C POR 34,5KV 100A 5KA PDE|UN|UC|CH FUS|21|3", "|")
    c.Add Split("122120026|CABO COB AL 150MM 15KV XLPE+HDPE+SMC|M|UC|COND PROT|22|3", "|")
    c.Add Split("122120003|CABO COB AL CA 150MM2 15KV XLPE CZ (BLQ)|M|UC|COND PROT|22|3", "|")
    ' versao 4 (o grupo 28 da tabela repete o 24 e ficou de fora; 23, 30, 31,
    ' 32 e 33 compartilham codigos e sao unidos na leitura: 6+21+23, 25+30,
    ' 4+32+33 e 31 com o par 105300003 = 105300005)
    c.Add Split("105310005|CHAVE FUS C 34,5KV 200A 150KV SUP(BLQ)|UN|UC|CH FUS|23|4", "|")
    c.Add Split("105310001|CHAVE FUS DT C POR 34,5KV 100A 5KA PDE|UN|UC|CH FUS|23|4", "|")
    c.Add Split("105000007|CHAVE SEC 1P 15KV 400A 12,5KA MHI(BLQ)|UN|UC|CH SU FACA|24|4", "|")
    c.Add Split("105000034|CHAV SEC 1F 15KV 400A MHI FC SL(BLQ)|UN|UC|CH SU FACA|24|4", "|")
    c.Add Split("102100170|TRAFO DT 1F 13,8KV FN 25KVA 220V OV|UN|UC|TRAFO|25|4", "|")
    c.Add Split("102100173|TRAFO DT 1F 13,8KV FN 25KVA 440/220(M)OV|UN|UC|TRAFO|25|4", "|")
    c.Add Split("101100044|REL,FUSE-SAVER,1F,27KV,200A,NTE04701,BLQ|UN|UC|RELIGADOR|26|4", "|")
    c.Add Split("101100027|RELIG 1F 24,2KV 100A 6,3KA S/CH PDE|UN|UC|RELIGADOR|26|4", "|")
    c.Add Split("110010002|REGU TS 1F 19,92KV 201A 400KVA C/C(BLQ)|UN|UC|REGULADOR|27|4", "|")
    c.Add Split("110010011|REGULAD TS 1F 19,92KV 201A 400KVA S/CONT|UN|UC|REGULADOR|27|4", "|")
    c.Add Split("102100315|TD,1F,7,97-0,44/0,22kV,15KVA,4TAPs(DAN)|UN|UC|TRAFO|29|4", "|")
    c.Add Split("102100168|TRAFO DT 1F 13,8KV FN 15KVA 440/220(M)OV|UN|UC|TRAFO|29|4", "|")
    c.Add Split("102100337|TD,1P,7.97-0.44/0.22kV,25KVA,BLQ|UN|UC|TRAFO|30|4", "|")
    c.Add Split("102100173|TRAFO DT 1F 13,8KV FN 25KVA 440/220(M)OV|UN|UC|TRAFO|30|4", "|")
    c.Add Split("105310016|CHAVE FUS C ZCORR 23,1KV 100A 6,3KA PDE|UN|COM|CH FUS|31|4", "|")
    c.Add Split("105300005|CHAVE FUS C 15KV 100A NBI 95KV 2KA(BLQ)|UN|COM|CH FUS|31|4", "|")
    c.Add Split("102100066|TD 3F 13,8KV 112,5KVA OV 3T 380/220V PDE|UN|UC|TRAFO|32|4", "|")
    c.Add Split("102100062|TRAFO DT 3F 13,8KV 112,5KVA 220/127V OV|UN|UC|TRAFO|32|4", "|")
    c.Add Split("102100350|TD,3F,13,8kV,380V,112KVA,4TAPs(DAN),BLQ|UN|UC|TRAFO|33|4", "|")
    c.Add Split("102100066|TD 3F 13,8KV 112,5KVA OV 3T 380/220V PDE|UN|UC|TRAFO|33|4", "|")
    c.Add Split("102110364|TD,1F,19,9-0,44/0,22kV,15KVA,4TAPs(DAN)|UN|UC|TRAFO|34|4", "|")
    c.Add Split("102110363|TD,1F,19,9-0,44/0,22kV,15KVA,4TAPs FP|UN|UC|TRAFO|34|4", "|")
    ' versao 1, renumerado depois: a tabela da area passou a usar o 6 (v2), o
    ' 13 (v3) e o 23 (v4); fica depois do ultimo grupo dela
    c.Add Split("105300003|||||35|1", "|")
    c.Add Split("105300005|||||35|1", "|")
    ' versao 5 (122120026 e 122120003 ja estao no grupo 22: 22 + 37 + 41 se unem na leitura)
    c.Add Split("102110178|TRAFO DT 1F 34,5KV FN 10KVA 220V OV|UM|UC|TRAFO|36|5", "|")
    c.Add Split("102110181|TRAFO DT 1F 34,5KV FN 10KVA 440/220(M)OV|UM|UC|TRAFO|36|5", "|")
    c.Add Split("122120026|CABO COB AL 150MM 15KV XLPE+HDPE+SMC|M|UC|COND PROT|37|5", "|")
    c.Add Split("122120028|CABO COB AL 150MM 35KV XLPE+HDPE+SMC|M|UC|COND PROT|37|5", "|")
    c.Add Split("133000127|POSTE CONC DT B-4,5 12M 2000DAN PDE|UN|UC|POSTE RD|38|5", "|")
    c.Add Split("133000041|POSTE CONC DT B-4,5 12M 2000DAN BLQ|UN|UC|POSTE RD|38|5", "|")
    c.Add Split("102100064|TD 3F 13,8KV 45KVA OV 3T 380/220V PDE|UN|UC|TRAFO|39|5", "|")
    c.Add Split("102100373|TD,3F,13,8-0,38/0,22kV,45KVA,4TAPs(DAN)|UN|UC|TRAFO|39|5", "|")
    c.Add Split("122120033|CABO COB AL 50MM 15KV XLPE+HDPE+SMC|M|UC|COND PROT|40|5", "|")
    c.Add Split("122120006|CABO COB AL CA 50MM2 15KV XLPE CZ BLQ|M|UC|COND PROT|40|5", "|")
    c.Add Split("122120003|CABO COB AL CA 150MM2 15KV XLPE CZ(BLQ)|M|UC|COND PROT|41|5", "|")
    c.Add Split("122120026|CABO COB AL 150MM 15KV XLPE+HDPE+SMC|M|UC|COND PROT|41|5", "|")
    Set EquivPadrao = c
End Function

Private Sub GarantirEquivalencias(wb As Workbook)
    Set mEquiv = CreateObject("Scripting.Dictionary")
    Set mEquivLista = CreateObject("Scripting.Dictionary")

    Dim padrao As Collection : Set padrao = EquivPadrao()
    Dim it As Variant, maxPad As Long
    Dim jaPadrao As Object : Set jaPadrao = CreateObject("Scripting.Dictionary")
    For Each it In padrao
        jaPadrao(CStr(it(0))) = True
        If Val(it(5)) > maxPad Then maxPad = Val(it(5))
    Next it

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("EQUIVALENCIAS")
    On Error GoTo 0
    Dim ver As Long : ver = 0
    On Error Resume Next
    ver = CLng(Mid(wb.Names("EQUIV_VERSAO").RefersTo, 2))
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "EQUIVALENCIAS"
        EscreverEquivalencias ws, padrao
    ElseIf NormStr(CStr(ws.Cells(1, 1).Value)) = "GRUPO" Then
        ' Migracao unica do layout antigo (A = GRUPO, B = COD MAT, C = DESCRICAO):
        ' a aba vira o layout novo com a lista padrao; linhas que o usuario tinha
        ' acrescentado sao mantidas, renumeradas depois do ultimo grupo padrao
        ' para nao se fundirem com eles.
        Dim linhas As Collection : Set linhas = New Collection
        For Each it In padrao
            linhas.Add it
        Next it
        Dim novoGrp As Object : Set novoGrp = CreateObject("Scripting.Dictionary")
        Dim prox As Long : prox = maxPad + 1
        Dim lastO As Long : lastO = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        Dim rO As Long, gOld As String, cOld As String, dOld As String
        For rO = 2 To lastO
            gOld = "" : cOld = "" : dOld = ""
            If Not IsError(ws.Cells(rO, 1).Value) Then gOld = Trim(CStr(ws.Cells(rO, 1).Value))
            If Not IsError(ws.Cells(rO, 2).Value) Then cOld = NormCod(ws.Cells(rO, 2).Value)
            If Not IsError(ws.Cells(rO, 3).Value) Then dOld = Trim(CStr(ws.Cells(rO, 3).Value))
            If gOld <> "" And cOld <> "" And Not jaPadrao.Exists(cOld) Then
                If Not novoGrp.Exists(gOld) Then novoGrp.Add gOld, CStr(prox) : prox = prox + 1
                linhas.Add Array(cOld, dOld, "", "", "", novoGrp(gOld), "0")
            End If
        Next rO
        ws.Cells.Clear
        EscreverEquivalencias ws, linhas
        LogAdd "EQUIVALENCIAS: aba migrada para o layout novo (" & linhas.Count & " codigo(s))"
    ElseIf ver < EQUIV_VERSAO_ATUAL Then
        ' aba sem marcador = criada pela versao 1 da lista
        AtualizarEquivalencias ws, padrao, IIf(ver = 0, 1, ver)
    End If
    On Error Resume Next
    wb.Names.Add Name:="EQUIV_VERSAO", RefersTo:="=" & EQUIV_VERSAO_ATUAL, Visible:=False
    On Error GoTo 0

    OcultarAba ws      ' aba de apoio: some das guias (Reexibir para editar)

    ' ---- leitura: colunas pelo cabecalho ----
    Dim cCod As Long, cGrp As Long
    ColunasEquiv ws, cCod, cGrp
    If cCod = 0 Or cGrp = 0 Then
        AvisoExec "EQUIVALENCIAS: cabecalho sem as colunas COD MAT e EQUIVALENCIA - " & _
                  "nenhuma equivalencia aplicada"
        Exit Sub
    End If

    ' Codigo em dois grupos une os grupos (union-find por numero de grupo):
    ' 102310194 nos grupos 2 e 11 deixa os tres TPs equivalentes entre si.
    Dim pai As Object : Set pai = CreateObject("Scripting.Dictionary")   ' grupo -> grupo pai
    Dim codG As Object : Set codG = CreateObject("Scripting.Dictionary")  ' cod -> 1o grupo
    Dim lastR As Long : lastR = ws.Cells(ws.Rows.Count, cCod).End(xlUp).Row
    Dim r As Long, grp As String, cod As String, ga As String, gb As String
    For r = 2 To lastR
        grp = "" : cod = ""
        If Not IsError(ws.Cells(r, cGrp).Value) Then grp = Trim(CStr(ws.Cells(r, cGrp).Value))
        If Not IsError(ws.Cells(r, cCod).Value) Then cod = NormCod(ws.Cells(r, cCod).Value)
        If grp <> "" And cod <> "" Then
            If Not pai.Exists(grp) Then pai.Add grp, grp
            If codG.Exists(cod) Then
                ga = RaizGrupo(pai, CStr(codG(cod))) : gb = RaizGrupo(pai, grp)
                If ga <> gb Then
                    pai(gb) = ga
                    LogAdd "EQUIVALENCIAS: grupos " & ga & " e " & gb & " unidos (codigo " & cod & " nos dois)"
                End If
            Else
                codG.Add cod, grp
            End If
        End If
    Next r
    Dim k As Variant, gR As String
    For Each k In codG.Keys
        gR = RaizGrupo(pai, CStr(codG(k)))
        mEquiv.Add k, gR
        If mEquivLista.Exists(gR) Then
            mEquivLista(gR) = mEquivLista(gR) & " = " & k
        Else
            mEquivLista.Add gR, CStr(k)
        End If
    Next k
    LogAdd "EQUIVALENCIAS: " & mEquiv.Count & " codigo(s) em " & mEquivLista.Count & " grupo(s)"
End Sub

' Grupo representante (raiz) no union-find de grupos de equivalencia.
Private Function RaizGrupo(pai As Object, ByVal g As String) As String
    Do While CStr(pai(g)) <> g
        g = CStr(pai(g))
    Loop
    RaizGrupo = g
End Function

' Colunas COD MAT e EQUIVALENCIA da aba EQUIVALENCIAS (0 quando ausentes).
Private Sub ColunasEquiv(ws As Worksheet, ByRef cCod As Long, ByRef cGrp As Long)
    Dim c As Long, hN As String
    cCod = 0 : cGrp = 0
    For c = 1 To 20
        If Not IsError(ws.Cells(1, c).Value) Then
            hN = NormStr(CStr(ws.Cells(1, c).Value))
            If hN = "COD MAT" And cCod = 0 Then cCod = c
            If hN = "EQUIVALENCIA" And cGrp = 0 Then cGrp = c
        End If
    Next c
End Sub

' Acrescenta numa aba EQUIVALENCIAS ja existente os grupos da lista padrao de
' versao mais nova que a da aba. Grupo com todos os codigos ja cadastrados e
' pulado; numero de grupo ja usado por outros codigos vira o proximo livre.
Private Sub AtualizarEquivalencias(ws As Worksheet, padrao As Collection, ByVal verAba As Long)
    Dim cCod As Long, cGrp As Long
    ColunasEquiv ws, cCod, cGrp
    If cCod = 0 Or cGrp = 0 Then Exit Sub

    ' o que a aba ja tem: codigos e codigos por grupo
    Dim temCod As Object : Set temCod = CreateObject("Scripting.Dictionary")
    Dim grpCods As Object : Set grpCods = CreateObject("Scripting.Dictionary")
    Dim lastR As Long : lastR = ws.Cells(ws.Rows.Count, cCod).End(xlUp).Row
    Dim r As Long, g As String, cod As String, maxG As Long
    For r = 2 To lastR
        g = "" : cod = ""
        If Not IsError(ws.Cells(r, cGrp).Value) Then g = Trim(CStr(ws.Cells(r, cGrp).Value))
        If Not IsError(ws.Cells(r, cCod).Value) Then cod = NormCod(ws.Cells(r, cCod).Value)
        If cod <> "" Then temCod(cod) = True
        If g <> "" Then
            If Not grpCods.Exists(g) Then grpCods.Add g, CreateObject("Scripting.Dictionary")
            If cod <> "" Then grpCods(g).Item(cod) = True
            If Val(g) > maxG Then maxG = Val(g)
        End If
    Next r

    ' O par 105300003 = 105300005 muda de numero quando a tabela da area passa a
    ' usar o dele (6 na v1, 13 na v2, 23 na v3). Se o grupo dele na aba ainda e exatamente
    ' esse par e esta fora do numero atual da lista padrao, vai para esse numero
    ' (ou para o proximo livre, se o usuario ja o usa).
    Dim itP As Variant, gAlvo As String, gPar As Variant, gNovo As String
    For Each itP In padrao
        If CStr(itP(0)) = "105300003" Then gAlvo = CStr(itP(5))
    Next itP
    For Each gPar In grpCods.Keys
        If CStr(gPar) <> gAlvo And grpCods(gPar).Count = 2 Then
            If grpCods(gPar).Exists("105300003") And grpCods(gPar).Exists("105300005") Then
                gNovo = gAlvo
                If grpCods.Exists(gNovo) Then maxG = maxG + 1 : gNovo = CStr(maxG)
                For r = 2 To lastR
                    If Not IsError(ws.Cells(r, cGrp).Value) Then
                        If Trim(CStr(ws.Cells(r, cGrp).Value)) = CStr(gPar) Then ws.Cells(r, cGrp).Value = CLng(gNovo)
                    End If
                Next r
                grpCods.Add gNovo, grpCods(gPar)
                grpCods.Remove gPar
                If Val(gNovo) > maxG Then maxG = Val(gNovo)
                LogAdd "EQUIVALENCIAS: par 105300003 = 105300005 movido do grupo " & gPar & " para o " & gNovo
                Exit For
            End If
        End If
    Next gPar

    ' grupos novos (versao > aba), na ordem da lista padrao
    Dim novos As Object : Set novos = CreateObject("Scripting.Dictionary")   ' grupo -> Collection
    Dim it As Variant
    For Each it In padrao
        If CLng(it(6)) > verAba Then
            If Not novos.Exists(CStr(it(5))) Then novos.Add CStr(it(5)), New Collection
            novos(CStr(it(5))).Add it
        End If
    Next it

    Dim gk As Variant, todos As Boolean, gDest As String, primeiro As Boolean, nAdd As Long
    For Each gk In novos.Keys
        todos = True
        For Each it In novos(gk)
            If Not temCod.Exists(CStr(it(0))) Then todos = False
        Next it
        If todos Then
            LogAdd "EQUIVALENCIAS: grupo padrao " & gk & " ja cadastrado - pulado"
        Else
            gDest = CStr(gk)
            If grpCods.Exists(gDest) Then
                maxG = maxG + 1 : gDest = CStr(maxG)
                LogAdd "EQUIVALENCIAS: grupo padrao " & gk & " ja usado na aba - gravado como " & gDest
            End If
            If Val(gDest) > maxG Then maxG = Val(gDest)
            grpCods.Add gDest, CreateObject("Scripting.Dictionary")
            primeiro = True
            For Each it In novos(gk)
                lastR = lastR + 1
                EscreverLinhaEquiv ws, lastR, it, gDest, primeiro
                primeiro = False
                temCod(CStr(it(0))) = True
                nAdd = nAdd + 1
            Next it
        End If
    Next gk
    LogAdd "EQUIVALENCIAS: lista padrao v" & verAba & " -> v" & EQUIV_VERSAO_ATUAL & _
           ", " & nAdd & " codigo(s) acrescentado(s)"
End Sub

' Escreve a aba EQUIVALENCIAS no layout da tabela da area. linhas: cada item e
' (COD, DESC, UND, TIPO, FAMILIA, EQUIVALENCIA, VERSAO); "PAR" marca a 1a
' linha de cada grupo, como na tabela de origem.
Private Sub EscreverEquivalencias(ws As Worksheet, linhas As Collection)
    Dim h As Variant
    h = Array("COD MAT", "DESC MAT", "UND", "TIPO", "FAMILIA", "EQUIVALENCIA", "PAR DE EQUIVALENCIA")
    Dim c As Long
    For c = 1 To 7
        ws.Cells(1, c).Value = h(c - 1)
    Next c
    With ws.Range("A1:G1")
        .Font.Bold = True : .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(17, 24, 39) : .HorizontalAlignment = xlCenter
    End With
    Dim r As Long : r = 1
    Dim it As Variant, grpAnt As String
    For Each it In linhas
        r = r + 1
        EscreverLinhaEquiv ws, r, it, CStr(it(5)), (CStr(it(5)) <> grpAnt)
        grpAnt = CStr(it(5))
    Next it
    ws.Cells(1, 9).Value = "Codigos com o mesmo numero em EQUIVALENCIA sao equivalentes: na ANALISE " & _
        "SAP x PRJ, quando 2+ codigos do grupo aparecem no mesmo PEP4, todos ficam ADERENTE."
    ws.Cells(1, 9).Font.Italic = True : ws.Cells(1, 9).Font.Color = RGB(110, 118, 130)
    ws.Columns(1).ColumnWidth = 13 : ws.Columns(2).ColumnWidth = 44
    ws.Columns(3).ColumnWidth = 6 : ws.Columns(4).ColumnWidth = 6
    ws.Columns(5).ColumnWidth = 14 : ws.Columns(6).ColumnWidth = 14
    ws.Columns(7).ColumnWidth = 22
    ws.Tab.Color = RGB(90, 90, 90)
End Sub

' Uma linha da aba EQUIVALENCIAS (codigo como texto, sem notacao cientifica).
Private Sub EscreverLinhaEquiv(ws As Worksheet, ByVal r As Long, it As Variant, _
                               ByVal grp As String, ByVal marcaPar As Boolean)
    Dim c As Long
    ws.Cells(r, 1).NumberFormat = "@"
    ws.Cells(r, 1).Value = CStr(it(0))
    For c = 1 To 4
        ws.Cells(r, c + 1).Value = it(c)
    Next c
    If IsNumeric(grp) Then ws.Cells(r, 6).Value = CLng(grp) Else ws.Cells(r, 6).Value = grp
    If marcaPar Then ws.Cells(r, 7).Value = "PAR"
    ws.Range(ws.Cells(r, 6), ws.Cells(r, 7)).HorizontalAlignment = xlCenter
End Sub

' Chave "PEP4|GRUPO" da linha da base quando o material esta em um grupo de
' equivalentes; "" quando nao esta (ou falta PEP4 / COD MAT).
Private Function ChaveEquiv(baseData As Variant, ByVal i As Long, _
                            ByVal cP4 As Long, ByVal cCod As Long) As String
    If mEquiv Is Nothing Then Exit Function
    If mEquiv.Count = 0 Or cP4 = 0 Or cCod = 0 Then Exit Function
    If IsError(baseData(i, cCod)) Or IsError(baseData(i, cP4)) Then Exit Function
    Dim cod As String : cod = NormCod(baseData(i, cCod))
    If cod = "" Then Exit Function
    If Not mEquiv.Exists(cod) Then Exit Function
    Dim p4 As String : p4 = NormStr(CStr(baseData(i, cP4)))
    If p4 = "" Then Exit Function
    ChaveEquiv = p4 & "|" & mEquiv(cod)
End Function

' SIT MAT de um grupo de EQUIVALENTES no PEP4, pela quantidade TOTAL: exige o
' MESMO numero no SAP e no PRJ, sem tolerancia. SAP 3 num codigo e PRJ 9 no
' outro = NAO ADERENTE; SAP 9 x PRJ 10 tambem (a tolerancia de 10% das demais
' linhas nao vale para equivalentes); SAP 9 x PRJ 9 = ADERENTE.
Private Function SitPorQuantidade(ByVal libV As Variant, ByVal prjV As Variant) As String
    SitPorQuantidade = IIf(Abs(Val0(libV) - Val0(prjV)) < 0.000001, "ADERENTE", "NAO ADERENTE")
End Function

' Aplica a regra do mesmo numero depois de AvaliarLinha: familias de cabo /
' condutor / ramal sao avaliadas por diferenca absoluta ou % (EhAderente) e
' ignorariam a SIT MAT recalculada. UC zerado e regras de OD prevalecem.
Private Sub AjustarAderEquiv(av As tAvalLinha, ByVal libT As Variant, ByVal prjT As Variant)
    If av.UcZerado Or av.ViolaOD Then Exit Sub
    av.EhAder = (SitPorQuantidade(libT, prjT) = "ADERENTE")
    av.ReprovaLinha = Not av.EhAder
End Sub

' "105000058 CH SEC 1P VHI ... = 105000029 CH SEC BYPAS ...": os codigos do
' grupo presentes no PEP4, cada um com a propria descricao da base.
Private Function TextoEquiv(d As Object) As String
    Dim k As Variant, t As String
    For Each k In d.Keys
        t = t & IIf(t <> "", " = ", "") & CStr(k)
        If CStr(d(k)) <> "" Then t = t & " " & CStr(d(k))
    Next k
    TextoEquiv = t
End Function

' ===========================================================================
'  CARREGAR BASES: importa as extracoes de analise de carga e roda a analise
' ===========================================================================

' Macro de entrada (Alt+F8 > CarregarBases, ou num botao). Escolhe um ou mais
' arquivos de base (xlsx/xlsm/xls/xlsb/csv), traz os dados para a aba de base
' deste arquivo e roda GerarInventario. Com varios arquivos pergunta:
'   SIM = JUNTAR: uma analise com todas as bases (coluna ARQUIVO ORIGEM)
'   NAO = SEPARAR: uma analise por base, cada uma salva ao lado do arquivo
'         como "ANALISE INVENTARIO - <base>.xlsx", com resumo no final
' Fluxo: LER (so leitura) -> PREVIA (nada alterado; OK/Cancelar) -> GRAVAR.
Public Sub CarregarBases()
    Dim wbDest As Workbook : Set wbDest = ActiveWorkbook

    Dim fd As Object : Set fd = Application.FileDialog(3)    ' 3 = msoFileDialogFilePicker
    With fd
        .Title = "Selecione a(s) base(s) de analise de inventario"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Bases de inventario (Excel / CSV)", "*.xlsx;*.xlsm;*.xls;*.xlsb;*.csv"
        If .Show <> -1 Then Exit Sub
    End With
    Dim arqs As Collection : Set arqs = New Collection
    Dim it As Variant
    For Each it In fd.SelectedItems
        If LCase(CStr(it)) <> LCase(wbDest.FullName) Then arqs.Add CStr(it)
    Next it
    If arqs.Count = 0 Then
        MsgBox "Nenhum arquivo de base selecionado (o proprio arquivo de analise nao conta).", _
               vbExclamation, "Carregar bases"
        Exit Sub
    End If

    Dim juntar As Boolean : juntar = True
    If arqs.Count > 1 Then
        Select Case MsgBox(arqs.Count & " bases selecionadas." & vbCrLf & vbCrLf & _
                "SIM = JUNTAR todas numa unica analise" & vbCrLf & _
                "NAO = SEPARAR: uma analise por base, salva ao lado de cada arquivo" & vbCrLf & vbCrLf & _
                "Antes de alterar qualquer coisa aparece uma previa da carga.", _
                vbYesNoCancel + vbQuestion, "Carregar bases")
            Case vbYes: juntar = True
            Case vbNo: juntar = False
            Case Else: Exit Sub
        End Select
    End If

    ' 1) LER: so leitura, nada muda neste arquivo
    Dim msgL As String
    Dim blocos As Collection : Set blocos = LerBases(arqs, msgL)
    If blocos.Count = 0 Then
        MsgBox "Nenhuma base foi lida:" & msgL, vbExclamation, "Carregar bases"
        Exit Sub
    End If

    ' 2) PREVIA (dry-run): o que vai entrar, por arquivo, antes de substituir a base
    If MsgBox(PreviaCarga(wbDest, blocos, msgL, juntar), vbOKCancel + vbQuestion, _
              "Carregar bases - previa") <> vbOK Then Exit Sub

    ' 3) GRAVAR e analisar
    If juntar Then
        Dim nLin As Long
        If Not GravarBase(wbDest, blocos, nLin) Then
            MsgBox "Falha ao gravar a base neste arquivo.", vbCritical, "Carregar bases"
            Exit Sub
        End If
        wbDest.Activate
        GerarInventario
    Else
        ProcessarLoteSeparado wbDest, blocos, msgL
    End If
End Sub

' Modo SEPARAR: para cada base ja lida, grava, roda GerarInventario em modo lote
' e exporta o resultado; no fim, um resumo com o que deu certo e o que falhou.
Private Sub ProcessarLoteSeparado(wbDest As Workbook, blocos As Collection, ByVal msgLeitura As String)
    Dim bl As Variant, um As Collection, resumo As String, nOk As Long
    Dim nL As Long, destino As String
    mModoLote = True
    On Error GoTo Falha
    For Each bl In blocos
        Set um = New Collection : um.Add bl
        nL = 0
        If GravarBase(wbDest, um, nL) Then
            wbDest.Activate
            GerarInventario
            If mUltimoOk Then
                destino = ExportarResultado(wbDest, CStr(bl(3)))
                resumo = resumo & vbCrLf & "  OK    " & bl(1) & "  ->  " & NomeArquivo(destino) & _
                         IIf(mFalhas <> "", "  (com avisos - ver SAUDE DA EXECUCAO no painel)", "")
                nOk = nOk + 1
            Else
                resumo = resumo & vbCrLf & "  ERRO  " & bl(1) & ": " & Replace(mAvisoLote, vbCrLf, " ")
            End If
        Else
            resumo = resumo & vbCrLf & "  ERRO  " & bl(1) & ": falha ao gravar a base"
        End If
    Next bl
    mModoLote = False
    If msgLeitura <> "" Then resumo = resumo & vbCrLf & vbCrLf & "Nao lidos:" & msgLeitura
    MsgBox nOk & " de " & blocos.Count & " base(s) analisada(s)." & vbCrLf & resumo, _
           IIf(nOk = blocos.Count And msgLeitura = "", vbInformation, vbExclamation), "Carregar bases"
    Exit Sub
Falha:
    mModoLote = False
    Application.ScreenUpdating = True : Application.EnableEvents = True
    Application.DisplayAlerts = True : Application.Calculation = xlCalculationAutomatic
    MsgBox "Erro " & Err.Number & " na carga em lote: " & Err.Description & vbCrLf & vbCrLf & _
           "Processadas ate aqui:" & resumo, vbCritical, "Carregar bases"
End Sub

' LER: abre cada arquivo so para leitura (sem eventos), acha a aba da base e
' le a tabela. Cada base valida vira Array(dados, nome do arquivo, nome da aba,
' caminho). msg recebe o que nao deu para ler. Nada muda neste arquivo.
Private Function LerBases(arqs As Collection, ByRef msg As String) As Collection
    Dim blocos As Collection : Set blocos = New Collection
    Dim it As Variant, wbS As Workbook, wsS As Worksheet, jaAberto As Boolean
    Dim d As Variant, nomeA As String
    Dim evOld As Boolean : evOld = Application.EnableEvents
    On Error GoTo Falha
    Application.ScreenUpdating = False
    Application.EnableEvents = False        ' sem macros de abertura dos arquivos de base

    For Each it In arqs
        nomeA = NomeArquivo(CStr(it))
        ' ja aberto (mesmo caminho): usa e NAO fecha; senao abre so para leitura
        Set wbS = Nothing : jaAberto = False
        On Error Resume Next
        Set wbS = Workbooks(nomeA)
        On Error GoTo Falha
        If Not wbS Is Nothing Then
            If LCase(wbS.FullName) = LCase(CStr(it)) Then jaAberto = True Else Set wbS = Nothing
        End If
        If wbS Is Nothing Then
            On Error Resume Next
            If LCase(Right(CStr(it), 4)) = ".csv" Then
                Set wbS = Workbooks.Open(CStr(it), ReadOnly:=True, Local:=True)   ' ; e virgula decimal
            Else
                Set wbS = Workbooks.Open(CStr(it), ReadOnly:=True, UpdateLinks:=0)
            End If
            On Error GoTo Falha
        End If

        If wbS Is Nothing Then
            msg = msg & vbCrLf & "  - " & nomeA & ": nao foi possivel abrir"
        Else
            Set wsS = AcharBaseInventario(wbS)
            If wsS Is Nothing Then
                msg = msg & vbCrLf & "  - " & nomeA & ": sem as colunas MAT LIB SAP / ATENDIDO SAP e " & _
                      "MAT PRJ CAD / TOMBADO PROJ"
            Else
                d = LerTabela(wsS)
                If IsEmpty(d) Then
                    msg = msg & vbCrLf & "  - " & nomeA & ": aba '" & wsS.Name & "' sem linhas de dados"
                Else
                    blocos.Add Array(d, nomeA, wsS.Name, CStr(it))
                End If
            End If
            If Not jaAberto Then wbS.Close SaveChanges:=False
        End If
    Next it
    Application.EnableEvents = evOld
    Set LerBases = blocos
    Exit Function
Falha:
    Application.EnableEvents = evOld
    msg = msg & vbCrLf & "  - erro " & Err.Number & " na leitura: " & Err.Description
    Set LerBases = blocos
End Function

' Colunas de destino: o cabecalho do 1o bloco, sem ARQUIVO ORIGEM (que vai
' sempre no fim) e sem coluna de cabecalho vazio.
Private Sub CabecalhoDestino(blocos As Collection, ByRef hdr() As String, ByRef nCol As Long)
    Dim b1 As Variant : b1 = blocos(1)
    Dim d As Variant : d = b1(0)
    Dim c As Long
    nCol = 0
    ReDim hdr(1 To UBound(d, 2))
    For c = 1 To UBound(d, 2)
        If Not IsError(d(1, c)) Then
            If Trim(CStr(d(1, c))) <> "" And NormStr(CStr(d(1, c))) <> "ARQUIVO ORIGEM" Then
                nCol = nCol + 1 : hdr(nCol) = Trim(CStr(d(1, c)))
            End If
        End If
    Next c
End Sub

' Casa as colunas de um bloco com o cabecalho de destino, por NOME (com os
' aliases antigos/novos): src(c) = coluna no bloco, 0 = ausente. faltam recebe
' os nomes ausentes, separados por virgula.
Private Function MapearBloco(d As Variant, hdr() As String, ByVal nCol As Long, _
                             ByRef faltam As String) As Long()
    Dim c As Long, k As String
    Dim hS() As Variant : ReDim hS(1 To 1, 1 To UBound(d, 2))
    For c = 1 To UBound(d, 2)
        hS(1, c) = d(1, c)
    Next c
    Dim mS As Object : Set mS = MapaDeCabecalho(hS)
    Dim src() As Long : ReDim src(1 To nCol)
    faltam = ""
    For c = 1 To nCol
        k = NormStr(hdr(c))
        If mS.Exists(k) Then
            src(c) = mS(k)
        ElseIf mS.Exists(CanonicoBase(k)) Then
            src(c) = mS(CanonicoBase(k))
        Else
            faltam = faltam & IIf(faltam <> "", ", ", "") & hdr(c)
        End If
    Next c
    MapearBloco = src
End Function

' PREVIA (dry-run): texto com o que a carga vai fazer - linhas por arquivo,
' colunas ausentes (modo juntar), arquivos nao lidos e a aba que sera
' substituida. Nao altera nada. Limitado a 8 arquivos para caber na MsgBox.
Private Function PreviaCarga(wbDest As Workbook, blocos As Collection, _
                             ByVal msgL As String, ByVal juntar As Boolean) As String
    Dim t As String, hdr() As String, nCol As Long, src() As Long, faltam As String
    Dim bl As Variant, nArq As Long, nTot As Long, nLinB As Long
    CabecalhoDestino blocos, hdr, nCol
    t = "PREVIA DA CARGA - nada foi alterado ainda" & vbCrLf & vbCrLf
    For Each bl In blocos
        nArq = nArq + 1
        nLinB = UBound(bl(0), 1) - 1
        nTot = nTot + nLinB
        If nArq <= 8 Then
            t = t & "- " & bl(1) & "  [aba '" & bl(2) & "']: " & Format(nLinB, "#,##0") & " linha(s)" & vbCrLf
            If juntar Then
                src = MapearBloco(bl(0), hdr, nCol, faltam)
                If faltam <> "" Then
                    If Len(faltam) > 90 Then faltam = Left(faltam, 87) & "..."
                    t = t & "      sem: " & faltam & " (fica vazia)" & vbCrLf
                End If
            End If
        End If
    Next bl
    If nArq > 8 Then t = t & "... e mais " & (nArq - 8) & " arquivo(s)" & vbCrLf
    If msgL <> "" Then
        If Len(msgL) > 300 Then msgL = Left(msgL, 297) & "..."
        t = t & vbCrLf & "NAO serao carregados:" & msgL & vbCrLf
    End If
    If juntar Then
        t = t & vbCrLf & "Total: " & Format(nTot, "#,##0") & " linha(s) numa unica analise"
    Else
        t = t & vbCrLf & nArq & " analise(s) separada(s), salvas ao lado de cada arquivo"
    End If
    Dim wsAt As Worksheet : Set wsAt = AcharBaseInventario(wbDest)
    If Not wsAt Is Nothing Then t = t & vbCrLf & "A aba de base atual '" & wsAt.Name & "' sera SUBSTITUIDA."
    t = t & vbCrLf & vbCrLf & "OK = carregar e analisar     Cancelar = nao altera nada"
    PreviaCarga = t
End Function

' GRAVAR: junta os blocos (colunas casadas por nome) e grava na aba de base
' deste arquivo (a que AcharBaseInventario encontra; sem ela, cria "BASE
' INVENTARIO"). Ultima coluna: ARQUIVO ORIGEM.
Private Function GravarBase(wbDest As Workbook, blocos As Collection, ByRef nLin As Long) As Boolean
    Dim hdr() As String, nCol As Long, c As Long
    On Error GoTo Falha
    CabecalhoDestino blocos, hdr, nCol
    If nCol = 0 Then Exit Function
    Dim bl As Variant
    nLin = 0
    For Each bl In blocos
        nLin = nLin + UBound(bl(0), 1) - 1
    Next bl

    Dim outD() As Variant : ReDim outD(1 To nLin + 1, 1 To nCol + 1)
    For c = 1 To nCol
        outD(1, c) = hdr(c)
    Next c
    outD(1, nCol + 1) = "ARQUIVO ORIGEM"
    Dim d As Variant, src() As Long, faltam As String, cOrig As Long
    Dim r As Long, rOut As Long, hO() As Variant, mO As Object
    rOut = 1
    For Each bl In blocos
        d = bl(0)
        src = MapearBloco(d, hdr, nCol, faltam)
        ' base que ja veio de uma juncao anterior mantem a origem dela
        ReDim hO(1 To 1, 1 To UBound(d, 2))
        For c = 1 To UBound(d, 2)
            hO(1, c) = d(1, c)
        Next c
        Set mO = MapaDeCabecalho(hO)
        cOrig = 0 : If mO.Exists("ARQUIVO ORIGEM") Then cOrig = mO("ARQUIVO ORIGEM")
        For r = 2 To UBound(d, 1)
            rOut = rOut + 1
            For c = 1 To nCol
                If src(c) > 0 Then outD(rOut, c) = d(r, src(c))
            Next c
            outD(rOut, nCol + 1) = bl(1)
            If cOrig > 0 Then
                If Not IsError(d(r, cOrig)) Then
                    If Trim(CStr(d(r, cOrig))) <> "" Then outD(rOut, nCol + 1) = d(r, cOrig)
                End If
            End If
        Next r
    Next bl

    Dim wsB As Worksheet : Set wsB = AcharBaseInventario(wbDest)
    If wsB Is Nothing Then
        On Error Resume Next
        Set wsB = wbDest.Worksheets("BASE INVENTARIO")
        On Error GoTo Falha
        If wsB Is Nothing Then
            Set wsB = wbDest.Worksheets.Add(Before:=wbDest.Worksheets(1))
            wsB.Name = "BASE INVENTARIO"
        End If
    End If
    If wsB.AutoFilterMode Then wsB.AutoFilterMode = False
    wsB.Cells.Clear
    wsB.Range(wsB.Cells(1, 1), wsB.Cells(nLin + 1, nCol + 1)).Value = outD
    wsB.Rows(1).Font.Bold = True
    wsB.Range(wsB.Columns(1), wsB.Columns(nCol + 1)).ColumnWidth = 14
    GravarBase = True
    Exit Function
Falha:
    GravarBase = False
End Function

' Tabela da aba (linha 1 = cabecalho) como array 2D; Empty sem linhas de dados.
' Find com xlFormulas acha a ultima celula mesmo em linhas ocultas por filtro.
Private Function LerTabela(ws As Worksheet) As Variant
    Dim uR As Range, uC As Range
    Set uR = ws.Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious)
    Set uC = ws.Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious)
    If uR Is Nothing Or uC Is Nothing Then Exit Function
    If uR.Row < 2 Or uC.Column < 2 Then Exit Function
    LerTabela = ws.Range(ws.Cells(1, 1), ws.Cells(uR.Row, uC.Column)).Value
End Function

' Copia as abas de resultado para um arquivo novo, salvo ao lado da base:
' "<pasta da base>\ANALISE INVENTARIO - <nome da base>.xlsx" (sobrescreve).
Private Function ExportarResultado(wb As Workbook, ByVal arqBase As String) As String
    Dim nomes As Variant
    nomes = Array("PAINEL DO GESTOR", "ANALISE SAP x PRJ", "RESUMO SAP x PRJ", "RACIONALIZACAO COM", _
                  "PRECO UNITARIO", "ALERTA CRITICO", "RANKING DE RISCO", "DIAGNOSTICO BASE")
    Dim tem() As String, n As Long, i As Long, ws As Worksheet
    ReDim tem(0 To UBound(nomes))
    For i = 0 To UBound(nomes)
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Worksheets(CStr(nomes(i)))
        On Error GoTo 0
        If Not ws Is Nothing Then tem(n) = CStr(nomes(i)) : n = n + 1
    Next i
    If n = 0 Then Exit Function
    ReDim Preserve tem(0 To n - 1)

    wb.Worksheets(tem).Copy                  ' cria um arquivo novo (ativo) com as abas
    Dim wbN As Workbook : Set wbN = ActiveWorkbook
    Dim pasta As String : pasta = Left(arqBase, Len(arqBase) - Len(NomeArquivo(arqBase)))
    Dim base As String : base = NomeArquivo(arqBase)
    If InStrRev(base, ".") > 0 Then base = Left(base, InStrRev(base, ".") - 1)
    Dim destino As String : destino = pasta & "ANALISE INVENTARIO - " & base & ".xlsx"
    Application.DisplayAlerts = False
    wbN.SaveAs destino, 51                   ' 51 = xlOpenXMLWorkbook (.xlsx)
    Application.DisplayAlerts = True
    wbN.Close SaveChanges:=False
    wb.Activate
    ExportarResultado = destino
End Function

' Nome do arquivo de um caminho completo (aceita \ e /).
Private Function NomeArquivo(ByVal caminho As String) As String
    Dim p As Long : p = InStrRev(caminho, "\")
    If InStrRev(caminho, "/") > p Then p = InStrRev(caminho, "/")
    NomeArquivo = Mid(caminho, p + 1)
End Function

' --------- Mapa de colunas da base: fonte UNICA do "onde esta cada coluna" ---------
' Devolve um dicionario NOME NORMALIZADO -> indice da coluna, lendo a linha de
' cabecalho em UM unico acesso a planilha. Antes, cada procedure repetia o seu
' proprio "For c = 1 To lastC / Select Case", com nomes de variavel diferentes:
' renomear uma coluna na extracao obrigava a editar todos os pontos.
' Cabecalho repetido: vence a primeira ocorrencia.
Private Function MapaColunas(wsBase As Worksheet) As Object
    Dim lastC As Long
    lastC = wsBase.Cells(1, wsBase.Columns.Count).End(xlToLeft).Column
    Dim hd As Variant
    If lastC = 1 Then
        ReDim hd(1 To 1, 1 To 1)
        hd(1, 1) = wsBase.Cells(1, 1).Value
    Else
        hd = wsBase.Range(wsBase.Cells(1, 1), wsBase.Cells(1, lastC)).Value
    End If
    Set MapaColunas = MapaDeCabecalho(hd)
End Function

' Mapa NOME NORMALIZADO -> coluna a partir da 1a linha de um array 2D (o
' cabecalho lido de uma aba), ja com os aliases. Base de MapaColunas e da
' carga de bases (CarregarBases), que casa colunas de arquivos diferentes.
Private Function MapaDeCabecalho(ByVal hd As Variant) As Object
    Dim m As Object : Set m = CreateObject("Scripting.Dictionary")
    Dim c As Long, h As String
    For c = 1 To UBound(hd, 2)
        h = ""
        If Not IsError(hd(1, c)) Then h = NormStr(CStr(hd(1, c)))
        If h <> "" And Not m.Exists(h) Then m.Add h, c
    Next c
    AplicarAliasesBase m
    Set MapaDeCabecalho = m
End Function

' Nome canonico de um cabecalho ja normalizado ("ATENDIDO SAP" -> "MAT LIB SAP");
' nome sem alias volta igual.
Private Function CanonicoBase(ByVal hNorm As String) As String
    CanonicoBase = hNorm
    Dim al As Variant, p() As String, alts() As String, k As Long
    For Each al In AliasesBase()
        p = Split(CStr(al), "=")
        alts = Split(p(1), "|")
        For k = 0 To UBound(alts)
            If NormStr(alts(k)) = hNorm Then
                CanonicoBase = p(0)
                Exit Function
            End If
        Next k
    Next al
End Function

' Nomes de cabecalho aceitos na base alem do nome CANONICO usado pelo codigo.
' Formato "CANONICO=ALIAS1|ALIAS2", com os aliases escritos como aparecem na
' planilha (NormStr e aplicado na comparacao: "_" e acentos nao importam).
' O layout novo da extracao renomeou 7 colunas; o antigo continua aceito.
Private Function AliasesBase() As Variant
    AliasesBase = Array( _
        "COD MAT=MATERIAL", _
        "DESC MAT=DESC_MATERIAL", _
        "UND=UNIDADE", _
        "CLASSE=CLASSE_CUSTO", _
        "MAT LIB SAP=ATENDIDO SAP", _
        "MAT PRJ CAD=TOMBADO PROJ", _
        "SIT MAT=SITUACAO_DO_MATERIAL")
End Function

' Completa o mapa de colunas com os nomes canonicos: se a base nao tem o
' canonico mas tem um alias, o canonico aponta para a coluna do alias. Assim
' todo o modulo segue usando ColIdx(mapa, "MAT LIB SAP") com qualquer layout.
' Se a base tiver os dois nomes, vale o canonico.
Private Sub AplicarAliasesBase(m As Object)
    Dim al As Variant, p() As String, alts() As String, k As Long, a As String
    For Each al In AliasesBase()
        p = Split(CStr(al), "=")
        If Not m.Exists(p(0)) Then
            alts = Split(p(1), "|")
            For k = 0 To UBound(alts)
                a = NormStr(alts(k))
                If m.Exists(a) Then m.Add p(0), m(a) : Exit For
            Next k
        End If
    Next al
End Sub

' Nome da coluna para mensagens ao usuario: "MAT LIB SAP (ou ATENDIDO SAP)".
Private Function NomeColunaBase(ByVal canonico As String) As String
    NomeColunaBase = canonico
    Dim al As Variant, p() As String
    For Each al In AliasesBase()
        p = Split(CStr(al), "=")
        If p(0) = canonico Then
            NomeColunaBase = canonico & " (ou " & Replace(p(1), "|", " / ") & ")"
            Exit Function
        End If
    Next al
End Function

' Indice da coluna pelo nome; 0 quando a coluna nao existe na base (os chamadores
' ja tratam 0 como "coluna ausente").
Private Function ColIdx(mapa As Object, nome As String) As Long
    If mapa Is Nothing Then Exit Function
    If mapa.Exists(nome) Then ColIdx = mapa(nome)
End Function

' --------- Validacao da base: lista todas as colunas obrigatorias ausentes ---------
Private Function ValidarBase(wsBase As Worksheet) As String
    Dim req As Variant
    req = Array("PEP3NIVEL", "PEP4NIVEL", "COD MAT", "VALOR", "FAMILIA", "TIPO", _
                "SIT MAT", "MAT LIB SAP", "MAT PRJ CAD")
    Dim have As Object : Set have = MapaColunas(wsBase)
    Dim falta As String, i As Long
    For i = 0 To UBound(req)
        If Not have.Exists(CStr(req(i))) Then
            falta = falta & IIf(falta <> "", ", ", "") & NomeColunaBase(CStr(req(i)))
        End If
    Next i
    ValidarBase = falta
End Function

' extrai o nome da familia do texto "Devolvido por divergencia de X"
Private Function MotivoFamilia(s As String) As String
    Dim pre As String : pre = "Devolvido por divergencia de "
    If Left(s, Len(pre)) = pre Then MotivoFamilia = Mid(s, Len(pre) + 1) Else MotivoFamilia = s
End Function

' --------- RESUMO por PEP3: 1 linha por obra (le a ANALISE ja gerada) ---------
Private Sub ProcessarResumoPEP3(wb As Workbook, wsDet As Worksheet)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("RESUMO SAP x PRJ")
    On Error GoTo 0
    If ws Is Nothing Then Set ws = wb.Worksheets.Add(After:=wsDet)
    ws.Name = "RESUMO SAP x PRJ"
    ws.Cells.Clear

    Dim hdrRow As Long, rr As Long, lastScan As Long
    lastScan = wsDet.Cells(wsDet.Rows.Count, 1).End(xlUp).Row
    For rr = 1 To lastScan
        If NormStr(CStr(wsDet.Cells(rr, 1).Value)) = "PEP3NIVEL" Then hdrRow = rr : Exit For
    Next rr
    Dim ds As Long : ds = hdrRow + 1
    Dim lastD As Long : lastD = wsDet.Cells(wsDet.Rows.Count, 1).End(xlUp).Row

    Dim sit As Object : Set sit = CreateObject("Scripting.Dictionary")
    Dim vObra As Object : Set vObra = CreateObject("Scripting.Dictionary")
    Dim vNA As Object : Set vNA = CreateObject("Scripting.Dictionary")
    Dim mot As Object : Set mot = CreateObject("Scripting.Dictionary")
    Dim ods As Object : Set ods = CreateObject("Scripting.Dictionary")

    If hdrRow > 0 And lastD >= ds Then
        Dim dd As Variant : dd = wsDet.Range(wsDet.Cells(ds, 1), wsDet.Cells(lastD, 16)).Value
        Dim i As Long, p3 As String, p4 As String, apv As String, s13 As String, m16 As String, v6 As Double
        For i = 1 To UBound(dd, 1)
            p3 = Trim(CStr(dd(i, 1)))
            If p3 <> "" Then
                If Not sit.Exists(p3) Then
                    sit.Add p3, False : vObra.Add p3, 0# : vNA.Add p3, 0# : mot.Add p3, ""
                    ods.Add p3, CreateObject("Scripting.Dictionary")
                End If
                apv = UCase(CStr(dd(i, 14)))
                If InStr(apv, "REPROVADO") > 0 Then sit(p3) = True
                v6 = Val0(dd(i, 6))
                vObra(p3) = vObra(p3) + v6
                s13 = UCase(CStr(dd(i, 13)))
                If InStr(s13, "NAO ADER") > 0 Then vNA(p3) = vNA(p3) + Abs(v6)
                m16 = Trim(CStr(dd(i, 16)))
                If m16 <> "" And mot(p3) = "" Then mot(p3) = m16
                p4 = Trim(CStr(dd(i, 2)))
                If p4 <> "" Then If Not ods(p3).Exists(p4) Then ods(p3).Add p4, True
            End If
        Next i
    End If

    With ws.Range("A1:G2")
        .Merge : .Value = "RESUMO POR OBRA (PEP3)  -  1 linha por obra"
        .Font.Name = "Segoe UI Semibold" : .Font.Size = 15 : .Font.Bold = True
        .Font.Color = RGB(255, 255, 255) : .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 20 : ws.Rows(2).RowHeight = 20
    Dim hdrs As Variant
    hdrs = Array("PEP3NIVEL", "SITUACAO", "ODs (PEP4)", "VALOR OBRA", "VALOR NAO ADER", _
                 "FAMILIAS REPROVADAS", "MOTIVO DEVOLUCAO")
    Dim c As Long
    For c = 1 To 7
        With ws.Cells(3, c)
            .Value = hdrs(c - 1)
            .Font.Bold = True : .Font.Color = RGB(255, 255, 255) : .Interior.Color = RGB(31, 78, 121)
            .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous : .Borders.Color = RGB(170, 170, 170)
        End With
    Next c
    ws.Rows(3).RowHeight = 24

    Dim n As Long : n = sit.Count
    If n > 0 Then
        Dim ks() As String, scr() As Double : ReDim ks(0 To n - 1) : ReDim scr(0 To n - 1)
        Dim ix As Long : ix = 0
        Dim kk As Variant
        For Each kk In sit.Keys
            ks(ix) = CStr(kk)
            scr(ix) = IIf(sit(kk), 1000000000000#, 0#) + vNA(kk)
            ix = ix + 1
        Next kk
        Dim a As Long, b As Long, best As Long, tS As String, tD As Double
        For a = 0 To n - 2
            best = a
            For b = a + 1 To n - 1
                If scr(b) > scr(best) Then best = b
            Next b
            If best <> a Then
                tD = scr(a) : scr(a) = scr(best) : scr(best) = tD
                tS = ks(a) : ks(a) = ks(best) : ks(best) = tS
            End If
        Next a

        Dim outA() As Variant : ReDim outA(1 To n, 1 To 7)
        For ix = 0 To n - 1
            p3 = ks(ix)
            outA(ix + 1, 1) = p3
            outA(ix + 1, 2) = IIf(sit(p3), "REPROVADO", "APROVADO")
            outA(ix + 1, 3) = ods(p3).Count
            outA(ix + 1, 4) = vObra(p3)
            outA(ix + 1, 5) = vNA(p3)
            outA(ix + 1, 6) = IIf(sit(p3), MotivoFamilia(CStr(mot(p3))), "")
            outA(ix + 1, 7) = mot(p3)
        Next ix
        ws.Range(ws.Cells(4, 1), ws.Cells(3 + n, 7)).Value = outA

        Dim wds As Variant : wds = Array(26, 13, 10, 15, 16, 40, 46)
        For c = 1 To 7 : ws.Columns(c).ColumnWidth = wds(c - 1) : Next c
        With ws.Range(ws.Cells(4, 1), ws.Cells(3 + n, 7))
            .Font.Name = "Segoe UI" : .Font.Size = 9 : .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous : .Borders.Color = RGB(225, 225, 225)
        End With
        ws.Range(ws.Cells(4, 4), ws.Cells(3 + n, 5)).NumberFormat = "#,##0.00"
        ws.Range(ws.Cells(4, 2), ws.Cells(3 + n, 3)).HorizontalAlignment = xlCenter

        ' Os REPROVADO vem primeiro (score com +1e12): dois blocos contiguos,
        ' formatados de uma vez em vez de linha a linha
        Dim nRepR As Long : nRepR = 0
        For ix = 1 To n
            If CStr(outA(ix, 2)) = "REPROVADO" Then nRepR = nRepR + 1
        Next ix
        If nRepR > 0 Then
            ws.Range(ws.Cells(4, 1), ws.Cells(3 + nRepR, 7)).Interior.Color = RGB(252, 226, 228)
            With ws.Range(ws.Cells(4, 2), ws.Cells(3 + nRepR, 2))
                .Interior.Color = RGB(192, 0, 0) : .Font.Color = RGB(255, 255, 255) : .Font.Bold = True
            End With
        End If
        If nRepR < n Then
            With ws.Range(ws.Cells(4 + nRepR, 2), ws.Cells(3 + n, 2))
                .Interior.Color = RGB(198, 239, 206) : .Font.Color = RGB(0, 97, 0)
            End With
        End If

        ' Tabela estruturada (ListObject) - filtros persistentes e referencia estavel
        On Error Resume Next
        Dim lo As ListObject
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(3, 1), ws.Cells(3 + n, 7)), , xlYes)
        lo.Name = "tblResumoPEP3" : lo.TableStyle = ""
        On Error GoTo 0
    Else
        ws.Cells(4, 1).Value = "Nenhuma obra encontrada." : ws.Cells(4, 1).Font.Italic = True
    End If

    ' Botao: exporta so as devolucoes (PEP3 reprovados) para envio a projetista
    ' Best-effort (so formatacao visual): falha aqui nao altera a analise.
    On Error Resume Next
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Cells(1, 9).Left, ws.Cells(1, 9).Top, 170, 34)
    btn.Name = "btnExportarDevol"
    btn.TextFrame2.TextRange.Text = "Exportar devolucoes"
    btn.Fill.ForeColor.RGB = RGB(192, 0, 0) : btn.Line.Visible = msoFalse
    btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    btn.TextFrame2.TextRange.Font.Bold = msoTrue
    btn.OnAction = "ExportarDevolucoes"
    ws.Tab.Color = RGB(31, 78, 121)
    On Error GoTo 0
End Sub

' --------- Exporta os PEP3 reprovados p/ nova planilha (botao / Alt+F8) ---------
Public Sub ExportarDevolucoes()
    Dim wb As Workbook : Set wb = ActiveWorkbook
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("RESUMO SAP x PRJ")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Aba 'RESUMO SAP x PRJ' nao encontrada. Rode GerarInventario primeiro.", _
               vbExclamation, "Exportar Devolucoes"
        Exit Sub
    End If
    Dim lastR As Long : lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 4 Then
        MsgBox "Sem dados na RESUMO SAP x PRJ.", vbInformation, "Exportar Devolucoes"
        Exit Sub
    End If
    Dim d As Variant : d = ws.Range(ws.Cells(4, 1), ws.Cells(lastR, 7)).Value
    Dim i As Long, cnt As Long
    For i = 1 To UBound(d, 1)
        If UCase(Trim(CStr(d(i, 2)))) = "REPROVADO" Then cnt = cnt + 1
    Next i
    If cnt = 0 Then
        MsgBox "Nenhum PEP3 reprovado para exportar.", vbInformation, "Exportar Devolucoes"
        Exit Sub
    End If
    Dim out() As Variant : ReDim out(1 To cnt + 1, 1 To 5)
    out(1, 1) = "PEP3NIVEL" : out(1, 2) = "SITUACAO" : out(1, 3) = "VALOR NAO ADER"
    out(1, 4) = "FAMILIAS REPROVADAS" : out(1, 5) = "MOTIVO DEVOLUCAO"
    Dim rw As Long : rw = 1
    For i = 1 To UBound(d, 1)
        If UCase(Trim(CStr(d(i, 2)))) = "REPROVADO" Then
            rw = rw + 1
            out(rw, 1) = d(i, 1) : out(rw, 2) = d(i, 2) : out(rw, 3) = d(i, 5)
            out(rw, 4) = d(i, 6) : out(rw, 5) = d(i, 7)
        End If
    Next i
    Dim nb As Workbook : Set nb = Workbooks.Add
    Dim ns As Worksheet : Set ns = nb.Worksheets(1) : ns.Name = "DEVOLUCOES"
    ns.Range(ns.Cells(1, 1), ns.Cells(cnt + 1, 5)).Value = out
    With ns.Range("A1:E1")
        .Font.Bold = True : .Interior.Color = RGB(192, 0, 0) : .Font.Color = RGB(255, 255, 255)
    End With
    ns.Range(ns.Cells(2, 3), ns.Cells(cnt + 1, 3)).NumberFormat = "#,##0.00"
    ns.Columns("A:E").AutoFit : ns.Columns(5).ColumnWidth = 55
    MsgBox cnt & " PEP3 reprovado(s) exportado(s) para uma nova planilha." & vbCrLf & _
           "Use 'Salvar Como' para enviar a projetista.", vbInformation, "Exportar Devolucoes"
End Sub

' --------- Historico entre execucoes (aba persistente HISTORICO) ---------
' ===========================================================================
'  DIAGNOSTICO DA BASE: mede a qualidade dos DADOS antes de confiar na
'  analise (garbage in, garbage out). Nota de saude 0-100 + verificacoes
'  com status OK / ATENCAO / CRITICO e recomendacao pratica para cada uma.
' ===========================================================================
Private Sub ProcessarDiagnosticoBase(wsBase As Worksheet, wb As Workbook)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("DIAGNOSTICO BASE")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "DIAGNOSTICO BASE"
    Else
        ws.Cells.Clear
    End If

    ' ---- mapeia colunas da base (tolerante a ausencia) ----
    Dim lastC As Long : lastC = wsBase.Cells(1, wsBase.Columns.Count).End(xlToLeft).Column
    Dim iPep3 As Long, iPep4 As Long, iCod As Long, iFam As Long, iTipo As Long
    Dim iValor As Long, iLib As Long, iPrj As Long, iTipoPep As Long, iUnd As Long
    Dim c As Long
    Dim mapa As Object : Set mapa = MapaColunas(wsBase)
    iPep3    = ColIdx(mapa, "PEP3NIVEL")
    iPep4    = ColIdx(mapa, "PEP4NIVEL")
    iCod     = ColIdx(mapa, "COD MAT")
    iFam     = ColIdx(mapa, "FAMILIA")
    iTipo    = ColIdx(mapa, "TIPO")
    iValor   = ColIdx(mapa, "VALOR")
    iLib     = ColIdx(mapa, "MAT LIB SAP")
    iPrj     = ColIdx(mapa, "MAT PRJ CAD")
    iTipoPep = ColIdx(mapa, "TIPO PEP")
    iUnd     = ColIdx(mapa, "UND")

    Dim refCol As Long : refCol = iPep4 : If refCol = 0 Then refCol = 1
    Dim lastR As Long : lastR = wsBase.Cells(wsBase.Rows.Count, refCol).End(xlUp).Row
    Dim n As Long : n = 0
    Dim d As Variant
    If lastR >= 2 Then
        d = wsBase.Range(wsBase.Cells(2, 1), wsBase.Cells(lastR, lastC)).Value
        n = lastR - 1
    End If

    ' ---- coleta metricas numa unica passada ----
    Dim dP3 As Object : Set dP3 = CreateObject("Scripting.Dictionary")
    Dim dP4 As Object : Set dP4 = CreateObject("Scripting.Dictionary")
    Dim dDup As Object : Set dDup = CreateObject("Scripting.Dictionary")
    Dim nFamVazia As Long, nTipoVazio As Long, nCodVazio As Long
    Dim nQtdInval As Long, nUC As Long, nUC0 As Long, nDup As Long, nSinalInv As Long
    Dim i As Long, s As String, tp As String, p4 As String, kDup As String
    For i = 1 To n
        If iPep3 > 0 Then
            s = Trim(CStr(d(i, iPep3))) : If s <> "" And Not dP3.Exists(s) Then dP3.Add s, True
        End If
        p4 = "" : If iPep4 > 0 Then p4 = Trim(CStr(d(i, iPep4)))
        If p4 <> "" And Not dP4.Exists(p4) Then dP4.Add p4, True
        If iFam > 0 Then If Trim(CStr(d(i, iFam))) = "" Then nFamVazia = nFamVazia + 1
        tp = "" : If iTipo > 0 Then tp = NormStr(CStr(d(i, iTipo)))
        If iTipo > 0 And tp = "" Then nTipoVazio = nTipoVazio + 1
        If iCod > 0 Then If Trim(CStr(d(i, iCod))) = "" Then nCodVazio = nCodVazio + 1
        ' quantidades com texto invalido (nao vazio e nao numerico)
        If iLib > 0 Then
            s = Trim(CStr(d(i, iLib)))
            If s <> "" And Not IsNumeric(d(i, iLib)) Then nQtdInval = nQtdInval + 1
        End If
        If iPrj > 0 Then
            s = Trim(CStr(d(i, iPrj)))
            If s <> "" And Not IsNumeric(d(i, iPrj)) Then nQtdInval = nQtdInval + 1
        End If
        ' UC zerados
        If tp = "UC" Then
            nUC = nUC + 1
            Dim lb As Variant, pj As Variant
            lb = "" : pj = ""
            If iLib > 0 Then lb = d(i, iLib)
            If iPrj > 0 Then pj = d(i, iPrj)
            If Val0(lb) = 0 And Val0(pj) = 0 Then nUC0 = nUC0 + 1
        End If
        ' duplicidade PEP4 + COD MAT (mesmo material lancado 2x na mesma obra)
        If p4 <> "" And iCod > 0 Then
            kDup = p4 & "|" & Trim(CStr(d(i, iCod)))
            If Trim(CStr(d(i, iCod))) <> "" Then
                dDup(kDup) = dDup(kDup) + 1
                If dDup(kDup) = 2 Then nDup = nDup + 1
            End If
        End If
        ' sinal invertido por tipo de OD (mesma premissa de AvaliarLinha):
        ' ODD exige negativo; ODI/ODS exigem positivo - em valor E quantidade
        If p4 <> "" Then
            Dim tpOdD As String : tpOdD = ""
            If iTipoPep > 0 Then tpOdD = CStr(d(i, iTipoPep))
            Dim lbD As Variant, vlD As Variant
            lbD = 0 : vlD = 0
            If iLib > 0 Then lbD = d(i, iLib)
            If iValor > 0 Then vlD = d(i, iValor)
            If SinalInvalidoOD(ClassificarOD(p4, tpOdD), lbD, vlD) <> "" Then nSinalInv = nSinalInv + 1
        End If
    Next i

    ' ---- monta as verificacoes (nome, resultado, status, recomendacao) ----
    ' folga proposital: DiagAdd ignora verificacoes acima deste limite em vez
    ' de derrubar a macro com "Subscript out of range" no meio da geracao
    Dim chk(1 To 32, 1 To 4) As String
    Dim nChk As Long, nCrit As Long, nAten As Long

    DiagAdd chk, nChk, nCrit, nAten, "Volume da base", _
        Format(n, "#,##0") & " linhas | " & Format(dP3.Count, "#,##0") & " obras (PEP3) | " & _
        Format(dP4.Count, "#,##0") & " PEP4", "OK", "Confira se o volume bate com o esperado da extracao."

    DiagAdd chk, nChk, nCrit, nAten, "FAMILIA em branco", _
        Format(nFamVazia, "#,##0") & " linha(s)" & PctDe(nFamVazia, n), _
        StatusPct(nFamVazia, n, 0.02, 0.1), _
        "Sem FAMILIA a linha fica fora das regras de aderencia. Reclassifique na origem."

    DiagAdd chk, nChk, nCrit, nAten, "TIPO em branco (UC/COM)", _
        Format(nTipoVazio, "#,##0") & " linha(s)" & PctDe(nTipoVazio, n), _
        StatusPct(nTipoVazio, n, 0.02, 0.1), _
        "Sem TIPO o item nao entra no veredito APROVADO/REPROVADO."

    DiagAdd chk, nChk, nCrit, nAten, "COD MAT em branco", _
        Format(nCodVazio, "#,##0") & " linha(s)", _
        IIf(nCodVazio > 0, "ATENCAO", "OK"), _
        "Sem codigo nao ha cruzamento NT.006 nem faixa de preco."

    DiagAdd chk, nChk, nCrit, nAten, "Quantidade com texto invalido", _
        Format(nQtdInval, "#,##0") & " celula(s)", _
        IIf(nQtdInval > 0, "CRITICO", "OK"), _
        "Texto em coluna numerica vira 0 na analise e distorce a aderencia. Corrija na base."

    If nUC > 0 Then
        DiagAdd chk, nChk, nCrit, nAten, "Materiais UC com valores zerados", _
            Format(nUC0, "#,##0") & " de " & Format(nUC, "#,##0") & " UC" & PctDe(nUC0, nUC), _
            IIf(nUC > 0 And nUC0 = nUC, "CRITICO", IIf(nUC0 / nUC > 0.2, "ATENCAO", "OK")), _
            "Zerado isolado e ignorado; obra 100% zerada e devolvida. Muitos zerados = extracao incompleta."
    End If

    DiagAdd chk, nChk, nCrit, nAten, "Material duplicado na mesma obra", _
        Format(nDup, "#,##0") & " chave(s) PEP4+material repetida(s)", _
        IIf(nDup > 0, "ATENCAO", "OK"), _
        "Duplicidade soma quantidades 2x e pode gerar falso NAO ADERENTE. Consolide os lancamentos."

    DiagAdd chk, nChk, nCrit, nAten, "Sinal invertido por tipo de OD", _
        Format(nSinalInv, "#,##0") & " linha(s)", _
        IIf(nSinalInv > 0, "ATENCAO", "OK"), _
        "ODD exige valor e quantidade negativos; ODI/ODS, positivos. Essas linhas viram NAO ADERENTE."

    Dim faltantes As String : faltantes = ""
    If iTipoPep = 0 Then faltantes = faltantes & "TIPO PEP; "
    If iUnd = 0 Then faltantes = faltantes & "UND; "
    If iValor = 0 Then faltantes = faltantes & "VALOR; "
    DiagAdd chk, nChk, nCrit, nAten, "Colunas opcionais presentes", _
        IIf(faltantes = "", "Todas presentes", "Faltam: " & Left(faltantes, Len(faltantes) - 2)), _
        IIf(faltantes = "", "OK", "ATENCAO"), _
        "Colunas opcionais enriquecem alertas (ODI/ODM, unidade, valores em R$)."

    ' base de precos: fonte usada nesta execucao e idade do arquivo externo
    ' (PRECO UNITARIO e ALERTA CRITICO ja rodaram e registraram a fonte)
    Dim maxDias As Double : maxDias = CfgD("PRECOS_MAX_DIAS", 90)
    Dim resP As String, stP As String
    If mFontePrecos = "" Then
        resP = "Nao consultada nesta execucao" : stP = "OK"
    ElseIf Left(mFontePrecos, 7) = "NENHUMA" Or InStr(mFontePrecos, "NAO ABRIU") > 0 _
           Or InStr(mFontePrecos, "vazia") > 0 Then
        resP = mFontePrecos : stP = "ATENCAO"
    ElseIf IsEmpty(mFontePrecosData) Then
        resP = "Aba interna (sem data para conferir)" : stP = "OK"
    Else
        Dim diasP As Long : diasP = DateDiff("d", mFontePrecosData, Now)
        resP = "Arquivo alterado em " & Format(mFontePrecosData, "dd/mm/yyyy") & " (" & diasP & " dias)"
        stP = IIf(diasP > maxDias, "ATENCAO", "OK")
    End If
    DiagAdd chk, nChk, nCrit, nAten, "Base de precos atualizada", resP, stP, _
        "Faixas MIN/MAX antigas distorcem PRECO UNITARIO e os alertas de preco. Atualize a " & _
        "BASE DE PRECOS (limite: " & maxDias & " dias - CONFIG PRECOS_MAX_DIAS)."

    Dim nota As Double : nota = NotaSaude(nCrit, nAten)
    Dim rotSaude As String, corSaude As Long
    If nota >= 90 Then
        rotSaude = "EXCELENTE" : corSaude = RGB(0, 130, 62)
    ElseIf nota >= 70 Then
        rotSaude = "BOA" : corSaude = RGB(46, 117, 62)
    ElseIf nota >= 50 Then
        rotSaude = "REQUER ATENCAO" : corSaude = RGB(214, 132, 0)
    Else
        rotSaude = "CRITICA" : corSaude = RGB(176, 0, 0)
    End If

    ' ---- desenha a aba ----
    With ws.Range("A1:D2")
        .Merge
        .Value = ChrW(&H2695) & "  DIAGNOSTICO DA BASE"
        .Font.Name = "Segoe UI Semibold" : .Font.Size = 18 : .Font.Bold = True
        .Font.Color = RGB(255, 255, 255) : .Interior.Color = RGB(17, 24, 39)
        .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 22 : ws.Rows(2).RowHeight = 22
    With ws.Range("A3:D3")
        .Merge
        .Value = "Base: " & wsBase.Name & "  |  Gerado: " & Format(Now, "dd/mm/yyyy hh:nn") & _
                 "  |  A nota mede a CONFIABILIDADE dos dados, nao o resultado das obras"
        .Font.Name = "Segoe UI" : .Font.Size = 9 : .Font.Italic = True
        .Font.Color = RGB(255, 255, 255) : .Interior.Color = RGB(46, 62, 92)
        .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
    End With

    With ws.Range("A5:A7")
        .Merge
        .Value = Format(nota, "0")
        .Font.Name = "Segoe UI Black" : .Font.Size = 34 : .Font.Bold = True
        .Font.Color = RGB(255, 255, 255) : .Interior.Color = corSaude
        .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
    End With
    With ws.Range("B5:D7")
        .Merge
        .Value = "SAUDE DOS DADOS: " & rotSaude & "   (" & _
                 nCrit & " critico(s), " & nAten & " atencao(oes))"
        .Font.Name = "Segoe UI Semibold" : .Font.Size = 14 : .Font.Bold = True
        .Font.Color = corSaude : .Interior.Color = RGB(244, 247, 250)
        .HorizontalAlignment = xlLeft : .VerticalAlignment = xlCenter : .IndentLevel = 1
    End With

    Dim hdr As Variant : hdr = Array("VERIFICACAO", "RESULTADO", "STATUS", "RECOMENDACAO")
    For c = 1 To 4
        With ws.Cells(9, c)
            .Value = hdr(c - 1)
            .Font.Name = "Segoe UI Semibold" : .Font.Bold = True : .Font.Size = 10
            .Font.Color = RGB(235, 240, 248) : .Interior.Color = RGB(17, 24, 39)
            .HorizontalAlignment = xlCenter : .VerticalAlignment = xlCenter
        End With
    Next c
    ws.Rows(9).RowHeight = 24

    Dim outD() As Variant : ReDim outD(1 To nChk, 1 To 4)
    For i = 1 To nChk
        For c = 1 To 4
            outD(i, c) = chk(i, c)
        Next c
    Next i
    ws.Range(ws.Cells(10, 1), ws.Cells(9 + nChk, 4)).Value = outD

    With ws.Range(ws.Cells(10, 1), ws.Cells(9 + nChk, 4))
        .Font.Name = "Segoe UI" : .Font.Size = 10
        .VerticalAlignment = xlCenter : .WrapText = False
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Color = RGB(225, 225, 225)
    End With
    For i = 1 To nChk
        With ws.Cells(9 + i, 3)
            .HorizontalAlignment = xlCenter : .Font.Bold = True : .Font.Size = 9
            Select Case chk(i, 3)
                Case "OK":      .Interior.Color = RGB(198, 239, 206) : .Font.Color = RGB(0, 97, 0)
                Case "ATENCAO": .Interior.Color = RGB(255, 235, 156) : .Font.Color = RGB(156, 101, 0)
                Case Else:      .Interior.Color = RGB(255, 199, 206) : .Font.Color = RGB(156, 0, 6)
            End Select
        End With
        ws.Rows(9 + i).RowHeight = 20
    Next i

    ws.Columns(1).ColumnWidth = 34 : ws.Columns(2).ColumnWidth = 38
    ws.Columns(3).ColumnWidth = 12 : ws.Columns(4).ColumnWidth = 78
    ws.Tab.Color = corSaude

    ' Replica os apontamentos CRITICO/ATENCAO na aba ALERTA CRITICO
    AnexarAtencoesAlerta wb, chk, nChk

    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

' Copia para a aba ALERTA CRITICO as verificacoes do DIAGNOSTICO DA BASE que
' ficaram com status CRITICO ou ATENCAO, como um bloco no fim da lista de
' alertas (os CRITICO primeiro).
' Roda depois de ProcessarAlertaCritico, entao a aba ja esta montada:
' linhas 1-2 = banda de titulo, linha 3 = cabecalho, linha 4+ = dados.
Private Sub AnexarAtencoesAlerta(wb As Workbook, chk() As String, nChk As Long)
    Dim wsA As Worksheet
    On Error Resume Next
    Set wsA = wb.Worksheets("ALERTA CRITICO")
    On Error GoTo 0
    If wsA Is Nothing Then Exit Sub

    Dim i As Long, nAt As Long, nCr As Long
    For i = 1 To nChk
        If chk(i, 3) = "CRITICO" Then nCr = nCr + 1
        If chk(i, 3) = "ATENCAO" Then nAt = nAt + 1
    Next i
    If nCr + nAt = 0 Then Exit Sub

    ' Aba sem cabecalho = ProcessarAlertaCritico abortou antes de monta-la (o
    ' motivo vai para a mensagem final e o LOG EXECUCAO). Monta o esqueleto -
    ' titulo + cabecalho - para as linhas abaixo ficarem legiveis e filtraveis,
    ' em vez de cairem soltas numa aba vazia.
    If LinhaCabecalhoAlerta(wsA) = 0 Then
        wsA.Cells.Clear
        EscreverTituloAlerta wsA
        EscreverCabecalhoAlerta wsA, 3
    End If

    Dim r0A As Long : r0A = PrimeiraLinhaAlerta(wsA)
    Dim lastA As Long : lastA = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row
    If lastA < r0A - 1 Then lastA = r0A - 1
    ' base sem alertas: a 1a linha de dados traz so o texto informativo, que sai
    If lastA = r0A Then
        If InStr(UCase(CStr(wsA.Cells(r0A, 1).Value)), "NENHUM ALERTA") > 0 Then
            wsA.Rows(r0A).ClearContents
            wsA.Cells(r0A, 1).Font.Italic = False
            wsA.Cells(r0A, 1).Font.Color = RGB(45, 52, 64)
            lastA = r0A - 1
        End If
    End If

    ' 1 TIPO | 5 DESC MAT = verificacao | 10 REFERENCIA = resultado | 11 MOTIVO
    Dim nTot As Long : nTot = nCr + nAt
    Dim outA() As Variant : ReDim outA(1 To nTot, 1 To 11)
    Dim k As Long
    Dim passe As Long, alvo As String
    For passe = 1 To 2
        alvo = IIf(passe = 1, "CRITICO", "ATENCAO")
        For i = 1 To nChk
            If chk(i, 3) = alvo Then
                k = k + 1
                outA(k, 1) = IIf(passe = 1, "QUALIDADE DA BASE - CRITICO", "QUALIDADE DA BASE")
                outA(k, 5) = chk(i, 1)
                outA(k, 10) = chk(i, 2)
                outA(k, 11) = chk(i, 4)
            End If
        Next i
    Next passe

    Dim r1 As Long : r1 = lastA + 1
    Dim r2 As Long : r2 = lastA + nTot
    wsA.Range(wsA.Cells(r1, 1), wsA.Cells(r2, 11)).Value = outA

    With wsA.Range(wsA.Cells(r1, 1), wsA.Cells(r2, 11))
        .Font.Name = "Segoe UI" : .Font.Size = 10 : .Font.Color = RGB(45, 52, 64)
        .VerticalAlignment = xlCenter
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Color = RGB(255, 255, 255)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(255, 255, 255)
    End With
    ' bloco CRITICO em vermelho, bloco ATENCAO em ambar
    If nCr > 0 Then
        wsA.Range(wsA.Cells(r1, 1), wsA.Cells(r1 + nCr - 1, 11)).Interior.Color = RGB(252, 226, 228)
        With wsA.Range(wsA.Cells(r1, 1), wsA.Cells(r1 + nCr - 1, 1))
            .Interior.Color = RGB(192, 0, 0)
            .Font.Bold = True : .Font.Size = 9 : .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With
    End If
    If nAt > 0 Then
        wsA.Range(wsA.Cells(r1 + nCr, 1), wsA.Cells(r2, 11)).Interior.Color = RGB(255, 244, 214)
        With wsA.Range(wsA.Cells(r1 + nCr, 1), wsA.Cells(r2, 1))
            .Interior.Color = RGB(176, 124, 0)
            .Font.Bold = True : .Font.Size = 9 : .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With
    End If
    wsA.Range(wsA.Cells(r1, 5), wsA.Cells(r2, 5)).HorizontalAlignment = xlLeft
    wsA.Range(wsA.Cells(r1, 10), wsA.Cells(r2, 11)).HorizontalAlignment = xlLeft
    Dim rr As Long
    For rr = r1 To r2
        wsA.Rows(rr).RowHeight = 20
    Next rr
    If wsA.Columns(10).ColumnWidth < 24 Then wsA.Columns(10).ColumnWidth = 24

    ' refaz o autofiltro para abranger as linhas acrescentadas
    On Error Resume Next
    If wsA.AutoFilterMode Then wsA.AutoFilterMode = False
    wsA.Range(wsA.Cells(r0A - 1, 1), wsA.Cells(r2, 11)).AutoFilter
    On Error GoTo 0
End Sub

' Acrescenta uma verificacao ao quadro do diagnostico e contabiliza o status.
Private Sub DiagAdd(chk() As String, ByRef nChk As Long, ByRef nCrit As Long, _
                    ByRef nAten As Long, nome As String, resultado As String, _
                    status As String, recom As String)
    ' guarda de limite: sem isto, a 1a verificacao alem do tamanho de chk()
    ' derrubaria GerarInventario com erro 9 em tempo de execucao
    If nChk >= UBound(chk, 1) Then Exit Sub
    nChk = nChk + 1
    chk(nChk, 1) = nome : chk(nChk, 2) = resultado
    chk(nChk, 3) = status : chk(nChk, 4) = recom
    If status = "CRITICO" Then nCrit = nCrit + 1
    If status = "ATENCAO" Then nAten = nAten + 1
End Sub

' Nota de saude dos dados: 100 - 25 por CRITICO - 10 por ATENCAO (minimo 0).
Private Function NotaSaude(nCrit As Long, nAten As Long) As Double
    Dim v As Double : v = 100 - 25 * nCrit - 10 * nAten
    If v < 0 Then v = 0
    NotaSaude = v
End Function

' Status por percentual: CRITICO acima de pCrit, ATENCAO acima de pAten.
Private Function StatusPct(qtd As Long, total As Long, pAten As Double, pCrit As Double) As String
    If total <= 0 Or qtd = 0 Then StatusPct = "OK" : Exit Function
    Dim p As Double : p = qtd / total
    If p > pCrit Then
        StatusPct = "CRITICO"
    ElseIf p > pAten Then
        StatusPct = "ATENCAO"
    Else
        StatusPct = "OK"
    End If
End Function

' " (x,x%)" do percentual, ou vazio se o total for 0.
Private Function PctDe(qtd As Long, total As Long) As String
    If total > 0 Then PctDe = " (" & Format(qtd / total, "0.0%") & ")" Else PctDe = ""
End Function

Private Sub AtualizarHistorico(wb As Workbook)
    Dim wsR As Worksheet
    On Error Resume Next
    Set wsR = wb.Worksheets("RESUMO SAP x PRJ")
    On Error GoTo 0
    If wsR Is Nothing Then Exit Sub
    Dim lastR As Long : lastR = wsR.Cells(wsR.Rows.Count, 1).End(xlUp).Row
    If lastR < 4 Then Exit Sub
    Dim d As Variant : d = wsR.Range(wsR.Cells(4, 1), wsR.Cells(lastR, 7)).Value

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("HISTORICO")
    On Error GoTo 0
    Dim prev As Object : Set prev = CreateObject("Scripting.Dictionary")
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "HISTORICO"
        ws.Cells(1, 1).Value = "RUN" : ws.Cells(1, 2).Value = "PEP3NIVEL"
        ws.Cells(1, 3).Value = "SITUACAO" : ws.Cells(1, 4).Value = "VALOR NAO ADER"
        ws.Cells(1, 5).Value = "MUDOU"
        With ws.Range("A1:E1")
            .Font.Bold = True : .Interior.Color = RGB(17, 24, 39) : .Font.Color = RGB(255, 255, 255)
        End With
        ws.Columns(1).ColumnWidth = 18 : ws.Columns(2).ColumnWidth = 26
        ws.Columns(3).ColumnWidth = 13 : ws.Columns(4).ColumnWidth = 16 : ws.Columns(5).ColumnWidth = 34
    Else
        Dim lh As Long : lh = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        If lh >= 2 Then
            Dim hh As Variant : hh = ws.Range(ws.Cells(2, 2), ws.Cells(lh, 3)).Value
            Dim r As Long, kp As String
            For r = 1 To UBound(hh, 1)
                kp = Trim(CStr(hh(r, 1)))
                If kp <> "" Then prev(kp) = UCase(Trim(CStr(hh(r, 2))))
            Next r
        End If
    End If

    Dim runId As String : runId = Format(Now, "dd/mm/yyyy hh:nn")
    Dim base As Long : base = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Dim i As Long, out() As Variant : ReDim out(1 To UBound(d, 1), 1 To 5)
    Dim p3 As String, si As String
    For i = 1 To UBound(d, 1)
        p3 = Trim(CStr(d(i, 1))) : si = UCase(Trim(CStr(d(i, 2))))
        out(i, 1) = runId : out(i, 2) = p3 : out(i, 3) = si : out(i, 4) = Val0(d(i, 5))
        If Not prev.Exists(p3) Then
            out(i, 5) = "NOVO"
        ElseIf CStr(prev(p3)) <> si Then
            out(i, 5) = "MUDOU (" & prev(p3) & " -> " & si & ")"
        Else
            out(i, 5) = "MANTEVE"
        End If
    Next i
    ws.Range(ws.Cells(base, 1), ws.Cells(base + UBound(d, 1) - 1, 5)).Value = out
    ws.Tab.Color = RGB(60, 60, 60)
End Sub

' --------- KPIs por execucao: base da comparacao "vs. execucao anterior" ---------
' O HISTORICO existente guarda uma linha por OBRA; para os cards do painel e
' preciso o agregado da RODADA, entao vai numa aba propria e oculta (ABA_KPI).

' Ultima rodada registrada: array (1 To 6) = PEPs, reprovados, alertas,
' divergencias, valor em risco, nota da base. Empty quando nao ha historico.
Private Function LerUltimoKpiRun(wb As Workbook) As Variant
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(ABA_KPI)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    Dim lastR As Long : lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Function
    Dim v(1 To 6) As Double, c As Long
    For c = 1 To 6
        v(c) = Val0(ws.Cells(lastR, c + 1).Value)
    Next c
    LerUltimoKpiRun = v
End Function

Private Sub RegistrarKpiRun(wb As Workbook, ByVal nPep3 As Long, ByVal nReprov As Long, _
        ByVal nAlertas As Long, ByVal nDiverg As Long, ByVal valRisk As Double, _
        ByVal nota As Double)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(ABA_KPI)   ' Nothing se a aba ainda nao existe
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = ABA_KPI
        Dim h As Variant
        h = Array("RUN", "PEPS", "REPROVADOS", "ALERTAS", "DIVERGENCIAS", _
                  "VALOR EM RISCO", "NOTA BASE")
        Dim c As Long
        For c = 1 To 7
            ws.Cells(1, c).Value = h(c - 1)
        Next c
        With ws.Range("A1:G1")
            .Font.Bold = True
            .Interior.Color = RGB(17, 24, 39) : .Font.Color = RGB(255, 255, 255)
        End With
        ws.Columns(1).ColumnWidth = 18
        For c = 2 To 7
            ws.Columns(c).ColumnWidth = 15
        Next c
    Else
        ws.Visible = xlSheetVisible
    End If
    Dim r As Long : r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value = Format(Now, "dd/mm/yyyy hh:nn")
    ws.Cells(r, 2).Value = nPep3
    ws.Cells(r, 3).Value = nReprov
    ws.Cells(r, 4).Value = nAlertas
    ws.Cells(r, 5).Value = nDiverg
    ws.Cells(r, 6).Value = valRisk
    ws.Cells(r, 7).Value = nota
    ws.Tab.Color = RGB(120, 120, 120)
    ' sai da aba antes de oculta-la (a recem-criada fica ativa)
    On Error Resume Next   ' Worksheets(1) pode estar oculta
    wb.Worksheets(1).Activate
    ws.Visible = xlSheetHidden
    On Error GoTo 0
End Sub

' Texto da variacao para o rodape do card. bomSeCai = True quando DIMINUIR e
' o resultado desejado (alertas, reprovados, valor em risco...).
' ByVal nos parametros de entrada: sem isso, passar um Long (nPep3, nAlertas...)
' para um parametro Double ByRef daria "ByRef argument type mismatch".
Private Function TxtDelta(ByVal atual As Double, ByVal ant As Variant, _
        ByVal bomSeCai As Boolean, ByRef cor As Long, _
        Optional ByVal moeda As Boolean = False) As String
    cor = RGB(130, 138, 150)
    If IsEmpty(ant) Then Exit Function
    Dim dif As Double : dif = atual - CDbl(ant)
    If Abs(dif) < 0.000001 Then
        TxtDelta = ChrW(&H2022) & " sem variacao"
        Exit Function
    End If
    Dim bom As Boolean
    If dif < 0 Then bom = bomSeCai Else bom = Not bomSeCai
    If bom Then cor = RGB(21, 115, 71) Else cor = RGB(192, 0, 0)
    Dim v As String
    If moeda Then v = FmtKPI(Abs(dif)) Else v = Format(Abs(dif), "#,##0")
    TxtDelta = IIf(dif > 0, ChrW(&H25B2), ChrW(&H25BC)) & " " & v & " vs. anterior"
End Function

' --------- Nomes definidos: referencia estavel para formulas do usuario ---------
Private Sub CriarNomesDefinidos(wb As Workbook)
    DefinirNome wb, "ANALISE SAP x PRJ", "dadosAnalise", "PEP3NIVEL", 1, 16
    DefinirNome wb, "PRECO UNITARIO", "dadosPrecoUnitario", "PEP3NIVEL", 1, 14
    DefinirNome wb, "RANKING DE RISCO", "dadosRanking", "PEP3NIVEL", 2, 11, True
End Sub

' contiguo = True: o intervalo termina no fim do bloco de dados (para abas com
' outra secao abaixo, como o RANKING DE RISCO); False: ultima linha preenchida.
Private Sub DefinirNome(wb As Workbook, sheetName As String, nome As String, _
                        token As String, colToken As Long, nCols As Long, _
                        Optional ByVal contiguo As Boolean = False)
    ' Best-effort: o nome definido e so atalho de navegacao; falha nao altera a analise.
    On Error Resume Next
    Dim ws As Worksheet : Set ws = wb.Worksheets(sheetName)
    If ws Is Nothing Then Exit Sub
    Dim hdrRow As Long, rr As Long, lastScan As Long
    lastScan = ws.Cells(ws.Rows.Count, colToken).End(xlUp).Row
    For rr = 1 To lastScan
        If NormStr(CStr(ws.Cells(rr, colToken).Value)) = token Then hdrRow = rr : Exit For
    Next rr
    If hdrRow = 0 Then Exit Sub
    Dim lastRow As Long : lastRow = ws.Cells(ws.Rows.Count, colToken).End(xlUp).Row
    If contiguo Then lastRow = UltimaLinhaContigua(ws, hdrRow + 1, colToken)
    If lastRow < hdrRow Then Exit Sub
    wb.Names(nome).Delete
    wb.Names.Add Name:=nome, RefersTo:="='" & ws.Name & "'!" & _
        ws.Range(ws.Cells(hdrRow, 1), ws.Cells(lastRow, nCols)).Address
    On Error GoTo 0
End Sub

' --------- Testes de logica de negocio (rode via Alt+F8) ---------
Public Sub TestarLogicaInventario()
    Dim nomes(1 To 200) As String, oks(1 To 200) As Boolean, det(1 To 200) As String
    Dim cnt As Long, pass As Long
    Dim mp As Object : Set mp = CriarMapaNT006()

    RegTest nomes, oks, det, cnt, "EhComCritico CH FUS = True", (EhComCritico(NormStr("CH FUS")) = True)
    RegTest nomes, oks, det, cnt, "EhComCritico PARA RAIO MT = True", (EhComCritico(NormStr("PARA RAIO MT")) = True)
    RegTest nomes, oks, det, cnt, "EhComCritico POSTE = False", (EhComCritico(NormStr("POSTE")) = False)
    RegTest nomes, oks, det, cnt, "CaboComoCOM 10m = True", (CaboComoCOM("CABO ISOLADO", 10) = True)
    RegTest nomes, oks, det, cnt, "CaboComoCOM 20m = False", (CaboComoCOM("CABO ISOLADO", 20) = False)
    RegTest nomes, oks, det, cnt, "EhAderente COND 100/100 = True", (EhAderente("COND", 100, 100, "") = True)
    RegTest nomes, oks, det, cnt, "EhAderente COND 100/105 = True", (EhAderente("COND", 100, 105, "") = True)
    RegTest nomes, oks, det, cnt, "EhAderente COND 100/50 = False", (EhAderente("COND", 100, 50, "") = False)
    ' COND NU dif<=5 resgata mesmo fora dos 10% (prj 10 x 10% = 1, dif 4 > 1):
    RegTest nomes, oks, det, cnt, "EhAderente COND NU 14/10 (dif 4) = True", (EhAderente("COND NU", 14, 10, "") = True)
    RegTest nomes, oks, det, cnt, "EhAderente COND (nao NU) 14/10 = False", (EhAderente("COND", 14, 10, "") = False)
    RegTest nomes, oks, det, cnt, "EhAderente COND NU 15/10 (dif 5) = True", (EhAderente("COND NU", 15, 10, "") = True)
    RegTest nomes, oks, det, cnt, "EhAderente COND NU 16/10 (dif 6) = False", (EhAderente("COND NU", 16, 10, "") = False)
    RegTest nomes, oks, det, cnt, "EhAderente COND NU 3/0 (dif 3, prj 0) = True", (EhAderente("COND NU", 3, 0, "") = True)
    RegTest nomes, oks, det, cnt, "EhCondNu tolera pontuacao (COND-NU) = True", (EhCondNu(NormStr("COND-NU")) = True)
    RegTest nomes, oks, det, cnt, "EhCondNu COND PROT = False", (EhCondNu("COND PROT") = False)
    ' CABO ISOLADO / CABO COBRE / COND COBRE: dif ABSOLUTA ate 15 = ADERENTE
    ' (inclusive com projetado = 0, como no caso real da base):
    RegTest nomes, oks, det, cnt, "EhAderente COND COBRE 2/0 (dif 2, prj 0) = True", (EhAderente("COND COBRE", 2, 0, "") = True)
    RegTest nomes, oks, det, cnt, "EhAderente CABO COBRE 12/0 (dif 12, prj 0) = True", (EhAderente("CABO COBRE", 12, 0, "") = True)
    RegTest nomes, oks, det, cnt, "EhAderente CABO ISOLADO 15/0 (dif 15, prj 0) = True", (EhAderente("CABO ISOLADO", 15, 0, "") = True)
    RegTest nomes, oks, det, cnt, "EhAderente COND COBRE 16/0 (dif 16, prj 0) = False", (EhAderente("COND COBRE", 16, 0, "") = False)
    RegTest nomes, oks, det, cnt, "EhAderente COND COBRE 120/100 (dif 20 > 15) = False", (EhAderente("COND COBRE", 120, 100, "") = False)
    RegTest nomes, oks, det, cnt, "EhAderente CABO ACO 12/0 (fora do set, prj 0) = False", (EhAderente("CABO ACO", 12, 0, "") = False)
    RegTest nomes, oks, det, cnt, "EhAderente UC ADERENTE = True", (EhAderente("PARAFUSO", "x", "y", "ADERENTE") = True)
    ' Material UC com valores zerados: nao e mais isento, vira NAO ADERENTE:
    RegTest nomes, oks, det, cnt, "MaterialUCZerado UC 0/0 = True", (MaterialUCZerado("UC", 0, 0) = True)
    RegTest nomes, oks, det, cnt, "MaterialUCZerado UC vazio/vazio = True", (MaterialUCZerado("UC", "", "") = True)
    RegTest nomes, oks, det, cnt, "MaterialUCZerado UC 0/10 = False", (MaterialUCZerado("UC", 0, 10) = False)
    RegTest nomes, oks, det, cnt, "MaterialUCZerado COM 0/0 = False", (MaterialUCZerado("COM", 0, 0) = False)
    ' AvaliarLinha: fonte unica da regra (alimenta veredito E exibicao):
    Dim avT As tAvalLinha
    avT = AvaliarLinha("UC", "COND COBRE", 2, 0, "")
    RegTest nomes, oks, det, cnt, "AvaliarLinha COND COBRE 2/0: aderente, nao reprova", (avT.EhAder And Not avT.ReprovaLinha And avT.Avaliavel)
    avT = AvaliarLinha("UC", "PARAFUSO", 0, 0, "NULO", True)
    RegTest nomes, oks, det, cnt, "AvaliarLinha UC zerado (obra toda zerada): reprova", (avT.UcZerado And avT.ReprovaLinha And Not avT.EhAder)
    avT = AvaliarLinha("UC", "PARAFUSO", 0, 0, "NULO", False)
    RegTest nomes, oks, det, cnt, "AvaliarLinha UC zerado isolado: ignorado, nao reprova", (Not avT.UcZerado And Not avT.ReprovaLinha)
    ' Premissa de sinal por tipo de OD (valor E quantidade):
    RegTest nomes, oks, det, cnt, "ClassificarOD sufixo .D = ODD", (ClassificarOD("1234.D", "") = "ODD")
    RegTest nomes, oks, det, cnt, "ClassificarOD sem sufixo, TIPO PEP S = ODS", (ClassificarOD("1234", "S") = "ODS")
    RegTest nomes, oks, det, cnt, "Sinal ODD -5/-100 = ok", (SinalInvalidoOD("ODD", -5, -100) = "")
    RegTest nomes, oks, det, cnt, "Sinal ODD qtd +5 = NAO ADERENTE", (SinalInvalidoOD("ODD", 5, -100) <> "")
    RegTest nomes, oks, det, cnt, "Sinal ODD valor +100 = NAO ADERENTE", (SinalInvalidoOD("ODD", -5, 100) <> "")
    RegTest nomes, oks, det, cnt, "Sinal ODI 5/100 = ok", (SinalInvalidoOD("ODI", 5, 100) = "")
    RegTest nomes, oks, det, cnt, "Sinal ODS valor -100 = NAO ADERENTE", (SinalInvalidoOD("ODS", 5, -100) <> "")
    RegTest nomes, oks, det, cnt, "Sinal ODI zero = ok (nao e negativo)", (SinalInvalidoOD("ODI", 0, 0) = "")
    RegTest nomes, oks, det, cnt, "Sinal ODM negativo = fora da premissa", (SinalInvalidoOD("ODM", -5, -100) = "")
    avT = AvaliarLinha("UC", "PARAFUSO", 5, 5, "ADERENTE", False, "ODD", 100)
    RegTest nomes, oks, det, cnt, "AvaliarLinha ODD positivo: NAO ADERENTE e reprova", (avT.ViolaOD And Not avT.EhAder And avT.ReprovaLinha)
    avT = AvaliarLinha("UC", "PARAFUSO", -5, -5, "ADERENTE", False, "ODD", -100)
    RegTest nomes, oks, det, cnt, "AvaliarLinha ODD negativo: sinal ok, aderente", (Not avT.ViolaOD And avT.EhAder)
    ' Alerta PU BAIXO EM ODI (faixa 0,01 a 10, sem condutores/cabos):
    RegTest nomes, oks, det, cnt, "PuBaixoODI ODI PU 5 = alerta", (PuBaixoODI("ODI", "PARAFUSO", 5, 1, 0.01, 10) = True)
    RegTest nomes, oks, det, cnt, "PuBaixoODI ODI PU 0,01 (limite) = alerta", (PuBaixoODI("ODI", "PARAFUSO", 0.01, 1, 0.01, 10) = True)
    RegTest nomes, oks, det, cnt, "PuBaixoODI ODI PU 10 (limite) = alerta", (PuBaixoODI("ODI", "PARAFUSO", 30, 3, 0.01, 10) = True)
    RegTest nomes, oks, det, cnt, "PuBaixoODI ODI PU 10,50 = sem alerta", (PuBaixoODI("ODI", "PARAFUSO", 10.5, 1, 0.01, 10) = False)
    RegTest nomes, oks, det, cnt, "PuBaixoODI ODI PU 0,005 = sem alerta", (PuBaixoODI("ODI", "PARAFUSO", 0.005, 1, 0.01, 10) = False)
    RegTest nomes, oks, det, cnt, "PuBaixoODI CABO ISOLADO = excluido", (PuBaixoODI("ODI", "CABO ISOLADO", 5, 1, 0.01, 10) = False)
    RegTest nomes, oks, det, cnt, "PuBaixoODI COND NU = excluido", (PuBaixoODI("ODI", "COND NU", 5, 1, 0.01, 10) = False)
    RegTest nomes, oks, det, cnt, "PuBaixoODI PEP ODS = fora (so ODI)", (PuBaixoODI("ODS", "PARAFUSO", 5, 1, 0.01, 10) = False)
    RegTest nomes, oks, det, cnt, "PuBaixoODI QTD zero = sem alerta", (PuBaixoODI("ODI", "PARAFUSO", 5, 0, 0.01, 10) = False)
    ' Equivalentes: mesmo numero no total do grupo no PEP4. Caso real
    ' (grupo 3): 105010022 SAP 0 / PRJ 9 + 105010003 SAP 3 / PRJ 0 = 3 x 9
    RegTest nomes, oks, det, cnt, "Equivalente total SAP 3 / PRJ 9 = NAO ADERENTE", (SitPorQuantidade(0 + 3, 9 + 0) = "NAO ADERENTE")
    RegTest nomes, oks, det, cnt, "Equivalente total SAP 3 / PRJ 3 = ADERENTE", (SitPorQuantidade(3, 3) = "ADERENTE")
    RegTest nomes, oks, det, cnt, "Equivalente SAP 9 / PRJ 10 = NAO ADERENTE (sem tolerancia)", (SitPorQuantidade(9, 10) = "NAO ADERENTE")
    avT = AvaliarLinha("UC", "COND PROT", 10, 10.5, "ADERENTE", False)
    AjustarAderEquiv avT, 10, 10.5
    RegTest nomes, oks, det, cnt, "Equivalente cabo 10 x 10,5: NAO ADERENTE (regra de cabo nao vale)", (Not avT.EhAder And avT.ReprovaLinha)
    avT = AvaliarLinha("UC", "COND PROT", 10, 10, "ADERENTE", False)
    AjustarAderEquiv avT, 10, 10
    RegTest nomes, oks, det, cnt, "Equivalente cabo 10 x 10: ADERENTE", (avT.EhAder And Not avT.ReprovaLinha)
    avT = AvaliarLinha("UC", "CH SU FACA", 3, 9, SitPorQuantidade(3, 9), False)
    RegTest nomes, oks, det, cnt, "AvaliarLinha equivalente 3 x 9: NAO ADERENTE e reprova", (Not avT.EhAder And avT.ReprovaLinha)
    avT = AvaliarLinha("UC", "CH SU FACA", 3, 3, SitPorQuantidade(3, 3), False, "ODD", 100)
    RegTest nomes, oks, det, cnt, "Equivalente 3 x 3 com sinal invertido: NAO ADERENTE", (Not avT.EhAder And avT.ViolaOD)
    Dim dEq As Object : Set dEq = CreateObject("Scripting.Dictionary")
    dEq.Add "105000058", "CH SEC 1P" : dEq.Add "105000029", "CH SEC BYPAS"
    RegTest nomes, oks, det, cnt, "TextoEquiv mantem os dois nomes", (TextoEquiv(dEq) = "105000058 CH SEC 1P = 105000029 CH SEC BYPAS")
    ' Codigo em dois grupos une os grupos (2 e 11 compartilham 102310194):
    Dim pT As Object : Set pT = CreateObject("Scripting.Dictionary")
    pT.Add "2", "2" : pT.Add "11", "2" : pT.Add "5", "5"
    RegTest nomes, oks, det, cnt, "RaizGrupo 11 unido ao 2 = 2", (RaizGrupo(pT, "11") = "2")
    RegTest nomes, oks, det, cnt, "RaizGrupo 5 isolado = 5", (RaizGrupo(pT, "5") = "5")
    ' Lista padrao: 7 campos por linha e sem os grupos 10 e 28 (repetiam 8 e 24)
    Dim epOk As Boolean : epOk = True
    Dim epIt As Variant
    For Each epIt In EquivPadrao()
        If UBound(epIt) <> 6 Then epOk = False
        If CStr(epIt(5)) = "10" Or CStr(epIt(5)) = "28" Then epOk = False
    Next epIt
    RegTest nomes, oks, det, cnt, "EquivPadrao: 7 campos por linha, sem grupos 10 e 28", epOk
    avT = AvaliarLinha("UC", "PARAFUSO", -10, -10, "ADERENTE", False, "ODD", -50, 4)
    RegTest nomes, oks, det, cnt, "AvaliarLinha sinal pela qtd da linha (qtdSinal)", (avT.ViolaOD)
    ' Cabecalho do layout novo da extracao resolve para os nomes canonicos:
    Dim mt As Object : Set mt = CreateObject("Scripting.Dictionary")
    mt.Add NormStr("MATERIAL"), 5 : mt.Add NormStr("ATENDIDO SAP"), 9
    mt.Add NormStr("TOMBADO PROJ"), 10 : mt.Add NormStr("SITUACAO_DO_MATERIAL"), 13
    AplicarAliasesBase mt
    RegTest nomes, oks, det, cnt, "Alias MATERIAL -> COD MAT", (ColIdx(mt, "COD MAT") = 5)
    RegTest nomes, oks, det, cnt, "Alias ATENDIDO SAP -> MAT LIB SAP", (ColIdx(mt, "MAT LIB SAP") = 9)
    RegTest nomes, oks, det, cnt, "Alias TOMBADO PROJ -> MAT PRJ CAD", (ColIdx(mt, "MAT PRJ CAD") = 10)
    RegTest nomes, oks, det, cnt, "Alias SITUACAO_DO_MATERIAL -> SIT MAT", (ColIdx(mt, "SIT MAT") = 13)
    Dim mt2 As Object : Set mt2 = CreateObject("Scripting.Dictionary")
    mt2.Add "MAT LIB SAP", 3 : mt2.Add "ATENDIDO SAP", 9
    AplicarAliasesBase mt2
    RegTest nomes, oks, det, cnt, "Alias: com os dois nomes, vale o canonico", (ColIdx(mt2, "MAT LIB SAP") = 3)
    ' Alerta PU 0,01 EM CABO/COND (qualquer PEP; so familias CABO*/COND*):
    RegTest nomes, oks, det, cnt, "PuSimbolicoCabo CABO ISOLADO 1,00/100 m = alerta", (PuSimbolicoCabo("CABO ISOLADO", 1, 100, 0.01) = True)
    RegTest nomes, oks, det, cnt, "PuSimbolicoCabo COND NU 0,012 (mostra 0,01) = alerta", (PuSimbolicoCabo("COND NU", 1.2, 100, 0.01) = True)
    RegTest nomes, oks, det, cnt, "PuSimbolicoCabo ODD (valor e qtd negativos) = alerta", (PuSimbolicoCabo("CABO COBRE", -1, -100, 0.01) = True)
    RegTest nomes, oks, det, cnt, "PuSimbolicoCabo PU 0,02 = sem alerta", (PuSimbolicoCabo("CABO ISOLADO", 2, 100, 0.01) = False)
    RegTest nomes, oks, det, cnt, "PuSimbolicoCabo PARAFUSO 0,01 = fora (so cabo/cond)", (PuSimbolicoCabo("PARAFUSO", 1, 100, 0.01) = False)
    RegTest nomes, oks, det, cnt, "PuSimbolicoCabo QTD zero = sem alerta", (PuSimbolicoCabo("CABO ISOLADO", 1, 0, 0.01) = False)
    ' ODD com valor abaixo de -100 = NAO ADERENTE ("abaixo" = mais negativo):
    RegTest nomes, oks, det, cnt, "ValorBaixoODD -150 = NAO ADERENTE", (ValorBaixoODD("ODD", -150, -100) <> "")
    RegTest nomes, oks, det, cnt, "ValorBaixoODD -100 (limite) = ok", (ValorBaixoODD("ODD", -100, -100) = "")
    RegTest nomes, oks, det, cnt, "ValorBaixoODD -50 = ok", (ValorBaixoODD("ODD", -50, -100) = "")
    RegTest nomes, oks, det, cnt, "ValorBaixoODD em ODI = fora (so ODD)", (ValorBaixoODD("ODI", -150, -100) = "")
    avT = AvaliarLinha("UC", "PARAFUSO", -5, -5, "ADERENTE", False, "ODD", -150)
    RegTest nomes, oks, det, cnt, "AvaliarLinha ODD valor -150: NAO ADERENTE e reprova", (avT.ViolaOD And Not avT.EhAder And avT.ReprovaLinha)
    ' Alerta POSTE ODI PU BAIXO (preco unitario de 0 a < 500, so ODI):
    RegTest nomes, oks, det, cnt, "PosteBaratoODI POSTE RD R$ 400 = alerta", (PosteBaratoODI("ODI", "POSTE RD", 400, 1, 500) = True)
    RegTest nomes, oks, det, cnt, "PosteBaratoODI 2 postes R$ 800 (PU 400) = alerta", (PosteBaratoODI("ODI", "POSTE RD", 800, 2, 500) = True)
    RegTest nomes, oks, det, cnt, "PosteBaratoODI PU 500 (limite) = sem alerta", (PosteBaratoODI("ODI", "POSTE RD", 500, 1, 500) = False)
    RegTest nomes, oks, det, cnt, "PosteBaratoODI PU 1.200 = sem alerta", (PosteBaratoODI("ODI", "POSTE RD", 1200, 1, 500) = False)
    RegTest nomes, oks, det, cnt, "PosteBaratoODI PEP ODD = fora (so ODI)", (PosteBaratoODI("ODD", "POSTE RD", 400, 1, 500) = False)
    RegTest nomes, oks, det, cnt, "PosteBaratoODI TRAFO = fora (so poste)", (PosteBaratoODI("ODI", "TRAFO", 400, 1, 500) = False)
    RegTest nomes, oks, det, cnt, "PosteBaratoODI PU negativo = fora (alerta de sinal)", (PosteBaratoODI("ODI", "POSTE RD", -400, 1, 500) = False)
    ' Alerta MEDIDOR EM PEP ODI (medidor e material de ODM):
    RegTest nomes, oks, det, cnt, "MedidorEmODI medidor em ODI = alerta", (MedidorEmODI("ODI", NormStr("MEDIDOR ELETRONICO MONOFASICO"), "MEDIDOR") = True)
    RegTest nomes, oks, det, cnt, "MedidorEmODI so pela familia = alerta", (MedidorEmODI("ODI", NormStr("ELETRONICO 1F 120V"), "MEDIDOR") = True)
    RegTest nomes, oks, det, cnt, "MedidorEmODI medidor em ODM = sem alerta", (MedidorEmODI("ODM", NormStr("MEDIDOR ELETRONICO"), "MEDIDOR") = False)
    RegTest nomes, oks, det, cnt, "MedidorEmODI CAIXA PARA MEDIDOR = sem alerta", (MedidorEmODI("ODI", NormStr("CAIXA PARA MEDIDOR"), "CAIXA MEDICAO") = False)
    RegTest nomes, oks, det, cnt, "MedidorEmODI LACRE P/ MEDIDOR = sem alerta", (MedidorEmODI("ODI", NormStr("LACRE P/ MEDIDOR"), "LACRE") = False)
    RegTest nomes, oks, det, cnt, "MedidorEmODI TRAFO = sem alerta", (MedidorEmODI("ODI", NormStr("TRAFO DT 1F 13,8KV"), "TRAFO") = False)
    ' Carga de bases: colunas de arquivos diferentes casadas por nome
    RegTest nomes, oks, det, cnt, "CanonicoBase ATENDIDO SAP -> MAT LIB SAP", (CanonicoBase(NormStr("ATENDIDO SAP")) = "MAT LIB SAP")
    RegTest nomes, oks, det, cnt, "CanonicoBase FAMILIA -> FAMILIA (sem alias)", (CanonicoBase("FAMILIA") = "FAMILIA")
    Dim hT(1 To 1, 1 To 2) As Variant
    hT(1, 1) = "MATERIAL" : hT(1, 2) = "TOMBADO PROJ"
    Dim mH As Object : Set mH = MapaDeCabecalho(hT)
    RegTest nomes, oks, det, cnt, "MapaDeCabecalho layout novo: COD MAT e MAT PRJ CAD", (ColIdx(mH, "COD MAT") = 1 And ColIdx(mH, "MAT PRJ CAD") = 2)
    RegTest nomes, oks, det, cnt, "NomeArquivo caminho Windows", (NomeArquivo("C:\bases\base jan.xlsx") = "base jan.xlsx")
    RegTest nomes, oks, det, cnt, "NomeArquivo caminho com /", (NomeArquivo("/Users/x/base.csv") = "base.csv")
    ' PEP3 reprovado por qualquer valor negativo em ODI:
    RegTest nomes, oks, det, cnt, "NegativoEmODI ODI valor -10 = reprova PEP3", (NegativoEmODI("ODI", -10) = True)
    RegTest nomes, oks, det, cnt, "NegativoEmODI ODI valor 0 = ok", (NegativoEmODI("ODI", 0) = False)
    RegTest nomes, oks, det, cnt, "NegativoEmODI ODI valor positivo = ok", (NegativoEmODI("ODI", 150) = False)
    RegTest nomes, oks, det, cnt, "NegativoEmODI ODD negativo = fora (so ODI)", (NegativoEmODI("ODD", -10) = False)
    ' SAP x PRJ com sinais opostos: linha NAO ADERENTE e PEP3 reprovado
    RegTest nomes, oks, det, cnt, "SinaisOpostos SAP +3 / PRJ -3 = viola", (SinaisOpostos(3, -3) <> "")
    RegTest nomes, oks, det, cnt, "SinaisOpostos SAP -3 / PRJ +3 = viola", (SinaisOpostos(-3, 3) <> "")
    RegTest nomes, oks, det, cnt, "SinaisOpostos SAP 3 / PRJ 3 = ok", (SinaisOpostos(3, 3) = "")
    RegTest nomes, oks, det, cnt, "SinaisOpostos SAP 0 / PRJ -3 = ok (zero nao tem sinal)", (SinaisOpostos(0, -3) = "")
    avT = AvaliarLinha("COM", "PARAFUSO", 3, -3, "ADERENTE", False)
    RegTest nomes, oks, det, cnt, "AvaliarLinha SAP 3 / PRJ -3: NAO ADERENTE", (avT.ViolaOD And Not avT.EhAder)
    ' Carga de bases: casa colunas por nome e lista as ausentes (previa)
    Dim dMB(1 To 2, 1 To 2) As Variant
    dMB(1, 1) = "MATERIAL" : dMB(1, 2) = "ATENDIDO SAP" : dMB(2, 1) = 105010022 : dMB(2, 2) = 3
    Dim hMB(1 To 3) As String
    hMB(1) = "COD MAT" : hMB(2) = "MAT LIB SAP" : hMB(3) = "FAMILIA"
    Dim fMB As String, sMB() As Long
    Dim vMB As Variant : vMB = dMB
    sMB = MapearBloco(vMB, hMB, 3, fMB)
    RegTest nomes, oks, det, cnt, "MapearBloco casa nomes novos e lista a ausente", _
            (sMB(1) = 1 And sMB(2) = 2 And sMB(3) = 0 And fMB = "FAMILIA")
    ' TodosUCZerados: devolve a obra so quando TODOS os UC do PEP3 estao zerados
    Dim dz1 As Object, dz2 As Object
    Set dz1 = CreateObject("Scripting.Dictionary") : Set dz2 = CreateObject("Scripting.Dictionary")
    dz1("P3A") = 3 : dz2("P3A") = 3
    dz1("P3B") = 8 : dz2("P3B") = 3
    RegTest nomes, oks, det, cnt, "TodosUCZerados 3 de 3 = True (devolve)", (TodosUCZerados(dz1, dz2, "P3A") = True)
    RegTest nomes, oks, det, cnt, "TodosUCZerados 3 de 8 = False (aprova se resto ok)", (TodosUCZerados(dz1, dz2, "P3B") = False)
    RegTest nomes, oks, det, cnt, "TodosUCZerados PEP3 sem UC = False", (TodosUCZerados(dz1, dz2, "P3C") = False)
    ' Diagnostico da base: nota de saude e status por percentual
    RegTest nomes, oks, det, cnt, "NotaSaude 0 crit 0 aten = 100", (NotaSaude(0, 0) = 100)
    RegTest nomes, oks, det, cnt, "NotaSaude 1 crit 2 aten = 55", (NotaSaude(1, 2) = 55)
    RegTest nomes, oks, det, cnt, "NotaSaude 5 crit = 0 (piso)", (NotaSaude(5, 0) = 0)
    RegTest nomes, oks, det, cnt, "StatusPct 0 de 100 = OK", (StatusPct(0, 100, 0.02, 0.1) = "OK")
    RegTest nomes, oks, det, cnt, "StatusPct 5 de 100 = ATENCAO", (StatusPct(5, 100, 0.02, 0.1) = "ATENCAO")
    RegTest nomes, oks, det, cnt, "StatusPct 15 de 100 = CRITICO", (StatusPct(15, 100, 0.02, 0.1) = "CRITICO")
    avT = AvaliarLinha("COM", "POSTE", 100, 0, "NAO ADERENTE")
    RegTest nomes, oks, det, cnt, "AvaliarLinha COM nao critico: nao avaliavel", (Not avT.Avaliavel)
    avT = AvaliarLinha("COM", "CH FUS", 100, 0, "NAO ADERENTE")
    RegTest nomes, oks, det, cnt, "AvaliarLinha COM critico CH FUS: avaliavel e reprova", (avT.Avaliavel And avT.ReprovaLinha)
    avT = AvaliarLinha("UC", "CABO ISOLADO", 10, 0, "")
    RegTest nomes, oks, det, cnt, "AvaliarLinha CABO ISOLADO 10m: isento (CaboComoCOM)", (avT.CaboIsento)
    RegTest nomes, oks, det, cnt, "EhUnidadeInteira M = False", (EhUnidadeInteira("M") = False)
    RegTest nomes, oks, det, cnt, "EhUnidadeInteira UN = True", (EhUnidadeInteira("UN") = True)
    RegTest nomes, oks, det, cnt, "NT006 133100007 e ancora", (GetMat(mp, "133100007").EhAncora = True)
    RegTest nomes, oks, det, cnt, "NT006 123140003 RazaoMin = 2", (GetMat(mp, "123140003").RazaoMin = 2)
    ' Discrepancia de valor de material (ODI): +/-30% da media do codigo
    RegTest nomes, oks, det, cnt, "DesvioMedia PU 130 x media 100 = +30%", (Abs(DesvioMedia(130, 100) - 0.3) < 0.000001)
    RegTest nomes, oks, det, cnt, "DesvioMedia PU 70 x media 100 = -30%", (Abs(DesvioMedia(70, 100) + 0.3) < 0.000001)
    RegTest nomes, oks, det, cnt, "EhDiscrepValor PU 131 x media 100 = discrepante", (EhDiscrepValor(131, 100, 0.3) = True)
    RegTest nomes, oks, det, cnt, "EhDiscrepValor PU 69 x media 100 = discrepante", (EhDiscrepValor(69, 100, 0.3) = True)
    RegTest nomes, oks, det, cnt, "EhDiscrepValor PU 130 (limite) = ok", (EhDiscrepValor(130, 100, 0.3) = False)
    RegTest nomes, oks, det, cnt, "EhDiscrepValor PU 70 (limite) = ok", (EhDiscrepValor(70, 100, 0.3) = False)
    RegTest nomes, oks, det, cnt, "EhDiscrepValor PU 115 = ok", (EhDiscrepValor(115, 100, 0.3) = False)
    RegTest nomes, oks, det, cnt, "EhDiscrepValor media zero = sem comparacao", (EhDiscrepValor(50, 0, 0.3) = False)
    ' Discrepancia de valor no SCORE do ranking: 2 por item, teto 12
    RegTest nomes, oks, det, cnt, "ParcelaRisco 2 discrepancias x 2 = 4", (ParcelaRisco(2, 2, 12) = 4)
    RegTest nomes, oks, det, cnt, "ParcelaRisco 6 discrepancias x 2 = 12 (teto)", (ParcelaRisco(6, 2, 12) = 12)
    RegTest nomes, oks, det, cnt, "ParcelaRisco 10 discrepancias x 2 = teto 12", (ParcelaRisco(10, 2, 12) = 12)
    RegTest nomes, oks, det, cnt, "ParcelaRisco sem discrepancia = 0", (ParcelaRisco(0, 2, 12) = 0)
    ' Ordenacao estavel do ranking/discrepancias: k1 desc, empate k2 desc, empate total = ordem de entrada
    Dim oiT(0 To 4) As Long, oiK1(0 To 4) As Double, oiK2(0 To 4) As Double, oiI As Long
    oiK1(0) = 5 : oiK1(1) = 9 : oiK1(2) = 5 : oiK1(3) = 9 : oiK1(4) = 5
    oiK2(0) = 1 : oiK2(1) = 2 : oiK2(2) = 3 : oiK2(3) = 2 : oiK2(4) = 3
    For oiI = 0 To 4 : oiT(oiI) = oiI : Next oiI
    OrdenarIdxDesc oiT, oiK1, oiK2, 0, 4
    RegTest nomes, oks, det, cnt, "OrdenarIdxDesc score desc, valor desc, empate estavel = 1,3,2,4,0", _
            (oiT(0) = 1 And oiT(1) = 3 And oiT(2) = 2 And oiT(3) = 4 And oiT(4) = 0)
    ' NormStr com cache devolve o mesmo resultado na 1a e na 2a chamada
    RegTest nomes, oks, det, cnt, "NormStr cache: 1a chamada", (NormStr(" Pe" & ChrW(231) & "a/Cabo-01 ") = "PECACABO01")
    RegTest nomes, oks, det, cnt, "NormStr cache: 2a chamada igual", (NormStr(" Pe" & ChrW(231) & "a/Cabo-01 ") = "PECACABO01")
    ' LACRE P/ MEDIDOR e lacre (antes caia em MEDIDOR na NT.006 e na regra ODM)
    Dim tmTstLac As tMaterial, tmTstMed As tMaterial, okTstCl As Boolean
    okTstCl = ClassificarDesc(NormStr("LACRE P/ MEDIDOR"), tmTstLac)
    RegTest nomes, oks, det, cnt, "ClassificarDesc LACRE P/ MEDIDOR = familia LACRE", (okTstCl And tmTstLac.Familia = "LACRE")
    okTstCl = ClassificarDesc(NormStr("MEDIDOR ELETRONICO MONOFASICO"), tmTstMed)
    RegTest nomes, oks, det, cnt, "ClassificarDesc MEDIDOR ELETRONICO = familia MEDIDOR", (okTstCl And tmTstMed.Familia = "MEDIDOR")
    RegTest nomes, oks, det, cnt, "EhMedidorODM MEDIDOR ELETRONICO = medidor", (EhMedidorODM(NormStr("MEDIDOR ELETRONICO") & " MEDIDOR") = True)
    RegTest nomes, oks, det, cnt, "EhMedidorODM LACRE P/ MEDIDOR = nao e medidor", (EhMedidorODM(NormStr("LACRE P/ MEDIDOR") & " LACRE") = False)
    RegTest nomes, oks, det, cnt, "EhMedidorODM CAIXA PARA MEDIDOR = nao e medidor", (EhMedidorODM(NormStr("CAIXA PARA MEDIDOR") & " CAIXA MEDICAO") = False)

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("TESTES")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        ws.Name = "TESTES"
    Else
        ws.Cells.Clear
    End If
    ws.Cells(1, 1).Value = "TESTE DE LOGICA - " & Format(Now, "dd/mm/yyyy hh:nn")
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(2, 1).Value = "TESTE" : ws.Cells(2, 2).Value = "RESULTADO"
    ws.Range("A2:B2").Font.Bold = True
    Dim i As Long
    For i = 1 To cnt
        ws.Cells(2 + i, 1).Value = nomes(i)
        ws.Cells(2 + i, 2).Value = det(i)
        If oks(i) Then pass = pass + 1
        ws.Cells(2 + i, 2).Interior.Color = IIf(oks(i), RGB(198, 239, 206), RGB(255, 199, 206))
    Next i
    ws.Columns(1).ColumnWidth = 46 : ws.Columns(2).ColumnWidth = 14
    ws.Cells(3 + cnt, 1).Value = pass & " de " & cnt & " testes passaram"
    ws.Cells(3 + cnt, 1).Font.Bold = True
    MsgBox pass & " de " & cnt & " testes passaram." & _
           IIf(pass = cnt, " Tudo OK.", " HA FALHAS - ver aba TESTES."), _
           IIf(pass = cnt, vbInformation, vbExclamation), "Testes de Logica"
End Sub

Private Sub RegTest(nomes() As String, oks() As Boolean, det() As String, _
                    cnt As Long, nome As String, ok As Boolean)
    If cnt >= UBound(nomes) Then Exit Sub   ' nao estoura o quadro de testes
    cnt = cnt + 1
    nomes(cnt) = nome : oks(cnt) = ok : det(cnt) = IIf(ok, "PASSOU", "FALHOU")
End Sub

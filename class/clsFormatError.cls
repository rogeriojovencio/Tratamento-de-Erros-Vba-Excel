VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormatError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pCode As Long
Private pNamesht As String
Private pMsn1 As String
Private pMsnFinaly As String
'----------------------------
Private pWS_Count As Integer
Private pI As Integer
'------------------------------------
Private pLastLine As Integer
Private pLine As Long
Private pNewLine As Long
Private pSistema As String
Private pProcedimento As String
Private pData As String
Private pCont As Long

Public Property Get code() As Variant
    code = pCode
End Property
Public Property Let code(ByVal vNewValue As Variant)
    pCode = vNewValue
End Property
Public Property Get Nameshtsht() As Variant
    Namesht = pNamesht
End Property
Public Property Let Namesht(ByVal vNewValue As Variant)
    pNamesht = vNewValue
End Property
Public Property Get Msn1() As Variant
    Msn1 = pMsn1
End Property
Public Property Let Msn1(ByVal vNewValue As Variant)
    pMsn1 = vNewValue
End Property
Private Property Get LastLine() As Variant
    LastLine = Worksheets("shtErr").range("A1048576").End(xlUp).Row
End Property
Private Property Let LastLine(ByVal vNewValue As Variant)
    pLastLine = vNewValue
End Property
Private Property Get NewLine() As Variant
    NewLine = pNewLine
End Property
Private Property Let NewLine(ByVal vNewValue As Variant)
    pNewLine = vNewValue
End Property

Public Property Get Sistema() As Variant
    Sistema = pSistema
End Property

Public Property Let Sistema(ByVal vNewValue As Variant)
    pSistema = vNewValue
End Property

Public Property Get Procedimento() As Variant
    Procedimento = pProcedimento
End Property

Public Property Let Procedimento(ByVal vNewValue As Variant)
    pProcedimento = vNewValue
End Property

Public Property Get data() As Variant
    data = pData
End Property

Public Property Let data(ByVal vNewValue As Variant)
    pData = vNewValue
End Property

Public Function fcnError_TableCreate() As Integer
    Application.ScreenUpdating = False
    fcnError_TableCreate = 0
    ThisWorkbook.Worksheets.Add After:=Worksheets(Worksheets.count)
    pWS_Count = ThisWorkbook.Worksheets.count
    For pI = 1 To pWS_Count
        pNamesht = ThisWorkbook.Worksheets(pI).Name
        Select Case Left(pNamesht, 4)
        Case "Shee", "Plan"
            ActiveSheet.Name = "shtErr"
            Call fcnFormatCopy(1)    'controe formato para  nova tabela
            GoTo px
        End Select
    Next pI
px:

End Function
Private Sub suVerifyObjcts()
    Application.ScreenUpdating = False
    If fcnError_SearchTable = 0 Then    'verifica se tabela existe
        Call fcnError_TableCreate      'Se não existe cria a tabela
    End If
End Sub
Private Function fcnError_InsertCodName(sCode As Long, _
                                        sDescription As String, _
                                        sProcedimento As String, _
                                        sdata As String)
    pLine = ProcurarLinha(CStr(sCode), 1)
    If pLine < 0 Then
        pNewLine = LastLine + 1
        '-------------------------------------------------------------------
        Worksheets("shtErr").Cells(pNewLine, 1).NumberFormat = "@"
        Worksheets("shtErr").Cells(pNewLine, 2).NumberFormat = "@"
        Worksheets("shtErr").Cells(pNewLine, 3).NumberFormat = "@"
        Worksheets("shtErr").Cells(pNewLine, 4).NumberFormat = "@"
        Worksheets("shtErr").Cells(pNewLine, 5).NumberFormat = "@"
        '-------------------------------------------------------------------
        Worksheets("shtErr").Cells(pNewLine, 1).value = sCode
        Worksheets("shtErr").Cells(pNewLine, 2).value = sDescription
        '---------------------------------------------------------------------
        Worksheets("shtErr").Cells(pNewLine, 4).value = sProcedimento
        Worksheets("shtErr").Cells(pNewLine, 5).value = sdata
        Worksheets("shtErr").Cells(pNewLine, 6) = 1
        '-------------------------------------------------------------------
    ElseIf pLine > 0 Then

        '---------------------------------------------------------------------
        Worksheets("shtErr").Cells(pLine, 4).value = sProcedimento
        Worksheets("shtErr").Cells(pLine, 5).value = sdata
        Worksheets("shtErr").Cells(pLine, 6).value = Worksheets("shtErr").Cells(pLine, 6) + 1
    End If
    Call fcnFormatCopy(2)
End Function
Private Function ProcurarLinha(EncontraString As String, _
                               opt As Integer)

    Dim Intervalo As range
    Select Case opt
    Case 1    'Codigo de Erro

        If Trim(EncontraString) <> "" Then
            EncontraString = Trim(EncontraString)
            With Worksheets("shtErr").range("A:A")
                Set Intervalo = .Find(What:=EncontraString, _
                                      After:=.Cells(1), _
                                      LookIn:=xlValues, _
                                      LookAt:=xlWhole, _
                                      SearchOrder:=xlByRows, _
                                      SearchDirection:=xlPrevious, _
                                      MatchCase:=False)
                If Not Intervalo Is Nothing Then
                    drow = Worksheets("shtErr").range("A:A").Find(What:=EncontraString).Row
                    ProcurarLinha = drow
                Else
                    ProcurarLinha = -1
                End If
            End With
        End If
    Case 2    'Codigo de Erro Mensagem 2

        If Trim(EncontraString) <> "" Then
            EncontraString = Trim(EncontraString)
            With wsSegmento.range("C:C")
                Set Intervalo = .Find(What:=EncontraString, _
                                      After:=.Cells(1), _
                                      LookIn:=xlValues, _
                                      LookAt:=xlWhole, _
                                      SearchOrder:=xlByRows, _
                                      SearchDirection:=xlPrevious, _
                                      MatchCase:=False)
                If Not Intervalo Is Nothing Then
                    drow = Worksheets("shtErr").range("C:C").Find(What:=EncontraString).Row
                    ProcurarLinha = drow
                Else
                    ProcurarLinha = -1
                End If
            End With
        End If
    End Select

End Function
Private Function fcnError_SearchTable() As Integer
    fcnError_SearchTable = 0
    pWS_Count = ThisWorkbook.Worksheets.count
    For pI = 1 To pWS_Count
        pNamesht = ThisWorkbook.Worksheets(pI).Name
        If pNamesht = "shtErr" Then
            fcnError_SearchTable = 1
            GoTo px
        Else
            fcnError_SearchTable = 0
        End If
    Next pI
px:
End Function

Private Function fcnError_MsnConstruct()
' Neste momento jã foi inserido o codigo de erro
    Application.ScreenUpdating = False
    pLine = ProcurarLinha(CStr(pCode), 1)
    Select Case pLine
    Case Is > 0
        If Trim(Worksheets("shtErr").Cells(pLine, 3)) = "" Then
            pMsnFinaly = pCode & " - " & Worksheets("shtErr").Cells(pLine, 2)
        Else
            pMsnFinaly = pCode & " - " & Worksheets("shtErr").Cells(pLine, 3)
        End If
    Case Else
        pMsnFinaly = pCode & " - " & Worksheets("shtErr").Cells(pLine, 2)
    End Select

End Function
Public Function fcnMessageError() As String
' verifica se tabela existe caso exista cria
    suVerifyObjcts
    Call fcnError_InsertCodName(pCode, pMsn1, pProcedimento, pData)  ' caso nao existir insere novo codigo de erro
    Call fcnFormatCopy(2)
    Call fcnError_MsnConstruct  ' constroe a mensagem de erro.
    Worksheets("Menu").Activate
    Worksheets("shtErr").Visible = False
    MsgBox pMsnFinaly, vbInformation, Sistema
End Function
Private Sub fcnFormatCopy(stipo As Integer)
    Application.ScreenUpdating = False
    Select Case stipo
    Case 1
        range("A1:F1").Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = -0.249977111117893
        End With
        With Selection.Font
            .Name = "Calibri"
            .Size = 10
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = -0.249977111117893
            .ThemeFont = xlThemeFontMinor
        End With
        range("A1").Select
        ActiveCell.FormulaR1C1 = "Codigo:"
        range("B1").Select
        ActiveCell.FormulaR1C1 = "Descrição:"
        range("C1").Select
        ActiveCell.FormulaR1C1 = "Descrição2:"
        range("D1").Select
        ActiveCell.FormulaR1C1 = "Procedimento:"
        range("E1").Select
        ActiveCell.FormulaR1C1 = "Data:"
        range("F1").Select
        ActiveCell.FormulaR1C1 = "Qtde:"
        '---------------------------------------------------------------------------
        range("A2:F2").Select
        With Selection.Font
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
        End With
        With Selection.Font
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
        End With
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        ThisWorkbook.Save

    Case 2

        Worksheets("shtErr").Visible = True
        Worksheets("shtErr").Activate
        ActiveWindow.DisplayGridlines = False
        Worksheets("shtErr").range("A2:F2").Select
        Selection.Copy
        Worksheets("shtErr").range("A2:F" & LastLine & "").Select
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                               SkipBlanks:=False, Transpose:=False
        Worksheets("shtErr").Columns("A:F").EntireColumn.AutoFit
        Application.CutCopyMode = False
        Worksheets("shtErr").range("A1").Select
    End Select

End Sub




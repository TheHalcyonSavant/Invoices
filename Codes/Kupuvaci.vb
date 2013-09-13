'======================= Sheet1 ==========================
Private Sub btnKartica_Click()
    If Not isFirmaSelected Then Exit Sub
    frmKartica.Show False
End Sub
Private Sub btnPlakane_Click()
    If Not isFirmaSelected Then Exit Sub
    frmPlakane.Show False
End Sub

'======================= Sheet2 ==========================
Private Sub btnBriziPlakane_Click()
    Dim answer As String, row As Integer
    
    row = ActiveCell.row
    If row = 1 Or ActiveCell.Rows.Count > 1 Then Exit Sub
    
    answer = MsgBox("Sigurno li sakas da go izbrises broj " & (row - 1) & " ?", vbQuestion + vbYesNo)
    If answer = vbNo Then Exit Sub
    
    Sheet1.Unprotect
    Sheet2.Unprotect
    Application.EnableEvents = False
    
    row = Sheet2.Cells(row, LOG_FIRMA_ID_COL) + 1
    Sheet1.Cells(row, FIRMI_SUM_DOLZI_COL) = Sheet1.Cells(row, FIRMI_SUM_DOLZI_COL) - Sheet2.Cells(ActiveCell.row, LOG_DOLZI_COL)
    Sheet1.Cells(row, FIRMI_SUM_POBARUVA_COL) = Sheet1.Cells(row, FIRMI_SUM_POBARUVA_COL) - Sheet2.Cells(ActiveCell.row, LOG_POBARUVA_COL)
    Sheet2.Rows(ActiveCell.row).Delete
    
    Application.EnableEvents = True
    Sheet2.Protect
    Sheet1.Protect
End Sub

'======================= ThisWorkbook ==========================
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Sub Workbook_BeforePrint(Cancel As Boolean)
    Cancel = True
    If ActiveSheet.Index = 3 Then Cancel = False
End Sub
Private Sub Workbook_Open()
    Dim dolzi As Long, files As Variant, pobaruva As Long, wb As Workbook
    Dim dictDolziSum As New Scripting.Dictionary
    Dim dictPobaruvaSum As New Scripting.Dictionary

    Sheet1.Unprotect
    Sheet2.Unprotect
    Sheet3.Unprotect
    Sheet1.EnableSelection = xlUnlockedCells
    Sheet2.EnableSelection = xlNoRestrictions
    Sheet3.EnableSelection = xlNoRestrictions
    Application.WindowState = xlMaximized
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    If Sheet2.UsedRange.Rows.Count > 1 Then
        For Each c In Sheet2.Range("H2:H" & Sheet2.UsedRange.Rows.Count)
            dolzi = Sheet2.Cells(c.row, LOG_DOLZI_COL)
            If dictDolziSum.Exists(c.Value) Then
                dictDolziSum(c.Value) = dictDolziSum(c.Value) + dolzi
            Else
                dictDolziSum(c.Value) = dolzi
            End If
            pobaruva = Sheet2.Cells(c.row, LOG_POBARUVA_COL)
            If dictPobaruvaSum.Exists(c.Value) Then
                dictPobaruvaSum(c.Value) = dictPobaruvaSum(c.Value) + pobaruva
            Else
                dictPobaruvaSum(c.Value) = pobaruva
            End If
        Next c
    End If
    
    For Each c In Sheet1.Range("A2:A" & Sheet1.UsedRange.Rows.Count)
        If c = "" Then Exit For
        dolzi = 0
        If dictDolziSum.Exists(c.Value) Then dolzi = dictDolziSum(c.Value)
        Sheet1.Cells(c.row, FIRMI_SUM_DOLZI_COL) = dolzi
        pobaruva = 0
        If dictPobaruvaSum.Exists(c.Value) Then pobaruva = dictPobaruvaSum(c.Value)
        Sheet1.Cells(c.row, FIRMI_SUM_POBARUVA_COL) = pobaruva
    Next c
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Sheet3.Protect
    Sheet2.Protect
    Sheet1.Protect
    
    If GetKeyState(144) = 0 Then SendKeys "{NUMLOCK}", True
    
    ThisWorkbook.Save
End Sub

'======================= frmKartica ==========================
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As String, ByVal p4 As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, p4 As Any) As Long
Private Sub btnNapraviKartica_Click()
    Dim data As Date, dolzi As Long, dolziCol As String
    Dim frmId As Integer, iRow As Integer, lastCell As Range
    Dim pobaruva As Long, pobaruvaCol As String, pretDolzi As Long, pretPobaruva As Long
    
    Sheet2.Unprotect
    Sheet3.Unprotect
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Sheet3.Range(KARTICA_FIRMA_CELL) = lblFirma.Caption
    Sheet3.Range(KARTICA_FIRMA_GRAD_CELL) = lblGrad.Caption
    Set lastCell = Sheet3.UsedRange.SpecialCells(xlCellTypeLastCell)
    If lastCell.row >= KARTICA_BEGIN_EDIT_ROW Then Sheet3.Range("A" & KARTICA_BEGIN_EDIT_ROW, lastCell).Clear

    iRow = KARTICA_BEGIN_EDIT_ROW
    pretDolzi = Sheet1.Cells(tbFirmaId.Value + 1, FIRMI_PRED_DOLZI_COL)
    pretPobaruva = 0
    Set lastCell = Sheet2.UsedRange.SpecialCells(xlCellTypeLastCell)
    If lastCell.row > 1 Then
        For Each c In Sheet2.Range("A2:A" & lastCell.row)
            frmId = Sheet2.Cells(c.row, LOG_FIRMA_ID_COL)
            data = Sheet2.Cells(c.row, LOG_DATA_COL)
            dolzi = Sheet2.Cells(c.row, LOG_DOLZI_COL)
            pobaruva = Sheet2.Cells(c.row, LOG_POBARUVA_COL)
            If frmId = tbFirmaId.Value And data <= mscalDo.Value Then
                If data >= mscalOd.Value Then
                    Sheet3.Cells(iRow, KARTICA_DATA_COL).NumberFormat = "dd - mm - yyyy;@"
                    Sheet3.Cells(iRow, KARTICA_DATA_COL) = data
                    Sheet3.Cells(iRow, KARTICA_DOLZI_COL).NumberFormat = "#,##0"
                    Sheet3.Cells(iRow, KARTICA_DOLZI_COL) = dolzi
                    Sheet3.Cells(iRow, KARTICA_POBARUVA_COL).NumberFormat = "#,##0"
                    Sheet3.Cells(iRow, KARTICA_POBARUVA_COL) = pobaruva
                    Sheet3.Cells(iRow, KARTICA_OPIS_COL).Font.Name = "MAC C Times"
                    Sheet3.Cells(iRow, KARTICA_OPIS_COL).InsertIndent 1
                    Sheet3.Cells(iRow, KARTICA_OPIS_COL) = Sheet2.Cells(c.row, LOG_OPIS_COL)
                    With Sheet3.Range("A" & iRow & ":E" & iRow)
                        .Borders(xlEdgeLeft).Weight = xlHairline
                        .Borders(xlInsideVertical).Weight = xlHairline
                        .Borders(xlEdgeRight).Weight = xlHairline
                    End With
                    iRow = iRow + 1
                Else
                    pretDolzi = pretDolzi + dolzi
                    pretPobaruva = pretPobaruva + pobaruva
                End If
            End If
        Next c
    End If
    
    Sheet3.Cells(KARTICA_BEGIN_EDIT_ROW - 1, KARTICA_DOLZI_COL) = pretDolzi
    Sheet3.Cells(KARTICA_BEGIN_EDIT_ROW - 1, KARTICA_POBARUVA_COL) = pretPobaruva

    iRow = iRow - 1
    If iRow >= KARTICA_BEGIN_EDIT_ROW Then
        If iRow > KARTICA_BEGIN_EDIT_ROW Then
            Sheet3.Range("A" & KARTICA_BEGIN_EDIT_ROW & ":E" & iRow).Sort Sheet3.Range("A" & KARTICA_BEGIN_EDIT_ROW & ":A" & iRow)
        End If
        dolziCol = colToChr(KARTICA_DOLZI_COL)
        pobaruvaCol = colToChr(KARTICA_POBARUVA_COL)
        For Each c In Sheet3.Range(KARTICA_SALDO_COL & KARTICA_BEGIN_EDIT_ROW & ":" & KARTICA_SALDO_COL & iRow)
            c.Formula = "=" & KARTICA_SALDO_COL & (c.row - 1) & "+" & dolziCol & c.row & "-" & pobaruvaCol & c.row
        Next c
        iRow = iRow + 1
        Sheet3.Cells(iRow, KARTICA_DATA_COL).Font.Name = "MAC C Times"
        Sheet3.Cells(iRow, KARTICA_DATA_COL).HorizontalAlignment = xlRight
        Sheet3.Cells(iRow, KARTICA_DATA_COL) = "suma:"
        Sheet3.Cells(iRow, KARTICA_DOLZI_COL).Formula = "=SUM(" & dolziCol & (KARTICA_BEGIN_EDIT_ROW - 1) & ":" & dolziCol & (iRow - 1) & ")"
        Sheet3.Cells(iRow, KARTICA_POBARUVA_COL).NumberFormat = "#,##0"
        Sheet3.Cells(iRow, KARTICA_POBARUVA_COL).Formula = "=SUM(" & pobaruvaCol & (KARTICA_BEGIN_EDIT_ROW - 1) & ":" & pobaruvaCol & (iRow - 1) & ")"
        Sheet3.Range(KARTICA_SALDO_COL & iRow).NumberFormat = "#,##0"
        Sheet3.Range(KARTICA_SALDO_COL & iRow).Formula = "=" & dolziCol & iRow & "-" & pobaruvaCol & iRow
        With Sheet3.Range("A" & iRow & ":E" & iRow)
            .Borders(xlEdgeLeft).Weight = xlHairline
            .Borders(xlEdgeBottom).Weight = xlHairline
            .Borders(xlInsideVertical).Weight = xlHairline
            .Borders(xlEdgeRight).Weight = xlHairline
            .Borders(xlEdgeTop).Weight = xlHairline
        End With
    End If

    Unload Me
    Sheet3.Activate
    'Sheet3.UsedRange.Select
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Sheet3.Protect
    Sheet2.Protect
    ThisWorkbook.Save
    
    SetSheetFocus
End Sub
Private Sub UserForm_Initialize()
    tbFirmaId.Value = CInt(Sheet1.Cells(ActiveCell.row, ActiveCell.Column - 1))
    lblFirma.Caption = ActiveCell
    lblGrad.Caption = Sheet1.Cells(ActiveCell.row, ActiveCell.Column + 1)
    mscalOd.Value = "01.01." & Year(Date)
    mscalDo.Value = Date
End Sub
Sub SetSheetFocus()
    SendMessage FindWindowEx(FindWindowEx(Application.Hwnd, 0&, "XLDESK", vbNullString), 0&, "EXCEL7", ActiveWindow.Caption), &H7, 0&, 0&
End Sub

'======================= frmPlakane ==========================
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As String, ByVal p4 As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, p4 As Any) As Long
Private Sub btnVnesi_Click()
    Dim i As Integer
    
    If tbDolzi.Value = "" Then tbDolzi.Value = 0
    If tbPobaruva.Value = "" Then tbPobaruva.Value = 0
    If tbDolzi.Value = 0 And tbPobaruva.Value = 0 Then Exit Sub
    If Right(tbOpis.Value, 1) = "/" Then Exit Sub

    Sheet1.Unprotect
    Sheet2.Unprotect
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    i = 2
    If Sheet2.UsedRange.Rows.Count > 1 Then
        i = Sheet2.UsedRange.Rows.Count + 1
        Sheet2.Cells(i, 1).Formula = "=A" & (i - 1) & "+1"
    Else
        Sheet2.Cells(i, 1) = 1
    End If
    Sheet2.Cells(i, LOG_DATA_COL).NumberFormat = "[$-42F]dddd, dd mmmm yyyy;@"
    Sheet2.Cells(i, LOG_DATA_COL) = kalendar.Value
    Sheet2.Cells(i, LOG_FIRMA_COL).Font.Name = "MAC C Times"
    Sheet2.Cells(i, LOG_FIRMA_COL) = lblFirma.Caption
    Sheet2.Cells(i, LOG_FIRMA_GRAD_COL).Font.Name = "MAC C Times"
    Sheet2.Cells(i, LOG_FIRMA_GRAD_COL) = lblGrad.Caption
    Sheet2.Cells(i, LOG_DOLZI_COL).NumberFormat = "#,##0 [$" & ChrW(1076) & ChrW(1077) & ChrW(1085) & ".-42F]"
    Sheet2.Cells(i, LOG_DOLZI_COL) = tbDolzi.Value
    Sheet2.Cells(i, LOG_POBARUVA_COL).NumberFormat = "#,##0 [$" & ChrW(1076) & ChrW(1077) & ChrW(1085) & ".-42F]"
    Sheet2.Cells(i, LOG_POBARUVA_COL) = tbPobaruva.Value
    Sheet2.Cells(i, LOG_OPIS_COL).Font.Name = "MAC C Times"
    Sheet2.Cells(i, LOG_OPIS_COL) = tbOpis.Value
    Sheet2.Cells(i, LOG_FIRMA_ID_COL).NumberFormat = ";;;"
    Sheet2.Cells(i, LOG_FIRMA_ID_COL) = tbFirmaId.Value
    Sheet1.Cells(ActiveCell.row, FIRMI_SUM_DOLZI_COL) = Sheet1.Cells(ActiveCell.row, FIRMI_SUM_DOLZI_COL).Value + tbDolzi.Value
    Sheet1.Cells(ActiveCell.row, FIRMI_SUM_POBARUVA_COL) = Sheet1.Cells(ActiveCell.row, FIRMI_SUM_POBARUVA_COL).Value + tbPobaruva.Value
    
    Unload Me
    Sheet2.Activate
    Sheet2.Range("A" & i & ":G" & i).Select
    Sheet1.Activate
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Sheet2.Protect
    Sheet1.Protect
    ThisWorkbook.Save
    
    SetSheetFocus
End Sub
Private Sub tbCena_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        btnPlati_Click
    ElseIf KeyCode = 27 Then
        Unload Me
        SetSheetFocus
    End If
End Sub
Private Sub UserForm_Initialize()
    tbFirmaId.Value = CInt(Sheet1.Cells(ActiveCell.row, ActiveCell.Column - 1))
    lblFirma.Caption = ActiveCell
    lblGrad.Caption = Sheet1.Cells(ActiveCell.row, ActiveCell.Column + 1)
    kalendar.Value = Date
End Sub
Sub SetSheetFocus()
    SendMessage FindWindowEx(FindWindowEx(Application.Hwnd, 0&, "XLDESK", vbNullString), 0&, "EXCEL7", ActiveWindow.Caption), &H7, 0&, 0&
End Sub

'======================= General ==========================
Public Const FIRMI_PRED_DOLZI_COL = 7
Public Const FIRMI_SUM_DOLZI_COL = 5
Public Const FIRMI_SUM_POBARUVA_COL = 6
Public Const KARTICA_BEGIN_EDIT_ROW = 6
Public Const KARTICA_DATA_COL = 1
Public Const KARTICA_DOLZI_COL = 2
Public Const KARTICA_FIRMA_CELL = "E1"
Public Const KARTICA_FIRMA_GRAD_CELL = "E2"
Public Const KARTICA_OPIS_COL = 5
Public Const KARTICA_POBARUVA_COL = 3
Public Const KARTICA_SALDO_COL = "D"
Public Const LOG_DATA_COL = 2
Public Const LOG_DOLZI_COL = 5
Public Const LOG_FIRMA_COL = 3
Public Const LOG_FIRMA_GRAD_COL = 4
Public Const LOG_FIRMA_ID_COL = 8
Public Const LOG_OPIS_COL = 7
Public Const LOG_POBARUVA_COL = 6
Public Function isFirmaSelected() As Boolean
    If ActiveCell = "" Or ActiveCell.Rows.Count > 1 Or ActiveCell.row = 1 Or ActiveCell.Column <> 2 Then
        MsgBox "Prvo selektiraj firma"
        isFirmaSelected = False
        Exit Function
    End If
    isFirmaSelected = True
End Function
Function colToChr(iCol As Integer) As String
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   If iAlpha > 0 Then colToChr = Chr(iAlpha + 64)
   If iRemainder > 0 Then colToChr = colToChr & Chr(iRemainder + 64)
End Function

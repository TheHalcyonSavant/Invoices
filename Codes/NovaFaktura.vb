'======================= Sheet1 ==========================
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range(FAKTURA_KOLICINA_CELLS)) Is Nothing Then
        ChangeDanok
    End If
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim c As String
    
    c = Target.Address(False, False)
    Unload frmArtikli
    Unload frmFirmi
    Unload frmKalendar
    If c = FAKTURA_FIRMA_CELLS Then
        frmFirmi.Show False
        frmFirmi.tbInput.SetFocus
    ElseIf c = FAKTURA_DATA_CELLS Then
        frmKalendar.kalendar.Value = Range(FAKTURA_DATA_CELL)
        frmKalendar.Show False
    ElseIf Not Intersect(Target, Range(FAKTURA_ARTIKLI_CELLS)) Is Nothing And Target.Rows.Count = 1 Then
        frmArtikli.Show False
        frmArtikli.tbInput.SetFocus
    End If
End Sub
Public Sub ChangeDanok()
    Dim r As Range, sum5(1 To 2) As Long, sum18(1 To 2) As Long, v As Variant
    
    Sheet1.Unprotect
    Application.EnableEvents = False
    For Each c In Range(FAKTURA_KOLICINA_CELLS)
        Set r = Cells(c.Row, FAKTURA_DDV_COL)
        If r.Value = "5" Then
            If Cells(c.Row, FAKTURA_IZNOS_NA_DDV_COL) <> "" Then sum5(1) = sum5(1) + Cells(c.Row, FAKTURA_IZNOS_NA_DDV_COL)
            If Cells(c.Row, FAKTURA_IZNOS_VKUPEN_COL) <> "" Then sum5(2) = sum5(2) + Cells(c.Row, FAKTURA_IZNOS_VKUPEN_COL)
        ElseIf r.Value = "18" Then
            If Cells(c.Row, FAKTURA_IZNOS_NA_DDV_COL) <> "" Then sum18(1) = sum18(1) + Cells(c.Row, FAKTURA_IZNOS_NA_DDV_COL)
            If Cells(c.Row, FAKTURA_IZNOS_VKUPEN_COL) <> "" Then sum18(2) = sum18(2) + Cells(c.Row, FAKTURA_IZNOS_VKUPEN_COL)
        End If
    Next c
    Cells(FAKTURA_5DDV_ROW, FAKTURA_IZNOS_NA_DDV_COL) = sum5(1)
    Cells(FAKTURA_5DDV_ROW, FAKTURA_IZNOS_VKUPEN_COL) = sum5(2)
    Cells(FAKTURA_18DDV_ROW, FAKTURA_IZNOS_NA_DDV_COL) = sum18(1)
    Cells(FAKTURA_18DDV_ROW, FAKTURA_IZNOS_VKUPEN_COL) = sum18(2)
    Application.EnableEvents = True
    Sheet1.Protect
End Sub


'======================= ThisWorkbook ==========================
Public dictArtikli As Scripting.Dictionary
Public dictFirmi As Scripting.Dictionary

Private Const VK_NUMLOCK As Long = 144
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Sub Workbook_BeforePrint(Cancel As Boolean)
    If ThisWorkbook.Name = "NovaFaktura.xlsm" Then
        MsgBox "Prvo stisni 'Save' (ili CTRL+S), za da se sozdade kopija vo 'arhiva'"
        Cancel = True
    End If
End Sub
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim fn As String, fso As New Scripting.FileSystemObject, i As Long
    Dim newWb As Workbook, kupWb As Workbook, r As Range, rootPath As String, thisWb As Workbook, s As Worksheet
    Dim val As String

    If ThisWorkbook.Name = "NovaFaktura2.xlsm" Then Exit Sub

    If Sheet1.Range(FAKTURA_FIRMA_CELL) = "" Then
        MsgBox "Izberi firma !"
        Cancel = True
        Exit Sub
    End If
    If Sheet1.Range(FAKTURA_KRAJNA_SUMA_CELL) = 0 Then
        MsgBox "Vnesi artikli !"
        Cancel = True
        Exit Sub
    End If
    
    rootPath = ThisWorkbook.Path
    If Right(ThisWorkbook.Path, 6) = "arhiva" Then rootPath = Left(ThisWorkbook.Path, Len(ThisWorkbook.Path) - 7)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Set kupWb = Workbooks.Open(rootPath & "\Kupuvaci.xlsm")
    Set s = kupWb.Sheets(2)
    s.Unprotect
    val = ""
    Set r = s.Range(s.Cells(2, FIRMI_LOG_OPIS_COL), s.Cells(s.UsedRange.Rows.Count + 1, FIRMI_LOG_OPIS_COL)).Find(What:=Sheet1.Range(FAKTURA_BR_CELL) & "/", LookAt:=xlWhole)
    If r Is Nothing Then
        i = 2
        If s.UsedRange.Rows.Count > 1 Then
            i = s.UsedRange.Rows.Count + 1
            s.Cells(i, 1).Formula = "=A" & (i - 1) & "+1"
        Else
            s.Cells(i, 1) = 1
        End If
        s.Cells(i, FIRMI_LOG_DATA_COL).NumberFormat = "[$-42F]dddd, dd mmmm yyyy;@"
        s.Cells(i, FIRMI_LOG_DATA_COL) = Sheet1.Range(FAKTURA_DATA_CELL)
        s.Cells(i, FIRMI_LOG_FIRMA_COL).Font.Name = "MAC C Times"
        s.Cells(i, FIRMI_LOG_FIRMA_COL) = Sheet1.Range(FAKTURA_FIRMA_CELL)
        s.Cells(i, FIRMI_LOG_FIRMA_GRAD_COL).Font.Name = "MAC C Times"
        s.Cells(i, FIRMI_LOG_FIRMA_GRAD_COL) = Sheet1.Range(FAKTURA_FIRMA_GRAD_CELL)
        s.Cells(i, FIRMI_LOG_DOLZI_COL).NumberFormat = "#,##0 [$" & ChrW(1076) & ChrW(1077) & ChrW(1085) & ".-42F]"
        s.Cells(i, FIRMI_LOG_DOLZI_COL) = Sheet1.Range(FAKTURA_KRAJNA_SUMA_CELL)
        s.Cells(i, FIRMI_LOG_POBARUVA_COL).NumberFormat = "#,##0 [$" & ChrW(1076) & ChrW(1077) & ChrW(1085) & ".-42F]"
        s.Cells(i, FIRMI_LOG_POBARUVA_COL) = 0
        s.Cells(i, FIRMI_LOG_OPIS_COL).Font.Name = "MAC C Times"
        s.Cells(i, FIRMI_LOG_OPIS_COL) = Sheet1.Range(FAKTURA_BR_CELL) & "/"
        s.Cells(i, FIRMI_LOG_FIRMA_ID_COL).NumberFormat = ";;;"
        s.Cells(i, FIRMI_LOG_FIRMA_ID_COL) = Sheet1.Range(FAKTURA_FIRMA_ID_CELL)
        s.Range("A" & i & ":G" & i).Interior.ColorIndex = 24
    Else
        s.Cells(r.Row, FIRMI_LOG_DATA_COL) = Sheet1.Range(FAKTURA_DATA_CELL)
        s.Cells(r.Row, FIRMI_LOG_DOLZI_COL) = Sheet1.Range(FAKTURA_KRAJNA_SUMA_CELL)
    End If
    s.Protect
    kupWb.Save
    kupWb.Close False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Sheet1.Unprotect
    i = 0
    For j = FAKTURA_5DDV_ROW To FAKTURA_ARTIKLI_START_ROW - 1 Step -1
        If Sheet1.Cells(j, FAKTURA_KOLICINA_COL) <> "" Then
            i = j
            Exit For
        End If
    Next j
    If i > FAKTURA_ARTIKLI_START_ROW Then
        With Range(Cells(FAKTURA_ARTIKLI_START_ROW + 1, FAKTURA_REDEN_BR_COL), Cells(i, FAKTURA_IZNOS_VKUPEN_COL))
            .Borders(xlEdgeBottom).LineStyle = xlDash
            .Borders(xlEdgeBottom).Weight = xlHairline
            .Borders(xlInsideHorizontal).LineStyle = xlDash
            .Borders(xlInsideHorizontal).Weight = xlHairline
            .Borders(xlInsideVertical).LineStyle = xlDash
            .Borders(xlInsideVertical).Weight = xlHairline
        End With
    End If
    Sheet1.Protect
    
    If ThisWorkbook.Name <> "NovaFaktura.xlsm" Then Exit Sub
    
    fn = rootPath & "\arhiva\" & Sheet1.Range(FAKTURA_BR_CELL) & ".xlsm"
    fso.CopyFile ThisWorkbook.FullName, fn, True
    Set thisWb = ThisWorkbook
    Set newWb = Workbooks.Open(fn)
    thisWb.Sheets(1).Unprotect
    newWb.Sheets(1).Unprotect
    thisWb.Sheets(1).UsedRange.Copy newWb.Sheets(1).Range("A1")
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    newWb.Save
    newWb.Close
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ret = Shell("EXCEL """ & fn & """", 1)
    
    Cancel = True
    Application.DisplayAlerts = False
    Application.Quit
End Sub
Private Sub Workbook_Open()
    Dim a(1 To 5) As String, bw As Workbook, dictKey As Integer, fn As Integer
    Dim maxFn As Integer, rootPath As String, ws As Worksheet, v As Variant
    
    Sheet1.Unprotect
    Sheet1.EnableSelection = xlUnlockedCells
    Application.DisplayAlerts = True
    Application.WindowState = xlMaximized
    ActiveWindow.DisplayHeadings = False
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
    If Application.CommandBars.Item("Ribbon").Height > 80 Then SendKeys "^{F1}", True
    Application.OnKey "~", "NextCell"
    Application.OnKey "{ENTER}", "NextCell"
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    If Right(ThisWorkbook.Path, 6) = "arhiva" Then
        rootPath = Left(ThisWorkbook.Path, Len(ThisWorkbook.Path) - 7)
    Else
        rootPath = ThisWorkbook.Path
        v = GetFileList(Path & "\arhiva\*.xlsm")
        maxFn = 0
        If IsArray(v) Then
            For i = LBound(v) To UBound(v)
                fn = CInt(Split(v(i), ".")(0))
                If maxFn < fn Then maxFn = fn
            Next i
        End If
        Sheet1.Range(FAKTURA_BR_CELL).Value = maxFn + 1
    End If

    Set bw = Workbooks.Open(rootPath & "\Kupuvaci.xlsm")
    Set ws = bw.Sheets(1)
    Set dictFirmi = New Scripting.Dictionary
    dictFirmi(0) = a
    For Each c In ws.Range("B2:B" & ws.UsedRange.Rows.Count)
        If c = "" Then Exit For
        a(1) = c
        a(2) = ws.Cells(c.Row, FIRMA_GRAD_COL)
        dictFirmi(CInt(ws.Cells(c.Row, 1))) = a
    Next c
    bw.Close False

    Set bw = Workbooks.Open(rootPath & "\Artikli.xlsm")
    Set ws = bw.Sheets(1)
    Set dictArtikli = New Scripting.Dictionary
    Erase a
    dictArtikli(0) = a
    For Each c In ws.Range("B2:B" & ws.UsedRange.Rows.Count)
        If c = "" Then Exit For
        a(1) = c
        a(2) = ws.Cells(c.Row, ARTIKAL_EDINICNA_MERA_COL)
        a(3) = ws.Cells(c.Row, ARTIKAL_CENA_BEZ_DDV_COL)
        a(4) = ws.Cells(c.Row, ARTIKAL_DDV_COL)
        a(5) = ws.Cells(c.Row, ARTIKAL_CENA_SO_DDV)
        dictArtikli(CInt(ws.Cells(c.Row, 1))) = a
    Next c
    bw.Close False
    
    If ThisWorkbook.Name = "NovaFaktura.xlsm" Then
        Sheet1.Range(FAKTURA_FIRMA_CELL).Interior.ColorIndex = 44
        Sheet1.Range(FAKTURA_BR_CELLS).Locked = False
        Sheet1.Range(FAKTURA_BR_CELLS).Interior.ColorIndex = 6
        Sheet1.Range(FAKTURA_DATA_CELLS).Interior.ColorIndex = 44
        Sheet1.Range(FAKTURA_DATA_CELL) = Date
        Sheet1.Range(FAKTURA_ARTIKLI_CELLS).Interior.ColorIndex = 44
        Sheet1.Range(FAKTURA_KOLICINA_CELLS).Interior.ColorIndex = 6
        Sheet1.Range(FAKTURA_ROK_CELL).Interior.ColorIndex = 6
    Else
        Sheet1.Range(FAKTURA_FIRMA_CELL).Interior.ColorIndex = xlColorIndexNone
        Sheet1.Range(FAKTURA_BR_CELLS).Locked = True
        Sheet1.Range(FAKTURA_BR_CELLS).Interior.ColorIndex = xlColorIndexNone
        Sheet1.Range(FAKTURA_DATA_CELLS).Interior.ColorIndex = xlColorIndexNone
        Sheet1.Range(FAKTURA_ARTIKLI_CELLS).Interior.ColorIndex = xlColorIndexNone
        Sheet1.Range(FAKTURA_KOLICINA_CELLS).Interior.ColorIndex = xlColorIndexNone
        Sheet1.Range(FAKTURA_ROK_CELL).Interior.ColorIndex = xlColorIndexNone
    End If
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Sheet1.Protect
    ThisWorkbook.Saved = True
    If ThisWorkbook.Name = "NovaFaktura.xlsm" Then
        Sheet1.Range(FAKTURA_FIRMA_CELL).Select
        If frmFirmi.Visible = False Then frmFirmi.Show
        If GetKeyState(144) = 0 Then SendKeys "{NUMLOCK}", True
    End If
End Sub


'======================= frmArtikli ==========================
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As String, ByVal p4 As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, p4 As Any) As Long
Private Sub cmdIzberi_Click()
    SelectItem
End Sub
Private Sub lbArtikli_Click()
    tbInput.Value = lbArtikli.Value
End Sub
Private Sub lbArtikli_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    SelectItem
End Sub
Private Sub lbArtikli_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    tbInput_KeyDown KeyCode, Shift
End Sub
Private Sub lbArtikli_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll
End Sub
Private Sub tbInput_Change()
    Dim k As Variant, s As String

    Application.EnableEvents = False

    If tbInput.Value = "" Then
        lbArtikli.Value = ""
    ElseIf IsNumeric(tbInput.Value) Then
        If tbInput.Value = 0 Then
            lbArtikli.Value = ""
        ElseIf tbInput.Value < lbArtikli.ListCount Then
            lbArtikli.Value = tbInput.Value
        Else
            lbArtikli.Value = lbArtikli.ListCount - 1
        End If
    Else
        k = ""
        For Each key In ThisWorkbook.dictArtikli
            If InStr(UCase(ThisWorkbook.dictArtikli(key)(1)), UCase(tbInput.Value)) > 0 Then
                k = key
                Exit For
            End If
        Next key
        s = tbInput.Value
        lbArtikli.Value = k
        tbInput.Value = s
    End If

    Application.EnableEvents = True
End Sub
Private Sub tbInput_Enter()
    With tbInput
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub tbInput_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        tbInput_Change
        SelectItem
    ElseIf KeyCode = 27 Then
        Unload Me
        SetSheetFocus
    End If
End Sub
Private Sub UserForm_Initialize()
    With lbArtikli
        .clear
        For Each key In ThisWorkbook.dictArtikli
            .AddItem key
            .List(key, 1) = ThisWorkbook.dictArtikli(key)(1)
            .List(key, 2) = ThisWorkbook.dictArtikli(key)(5) & " den"
        Next key
        .List(0, 0) = ""
        .List(0, 2) = ""
    End With
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
     UnhookListBoxScroll
End Sub
Sub SelectItem()
    Dim k As Integer, v As Variant

    k = 0
    If IsNumeric(lbArtikli.Value) Then k = CInt(lbArtikli.Value)

    Application.EnableEvents = False
    Sheet1.Unprotect

    Sheet1.Cells(ActiveCell.Row, FAKTURA_EDINICNA_MERA_COL) = ThisWorkbook.dictArtikli(k)(2)
    v = ThisWorkbook.dictArtikli(k)(3)
    If v <> "" Then v = CDbl(v)
    Sheet1.Cells(ActiveCell.Row, FAKTURA_CENA_BEZ_DDV_COL) = v
    Sheet1.Cells(ActiveCell.Row, FAKTURA_DDV_COL) = ThisWorkbook.dictArtikli(k)(4)
    v = ThisWorkbook.dictArtikli(k)(5)
    If v <> "" Then v = CDbl(v)
    Sheet1.Cells(ActiveCell.Row, FAKTURA_CENA_SO_DDV_COL) = v
    ActiveCell = ThisWorkbook.dictArtikli(k)(1)

    Sheet1.Protect
    Application.EnableEvents = True

    Unload Me
    SetSheetFocus
    Sheet1.ChangeDanok
    ActiveCell.Next.Select
End Sub
Sub SetSheetFocus()
    SendMessage FindWindowEx(FindWindowEx(Application.hwnd, 0&, "XLDESK", vbNullString), 0&, "EXCEL7", ActiveWindow.Caption), &H7, 0&, 0&
End Sub


'======================= frmFirmi ==========================
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As String, ByVal p4 As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, p4 As Any) As Long
Private Sub cmdIzberi_Click()
    SelectItem
End Sub
Private Sub lbFirmi_Click()
    tbInput.Value = lbFirmi.Value
End Sub
Private Sub lbFirmi_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    SelectItem
End Sub
Private Sub lbFirmi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    tbInput_KeyDown KeyCode, Shift
End Sub
Private Sub lbFirmi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll
End Sub
Private Sub tbInput_Change()
    Dim k As Variant, s As String
    
    Application.EnableEvents = False
    
    If tbInput.Value = "" Then
        lbFirmi.Value = ""
    ElseIf IsNumeric(tbInput.Value) Then
        If tbInput.Value = 0 Then
            lbFirmi.Value = ""
        ElseIf tbInput.Value < lbFirmi.ListCount Then
            lbFirmi.Value = tbInput.Value
        Else
            lbFirmi.Value = lbFirmi.ListCount - 1
        End If
    Else
        k = ""
        For Each key In ThisWorkbook.dictFirmi
            If InStr(UCase(ThisWorkbook.dictFirmi(key)(1)), UCase(tbInput.Value)) > 0 Then
                k = key
                Exit For
            End If
        Next key
        s = tbInput.Value
        lbFirmi.Value = k
        tbInput.Value = s
    End If
    
    Application.EnableEvents = True
End Sub
Private Sub tbInput_Enter()
    With tbInput
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub tbInput_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        tbInput_Change
        SelectItem
    ElseIf KeyCode = 27 Then
        Unload Me
        SetSheetFocus
    End If
End Sub
Private Sub UserForm_Initialize()
    With lbFirmi
        .clear
        For Each key In ThisWorkbook.dictFirmi
            .AddItem key
            .List(key, 1) = ThisWorkbook.dictFirmi(key)(1)
            .List(key, 2) = ThisWorkbook.dictFirmi(key)(2)
        Next key
        .List(0, 0) = ""
    End With
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    UnhookListBoxScroll
End Sub
Sub SelectItem()
    Dim k As Integer
    
    k = 0
    If IsNumeric(lbFirmi.Value) Then k = CInt(lbFirmi.Value)
    
    Application.EnableEvents = False
    Sheet1.Unprotect
    
    Sheet1.Range(FAKTURA_FIRMA_ID_CELL) = k
    ActiveCell = ThisWorkbook.dictFirmi(k)(1)
    Sheet1.Range(FAKTURA_FIRMA_GRAD_CELL) = ThisWorkbook.dictFirmi(k)(2)
    
    Sheet1.Protect
    Application.EnableEvents = True
    
    Unload Me
    SetSheetFocus
    NextCell
    NextCell
End Sub
Sub SetSheetFocus()
    SendMessage FindWindowEx(FindWindowEx(Application.hwnd, 0&, "XLDESK", vbNullString), 0&, "EXCEL7", ActiveWindow.Caption), &H7, 0&, 0&
End Sub

'======================= frmKalendar ==========================
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As String, ByVal p4 As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, p4 As Any) As Long
Private Sub kalendar_Click()
    Sheet1.Range(FAKTURA_DATA_CELL) = kalendar.Value
    Unload Me
    SetSheetFocus
    Sheet1.Range(FAKTURA_DATA_CELL).Next.Select
End Sub
Private Sub kalendar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        kalendar_Click
    ElseIf KeyAscii = 27 Then
        Unload frmKalendar
        SetSheetFocus
    End If
End Sub
Sub SetSheetFocus()
    SendMessage FindWindowEx(FindWindowEx(Application.hwnd, 0&, "XLDESK", vbNullString), 0&, "EXCEL7", ActiveWindow.Caption), &H7, 0&, 0&
End Sub

'======================= General ==========================
Private Const GWL_HINSTANCE As Long = (-6)
Private Const HC_ACTION As Long = 0
Private Const VK_DOWN As Long = &H28
Private Const VK_UP As Long = &H26
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WH_MOUSE_LL As Long = 14
Private Const WM_MOUSEWHEEL As Long = &H20A

Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Type POINTAPI
     X As Long
     Y As Long
End Type
Private Type MOUSEHOOKSTRUCT
     pt As POINTAPI
     hwnd As Long
     wHitTestCode As Long
     dwExtraInfo As Long
End Type

Public Const ARTIKAL_CENA_BEZ_DDV_COL = 6
Public Const ARTIKAL_CENA_SO_DDV = 4
Public Const ARTIKAL_DDV_COL = 5
Public Const ARTIKAL_EDINICNA_MERA_COL = 3
Public Const FAKTURA_5DDV_ROW = 36
Public Const FAKTURA_18DDV_ROW = 37
Public Const FAKTURA_ARTIKLI_CELLS = "B12:B35"
Public Const FAKTURA_ARTIKLI_START_ROW = 12
Public Const FAKTURA_BR_CELL = "C8"
Public Const FAKTURA_BR_CELLS = "C8:D8"
Public Const FAKTURA_CENA_BEZ_DDV_COL = 5
Public Const FAKTURA_CENA_SO_DDV_COL = 8
Public Const FAKTURA_DATA_CELL = "F8"
Public Const FAKTURA_DATA_CELLS = "F8:H8"
Public Const FAKTURA_DDV_COL = 6
Public Const FAKTURA_EDINICNA_MERA_COL = 3
Public Const FAKTURA_FIRMA_CELL = "E2"
Public Const FAKTURA_FIRMA_CELLS = "E2:I4"
Public Const FAKTURA_FIRMA_GRAD_CELL = "E5"
Public Const FAKTURA_FIRMA_ID_CELL = "B7"
Public Const FAKTURA_IZNOS_NA_DDV_COL = 7
Public Const FAKTURA_IZNOS_VKUPEN_COL = 9
Public Const FAKTURA_KRAJNA_SUMA_CELL = "I38"
Public Const FAKTURA_KOLICINA_CELLS = "D12:D35"
Public Const FAKTURA_KOLICINA_COL = 4
Public Const FAKTURA_REDEN_BR_COL = 1
Public Const FAKTURA_ROK_CELL = "D40"
Public Const FIRMA_GRAD_COL = 3
Public Const FIRMI_LOG_DATA_COL = 2
Public Const FIRMI_LOG_DOLZI_COL = 5
Public Const FIRMI_LOG_FIRMA_COL = 3
Public Const FIRMI_LOG_FIRMA_GRAD_COL = 4
Public Const FIRMI_LOG_FIRMA_ID_COL = 8
Public Const FIRMI_LOG_OPIS_COL = 7
Public Const FIRMI_LOG_POBARUVA_COL = 6

Private mLngMouseHook As Long
Private mListBoxHwnd As Long
Private mbHook As Boolean
Function GetFileList(FileSpec As String) As Variant
    Dim FileArray() As Variant, FileCount As Integer, FileName As String
    
    FileCount = 0
    FileName = Dir(FileSpec)
    If FileName = "" Then GoTo NoFilesFound
    Do While FileName <> ""
        FileCount = FileCount + 1
        ReDim Preserve FileArray(1 To FileCount)
        FileArray(FileCount) = FileName
        FileName = Dir()
    Loop
    GetFileList = FileArray
    Exit Function
NoFilesFound:
    GetFileList = False
End Function
Function MouseProc(ByVal nCode As Long, ByVal wParam As Long, ByRef lParam As MOUSEHOOKSTRUCT) As Long
    On Error GoTo errH
    If (nCode = HC_ACTION) Then
        If WindowFromPoint(lParam.pt.X, lParam.pt.Y) = mListBoxHwnd Then
            If wParam = WM_MOUSEWHEEL Then
                MouseProc = True
                If lParam.hwnd > 0 Then
                    PostMessage mListBoxHwnd, WM_KEYDOWN, VK_UP, 0
                Else
                    PostMessage mListBoxHwnd, WM_KEYDOWN, VK_DOWN, 0
                End If
                PostMessage mListBoxHwnd, WM_KEYUP, VK_UP, 0
                Exit Function
            End If
        Else
            UnhookListBoxScroll
        End If
    End If
    MouseProc = CallNextHookEx(mLngMouseHook, nCode, wParam, ByVal lParam)
    Exit Function
errH:
    UnhookListBoxScroll
End Function
Sub HookListBoxScroll()
    Dim lngAppInst As Long
    Dim hwndUnderCursor As Long
    Dim tPT As POINTAPI
    GetCursorPos tPT
    hwndUnderCursor = WindowFromPoint(tPT.X, tPT.Y)
    If mListBoxHwnd <> hwndUnderCursor Then
        UnhookListBoxScroll
        mListBoxHwnd = hwndUnderCursor
        lngAppInst = GetWindowLong(mListBoxHwnd, GWL_HINSTANCE)
        PostMessage mListBoxHwnd, WM_LBUTTONDOWN, 0&, 0&
        If Not mbHook Then
            mLngMouseHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProc, lngAppInst, 0)
            mbHook = mLngMouseHook <> 0
        End If
     End If
End Sub
Sub NextCell()
    ActiveCell.Next.Select
End Sub
Sub UnhookListBoxScroll()
    If mbHook Then
        UnhookWindowsHookEx mLngMouseHook
        mLngMouseHook = 0
        mListBoxHwnd = 0
        mbHook = False
    End If
End Sub
Sub test()
    MsgBox frmArtikli.Top & " " & frmArtikli.Left
End Sub

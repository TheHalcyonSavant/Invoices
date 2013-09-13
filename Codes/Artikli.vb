'======================= ThisWorkbook ==========================
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Sheet1.Unprotect
    Application.ScreenUpdating = False
    
    'Dim mkdAlphabet As Variant, n As Long
    'mkdAlphabet = Array( _
    '    "a", "b", "v", "g", "d", "|", "e", "`", "z", "y", "i", "j", _
    '    "k", "l", "q", "m", "n", "w", "o", "p", "r", "s", "t", "}", _
    '    "u", "f", "h", "c", "~", "x", "{")
    'n = Application.GetCustomListNum(mkdAlphabet)
    'If Application.CustomListCount > 4 Then Application.DeleteCustomList n
    'Application.AddCustomList mkdAlphabet
    'n = Application.GetCustomListNum(mkdAlphabet)
    
    Sheet1.Range("B2", Sheet1.UsedRange.SpecialCells(xlCellTypeLastCell)).Sort Key1:=Sheet1.Range("B:B"), Key2:=Sheet1.Range("D:D")
        ', Ordercustom:=n
    
    Application.ScreenUpdating = True
    Sheet1.Protect
End Sub
Private Sub Workbook_Open()
    Sheet1.Unprotect
    Sheet1.EnableSelection = xlUnlockedCells
    Application.ScreenUpdating = False
    
    ActiveWindow.DisplayHeadings = False
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
    If Application.CommandBars.Item("Ribbon").Height > 80 Then SendKeys "^{F1}", True
    
    Application.ScreenUpdating = True
    Sheet1.Protect
    
    ThisWorkbook.Saved = True
End Sub

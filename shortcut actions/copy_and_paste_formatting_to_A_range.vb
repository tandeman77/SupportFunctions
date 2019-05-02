Option Explicit

Sub copyAndPasteFormatingToARange()
' copyAndPasteFormatingToARange Macro
    MsgBox ("this is how it works. 1. input the range you're copying from. 2. input the range you're pasting too(enter only the first cell of each range)")
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    Dim inputRange As String
    inputRange = InputBox("Type in the cells to copy from in a range. Eg. A1:C1")

    Dim originalRange As Range
    Set originalRange = ActiveSheet.Range(inputRange)
    originalRange.Select
    Selection.Copy


    Dim pasteRange As Range
    inputRange = InputBox("type in the range you need to paste to. e.g. C2:C10")
    Set pasteRange = ActiveSheet.Range(inputRange)
    Dim i As Variant

    For Each i In pasteRange
        i.Select
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
    Next i
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
End Sub

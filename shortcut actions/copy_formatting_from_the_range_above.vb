Option Explicit

Sub copyformatfromcellabove()
'
' copyformatfromcellabove Macro
'
' Keyboard Shortcut: Ctrl+Shift+F
'
    Dim rng As Range
    Set rng = Selection
    rng.Offset(-1, 0).Select
    Selection.Copy
    rng.Select
    Selection.PasteSpecial Paste:=xlPasteFormats

End Sub

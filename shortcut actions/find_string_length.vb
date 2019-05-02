Option Explicit

Sub getLength()
'
' getLength Macro
'
' Keyboard Shortcut: Ctrl+Shift+Q
'

Dim cell As Range
Set cell = Selection
MsgBox (Len(cell.text))
End Sub

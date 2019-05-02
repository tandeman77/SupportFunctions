Option Explicit

Sub RemoveDuplicateNoHeader()
'
' RemoveDuplicateNoHeader Macro
' 'remove Dupe
'
' Keyboard Shortcut: Ctrl+Shift+D
'
    Dim selectedRange As range
    Set selectedRange = Selection
    Dim columnCount As Integer
    columnCount = UBound(selectedRange.Value2, 2)
    ActiveSheet.range(selectedRange.Address).RemoveDuplicates Columns:=1, header:=xlNo
End Sub

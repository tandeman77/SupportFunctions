Option Explicit

Sub getPercentageChange()
    Dim oldValue As Variant
    Dim newValue As Variant
    oldValue = InputBox("original value = ?")
    newValue = InputBox("latest value = ?")

    MsgBox ((newValue - oldValue) / oldValue * 100 & "%")
End Sub

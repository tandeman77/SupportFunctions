Option Explicit

Sub dataInAColumn()
    Dim data As Variant
    data = Selection
    Dim i, j, k As Integer
    Dim output As Variant
    Dim outputRange As String
    outputRange = InputBox("what is the output sheet name?")

    k = 1
    ReDim output(1 To UBound(data) * UBound(data, 2), 1 To 1)
    For i = LBound(data) To UBound(data)
        For j = LBound(data, 2) To UBound(data, 2)
            output(k, 1) = data(i, j)
            k = k + 1
        Next j
    Next i

    Call pasteArrayToSheet(output, outputRange, 1, 1)
End Sub

Option Explicit

Sub HideEmptyRow()
    Dim rng As Range
    Set rng = Selection
    Dim i As Integer
    Dim cell As Variant
    For Each cell In rng
        If cell.Value2 = Empty Then
            cell.EntireRow.Hidden = True
        End If
    Next cell
End Sub

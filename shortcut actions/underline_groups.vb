Option Explicit
Sub underlineRowGroups()
    Dim inputRange As range
    Set inputRange = Selection
    Dim rangeAddress As String: rangeAddress = inputRange.Address
    Dim row, j As Variant
    Dim holder As String
    For Each row In inputRange.Rows
        If row.Value2 <> holder Then
            row.EntireRow.Select
            With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
            End With
        End If
        holder = row.Value2
    Next row
End Sub

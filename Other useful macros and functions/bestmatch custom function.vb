Function BestMatch(lookup_value As Range, lookup_range As Range) As Integer
    Dim cell As Variant
    Dim i, best As Integer
    i = 1
    Dim holder, ss As Double
    holder = 0

    For Each cell In lookup_range
        ss = Similarity(cell.Value, lookup_value.Value)
        If holder < ss Then
            holder = ss
            best = i
        End If
        i = i + 1
    Next cell
    BestMatch = best

End Function
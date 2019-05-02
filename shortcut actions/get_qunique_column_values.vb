Option Explicit

Sub getUniqueValuesOfEachColumn()
    Dim inputValues As Range
    Set inputValues = Selection
    Dim rowCount As Integer
    Dim columnCount As Integer
    rowCount = UBound(inputValues.Value2)
    columnCount = UBound(inputValues.Value2, 2)
    Dim holderArray As Variant
    Dim uniqueArray As Variant
    ReDim holderArray(1 To rowCount)
    Dim i As Integer
    Dim j As Long
    Dim outputSheet As String
    outputSheet = InputBox("what's the output sheet name?")
    Dim header As Integer
    header = InputBox("does you data include the header row? 1 = yes, 0 = no")
    Dim startingRow As Integer
    startingRow = InputBox("what is the first row you want to paste your data to?")

    For i = 1 To columnCount
        For j = 1 To rowCount
            holderArray(j) = inputValues(j, i)
        Next j
        uniqueArray = getUniqueValuesFromRange(holderArray)
        Call QuickSort(uniqueArray, header + 1, UBound(uniqueArray))
        uniqueArray = TransformArrayForExcelSheetWithStartingPoint(uniqueArray, 1)
        Call pasteArrayToSheet(uniqueArray, outputSheet, i, startingRow)
    Next i
End Sub

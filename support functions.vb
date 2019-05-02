
Public Function countDataRow(sheetName)
    'count how many rows of data there are in the active sheet
    Dim Ws As Worksheet
    Set Ws = Sheets(sheetName)
    Dim i As Integer
    i = 0

    Do While Ws.Cells(i + 1, 1) <> ""
        i = i + 1
    Loop

    countDataRow = i
End Function

Function getCorrectDate(dateText)
    Dim output As String
    Dim holder As Variant

    If VarType(dateText) = 7 Then
        holder = Split(dateText, "/")
        output = holder(1) & "/" & holder(0) & "/" & holder(2)
    ElseIf VarType(dateText) = 8 Then
        output = dateText
    End If

    getCorrectDate = output
End Function

Function TransformArrayForExcelSheet(inputArray As Variant)
    'get a 1d array ready to be transferred back to a spreadsheet
    Dim output As Variant
    ReDim output(1 To UBound(inputArray) + 1, 1 To 1)
    Dim i As Integer
    i = 0

    For i = 0 To UBound(inputArray)
        output(i + 1, 1) = inputArray(i)
    Next i
    TransformArrayForExcelSheet = output

End Function

Function TransformArrayForExcelSheetWithStartingPoint(inputArray As Variant, startingIndex As Integer) As Variant
    'get a 1d array ready to be transferred back to a spreadsheet
    Dim output As Variant
    ReDim output(1 To UBound(inputArray) + 1 - startingIndex, 1 To 1)
    Dim i As Integer
    i = 0

    For i = startingIndex To UBound(inputArray)
        output(i, 1) = inputArray(i)
    Next i

    TransformArrayForExcelSheetWithStartingPoint = output

End Function


Sub pasteArrayToSheet(outputArray As Variant, Sheet As String, columnNo, startingRow As Integer)
    'array needs to be 2 dimensional already

    Dim Ws As Worksheet
    Set Ws = Sheets(Sheet)
    Dim startColumn, endColumn As Variant
    startColumn = Number2Letter(columnNo)

    endColumn = Number2Letter(columnNo + UBound(outputArray, 2) - 1)
    'columnNo = UBound(outputArray, 2)
    Ws.range(startColumn & startingRow & ":" & endColumn & UBound(outputArray) + startingRow - 1) = outputArray
End Sub

Function Number2Letter(number As Variant) As String
    'PURPOSE: Convert a given number into it's corresponding Letter Reference
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
    'Convert To Column Letter
    Number2Letter = Split(Cells(1, number).Address, "$")(1)

End Function

Sub Letter2Number()
'PURPOSE: Convert a given letter into it's corresponding Numeric Reference
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim ColumnNumber As Long
Dim ColumnLetter As String

'Input Column Letter
  ColumnLetter = "AG"

'Convert To Column Number
   ColumnNumber = range(ColumnLetter & 1).Column

'Display Result
  MsgBox "Column " & ColumnLetter & " = Column " & ColumnNumber

End Sub
Function IsInArray2d(stringToBeFound, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i, 1) = stringToBeFound Then
            IsInArray2d = True
            Exit Function
        End If
    Next i
    IsInArray2d = False

End Function
Sub removePunctuationFromSelection()
    Dim inputValues As Variant
    Dim output As Variant
    inputValues = Selection
    ReDim output(1 To UBound(inputValues), 1 To 1)
    Dim i As Integer
    i = 0
    Stop
    Dim text As Variant

    For i = 1 To UBound(inputValues)
        If UBound(Split(inputValues(i, 1), " ")) > 0 Then
            text = Split(inputValues(i, 1), " ")
            output(i, 1) = removePunctuations(text)
        End If
    Next i

End Sub

Function removePunctuations(text As String) As String
    Dim punctuations As Object
    Set punctuations = CreateObject("vbscript.regexp")
    punctuations.Pattern = "\b(for|at|on|in|is|to|are|the|of|an|a)\b"
    punctuations.Global = True
    punctuations.ignorecase = True
    punctuations.MultiLine = True
    removePunctuations = punctuations.Replace(text, "")
End Function


Function IsInArray(stringToBeFound, arr As Variant) As Boolean
    Dim i As Long
    On Error GoTo emptyArray
        For i = LBound(arr) To UBound(arr)
            If arr(i) = stringToBeFound Then
                IsInArray = True
                Exit Function
            End If
        Next i
        IsInArray = False
emptyArray:
    IsInArray = False
End Function


Function removePunctuations2(textInput As String)
    Dim i As Integer
    i = 0
    Dim processString As Variant
    processString = Split(text)
    Dim punctuations As Variant
    punctuations = Array("for", "on", "in", "is", "to", "are", "the", "a", "an", "of", "at")

    For i = 0 To UBound(processString)
        If UBound(processString) > 0 Then
            If IsInArray(processString(i), punctuations) Then
                processString(i) = ""
            End If
        End If
    Next i
    removePunctuations = Trim(Join(processString, " "))

End Function

Function RemovePlurals(text As String) As String
    Dim re As Object
    Set re = CreateObject("vbscript.regexp")
    re.Pattern = "(s\b|es\b|ies\b)"
    re.Global = True
    re.ignorecase = True
    RemovePlurals = re.Replace(text, "")
End Function

Function RemovePluralsWithExceptions(text As String) As String
    Dim re As Object
    Set re = CreateObject("vbscript.regexp")
    re.Pattern = "(s\b|es\b|ies\b)"
    re.Global = True
    re.ignorecase = True
    RemovePlurals = re.Replace(text, "")
End Function

Function trimText(text As String) As String
    trimText = Trim(text)
End Function
Function trimtextinarray(inputArray As Variant) As Variant
    Dim i
    i = 0
    Dim output As Variant
    ReDim output(LBound(inputArray) To UBound(inputArray))
    For i = LBound(inputArray) To UBound(inputArray)
        output(i) = Trim(inputArray(i))
    Next i
    trimtextinarray = output
End Function

Function isIgnored(text As String, ignoreList As Variant) As Boolean
    'check if a word is a punctuation
    Dim i As Integer
    For i = 0 To UBound(ignoreList)
        If text = ignoreList(i) Then
            isIgnored = True
            Exit Function
        End If
    Next i
End Function

Function getUniqueValuesFromRange(inputArray As Variant) As Variant
    Dim i As Long
    Dim output As Variant
    For i = 1 To UBound(inputArray)
        If Not IsInArray(inputArray(i), output) Then
            output = Push(inputArray(i), output, 0)
        End If
    Next i
    getUniqueValuesFromRange = output
End Function

Function getUniqueValuesFromRange2d(inputArray As Variant) As Variant
    'col is the default from range =  1
    Dim i As Long
    Dim output As Variant
    For i = 1 To UBound(inputArray)
        If Not IsInArray(inputArray(i, 1), output) Then
            output = Push(inputArray(i), output, 1)
        End If
    Next i
    getUniqueValuesFromRange = output
End Function

Function Push(value As Variant, outputArray As Variant, defaultSecondDimension As Integer) As Variant
    On Error GoTo emptyArray
        Select Case defaultSecondDimension
            Case 0
                ReDim Preserve outputArray(LBound(outputArray) To UBound(outputArray) + 1)
                outputArray(UBound(outputArray)) = value
            Case 1
                ReDim Preserve outputArray(LBound(outputArray) To UBound(outputArray) + 1, 1 To 1)
                outputArray(UBound(outputArray), 1) = value

            Case 2
                ReDim Preserve outputArray(LBound(outputArray) To UBound(outputArray) + 1, 1 To 2)
                outputArray(UBound(outputArray), 2) = value
        End Select
    Push = outputArray
    Exit Function
emptyArray:
    Select Case defaultSecondDimension
        Case 0
            ReDim outputArray(1 To 1)
            outputArray(1) = value
        Case 1
            ReDim outputArray(1 To 1, 1 To 1)
            outputArray(1, 1) = value
        Case 2
            ReDim outputArray(1 To 1, 1 To defaultSecondDimension)
            outputArray(1, defaultSecondDimension) = value
    End Select
    Push = outputArray
End Function

Public Sub QuickSortArray(ByRef SortArray As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1, Optional lngColumn As Long = 0)
    On Error Resume Next

    'Sort a 2-Dimensional array

    ' SampleUsage: sort arrData by the contents of column 3
    '
    '   QuickSortArray arrData, , , 3

    '
    'Posted by Jim Rech 10/20/98 Excel.Programming

    'Modifications, Nigel Heffernan:

    '       ' Escape failed comparison with empty variant
    '       ' Defensive coding: check inputs

    Dim i As Long
    Dim j As Long
    Dim varMid As Variant
    Dim arrRowTemp As Variant
    Dim lngColTemp As Long

    If IsEmpty(SortArray) Then
        Exit Sub
    End If
    If InStr(TypeName(SortArray), "()") < 1 Then  'IsArray() is somewhat broken: Look for brackets in the type name
        Exit Sub
    End If
    If lngMin = -1 Then
        lngMin = LBound(SortArray, 1)
    End If
    If lngMax = -1 Then
        lngMax = UBound(SortArray, 1)
    End If
    If lngMin >= lngMax Then    ' no sorting required
        Exit Sub
    End If

    i = lngMin
    j = lngMax

    varMid = Empty
    varMid = SortArray((lngMin + lngMax) \ 2, lngColumn)

    ' We  send 'Empty' and invalid data items to the end of the list:
    If IsObject(varMid) Then  ' note that we don't check isObject(SortArray(n)) - varMid *might* pick up a valid default member or property
        i = lngMax
        j = lngMin
    ElseIf IsEmpty(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf IsNull(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf varMid = "" Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) = vbError Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) > 17 Then
        i = lngMax
        j = lngMin
    End If

    While i <= j
        While SortArray(i, lngColumn) < varMid And i < lngMax
            i = i + 1
        Wend
        While varMid < SortArray(j, lngColumn) And j > lngMin
            j = j - 1
        Wend

        If i <= j Then
            ' Swap the rows
            ReDim arrRowTemp(LBound(SortArray, 2) To UBound(SortArray, 2))
            For lngColTemp = LBound(SortArray, 2) To UBound(SortArray, 2)
                arrRowTemp(lngColTemp) = SortArray(i, lngColTemp)
                SortArray(i, lngColTemp) = SortArray(j, lngColTemp)
                SortArray(j, lngColTemp) = arrRowTemp(lngColTemp)
            Next lngColTemp
            Erase arrRowTemp

            i = i + 1
            j = j - 1
        End If
    Wend

    If (lngMin < j) Then Call QuickSortArray(SortArray, lngMin, j, lngColumn)
    If (i < lngMax) Then Call QuickSortArray(SortArray, i, lngMax, lngColumn)

End Sub

Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)
  'inlow and inhi are the boundaries of the array. you can use lbound and ubound for simplicity.
  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)
     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If
  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
End Sub

Function getWordCount(text As String) As Integer
    Dim oriLen As Integer
    oriLen = Len(text)
    Dim newlen As Integer
    newlen = Len(Replace(text, " ", ""))
    getWordCount = oriLen - newlen + 1
End Function
Function SubstituteArray(text, SubstituteArray, Replacement) As String
    Dim i As Variant
    SubstituteArray = text
    For Each i In SubstituteArray
        SubstituteArray = replce(SubstituteArray, i, "", 1, 1, vbTextCompare)
    Next i
 End Function

 Function properProper(text As String)
    Dim holder As Variant
    holder = Split(text, " ")

    Dim i As Integer
    For i = 0 To UBound(holder)
        If holder(i) = UCase(holder(i)) Then
            'do nothing
        Else
            holder(i) = StrConv(holder(i), 3)
        End If
    Next i

    Dim output As String
    output = Join(holder, " ")
    properProper = output
End Function

Function quickSortString(text As String) As String
    If InStr(1, " ", text) = 0 Then
        quickSortString = text
        Exit Function
    End If
    Dim process As Variant
    process = Split(text)
    quickSortString = QuickSort(process, LBound(process), UBound(process))
End Function
Public Function FindMax2D(arr As Variant, col As Long) As Long
  Dim myMax As Double
  Dim i As Long
  For i = LBound(arr, 1) To UBound(arr, 1)
    If arr(i, col) > myMax Then
      myMax = arr(i, col)
      FindMax2D = arr(i, col)
    End If
  Next i
End Function
Public Function FindMaxIndex2D(arr As Variant, col As Long) As Long
  Dim myMax As Double
  Dim i As Long
  For i = LBound(arr, 1) To UBound(arr, 1)
    If arr(i, col) > myMax Then
      myMax = arr(i, col)
      FindMaxIndex2D = i
    End If
  Next i
End Function
Public Function FindMax1D(arr As Variant) As Double
  Dim myMax As Double
  Dim i As Double

  For i = LBound(arr) To UBound(arr)
    If arr(i) > myMax Then
      myMax = i
    End If
  Next i
  FindMax1D = i
End Function

Function isPunctuation(text As String) As Boolean
    'check if a word is a punctuation
    Dim punctuations As Variant: punctuations = Array("for", "on", "in", "is", "to", "are", "the", "a", "an", "of", "at", "and", "with")
    Dim i As Integer
    For i = 0 To UBound(punctuations)
        If text = punctuations(i) Then
            isPunctuation = True
            Exit Function
        End If
    Next i
End Function


Function arrayTo1D(inputArray) As Variant
    Dim i As Integer
    Dim output As Variant: ReDim output(0 To UBound(inputArray) - LBound(inputArray))
    For i = LBound(inputArray) To UBound(inputArray)
        output(i - LBound(inputArray)) = inputArray(i, 1)
    Next i
    arrayTo1D = output
End Function

Function vbaTrim(textInput As String) As String
    Dim re As Object
    Set re = CreateObject("vbscript.regexp")
    re.Pattern = "\s+"
    re.Global = True
    re.ignorecase = True
    re.MultiLine = True
    vbaTrim = re.Replace(textInput, " ")

End Function
Function regexReplace(text, ReplaceString, replaceWith As String) As String
    Dim re As Object
    Set re = CreateObject("vbscript.regexp")
    re.Pattern = ("\b(" & ReplaceString & ")\b")
    re.Global = True
    re.ignorecase = True
    re.MultiLine = True
    regexReplace = re.Replace(text, replaceWith)
End Function

Function cleanUpLeftoverS(text) As String
' use this when after replacing a text leaves 's in the word
    Dim re As Object
    Set re = CreateObject("vbscript.regexp")
    re.Pattern = "\W(s|ies|es|s)\b"
    re.Global = True
    re.ignorecase = True
    re.MultiLine = True
    cleanUpLeftoverS = re.Replace(text, "")
End Function

Function VBAroundup(number) As Integer
    Dim newNo As Integer
    newNo = Int(number)
    If newNo < number Then
        VBAroundup = newNo + 1
    Else
        VBAroundup = number
    End If
End Function

Function TrimArray(arrayInput() As Variant, col As Integer) As Variant
    Dim i, count, j, k As Double
    Dim output As Variant
    For i = LBound(arrayInput) To UBound(arrayInput)
        If arrayInput(i, col) <> Empty Then
            count = count + 1
        Else
            ReDim output(LBound(arrayInput) To count, LBound(arrayInput, 2) To UBound(arrayInput, 2))
            For j = LBound(output) To UBound(output)
                For k = LBound(output, 2) To UBound(output, 2)
                    output(j, k) = arrayInput(j, k)
                Next k
            Next j
            TrimArray = output
            Exit Function
        End If
    Next i
End Function

Function regexExtract(text As String, regex As String) As String
    Dim re As Object
    Set re = CreateObject("vbscript.regexp")
    re.Pattern = regex
    re.Global = True
    re.ignorecase = True
    re.MultiLine = True
    regexExtract = re.Execute(text)(0)
End Function
Function AverageInArray1d(inputArray As Variant) As Double
    Dim i As Integer
    Dim total As Double
    total = 0
    For i = LBound(inputArray) To UBound(inputArray)
        total = inputArray(i) + total
    Next i
    AverageInArray1d = total / (UBound(inputArray) - LBound(inputArray) + 1)
End Function

Function sortRangeAndConcat(rng As range) As String
    Dim i As Variant
    Dim text As String
    For Each i In rng.value
        If i <> "" Then
            text = text & i
        End If
    Next i
    sortRangeAndConcat = text
End Function

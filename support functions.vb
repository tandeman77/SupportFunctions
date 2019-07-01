
Public Function countDataRow(sheetName)
    'count how many rows of data there are in the active sheet
    Dim ws As Worksheet
    Set ws = Sheets(sheetName)
    Dim i As Integer
    i = 0
    
    Do While ws.Cells(i + 1, 1) <> ""
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

    Dim ws As Worksheet
    Set ws = Sheets(Sheet)
    Dim startColumn, endColumn As Variant
    startColumn = Number2Letter(columnNo)
    
    endColumn = Number2Letter(columnNo + UBound(outputArray, 2) - 1)
    'columnNo = UBound(outputArray, 2)
    ws.Range(startColumn & startingRow & ":" & endColumn & UBound(outputArray) + startingRow - 1) = outputArray
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
   ColumnNumber = Range(ColumnLetter & 1).Column
   
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
    punctuations.IgnoreCase = True
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

Function RemovePlurals(text As String) As String
    Dim re As Object
    Set re = CreateObject("vbscript.regexp")
    re.Pattern = "(s\b|es\b|ies\b)"
    re.Global = True
    re.IgnoreCase = True
    RemovePlurals = re.Replace(text, "")
End Function

Function RemovePluralsWithExceptions(text As String) As String
    Dim re As Object
    Set re = CreateObject("vbscript.regexp")
    re.Pattern = "(s\b|es\b|ies\b)"
    re.Global = True
    re.IgnoreCase = True
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
    getUniqueValuesFromRange2d = output
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
    re.IgnoreCase = True
    re.MultiLine = True
    vbaTrim = re.Replace(textInput, " ")

End Function
Function regexReplace(text, ReplaceString, replaceWith As String) As String
    Dim re As Object
    Set re = CreateObject("vbscript.regexp")
    re.Pattern = ("\b(" & ReplaceString & ")\b")
    re.Global = True
    re.IgnoreCase = True
    re.MultiLine = True
    regexReplace = re.Replace(text, replaceWith)
End Function

Function cleanUpLeftoverS(text) As String
' use this when after replacing a text leaves 's in the word
    Dim re As Object
    Set re = CreateObject("vbscript.regexp")
    re.Pattern = "\W(s|ies|es|s)\b"
    re.Global = True
    re.IgnoreCase = True
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
            ReDim output(LBound(arrayInput) To count - 1 + LBound(arrayInput), LBound(arrayInput, 2) To UBound(arrayInput, 2))
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

Function trimArrayMultiD(inputarr As Variant, col As Integer) As Variant
    Dim row, i As Long
    Dim colc As Integer
    Dim output As Variant
    For i = LBound(inputarr) To UBound(inputarr)
        If inputarr(i, col) = Empty Then
            Exit For
        End If
    Next i
    ReDim output(LBound(inputarr) To i - 1, LBound(inputarr, 2) To UBound(inputarr, 2))
    For row = LBound(inputarr) To i - 1
        For colc = LBound(inputarr, 2) To UBound(inputarr, 2)
            output(row, colc) = inputarr(row, colc)
        Next colc
    Next row
    trimArrayMultiD = output
End Function

Function regexExtract(text As String, regex As String) As String
    Dim re As Object
    Set re = CreateObject("vbscript.regexp")
    re.Pattern = regex
    re.Global = True
    re.IgnoreCase = True
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

Function sortRangeAndConcat(rng As Range) As String
    Dim i As Variant
    Dim text As String
    For Each i In rng.value
        If i <> "" Then
            text = text & i
        End If
    Next i
    sortRangeAndConcat = text
End Function

Function indexOf(arr, value) As Long
    'find the location of a value in an array
    Dim val As Variant
    Dim count As Long
    count = LBound(arr)
    For Each val In arr
        If val = value Then
            indexOf = count
            Exit Function
        End If
        count = count + 1
    Next val
    indexOf = -1
End Function

Function indexOfInCol(arr, value, row) As Long
    'find the location of a value in an array
    Dim i As Integer
    For i = LBound(arr, 2) To UBound(arr, 2)
        If arr(row, i) = value Then
            indexOfInCol = i
            Exit Function
        End If
    Next i
End Function

Function AddToArrayWhereValueIsEmpty(arr, value, col) As Variant
    'find an empty spot in an array and add a given value
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i, col) = Empty Then
            arr(i, col) = value
            AddToArrayWhereValueIsEmpty = arr
            Exit Function
        End If
    Next i
End Function

Function invertArray(arr As Variant) As Variant
    'invert array
    Dim val As Variant
    Dim i, count As Long
    Dim j As Integer
    Dim output As Variant
    ReDim output(LBound(arr) To UBound(arr), LBound(arr, 2) To UBound(arr, 2))
    count = UBound(arr)
    For i = LBound(arr) To UBound(arr)
        For j = LBound(arr, 2) To UBound(arr, 2)
        output(i, j) = arr(count, j)
        Next j
        count = count - 1
    Next i
    invertArray = output
End Function

Function arrayLoading(newArr, oldArr As Variant, newCol As Integer, oldCol As Integer, redimTF As Boolean, Optional dimension As Integer) As Variant
    Dim i As Long
    If redimTF Then
        ReDim newArr(LBound(oldArr) To UBound(oldArr), LBound(oldArr, 2) To LBound(oldArr, 2) + dimension)
    End If
    For i = LBound(oldArr) To UBound(oldArr)
        newArr(i, newCol) = oldArr(i, oldCol)
    Next i
    arrayLoading = newCol
End Function

Function LoadMultiArrayTo1dArray(inputarr As Variant, outputArr As Variant, col As Integer, Optional startingIndex As Integer, Optional startingRow As Integer) As Variant
    'get a column of a multidimensional array into 1d array.
    'use startingrow if you have to skip a header row in the data or something like that.
    ' startingindex is for the output array, e.g. starting the array at index 0 or 1.
    Dim i, j, k As Long
    If startingRow = Empty Then
        k = 0
    End If
    If startingIndex = Empty Then
        j = 0
    End If
    ReDim outputArr(UBound(inputarr) + startingIndex - startingRow - 1)
    For i = LBound(inputarr) + k To UBound(inputarr)
        outputArr(j) = inputarr(i, col)
        j = j + 1
    Next i
    LoadMultiArrayTo1dArray = outputArr
End Function

Function BubbleSort2D(ByVal List As Variant, ByVal SortCol As Long, Optional ByVal SortColNumeric As Boolean = False, _
Optional ByVal Order As XlSortOrder = xlAscending) As Variant
' Sorts an array using bubble sort algorithm
    Dim First As Integer, Last As Integer
    Dim i As Integer, j As Integer, k As Long
    Dim iColumn
    Dim Temp
    First = LBound(List, 1)
    Last = UBound(List, 1)
    iColumn = LBound(List, 2) + SortCol - 1
    For i = First To Last - 1
        For j = i + 1 To Last
            If Order = xlAscending Then
                If SortColNumeric Then
                    If CDbl(List(i, iColumn)) > CDbl(List(j, iColumn)) Then
                        For k = LBound(List, 2) To UBound(List, 2)
                            Temp = List(j, k)
                            List(j, k) = List(i, k)
                            List(i, k) = Temp
                        Next k
                    End If
                Else
                    If List(i, iColumn) > List(j, iColumn) Then
                        For k = LBound(List, 2) To UBound(List, 2)
                            Temp = List(j, k)
                            List(j, k) = List(i, k)
                            List(i, k) = Temp
                        Next k
                    End If
                End If
            Else
                If SortColNumeric Then
                    If CDbl(List(i, iColumn)) < CDbl(List(j, iColumn)) Then
                        For k = LBound(List, 2) To UBound(List, 2)
                            Temp = List(j, k)
                            List(j, k) = List(i, k)
                            List(i, k) = Temp
                        Next k
                    End If
                Else
                    If List(i, iColumn) < List(j, iColumn) Then
                        For k = LBound(List, 2) To UBound(List, 2)
                            Temp = List(j, k)
                            List(j, k) = List(i, k)
                            List(i, k) = Temp
                        Next k
                    End If
                End If
            End If
        Next j
    Next i
    BubbleSort2D = List
End Function

Function getRootKeyword(ByVal text As String) As String
    getRootKeyword = Trim(Replace(text, "+", ""))
End Function

Function countValueInArray(arr, value, col) As Long
    Dim count As Long
    count = 0
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If LCase(arr(i, col)) = LCase(value) Then
            count = count + 1
        End If
    Next i
    countValueInArray = count
End Function
Function multipleRegexExtract(text As String, regexPattern As String) As Variant
    Dim regex As New VBScript_RegExp_55.RegExp
    Dim match As Variant
    regex.Pattern = regexPattern
    regex.Global = True
    Set match = regex.Execute(text)
    Dim item, output As Variant
    ReDim output(match.count - 1)
    Dim i As Integer
    i = 0
    For Each item In match
        output(i) = item.value
        i = i + 1
    Next item
    multipleRegexExtract = output
End Function
Sub getStringInDictionary(arr As Variant, ByRef dict As Scripting.Dictionary)
    Dim text As Variant
    For Each text In arr
    dict.CompareMode = TextCompare
        If Not dict.Exists(text) Then
            dict.Add text, 1
        End If
    Next text
End Sub
Sub splitStringIn2WordPhrase(ByVal dict, originalString As Variant, delimiter)
    Dim i As Variant
    Dim text, textHolder As String
    Dim regex As New VBScript_RegExp_55.RegExp
    Dim match As Variant
    regex.Pattern = "\w+\s\w+"
    regex.Global = True
    
    text = originalString
    Set match = regex.Execute(text)
    
    For Each i In match
        dict(i.value) = 1
    Next i
    
    'shift 1 to the right
    text = Right(originalString, Len(originalString) - InStr(1, originalString, delimiter))
    Set match = regex.Execute(text)
    For Each i In match
        dict(i.value) = 1
    Next i
End Sub

Sub splitStringIn3WordPhrase(ByVal dict, originalString As Variant, delimiter)
    Dim i As Variant
    Dim text As String
    Dim regex As New VBScript_RegExp_55.RegExp
    Dim match As Variant
    regex.Pattern = "\w+\s\w+\s\w+"
    regex.Global = True
    
    text = originalString
    Set match = regex.Execute(text)
    
    For Each i In match
        dict(i.value) = 1
    Next i
    
    'shift 1 to the right
    text = Right(originalString, Len(originalString) - InStr(1, originalString, delimiter))
    Set match = regex.Execute(text)
    
    For Each i In match
        dict(i.value) = 1
    Next i
    
    text = Right(originalString, Len(text) - InStr(1, text, delimiter))
    Set match = regex.Execute(text)
    
    For Each i In match
        dict(i.value) = 1
    Next i
End Sub


Function regexReplaceRangeArray(text As Variant, ReplaceString As Range, replaceWith As Range) As String
    'replace regex with something.
    'inputs need to be a range
    Dim re As Object
    Set re = CreateObject("vbscript.regexp")
    re.Global = True
    re.IgnoreCase = True
    re.MultiLine = True
    Dim i As Integer
    Dim output As String
    output = text
    For i = 1 To ReplaceString.count
    re.Pattern = ReplaceString.Value2(i, 1)
    output = re.Replace(output, replaceWith.Value2(i, 1))
    Next i
    regexReplaceRangeArray = output
End Function
Function preventDivideByZeroError(number1, number2) As Double
    'when dividing 2 numbers with a tendency to divide by 0.
    'this save you from using on error in your code.
    On Error GoTo error
        preventDivideByZeroError = number1 / number2
        Exit Function
error:
    If Err.number = 11 Then
        Resume Next
    Else
        Err.Message
    End If
End Function

Function SubstituteArray(text, SubstituteArray, Replacement) As String
   Dim i As Variant
   SubstituteArray = text
   For Each i In SubstituteArray
       SubstituteArray = replce(SubstituteArray, i, "", 1, 1, vbTextCompare)
   Next i
End Function


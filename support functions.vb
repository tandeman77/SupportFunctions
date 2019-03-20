'===================================================================================================='
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
'===================================================================================================='
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
'===================================================================================================='
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
'===================================================================================================='
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

'===================================================================================================='
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
'===================================================================================================='
Function Number2Letter(number As Variant) As String
    'PURPOSE: Convert a given number into it's corresponding Letter Reference
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
    'Convert To Column Letter
    Number2Letter = Split(Cells(1, number).Address, "$")(1)
  
End Function
'===================================================================================================='
Sub Letter2Number()
'PURPOSE: Convert a given letter into it's corresponding Numeric Reference
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim ColumnNumber As Long
Dim ColumnLetter As String

'Input Column Letter
  ColumnLetter = "AG"
  
'Convert To Column Number
   ColumnNumber = Range(ColumnLetter & 1).column
   
'Display Result
  MsgBox "Column " & ColumnLetter & " = Column " & ColumnNumber
    
End Sub

'===================================================================================================='
Function IsInArray2d(stringToBeFound, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i, 1) = stringToBeFound Then
            IsInArray2d = True
            Exit Function
        End If
    Next i
    IsInArray2d = False
'===================================================================================================='
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
'===================================================================================================='
Function removePunctuations(text As String) As String
    Dim punctuations As Object
    Set punctuations = CreateObject("vbscript.regexp")
    punctuations.Pattern = "\b(for|at|on|in|is|to|are|the|of|an|a)\b"
    punctuations.Global = True
    punctuations.IgnoreCase = True
    punctuations.MultiLine = True
    removePunctuations = punctuations.Replace(text, "")
End Function

'===================================================================================================='
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
'===================================================================================================='

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
'===================================================================================================='
Function RemovePlurals(text As String) As String
    Dim RE As Object
    Set RE = CreateObject("vbscript.regexp")
    RE.Pattern = "(s\b|es\b|ies\b)"
    RE.Global = True
    RE.IgnoreCase = True
    RemovePlurals = RE.Replace(text, "")
End Function
'===================================================================================================='
Function trimText(text As String) As String
    trimText = Trim(text)
End Function

'===================================================================================================='
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
'===================================================================================================='
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
'===================================================================================================='
Function getUniqueValuesFromRange(inputArray As Variant) As Variant
    Dim i As Long
    Dim output As Variant
    For i = 1 To UBound(inputArray)
        If Not IsInArray(inputArray(i), output) Then
            output = push1D(inputArray(i), output, 0)
        End If
    Next i
    getUniqueValuesFromRange = output
End Function
'===================================================================================================='
Function push1D(value As Variant, outputArray As Variant, defaultSecondDimension As Integer) As Variant
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
    push1D = outputArray
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
    push1D = outputArray
End Function
'===================================================================================================='
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
'===================================================================================================='
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



dim fullAddress as stringToBeFound
dim LocAddress as variant
locAddress = json("results")(1)("address_components")
dim component as variant
for each component in locAddress
    
next component

'===================================================================================================='
'border for rows and columns
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
If Target.Cells.Count > 1 Then Exit Sub
Application.ScreenUpdating = False
'Clear all cell colours
Cells.Borders.LineStyle = xlNone

With Target.EntireRow.Borders(xlEdgeTop)
'Row and column highlightingfor the active cell
.LineStyle = xlEdgeBottom
.Weight = xlMedium
.Color = -16776961
End With

    With Target.EntireColumn.Borders(xlEdgeTop)
.LineStyle = xlEdgeBottom
.Weight = xlMedium
.Color = -16776961
    End With

    With Target.EntireRow.Borders(xlEdgeBottom)
    'Row and column highlightingfor the active cell
    .LineStyle = xlEdgeBottom
    .Weight = xlMedium
    .Color = -16776961
    End With

    With Target.EntireColumn.Borders(xlEdgeBottom)
    .LineStyle = xlEdgeBottom
    .Weight = xlMedium
    .Color = -16776961
    End With

    With Target.EntireRow.Borders(xlEdgeRight)
    'Row and column highlightingfor the active cell
    .LineStyle = xlEdgeBottom
    .Weight = xlMedium
    .Color = -16776961
    End With

    With Target.EntireColumn.Borders(xlEdgeRight)
    .LineStyle = xlEdgeBottom
    .Weight = xlMedium
    .Color = -16776961
    End With

    With Target.EntireRow.Borders(xlEdgeLeft)
    'Row and column highlightingfor the active cell
    .LineStyle = xlEdgeBottom
    .Weight = xlMedium
    .Color = -16776961
    End With

    With Target.EntireColumn.Borders(xlEdgeLeft)
    .LineStyle = xlEdgeBottom
    .Weight = xlMedium
    .Color = -16776961
    End With

Application.ScreenUpdating = True
End Sub

'===================================================================================================='
Sub copyAndPasteFormatingToARange()
' copyAndPasteFormatingToARange Macro
    MsgBox ("this is how it works. 1. input the range you're copying from. 2. input the range you're pasting too(enter only the first cell of each range)")
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    Dim inputRange As String
    inputRange = InputBox("Type in the cells to copy from in a range. Eg. A1:C1")

    Dim originalRange As Range
    Set originalRange = ActiveSheet.Range(inputRange)
    originalRange.Select
    Selection.Copy


    Dim pasteRange As Range
    inputRange = InputBox("type in the range you need to paste to. e.g. C2:C10")
    Set pasteRange = ActiveSheet.Range(inputRange)
    Dim i As Variant

    For Each i In pasteRange
        i.Select
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
    Next i
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
End Sub

'===================================================================================================='
'optionsation setting for macro
Sub macroOptimisationToggle()
    If Application.EnableEvents Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.DisplayAlerts = False
    Else
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayAlerts = True
    End If
End Sub

'===================================================================================================='

Sub MergeCells()
' MergeCells Macro
' Keyboard Shortcut: Ctrl+Shift+M
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
End Sub

'===================================================================================================='
Sub MergeCellsInARange()
'
' MergeCellsInARange Macro
'

'
Dim rng As Range
Set rng = Selection
Stop
rng.FormulaR1C1Local(1).Select
Selection.Merge
Stop
    Range("C169:E169").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("C170:E170").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("C171:E171").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("C172:E172").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
End Sub

'===================================================================================================='
Sub modifyTextInARange()
    'adjust all strings in a range to have proper capitalisation
    ' can only do 1 column at a time
    '============================================================================================================================
    'save first just in case.
    Application.ActiveWorkbook.Save
    '============================================================================================================================
    'get variables ready for maco
    Dim userInput As String
    userInput = InputBox("what do you want to do?" & vbNewLine & _
    "input 1 for proper text" & vbNewLine & _
    "input 2 to trim text" & vbNewLine & _
    "input 3 for lower case text" & vbNewLine & _
    "input 4 for UPPER CASE text" & vbNewLine & _
    "input 5 append(front) something to the existing text" & vbNewLine & _
    "input 6 append(back) something to the existing text" & vbNewLine & _
    "input 7 for highlight above/below average value" & vbNewLine & _
    "input 8 for highlight value higher/lower than a specific number")

    Dim inputRange As Variant
    inputRange = Selection
    Dim output As Variant
    ReDim output(1 To UBound(inputRange), 1 To 1)
    Dim i As Variant
    Dim j As Long
    j = 1
    Dim compareDecision As String
    Dim selectionCol As Integer
    selectionCol = Selection.Column
    Dim highlightColour As String
    Dim checkValue As Variant
    '============================================================================================================================
    Select Case userInput
        '========================================
        Case 1
            For Each i In inputRange
                output(j, 1) = Application.WorksheetFunction.Proper(i)
                j = j + 1
            Next i
        '========================================
        Case 2
            For Each i In inputRange
                output(j, 1) = Application.WorksheetFunction.Trim(i)
                j = j + 1
            Next i
        '========================================
        Case 3
            For Each i In inputRange
                output(j, 1) = LCase(i)
                j = j + 1
            Next i
        '========================================
        Case 4
            For Each i In inputRange
                output(j, 1) = UCase(i)
                j = j + 1
            Next i
        '========================================
        Case 5
            Dim text As String
            text = InputBox("what text to append")
            For Each i In inputRange
                output(j, 1) = text & i
                j = j + 1
            Next i
        '========================================
        Case 6
            Dim text1 As String
            text = InputBox("what text to append")
            For Each i In inputRange
                output(j, 1) = i & text1
                j = j + 1
            Next i
        '========================================
        Case 7
            ' highlight cells with above/below average value.
            Dim average As Variant
            average = Application.WorksheetFunction.average(inputRange)
            compareDecision = InputBox("input 1 to highlight above average and 2 for below average")
            j = Selection.Row
            highlightColour = InputBox("choose from:" & vbNewLine & "1 = red" & vbNewLine & "2 = yellow" & vbNewLine & "3 = green")
            For Each i In inputRange
                Select Case compareDecision
                    Case 1
                        If i > average Then
                            Select Case highlightColour
                                Case 1
                                    ActiveSheet.Cells(j, selectionCol).Select
                                    With Selection.Interior
                                        .Pattern = xlSolid
                                        .PatternColorIndex = xlAutomatic
                                        .Color = 255
                                        .TintAndShade = 0
                                        .PatternTintAndShade = 0
                                    End With
                                Case 2
                                    ActiveSheet.Cells(j, selectionCol).Select
                                    With Selection.Interior
                                        .Pattern = xlSolid
                                        .PatternColorIndex = xlAutomatic
                                        .Color = 65535
                                        .TintAndShade = 0
                                        .PatternTintAndShade = 0
                                    End With
                                Case 3
                                    ActiveSheet.Cells(j, selectionCol).Select
                                    With Selection.Interior
                                        .Pattern = xlSolid
                                        .PatternColorIndex = xlAutomatic
                                        .Color = 5287936
                                        .TintAndShade = 0
                                        .PatternTintAndShade = 0
                                    End With
                            End Select
                        End If
                    Case 2
                        If i < average Then
                            Select Case highlightColour
                                Case 1
                                    ActiveSheet.Cells(j, selectionCol).Select
                                    With Selection.Interior
                                        .Pattern = xlSolid
                                        .PatternColorIndex = xlAutomatic
                                        .Color = 255
                                        .TintAndShade = 0
                                        .PatternTintAndShade = 0
                                    End With
                                Case 2
                                    ActiveSheet.Cells(j, selectionCol).Select
                                    With Selection.Interior
                                        .Pattern = xlSolid
                                        .PatternColorIndex = xlAutomatic
                                        .Color = 65535
                                        .TintAndShade = 0
                                        .PatternTintAndShade = 0
                                    End With
                                Case 3
                                    ActiveSheet.Cells(j, selectionCol).Select
                                    With Selection.Interior
                                        .Pattern = xlSolid
                                        .PatternColorIndex = xlAutomatic
                                        .Color = 5287936
                                        .TintAndShade = 0
                                        .PatternTintAndShade = 0
                                    End With
                            End Select
                        End If
                End Select
                j = j + 1
            Next i
            Exit Sub
        '========================================
        Case 8
            ' highlight cells with above/below a specific value.
            compareDecision = InputBox("input 1 to highlight higher values and 2 to highlight lower values")
            checkValue = InputBox("what is the value?")
            j = Selection.Row
            highlightColour = InputBox("choose from:" & vbNewLine & "1 = red" & vbNewLine & "2 = yellow" & vbNewLine & "3 = green")
            For Each i In inputRange
                Select Case compareDecision
                    Case 1
                        If i > Int(checkValue) Then
                            Select Case highlightColour
                                Case 1
                                    ActiveSheet.Cells(j, selectionCol).Select
                                    With Selection.Interior
                                        .Pattern = xlSolid
                                        .PatternColorIndex = xlAutomatic
                                        .Color = 255
                                        .TintAndShade = 0
                                        .PatternTintAndShade = 0
                                    End With
                                Case 2
                                    ActiveSheet.Cells(j, selectionCol).Select
                                    With Selection.Interior
                                        .Pattern = xlSolid
                                        .PatternColorIndex = xlAutomatic
                                        .Color = 65535
                                        .TintAndShade = 0
                                        .PatternTintAndShade = 0
                                    End With
                                Case 3
                                    ActiveSheet.Cells(j, selectionCol).Select
                                    With Selection.Interior
                                        .Pattern = xlSolid
                                        .PatternColorIndex = xlAutomatic
                                        .Color = 5287936
                                        .TintAndShade = 0
                                        .PatternTintAndShade = 0
                                    End With
                            End Select
                        End If
                    Case 2
                        If i < Int(checkValue) Then
                            Select Case highlightColour
                                Case 1
                                    ActiveSheet.Cells(j, selectionCol).Select
                                    With Selection.Interior
                                        .Pattern = xlSolid
                                        .PatternColorIndex = xlAutomatic
                                        .Color = 255
                                        .TintAndShade = 0
                                        .PatternTintAndShade = 0
                                    End With
                                Case 2
                                    ActiveSheet.Cells(j, selectionCol).Select
                                    With Selection.Interior
                                        .Pattern = xlSolid
                                        .PatternColorIndex = xlAutomatic
                                        .Color = 65535
                                        .TintAndShade = 0
                                        .PatternTintAndShade = 0
                                    End With
                                Case 3
                                    ActiveSheet.Cells(j, selectionCol).Select
                                    With Selection.Interior
                                        .Pattern = xlSolid
                                        .PatternColorIndex = xlAutomatic
                                        .Color = 5287936
                                        .TintAndShade = 0
                                        .PatternTintAndShade = 0
                                    End With
                            End Select
                        End If
                End Select
                j = j + 1
            Next i
            Exit Sub
        '========================================
        Case Else
            MsgBox ("invalid input action")
            Exit Sub
    End Select
    '============================================================================================================================
    'write output
    Selection = output
    MsgBox ("script completed")
End Sub


'===================================================================================================='
Sub removeKeywordsWithPlusSign()
'
' Keyword_adjustment_Monthly Macro
'

'
    Cells.Find(What:="keyword", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.Replace What:="=", Replacement:="+", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="++", Replacement:="+", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub

'===================================================================================================='
Sub RemoveAllHyperLinks()
'
' RemoveAllHyperlinks Macro
'

'
    Cells.Select
    Selection.Hyperlinks.Delete
End Sub

'===================================================================================================='
Function removePlurals(text As String) As String
    Dim final As String
    Dim regex As Object
    Set regex = CreateObject("vbscript.RegExp")
    regex.Pattern = "(ies\b)|(es\b)|(s\b)"
    regex.Global = True
    regex.IgnoreCase = False

    removePlurals = regex.Replace(text, "")

End Function


'===================================================================================================='
'This is a code bracket to trigger macro when a document is saved.
'put the code in "ThisWorkbook"

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'Your Code Here
End Sub

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
                output(j, 1) = properProper(CStr(i))
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
            text1 = InputBox("what text to append")
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
            j = Selection.row
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
            j = Selection.row
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

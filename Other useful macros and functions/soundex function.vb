Option Explicit

Function SOUNDEX(Surname As String) As String
' Developed by Richard J. Yanco
' This function follows the Soundex rules given at
' http://home.utah-inter.net/kinsearch/Soundex.html
'http://spreadsheetpage.com/index.php/tip/searching_using_soundex_codes/

    Dim Result As String, c As String * 1
    Dim Location As Integer

    Surname = UCase(Surname)

'   First character must be a letter
    If Asc(Left(Surname, 1)) < 65 Or Asc(Left(Surname, 1)) > 90 Then
        SOUNDEX = ""
        Exit Function
    Else
'       St. is converted to Saint
        If Left(Surname, 3) = "ST." Then
            Surname = "SAINT" & Mid(Surname, 4)
        End If

'       Convert to Soundex: letters to their appropriate digit,
'         A,E,I,O,U,Y ("slash letters") to slashes
'         H,W, and everything else to zero-length string

        Result = Left(Surname, 1)
        For Location = 2 To Len(Surname)
            Result = Result & Category(Mid(Surname, Location, 1))
        Next Location

'       Remove double letters
        Location = 2
        Do While Location < Len(Result)
            If Mid(Result, Location, 1) = Mid(Result, Location + 1, 1) Then
                Result = Left(Result, Location) & Mid(Result, Location + 2)
            Else
                Location = Location + 1
            End If
        Loop

'       If category of 1st letter equals 2nd character, remove 2nd character

        If Category(Left(Result, 1)) = Mid(Result, 2, 1) Then
            Result = Left(Result, 1) & Mid(Result, 3)
        End If

'       Remove slashes
        For Location = 2 To Len(Result)
            If Mid(Result, Location, 1) = "/" Then
                Result = Left(Result, Location - 1) & Mid(Result, Location + 1)
            End If
        Next

'       Trim or pad with zeroes as necessary
        Select Case Len(Result)
            Case 4
                SOUNDEX = Result
            Case Is < 4
                SOUNDEX = Result & String(4 - Len(Result), "0")
            Case Is > 4
                SOUNDEX = Left(Result, 4)
        End Select
    End If
End Function

Private Function Category(c) As String
'   Returns a Soundex code for a letter
    Select Case True
        Case c Like "[AEIOUY]"
            Category = "/"
        Case c Like "[BPFV]"
            Category = "1"
        Case c Like "[CSKGJQXZ]"
            Category = "2"
        Case c Like "[DT]"
            Category = "3"
        Case c = "L"
            Category = "4"
        Case c Like "[MN]"
            Category = "5"
        Case c = "R"
            Category = "6"
        Case Else 'This includes H and W, spaces, punctuation, etc.
            Category = ""
    End Select
End Function

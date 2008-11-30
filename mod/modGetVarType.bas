Attribute VB_Name = "modGetVarType"
Option Explicit

'Creators are Sebastian Ewe and Ruben Wisniewski

Function getVarType(Var As Variant) As Byte

    'Bedeutung der ausgegebenen Zahlen :
    ' 1 = positive Ganzzahl
    ' 2 = positive Kommazahl
    ' 3 = boolsche Zahl
    ' 4 = Datum
    ' 5 = Uhrzeit
    ' 6 = Feld
    ' 7 = negative Ganzzahl
    ' 8 = negative Kommazahl
    ' 9 = leer
    '98 = 0
    '99 = Text
    
    Dim Pattern As Variant
    Dim oRegex As RegExp
    Dim m As Match
    Dim counter As Boolean
    Dim i As Integer
    Dim RegEx(4) As String
    
    'Definition: was gilt als Zahl?
    RegEx(0) = "^[-+]{0,1}(\d{1,3}[.]){0,1}(\d{3}[.]){0,}\d{3}([,]\d{0,}){0,1}$"    'Pattern With Dots
    RegEx(1) = "^[-+]{0,1}\d{0,}[.]\d{0,}$"                                         'Pattern With Dot
    RegEx(2) = "^[-+]{0,1}\d{0,}[,]\d{0,}$"                                         'Pattern With Comma
    RegEx(3) = "^[-+]{0,1}(\d{1,3}[,]){0,1}(\d{3}[,]){0,}\d{3}([.]\d{0,}){0,1}$"    'Pattern With Commas
    RegEx(4) = "^[-+]{0,1}\d{1,}$"                                                  'Pattern simple
    On Error Resume Next
    'Abfrage ob die Eingabe 0 ist
    If IsNull(Var) Then
        getVarType = 9
    Else
        'Abfrage ob die Eingabe ein Array ist
        If IsArray(Var) Then
            getVarType = 6
        Else
            'Abfrage ob das Feld leer ist
            If Var = "" Then
                getVarType = 9
            Else
                'Abfrage ob die Eingabe unter die vorangegangenen Definitionen einer Zahl fällt
                Set oRegex = New RegExp
                counter = False
                For Each Pattern In RegEx
                    With oRegex
                        .Pattern = Pattern
                        .Global = True
                        For Each m In .Execute(Trim(Var))
                            counter = True
                        Next m
                    End With
                Next Pattern
                If counter Then
                    'Handelt es sich bei der Zahl um eine Ganzzahl?
                    If CLng(Val(Var)) = Var Then
                        'Ist die Zahl positiv?
                        If Var > 0 Then
                            getVarType = 1
                        Else
                            'Ist die Zahl negativ?
                            If Var < 0 Then
                                getVarType = 7
                            Else
                                getVarType = 98
                            End If
                        End If
                    'Ansonsten ist die Zahl eine Kommazahl
                    Else
                        'Ist die Zahl positiv?
                        If Var >= 0 Then
                            getVarType = 2
                        Else
                            'Ist die Zahl negativ?
                            If Var < 0 Then
                                getVarType = 8
                            End If
                        End If
                    End If
                Else
                    'Datumsformate
                    If IsDate(Var) Then
                        If CDbl(CDate(Var)) < 1 And CDbl(CDate(Var)) > 0 Then
                            getVarType = 5
                        Else
                            getVarType = 4
                        End If
                    Else
                    'boolische werte
                        If Var = True Or Var = False Then
                            getVarType = 3
                        Else
                            getVarType = 99
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

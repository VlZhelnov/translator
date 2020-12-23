Attribute VB_Name = "re_to_postfix"

' ѕреобразование регул€рного выражени€ в поствиксный вид
' до создани€ аналогичной грамматики

Function re2post(ByVal re As String) As String
    Dim nalt As Integer, natom As Integer
    Dim paren(0 To 99, 0 To 1) As Integer, p As Integer
    For i = 1 To Len(re)
        Select Case Mid(re, i, 1)
            Case "("
                If natom > 1 Then
                    natom = natom - 1
                    re2post = re2post & "."
                End If
                If p > 99 Then
                    re2post = Empty
                    Exit Function
                End If
                paren(p, 0) = nalt
                paren(p, 1) = natom
                p = p + 1
                nalt = 0: natom = 0
            Case "|"
                If natom = 0 Then
                    re2post = Empty
                    Exit Function
                End If
                natom = natom - 1
                While natom > 0
                    natom = natom - 1
                    re2post = re2post & "."
                Wend
                nalt = nalt + 1
            Case ")"
                If p = 0 Or natom = 0 Then
                    re2post = Empty
                    Exit Function
                End If
                natom = natom - 1
                While natom > 0
                    natom = natom - 1
                    re2post = re2post & "."
                Wend
                While nalt > 0
                    nalt = nalt - 1
                    re2post = re2post & "|"
                Wend
                p = p - 1
                nalt = paren(p, 0)
                natom = paren(p, 1)
                natom = natom + 1
            Case "*", "+", "?"
                If natom = 0 Then
                    re2post = Empty
                    Exit Function
                End If
                re2post = re2post & Mid(re, i, 1)
            Case "\"
                If natom > 1 Then
                    natom = natom - 1
                    re2post = re2post & "."
                End If
                re2post = re2post & Mid(re, i, 2)
                i = i + 1
                natom = natom + 1
            Case Else
                If natom > 1 Then
                    natom = natom - 1
                    re2post = re2post & "."
                End If
                re2post = re2post & Mid(re, i, 1)
                natom = natom + 1
        End Select
    Next
    If p <> 0 Then
        re2post = Empty
        Exit Function
    End If
    natom = natom - 1
    While natom > 0
        natom = natom - 1
        re2post = re2post & "."
    Wend
    While nalt > 0
        nalt = nalt - 1
        re2post = re2post & "|"
    Wend
End Function

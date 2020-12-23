Attribute VB_Name = "NFA"
Const match_ As Integer = 256
Const split_ As Integer = 257
Const digits_ As Integer = 258
Const chars_ As Integer = 259
Const spaces_ As Integer = 260


'Функции генерации НКА

Function state(ByVal char As String, Optional ByRef out As Object = Nothing, Optional ByRef out1 As Object = Nothing) As Object
    Set state = dict()
    state.add "c", char
    If ObjPtr(out) <> 0 Then state.add "out", out
    If ObjPtr(out1) <> 0 Then state.add "out1", out1
End Function
Function Frag(ByRef start As Object, ByRef out As Object) As Object
    Set Frag = dict()
    Frag.add "start", start
    Frag.add "out", out
End Function
Function list1(ByRef s As Object, out As String) As Object
    Set list1 = dict
    list1.add "s", s
    list1.add "out", out
    list1.add "next", Nothing
End Function
Sub patch(ByRef l As Object, ByRef s As Object, ByRef stateid As Long)
    While ObjPtr(l) <> 0
        If s.item("c") = match_ Then
            Set s = state(match_)
            s.item("stateid") = stateid
            stateid = stateid + 1
        End If
        Set l.item("s").item(l.item("out")) = s
        Set l = l.item("next")
    Wend
End Sub
Function append(ByRef L1 As Object, Optional ByRef L2 As Object) As Object
    Set append = L1
    While ObjPtr(L1.item("next")) > 0
        Set L1 = L1.item("next")
    Wend
    Set L1.item("next") = L2
End Function
Function post2nfa(postfix As String) As Object
    Dim stateid As Long: stateid = 0
    Dim stk As Object: Set stk = stack()
    Dim curchr As String, s As Object, e2 As Object, e1 As Object, e As Object
    For i = 1 To Len(postfix)
        curchr = Mid(postfix, i, 1)
        Select Case curchr
            Case "."
                Set e2 = Stack_pop(stk)
                Set e1 = Stack_pop(stk)
                patch e1.item("out"), e2.item("start"), stateid
                Stack_push stk, Frag(e1.item("start"), e2.item("out"))
            Case "|"
                Set e2 = Stack_pop(stk)
                Set e1 = Stack_pop(stk)
                Set s = state(split_, e1.item("start"), e2.item("start"))
                Stack_push stk, Frag(s, append(e1.item("out"), e2.item("out")))
            Case "?"
                Set e = Stack_pop(stk)
                Set s = state(split_, e.item("start"))
                Stack_push stk, Frag(s, append(e.item("out"), list1(s, "out1")))
            Case "*"
                Set e = Stack_pop(stk)
                Set s = state(split_, e.item("start"))
                patch e.item("out"), s, stateid
                Stack_push stk, Frag(s, list1(s, "out1"))
            Case "+"
                Set e = Stack_pop(stk)
                Set s = state(split_, e.item("start"))
                patch e.item("out"), s, stateid
                Stack_push stk, Frag(e.item("start"), list1(s, "out1"))
            Case "\"
                Select Case Mid(postfix, i, 2)
                    Case "\d": Set s = state(digits_)
                    Case "\w": Set s = state(chars_)
                    Case "\s": Set s = state(spaces_)
                    Case Else: Set s = state(Asc(Mid(postfix, i + 1, 1)))
                End Select
                Stack_push stk, Frag(s, list1(s, "out"))
                i = i + 1
            Case Else
                Set s = state(Asc(curchr))
                Stack_push stk, Frag(s, list1(s, "out"))
        End Select
    Next
    Set e = Stack_pop(stk)
    patch e.item("out"), state(match_), stateid
    Set post2nfa = e.item("start")
    End Function


'функции поиска соответствий

Function list() As Object
    Set list = dict()
    list.add "type", "List"
    list.add "s", dict()
    list.add "n", 0
End Function
Function startlist(ByRef start As Object, ByRef l As Object, _
                   ByRef listid As Long) As Object
    l.item("n") = 0
    listid = listid + 1
    addstate l, start, listid
    Set startlist = l
End Function
Function isnum(val As Integer) As Boolean
    Select Case val
        Case 48 To 57: isnum = True
        Case Else: isnum = False
    End Select
End Function
Function ischar(val As Integer) As Boolean
    Select Case val
        Case 65 To 90, 95, 97 To 122: ischar = True
        Case Else: ischar = False
    End Select
End Function
Function isspace(val As Integer) As Boolean
    Select Case val
        Case 9, 13, 32: isspace = True
        Case Else: isspace = False
    End Select
End Function
Sub addstate(ByRef l As Object, ByRef s As Object, ByRef listid As Long)
    If ObjPtr(s) = 0 Then Exit Sub
    s.item("lastlist") = listid
    If (s.item("c") = split_) Then
        addstate l, s.item("out"), listid
        addstate l, s.item("out1"), listid
        Exit Sub
    End If
    Set l.item("s").item(l.item("n")) = s
    l.item("n") = l.item("n") + 1
End Sub
Function step(ByRef clist As Object, ByVal C As Integer, _
              ByRef nlist As Object, ByRef listid As Long) As Object
              
    Dim i As Integer, s As Object, add As Boolean
    listid = listid + 1: nlist.item("n") = 0
    For i = 0 To clist.item("n") - 1
        Set s = clist.item("s").item(i)
        If s.item("c") = match_ Then
            Set step = dict()
            step.item("l") = s.item("lastlist")
            step.item("t") = s.item("stateid")
        Else
            add = False
            Select Case s.item("c")
                Case C: add = True
                Case spaces_: If isspace(C) Then add = True
                Case chars_: If ischar(C) Then add = True
                Case digits_: If isnum(C) Then add = True
            End Select
            If add Then addstate nlist, s.item("out"), listid
        End If
    Next
    
End Function
Function match(ByRef start As Object, ByRef lang As Object, ByVal s As String) As Object
    
    Dim listid As Long: listid = 0
    Dim i As Long, curchr As String
    Dim clist As Object, nlist As Object, t As Object
    
    Set clist = startlist(start, list, listid)
    Set nlist = list
    Set match = dict()
    
    
    s = s + "$"
    For i = 1 To Len(s)
        curchr = Mid(s, i, 1)
        Set t = step(clist, Asc(curchr), nlist, listid)
        If ObjPtr(t) Then Set match.item(t.item("t")) = t
        Set t = clist: Set clist = nlist: Set nlist = t
    Next
    
	    
    'вычисление приоритета шаблонов

    If match.Count = 0 Then
        Set t = Nothing
    ElseIf match.Count = 1 Then
        Set t = match.items()(0)
    Else
        Set t = match.items()(match.Count - 1)
        For i = match.Count - 2 To 0
            If t.item("l") < match.items()(i).item("l") Then
                Set t = match.items()(i)
            End If
        Next
    End If
    Set match = t
    If ObjPtr(match) Then
        match.item("t") = lang.keys()(match.item("t"))
        match.item("l") = Trim(Mid(s, 1, match.item("l") - 1))
    End If
    
End Function





Attribute VB_Name = "LR1"

Sub create_stack(ByRef pcode As Object)
    dict_append pcode, "Dim stack as Object, top as Integer"
    dict_append pcode, "Sub stack_init(): Set stack = CreateObject(" & Chr(34) & "Scripting.Dictionary" & Chr(34) & "): top = 0: End Sub"
    dict_append pcode, "Sub push(byval t): top = top + 1: stack.item(top) = t: End Sub"
    dict_append pcode, "Sub pop(byref t): t = stack.item(top): top = top - 1: End Sub"
    dict_append pcode, "Sub run(): stack_init: main: End Sub"
End Sub


Sub analize(ByRef src_str As Object, ByRef parseTable As Object, _
            ByRef rules As Object, ByVal outfile As String)
    
    Dim index As Integer, action As String, rule As Variant
    Dim symbols As Object: Set symbols = stack()
    Dim states As Object: Set states = stack(): Stack_push states, 0
    Dim num_product As Integer
    
    Dim vars As Object: Set vars = dict()
    Dim pcode As Object: Set pcode = dict()
    
    create_stack pcode    'for translator
    
    While True
        action = parseTable(stack_top(states))(src_str.item("t"))
        Select Case Mid(action, 1, 1)
            Case "r"
                rule_num = Int(Mid(action, 2))
                If rule_num = 0 Then ' accepts
                    'ReToPostfix 0, symbols, pcode, vars
                    Compile rule_num, symbols, pcode, vars                  'for translator
                    save_to_file outfile, Join(pcode.items(), vbNewLine)    'for translator
                    Debug.Print "accept"
                    Exit Sub
                End If
                    rule = rules.items()(rule_num)
                Dim ret As Object: Set ret = Compile(rule_num, symbols, pcode, vars)
              ' Dim ret As String: ret = ReToPostfix(rule_num, symbols, pcode, vars)
                If UBound(rule) > 0 Then
                    For index = 1 To UBound(rule)
                        Stack_pop states: Stack_pop symbols
                    Next
                End If
                
                Stack_push states, Int(Mid(parseTable(stack_top(states))(rule(0)), 2))  ' GOTO ACTION
                Stack_push symbols, ret
            Case "s"
                Stack_push states, Int(Mid(action, 2))
                Stack_push symbols, src_str.item("l")
                Set src_str = src_str.item("next")
            Case Else
                Debug.Print "sintax error"
                Debug.Print "no action for state " & stack_top(states) & " and char " & src_str.item("t")
                Exit Sub
        End Select
    Wend
   
End Sub

Function CreateTable(ByRef grammar As Object, Optional create_info_file = False) As Object

    Dim rules As Variant: rules = grammar.item("rules").items()
    Dim nterms As Object: Set nterms = grammar.item("nterm")
    Dim terms As Object: Set terms = grammar.item("term")

    Dim sym As Variant
    
    For Each sym In nterms.keys()
        nullable sym, grammar
    Next
    For Each sym In nterms.keys()
        nterms.item(sym).item("first") = nfirst(sym, grammar).keys()
    Next
    For Each sym In nterms.keys()
        nterms.item(sym).item("follow") = nfollow(sym, grammar).keys()
    Next

    Dim core As Object: Set core = dict()
    core.item(Int(Join(Array(0, 1, 0), " "))) = Array(0, 1, 0)
    
    Dim set_items As Object: Set set_items = dict()
    dict_append set_items, closure(core, grammar)
    
    Dim unic_key As String, unic As Object: Set unic = dict()
    unic_key = Join(sort(set_items.item(0).keys()), "")
    unic.item(unic_key) = 0
    
    
    Dim symbols As Variant, symbol As Variant, N As Integer
    Dim states As Object: Set states = dict()
    Dim goto_set As Object, action_type As String
    
    Do
        For Each symbols In Array(nterms.keys(), terms.keys())
            For Each symbol In symbols
                Set goto_set = ngoto(set_items.item(N), symbol, grammar)
                If goto_set.Count() > 0 Then
                    action_type = IIf(nterms.exists(symbol), "g", "s")
                    unic_key = Join(sort(goto_set.keys()), " ")
                    If Not unic.exists(unic_key) Then
                        unic.item(unic_key) = set_items.Count()
                        dict_append states, Array(action_type, N, set_items.Count(), symbol)
                        dict_append set_items, goto_set
                        For Each item In goto_set.items()
                            If item(1) = UBound(rules(item(0))) + 1 Then
                                dict_append states, _
                                    Array("r", set_items.Count() - 1, item(0), terms.keys()(item(2)))
                            End If
                        Next
                    Else: dict_append states, Array(action_type, N, unic.item(unic_key), symbol)
                    End If
                End If
            Next
        Next
        N = N + 1
    Loop Until N >= set_items.Count()
      
    Set CreateTable = dict()
    For Each s In states.items()
        If Not CreateTable.exists(s(1)) Then Set CreateTable.item(s(1)) = dict()
        CreateTable.item(s(1))(s(3)) = s(0) & s(2)
    Next
    
    If create_info_file Then
        Dim info As Object: Set info = dict()
        Set info.item("grammar") = grammar
        Set info.item("set_items") = set_items
        Set info.item("states") = states
        Set info.item("table") = CreateTable
        dict_to_file info, "info.txt"
    End If
    
    
End Function


Function sort(arr As Variant) As Variant
    Dim tmp As Long, i As Long, j As Long
    i = 1
    While (i <= UBound(arr))
        j = i
        While (j > 0)
            If (arr(j) < arr(j - 1)) Then
                tmp = arr(j)
                arr(j) = arr(j - 1)
                arr(j - 1) = tmp
            End If
            j = j - 1
        Wend
        i = i + 1
    Wend
    sort = arr
End Function

Function closure(ByVal set_points As Object, ByRef grammar As Object) As Object

    Dim n_point As Long: Set closure = set_points
    
    Dim rules As Variant: rules = grammar.item("rules").items()
    Dim nterms As Object: Set nterms = grammar.item("nterm")
    Dim terms As Object: Set terms = grammar.item("term")
    
    Dim unic As Variant, cur_point As Variant
    Dim cur_symbol As Variant, next_symbol As Variant
    
    While True
        cur_point = set_points.items()(n_point)
        If Not cur_point(1) = UBound(rules(cur_point(0))) + 1 Then
            cur_symbol = rules(cur_point(0))(cur_point(1))
            If nterms.exists(cur_symbol) Then
                For n_rule = 0 To UBound(rules)
                    If rules(n_rule)(0) = cur_symbol Then
                        If cur_point(1) = UBound(rules(cur_point(0))) Then
                            unic = Array(n_rule, 1, cur_point(2))
                            closure.item(Int(Join(unic, ""))) = unic
                        Else
                            next_symbol = rules(cur_point(0))(cur_point(1) + 1)
                            If nterms.exists(next_symbol) Then
                                For Each e In nterms.item(next_symbol).item("first")
                                    unic = Array(n_rule, 1, terms.item(e))
                                    closure.item(Int(Join(unic, ""))) = unic
                                Next
                                If nterms.item(cur_symbol).item("nullable") Then
                                    For Each e In nterms.item(next_symbol).item("follow")
                                        unic = Array(n_rule, 1, terms.item(e))
                                        closure.item(Int(Join(unic, ""))) = unic
                                    Next
                                End If
                            Else
                                unic = Array(n_rule, 1, terms.item(next_symbol))
                                closure.item(Int(Join(unic, ""))) = unic
                            End If
                        End If
                    End If
                Next
            End If
        End If
        If n_point = closure.Count() - 1 Then Exit Function
        n_point = n_point + 1
    Wend
    
End Function
    
Function ngoto(ByVal set_points As Object, ByVal sym As String, ByVal grammar As Object) As Object
    Dim point As Variant: Set ngoto = dict()
    Dim rules As Variant: rules = grammar.item("rules").items()
    
    For Each point In set_points.items()
        If point(1) <= UBound(rules(point(0))) Then
            If rules(point(0))(point(1)) = sym Then
                ngoto.item(Int(Join(Array(point(0), point(1) + 1, point(2)), ""))) _
                    = Array(point(0), point(1) + 1, point(2))
            End If
        End If
    Next
    If ngoto.Count() > 0 Then
    Set ngoto = closure(ngoto, grammar)
    End If
End Function

Sub nullable(ByRef sym As Variant, ByRef grammar As Object)

    Dim rules As Variant: rules = grammar("rules").items()
    Dim is_nullable As Boolean: is_nullable = True
    For Each rule In rules
        If rule(0) = sym And UBound(rule) = 0 Then
            With grammar.item("nterm").item(rule(0))
                .item("nullable") = True
            End With
            Exit Sub
        End If
    Next
    For Each rule In rules
        If rule(0) = sym Then
            For n_sym = UBound(rule) To 1 Step -1
                With grammar.item("nterm")
                    If .exists(rule(n_sym)) Then
                        With .item(rule(n_sym))
                            If Not .exists("nullable") Then
                                .item("nullable") = "None"
                                nullable rule(n_sym), grammar
                            End If
                            If .item("nullable") <> "None" Then
                                is_nullable = is_nullable And .item("nullable")
                            End If
                        End With
                    Else
                        is_nullable = is_nullable And False
                    End If
                End With
            Next
        End If
    Next
    With grammar.item("nterm").item(sym)
        .item("nullable") = is_nullable
    End With
End Sub

Function nfirst(ByVal sym As String, ByRef grammar As Object) As Object

    Dim terms As Object: Set terms = grammar.item("term")
    Dim rules As Variant: rules = grammar.item("rules").items()
    Dim tmp As Object, tmp_arr As Variant: Set nfirst = dict()
    
    If grammar.item("term").exists(sym) Then
        nfirst.item(sym) = Empty: Exit Function
    End If
    
    For Each rule In rules
        If rule(0) = sym Then
            For n_sym = 1 To UBound(rule)
                If grammar.item("term").exists(rule(n_sym)) Then
                    nfirst.item(rule(n_sym)) = Empty: Exit For
                End If
                If rule(n_sym) = sym Then Exit For
                Set tmp = nfirst(CStr(rule(n_sym)), grammar)
                For Each e In tmp
                    nfirst.item(e) = Empty
                Next
                If Not grammar.item("nterm").item(rule(n_sym)).item("nullable") Then Exit For
            Next
        End If
    Next
End Function

Function nfollow(ByVal sym As String, ByRef grammar As Object)
    Dim rules As Variant: rules = grammar.item("rules").items()
    Dim tmp_arr As Variant, n_term As Integer: Set nfollow = dict()
    
    If grammar.item("nterm").exists(sym) Then
        For Each rule In rules
            For n_sym = 1 To UBound(rule)
                If rule(n_sym) = sym Then
                    For n_sym_find = n_sym + 1 To UBound(rule)
                         If grammar.item("nterm").exists(rule(n_sym_find)) Then
                            For Each e In grammar.item("nterm").item(rule(n_sym_find)).item("first")
                                nfollow.item(e) = Empty
                            Next
                            If Not grammar.item("nterm").item(rule(n_sym_find)).item("nullable") Then Exit For
                        Else
                            nfollow.item(rule(n_sym_find)) = Empty
                            Exit For
                        End If
                    Next
                    If n_sym_find = UBound(rule) + 1 Then
                        If Not grammar.item("nterm").item(sym).exists("follow") Then
                            grammar.item("nterm").item(sym).item("follow") = "None"
                            For Each e In nfollow(sym, grammar)
                                nfollow.item(e) = Empty
                            Next
                            If grammar.item("nterm").item(rule(0)).exists("follow") Then
                                For Each e In grammar.item("nterm").item(rule(0)).item("follow")
                                    nfollow.item(e) = Empty
                                Next
                            End If
                        End If
                    End If
                End If
            Next
        Next
    End If
    
End Function

Function read_grammar(ByRef name_grammar_file As String) As Object

    Dim grammar_text As String: grammar_text = get_file_text(name_grammar_file)
    grammar_text = Replace(grammar_text, vbNewLine, " ")
    grammar_text = Replace(grammar_text, vbTab, " ")
    
    Dim grammar As Object: Set grammar = dict()
    
    Set grammar.item("nterm") = dict()
    Set grammar.item("term") = dict(): grammar.item("term").item("$") = 0
    Set grammar.item("rules") = dict()
    
    Dim tmp_dict As Object, cur_rule As Integer, cur_head As String
    
    Dim state As Integer
    For Each word In Split(grammar_text, " ")
        Select Case state
            Case 0
                Select Case word
                    Case "": state = 0
                    Case ":", "|":
                        Debug.Print "FormatError: No head nterm for production"
                        Exit Function
                    Case ";"
                        Debug.Print "FormatError: No enter production"
                        Exit Function
                    Case Else
                        If Mid(word, 1, 1) = Chr(34) And _
                           Mid(word, Len(word), 1) = Chr(34) Then
                            Debug.Print "head symbol must be nterm"
                            Exit Function
                        Else
                            If Not grammar.item("nterm").exists(word) Then
                                Set grammar.item("nterm").item(word) = dict()
                            End If
                            cur_rule = 0: cur_head = word
                            Set tmp_dict = dict(): Set tmp_dict.item(cur_rule) = dict()
                            dict_append tmp_dict.item(cur_rule), word
                            state = 1
                        End If
                End Select
            Case 1
                Select Case word
                    Case "": state = 1
                    Case ":": state = 2
                    Case Else
                        Debug.Print "FormatError: No sybmol ':' after head nterm"
                        Exit Function
                End Select
            Case 2
                Select Case word
                    Case "": state = 2
                    Case ":":
                        Debug.Print "FormatError: Use ':' in body production"
                        Exit Function
                    Case "|"
                        cur_rule = cur_rule + 1
                        Set tmp_dict.item(cur_rule) = dict()
                        dict_append tmp_dict.item(cur_rule), cur_head
                    Case ";"
                        For Each rule In tmp_dict.items()
                            grammar.item("rules").item(grammar.item("rules").Count()) = rule.items()
                        Next
                        state = 0
                    Case Else
                        If Mid(word, 1, 1) = Chr(34) And _
                           Mid(word, Len(word), 1) = Chr(34) Then
                            word = Mid(word, 2, Len(word) - 2)
                            If Not grammar.item("term").exists(word) Then
                                grammar.item("term").item(word) = grammar.item("term").Count()
                            End If
                            dict_append tmp_dict.item(cur_rule), word
                            state = 2
                        Else
                            If Not grammar.item("nterm").exists(word) Then
                                Set grammar.item("nterm").item(word) = dict()
                            End If
                           dict_append tmp_dict.item(cur_rule), word
                            state = 2
                        End If
                End Select
        
        End Select
    Next
    Set read_grammar = grammar
End Function

Sub save_table(ByRef table As Object, ByVal name_file As String)
    Dim fso As Object, text_stream As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim tmp_str As String
    Dim d As Object: Set d = dict()
    
    For cnt = 0 To table.Count() - 1
        tmp_str = ""
        For Each k In table.item(cnt)
            tmp_str = tmp_str & " | " & k & " " & table.item(cnt).item(k)
        Next
        dict_append d, Mid(tmp_str, 4)
    Next
    
    Set text_stream = fso.OpenTextFile(name_file, 2, True)
    text_stream.Write Join(d.items(), vbNewLine) & vbNewLine
    text_stream.Close
End Sub
Function read_table(ByVal name_file As String) As Object
    
    Dim text As String: text = get_file_text(name_file)
    text = Replace(text, vbNewLine, " EndRow ")
    
    Set read_table = dict()
    Dim state As Integer, symbol As String, row As Integer
    
    For Each word In Split(text, " ")
        Select Case state
            Case 0
                 Select Case word
                    Case "EndRow": state = 0
                    Case Else
                        row = read_table.Count()
                        Set read_table.item(row) = dict()
                        sympol = word
                        state = 1
                 End Select
            Case 1
                 Select Case word
                    Case "EndRow": Debug.Print "FormatError": Exit Function
                    Case Else
                        read_table.item(row).item(sympol) = word
                        state = 2
                 End Select
            Case 2
                 Select Case word
                    Case "EndRow": state = 0
                    Case "|": state = 3
                    Case Else
                        read_table.item(row).item(symbol) = word
                        state = 2
                 End Select
            Case 3
                 Select Case word
                    Case "EndRow": state = 0
                    Case Else
                        sympol = word
                        state = 1
                 End Select
        End Select
    Next
    
End Function
Sub table_to_range(ByRef table As Object, ByRef grammar As Object)
    Dim symbols As Variant, symbol As Integer, state As Integer
    
    Dim nterms As Object: Set nterms = grammar.item("nterm")
    Dim terms As Object: Set terms = grammar.item("term")

    
    For Each symbols In Array(nterms.keys(), terms.keys())
        For symbol = 0 To UBound(symbols)
            Range("A1").Offset(, symbol + 1) = symbols(symbol)
            For state = 0 To table.Count - 1
                Range("A2").Offset(state, symbol + 1) = _
                    table.item(state).item(symbols(symbol))
                Range("A1").Offset(state + 1) = state
            Next
        Next
    Next
End Sub

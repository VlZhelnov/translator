Attribute VB_Name = "LR1"
Const left As Integer = 0
Const right As Integer = 1
Const lookahead As Integer = 2

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
                    Exit Sub
                End If
                rule = Split(rules.keys()(rule_num), "->", 2)
                Dim ret As Object: Set ret = Compile(rule_num, symbols, pcode, vars)
              ' Dim ret As String: ret = ReToPostfix(rule_num, symbols, pcode, vars)
                If rule(1) <> " ." Then
                    For index = 1 To UBound(Split(rule(1), " "))
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

'Функции для генерации таблицы разбора

Function CreateTable(ByRef grammar As Object, Optional create_info_file = False) As Object

    Dim products As Object: Set products = grammar("products")
    Dim terms As Variant: terms = grammar("terms")
    Dim rules As Object: Set rules = grammar("rules")
    Dim symbols As Variant, N As Integer
    
    Set CreateTable = dict()
    
    Dim first_set As Object: Set first_set = dict()
    Dim follow_set As Object: Set follow_set = dict()
    Dim nmemory As Object: Set nmemory = dict()
    
    For Each nterm In products.keys()
        nullable nterm, grammar, nmemory
        Set first_set.item(nterm) = first(nterm, grammar, nmemory)
        Set follow_set.item(nterm) = follow(nterm, grammar, first_set, nmemory)
    Next
    
    symbols = Split(Join(products.keys(), "|") & "|" & Join(terms, "|"), "|")
    start_rule = Array(products.keys()(0), "." & products.items()(0)(0), "$")
    
    Dim tmp_set As Object: Set tmp_set = dict(): dict_append tmp_set, start_rule
    Dim C As Object: Set C = dict(): dict_append C, closure(tmp_set, grammar, first_set, follow_set)
    Dim goto_set As Object, states As Object: Set states = dict()
    
    Do
        For Each symbol In symbols
            Set goto_set = fgoto(C(N), symbol, grammar, first_set, follow_set)
            If goto_set.Count() > 0 Then
                unic = True: l = 0
                For Each p In C.items()
                    If to_str(goto_set) = to_str(p) Then
                        unic = False: Exit For
                    Else: l = l + 1
                    End If
                Next
                action_type = IIf(products.exists(symbol), "g", "s")
                If unic Then
                    dict_append states, Array(action_type, N, C.Count(), symbol)
                    dict_append C, goto_set
                    
                    For Each item In goto_set.items()
                        If Mid(item(right), Len(item(right)), 1) = "." Or Mid(item(right), Len(item(right)), 1) = "$" Then
                            rule = item(left) & "->" & IIf(item(right) = ".", " .", item(right))
                            If rules.exists(rule) Then
                                dict_append states, Array("r", C.Count() - 1, rules(rule), item(lookahead))
                            End If
                        End If
                    Next
                Else: dict_append states, Array(action_type, N, l, symbol)
                End If
            End If
        Next
        N = N + 1
    Loop Until N >= C.Count()


    Set CreateTable = dict()
    For Each s In states.items()
        If Not CreateTable.exists(s(1)) Then Set CreateTable.item(s(1)) = dict()
        CreateTable.item(s(1))(s(3)) = s(0) & s(2)
    Next
    
    If create_info_file Then
        Dim info As Object: Set info = dict()
        Set info.item("grammar") = grammar
        Set info.item("nullable") = nmemory
        Set info.item("first") = first_set
        Set info.item("follow") = follow_set
        Set info.item("items") = C
        Set info.item("states") = states
        Set info.item("table") = CreateTable
        dict_to_file info, "info.txt"
    End If
    
End Function
Function fgoto(ByRef i As Object, ByVal x As String, ByVal grammar As Object, ByRef first_set As Object, ByRef follow_set As Object) As Object
    Set fgoto = dict()
    'If x = "$" Then Exit Function
    Dim point_position As Integer
    For Each Point In i.items
        point_position = InStr(1, Point(right), "." & x)
        If point_position < Len(Point(right)) And point_position > 0 Then
            If point_position + Len(x) = Len(Point(right)) Or _
               Mid(Point(right), point_position + Len(x) + 1, 1) = " " Then
                    next_point = Trim(Replace(Point(right) & " ", "." & x & " ", x & " ."))
                    dict_append fgoto, Array(Point(left), next_point, Point(lookahead))
            End If
        End If
    Next
   If fgoto.Count() > 0 Then Set fgoto = closure(fgoto, grammar, first_set, follow_set)
End Function
Sub nullable(ByVal nterm As String, ByRef grammar As Object, ByRef memory As Object)
  
    Dim products As Object: Set products = grammar("products")
    
    If products.exists(nterm) Then
        For Each product In products.item(nterm)
            If product = "" Then
                 memory.item(nterm) = True: Exit Sub
            End If
        Next
        For Each product In products.item(nterm)
            Words = Split(product, " ")
            For n_word = UBound(Words) To 0 Step -1
                If products.exists(Words(n_word)) Then
                    If Not memory.exists(Words(n_word)) Then
                        memory.item(Words(n_word)) = False
                        nullable Words(n_word), grammar, memory
                        If Not memory.item(Words(n_word)) Then
                            memory.item(nterm) = False: Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                Else
                     memory.item(nterm) = False: Exit Sub
                End If
            Next
        Next
            memory.item(nterm) = True
    End If
    
End Sub

Function closure(ByVal set_items As Object, ByRef grammar As Object, ByRef first_set As Object, ByRef follow_set As Object) As Object
    Dim n_item As Long: Set closure = set_items
    Dim next_symbol, point_sympol As String, next_point As Variant
    Dim products As Object: Set products = grammar("products")
    Dim point_position As Integer, tmp_point As String
    Dim d As Object: Set d = dict()
    While True
        tmp_point = set_items(n_item)(right)
        point_position = InStr(1, tmp_point, ".")
        If point_position <> Len(tmp_point) Then
            point_symbol = Split(Split(tmp_point, ".")(1), " ")(0)
            If products.exists(point_symbol) Then
                If UBound(products.item(point_symbol)) = -1 Then
                    For Each term In follow_set.item(point_symbol)
                        next_point = Array(point_symbol, ".", term)
                        If check(set_items, next_point) Then dict_append set_items, next_point
                    Next
                Else
                    For Each product In products.item(point_symbol)
                        If point_position + Len(point_symbol) = Len(tmp_point) Then
                            next_point = Array(point_symbol, "." & product, set_items(n_item)(lookahead))
                            If check(set_items, next_point) Then dict_append set_items, next_point
                        Else
                            next_symbol = Split(Split(tmp_point, ".")(1), " ")(1)
                            Set d = dict()
                            For Each term In first_set.item(next_symbol)
                                d.item(term) = Null
                            Next
                            If follow_set.exists(next_symbol) Then
                                For Each term In follow_set.item(next_symbol)
                                    d.item(term) = Null
                                Next
                            End If
                            For Each term In d.keys
                                next_point = Array(point_symbol, "." & product, term)
                                If check(set_items, next_point) Then dict_append set_items, next_point
                            Next
                        End If
                    Next
                End If
            End If
        End If
        If n_item = set_items.Count() - 1 Then Exit Function
        n_item = n_item + 1
    Wend

End Function

Function check(ByRef set_items As Object, ByRef item As Variant) As Boolean
    Dim i As Variant: check = True
    For Each i In set_items.items()
        If to_str(item) = to_str(i) Then
            check = False: Exit Function
        End If
    Next
End Function

Function first(ByVal s As String, ByRef grammar As Object, ByRef nmemory As Object) As Object

    Dim terms As Variant: terms = grammar.item("terms")
    Dim products As Object: Set products = grammar.item("products")
    Dim tmp As Object, tmp_arr As Variant: Set first = dict()
    
    If InStr(1, to_str(terms), s) Then
        first.item(s) = Empty: Exit Function
    End If
    
    If products.exists(s) Then
        For Each pr In products.item(s)
            tmp_arr = Split(pr, " ")
            For cnt = 0 To UBound(tmp_arr)
            If InStr(1, to_str(terms), tmp_arr(cnt)) Then
                first.item(tmp_arr(cnt)) = Empty: Exit For
            End If
            If tmp_arr(cnt) <> s Then
                    Set tmp = first(CStr(tmp_arr(cnt)), grammar, nmemory)
                    For Each e In tmp
                        first.item(e) = Empty
                    Next
                    If Not nmemory.exists(tmp_arr(cnt)) Then
                        nullable tmp_arr(cnt), grammar, nmemory
                    End If
                    If Not nmemory.item(tmp_arr(cnt)) Then Exit For
                End If
            Next
        Next
    End If
End Function
Function follow(ByVal s As String, ByRef grammar As Object, ByRef first_memory As Object, ByRef nmemory As Object, Optional ByRef checked As Object = Nothing)
    Dim products As Object: Set products = grammar.item("products")
    Dim tmp_arr As Variant, n_term As Integer: Set follow = dict()
    
    If TypeName(checked) = "Nothing" Then Set checked = dict()
    
    If products.exists(s) Then
        For Each product In products.keys():
            For Each r In products.item(product)
                If InStr(1, r & " ", s & " ") > 0 Then
                    tmp_arr = Split(Trim(Split(r & " ", s & " ")(1)), " ")
                    For n_nterm = 0 To UBound(tmp_arr)
                        If Not first_memory.exists(tmp_arr(n_nterm)) Then
                            Set first_memory.item(tmp_arr(n_nterm)) = first(tmp_arr(n_nterm), grammar, nmemory)
                        End If
                        For Each e In first_memory.item(tmp_arr(n_nterm))
                            follow.item(e) = Empty
                        Next
                        If products.exists(tmp_arr(n_nterm)) Then
                            If Not nmemory.exists(tmp_arr(n_nterm)) Then
                                nullable tmp_arr(n_nterm), grammar, nmemory
                            End If
                            If Not nmemory.item(tmp_arr(n_nterm)) Then Exit For
                        Else
                            Exit For
                        End If
                     Next
                     If n_nterm = UBound(tmp_arr) + 1 Then
                        If Not checked.exists(product) Then
                            checked.item(product) = Null
                            For Each e In follow(product, grammar, first_memory, nmemory, checked)
                                follow.item(e) = Empty
                            Next
                        End If
                    End If
                End If
            Next
        Next
    End If
End Function

' функции сохранения и загрузки грамматики и таблицы в разных форматах

Function load_grammar(ByVal name_file As String) As Object

    Set load_grammar = dict()
    
    Dim products As Object: Set products = dict()
    Dim terms As Object: Set terms = dict()
    Dim rules As Object: Set rules = dict()
    Dim nullabls As Object: Set nullabls = dict()
    Dim form As Object, simbol As Variant
    Dim tmp_arr As Variant, line As Variant
    Dim product As Variant, char As Variant

    For Each line In Split(get_file_text(name_file), ";" & vbNewLine)
        If InStr(1, line, "->") Then
            If line <> Empty Then
                tmp_arr = Split(Replace(line, vbNewLine, Empty), "->")
                For Each product In Split(tmp_arr(1), "|")
                    For Each char In Split(product, " ")
                        If char Like Chr(34) & "*" & Chr(34) Then
                            If char <> Empty Then
                                terms(Trim(Replace(char, Chr(34), Empty))) = Null
                            End If
                        End If
                    Next
                Next
                Set form = dict()
                For Each simbol In Split(Replace(tmp_arr(1), Chr(34), Empty), "|")
                    dict_append form, Trim(simbol)
                Next
                products.item(Trim(tmp_arr(0))) = form.items()
            End If
        End If
    Next
    For Each prod In products.keys()
        If UBound(products.item(prod)) = -1 Then
            rules.item(prod & "-> .") = rules.Count()
        Else
            For Each elem In products(prod)
                rules.item(Join(Array(prod, Replace(elem, ".", Empty) & " ."), "->")) _
                        = rules.Count()
            Next
        End If
    Next
    
    Set load_grammar.item("products") = products
    Set load_grammar.item("rules") = rules
    load_grammar.item("terms") = terms.keys()
    
End Function




Sub save_table(ByRef table As Object, ByVal name_file As String)
    Dim fso As Object, text_stream As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim tmp_str As String
    Dim d As Object: Set d = dict()
    
    For cnt = 0 To table.Count() - 1
        tmp_str = ""
        For Each k In table.item(cnt)
            tmp_str = tmp_str & "|" & k & " " & table.item(cnt).item(k)
        Next
        dict_append d, Mid(tmp_str, 2)
    Next
    
    Set text_stream = fso.OpenTextFile(name_file, 2, True)
    text_stream.Write Join(d.items(), vbNewLine) & vbNewLine
    text_stream.Close
End Sub
Function read_table(ByVal name_file As String) As Object
    Set read_table = dict()
    Dim tmp_dict As Object
    Dim text As String: text = get_file_text(name_file)
    For Each line In Split(text, vbNewLine)
        If line <> Empty Then
            Set tmp_dict = dict()
            For Each pair In Split(line, "|")
                pair = Split(pair, " ")
                tmp_dict.item(pair(0)) = pair(1)
            Next
            dict_append read_table, tmp_dict
        End If
    Next
End Function
Sub table_to_range(ByRef table As Object, ByRef grammar As Object)
    Dim symbols As Variant
    symbols = Split(Join(grammar.item("terms"), "|") & "|$|" & _
                    Join(grammar.item("products").keys(), "|"), "|")
            
    For t = 0 To UBound(symbols)
        Range("A1").Offset(, t + 1) = symbols(t)
    Next
    For cnt = 0 To table.Count - 1
        For t = 0 To UBound(symbols)
                Range("A2").Offset(cnt, t + 1) = _
                    table.item(cnt).item(symbols(t))
        Next
        Range("A1").Offset(cnt + 1) = cnt
    Next
End Sub




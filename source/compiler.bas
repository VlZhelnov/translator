Attribute VB_Name = "compiler"

Sub main() 'Точка входа анализа

    Dim current_path As String: current_path = ThisWorkbook.Path
    
    Dim grammar As Object: Set grammar = load_grammar(current_path & "\grammar.txt")
    Dim parse_table As Object: Set parse_table = read_table(current_path & "\parse_table.txt")
    Dim lex_rule As Object: Set lex_rule = load_lex_re(current_path & "\lex_rule.txt")
    Dim source_text As String: source_text = get_file_text(current_path & "\source_code.txt")
    Dim lexer As Object: Set lexer = lex(source_text & "$", lex_rule)

    analize lexer, parse_table, grammar.item("rules"), current_path & "\out.txt"
    
End Sub


Sub gen() 'Точка входа анализа
    Dim current_path As String: current_path = ThisWorkbook.Path
    Dim grammar As Object: Set grammar = load_grammar(current_path & "\grammar.txt")
    Dim parse_table As Object: Set parse_table = CreateTable(grammar, True)
    save_table parse_table, current_path & "\parse_table.txt"
    table_to_range parse_table, grammar
End Sub


' функции вызываемые во время анализа при свертках для разных грамматик

Function ReToPostfix(ByVal rule_num As Integer, ByRef args As Object, _
                     ByRef pcode As Object, ByRef vars As Object) As String
    Dim t As Integer: t = args.item("top")
    Select Case rule_num
        Case 0: Debug.Print args.item(t)
        Case 1: ReToPostfix = Join(Array(args(t - 2), args(t), args(t - 1)))
        Case 3: ReToPostfix = Join(Array(args(t - 1), args(t), "&"))
        Case 5, 6: ReToPostfix = Join(Array(args(t - 1), args(t)))
        Case 8: ReToPostfix = args(t - 1)
        Case 2, 4, 7, 9: ReToPostfix = args.item(t)
    End Select
End Function

Function Compile(ByVal rule_num As Integer, ByRef args As Object, _
                 ByRef pcode As Object, ByRef vars As Object) As Object
                 
    Set Compile = dict(): Dim t As Integer: t = args.item("top")
    Dim expr As String, tmp_var As Variant
    Select Case rule_num
        Case 0 'program->fundefs .
            If Not vars.exists("main") Then
                pcode.item(5) = "'Error: Function main not defained" _
                                & vbNewLine & pcode.item(5)
            End If
            Debug.Print dict_to_string(vars)
            
        Case 4 ' fundef->basic id F ( iargs ) block .
            dict_append pcode, "End Sub"
            
        Case 8 ' iarg->type id . 8
            dict_append pcode, "pop " & args.item(t)
            
        Case 12 'decl->type id ;
            If Not vars(vars.item("cur_env")).exists(args.item(t - 1)) Then
                If args.item(t - 2).item("dim") <> Empty Then
                    tmp_var = "Dim " & args.item(t - 1) _
                    & "(" & Mid(args.item(t - 2).item("dim"), 3) & ")"
                    Select Case args.item(t - 2).item("type")
                        Case "int": tmp_var = tmp_var & " as Integer"
                        Case "float": tmp_var = tmp_var & " as Single"
                        Case "long": tmp_var = tmp_var & " as Long"
                        Case Else
                            tmp_var = tmp_var & " 'Warning: undefined size array " _
                                              & args.item(t - 2).item("type")
                    End Select
                 Else
                    tmp_var = "Dim " & args.item(t - 1)
                    Select Case args.item(t - 2).item("type")
                        Case "int": tmp_var = tmp_var & " as Integer"
                        Case "float": tmp_var = tmp_var & " as Single"
                        Case "long": tmp_var = tmp_var & " as Long"
                        Case "void", "void[]"
                            tmp_var = tmp_var & " 'Variable Type Error: " _
                                              & args.item(t - 2).item("type")
                        Case Else: tmp_var = tmp_var & " as Variant"
                    End Select
                End If
                vars.item(vars.item("cur_env")).item(args.item(t - 1)) = tmp_var
                dict_append pcode, tmp_var
            Else
                dict_append pcode, vars.item(vars.item("cur_env")).item(args.item(t - 1)) & _
                                   " 'Warning: id " & args.item(t - 1) & _
                                   " has be declared in this function " & tmp_var
            End If
             
        Case 13 ' type->type [ num ]
            Set Compile = args.item(t - 3)
            Compile.item("dim") = Join(Array(Compile.item("dim"), _
                                       "0 To " & val(args.item(t - 1)) - 1), ", ")
        Case 14 'type->basic
            Compile.item("type") = args.item(t)
                
        Case 17 'stmt->loc = bool ;
            dict_append pcode, args.item(t - 3).item("id") _
                               & " = " & args.item(t - 1).item("id")
                               
        Case 18, 19: 'if
            dict_append pcode, args(t - 1).item("label") & ":"
        
        Case 20, 21 'while
            tmp_var = IIf(rule_num = 20, 5, 7)
            dict_append pcode, "GoTo " & args(t - tmp_var).item("label")
            dict_append pcode, args(t - 1).item("label") & ":"
        
        Case 22 'stmt->return arg ; .
            dict_append pcode, "push " & args.item(t - 1).item("id")
            dict_append pcode, "Exit Sub"
        Case 24 'args->args , arg .
            Compile.item("id") = args(t).item("id") _
                                 & "," & args(t - 2).item("id")
            
        Case 27 ' M
            With vars.item(vars.item("cur_env"))
                tmp_var = "L" & .item("lbl_num")
                .item("lbl_num") = .item("lbl_num") + 1
            End With
            
            If args.item(t) = "else" Then
                dict_append pcode, "GoTo " & tmp_var
                dict_append pcode, args(t - 2).item("label") & ":"
            Else
                dict_append pcode, "If Not " & args.item(t - 1).item("id") _
                                             & " then GoTo " & tmp_var
            End If
            
            Compile.item("label") = tmp_var
            
        Case 28 ' N
            With vars.item(vars.item("cur_env"))
                tmp_var = "L" & .item("lbl_num")
                .item("lbl_num") = .item("lbl_num") + 1
            End With
            
            dict_append pcode, tmp_var & ":"
            
            Compile.item("label") = tmp_var
            
        Case 29 'F
            tmp_var = "Sub " & args.item(t) & "()"
            If Not vars.exists(args.item(t)) Then
                vars.item("cur_env") = args.item(t)
                Set vars.item(args.item(t)) = dict()
                With vars.item(args.item(t))
                    .item("tmp_num") = 0: .item("lbl_num") = 0
                End With
                
            Else
                tmp_var = tmp_var & " 'Error: double definition function"
            End If
            dict_append pcode, tmp_var
        
        Case 30 'loc->loc [ bool ]
            Compile.item("dim") = Join(Array(args.item(t - 3).item("dim"), _
                                             args.item(t - 1).item("id")), ", ")
            If args.item(t - 3).item("dim") <> Empty Then
                Compile.item("id") = Split(args.item(t - 3).item("id"), "(")(0) _
                                     & "(" & Join(Array(Mid(args.item(t - 3).item("dim"), 3), _
                                                        args.item(t - 1).item("id")), ", ") & ")"
            Else
                            Compile.item("id") = args.item(t - 3).item("id") _
                                                 & "(" & args.item(t - 1).item("id") & ")"
            End If

        Case 54: '#-> ( # )
            Set Compile = args.item(t - 1)

        Case 56 'factor->id ( args ) .
            With vars.item(vars.item("cur_env"))
                tmp_var = "t" & .item("tmp_num")
                .item("tmp_num") = .item("tmp_num") + 1
            End With
            
            Compile.item("id") = tmp_var
            
            For Each p In Split(args(t - 1).item("id"), ",")
                dict_append pcode, "push " & p
            Next
            dict_append pcode, "call " & args.item(t - 3) & "()"
            dict_append pcode, "pop " & tmp_var
            
        Case 25, 33, 35, 38, 43, 46, 50, 53, 55, 61, 62: '  #->#
            Set Compile = args.item(t)
            
        Case 31, 57, 58, 59, 60: '#->val
            Compile.item("id") = args.item(t)
            
        Case 32, 34, 36, 37, 39, 40, 41, 42, 44, 45, 47, 48, 49, 51, 52 '#-># % #
            If rule_num = 51 Or rule_num = 52 Then ' #-> % #
                expr = Join(Array(args.item(t - 1), _
                                  args.item(t).item("id")))
            Else
                expr = args.item(t - 1)
                If rule_num = 32 Then expr = "or"
                If rule_num = 34 Then expr = "and"
                If rule_num = 36 Then expr = "="
                If rule_num = 37 Then expr = "<>"
                If rule_num = 49 Then expr = "mod"
                expr = Join(Array(args.item(t - 2).item("id"), _
                                  expr, _
                                  args.item(t).item("id")))
            End If
            
            With vars.item(vars.item("cur_env"))
                If Not .exists(expr) Then
                    tmp_var = "t" & .item("tmp_num")
                    .item("tmp_num") = .item("tmp_num") + 1
                    .item(expr) = tmp_var
                    dict_append pcode, tmp_var & " = " & expr
                Else: tmp_var = .item(expr)
                End If
                Compile.item("id") = tmp_var
            End With
    End Select
    
End Function





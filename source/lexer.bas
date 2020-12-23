Attribute VB_Name = "lexer"
Function lex(ByRef src As String, ByRef lang As Object) As Object
    Set lex = CreateObject("Scripting.Dictionary")
    Dim nfa As Object: Set nfa = post2nfa(re2post(Join(lang.items(), "|")))
    Dim lex_tmp As Object, buffer As Variant: Set lex_tmp = lex
    
    For Each buffer In Split(src, vbNewLine)
        While buffer <> Empty
            Set lex_tmp.item("next") = match(nfa, lang, CStr(buffer))
            If ObjPtr(lex_tmp.item("next")) Then
                buffer = Mid(buffer, Len(lex_tmp.item("next").item("l")) + 1)
                If lex_tmp.item("next").item("t") <> "ws" Then
                    Set lex_tmp = lex_tmp.item("next")
                End If
            Else: buffer = Mid(buffer, 2)
            End If
        Wend
    Next
    Set lex_tmp.item("next") = Nothing
    Set lex = lex.item("next")
End Function

Function load_lex_re(ByVal name_file As String) As Object
    Dim temp_str As String: temp_str = get_file_text(name_file)
    Set load_lex_re = dict()
    For Each line In Split(temp_str, vbNewLine)
        tmp_arr = Split(line, "=>")
        load_lex_re.item(tmp_arr(0)) = tmp_arr(1)
    Next
End Function

Sub lex_for_grammar(ByRef grammar As Object, ByVal name_file As String)
    save_to_file name_file, Join(grammar.item("terms"), "=>" & vbNewLine)
End Sub

Function get_file_text(ByRef file_path As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim text_stream As Object: Set text_stream = fso.OpenTextFile(file_path, 1, True)
    get_file_text = text_stream.ReadAll()
End Function

Sub save_to_file(ByVal outfile As String, ByRef text As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim text_stream As Object: Set text_stream = fso.OpenTextFile(outfile, 2, True)
    text_stream.Write text & vbNewLine: text_stream.Close
End Sub





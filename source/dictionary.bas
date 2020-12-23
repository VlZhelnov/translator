Attribute VB_Name = "dictionary"
Function dict() As Object
    Set dict = CreateObject("Scripting.Dictionary")
End Function
Sub dict_append(ByRef dict As Object, ByVal e As Variant)
    If TypeName(e) = "Dictionary" Then
        Set dict(dict.Count()) = e
    Else: dict(dict.Count()) = e
    End If
End Sub
Function to_str(ByRef s As Variant) As String
    Select Case VarType(s)
        Case 9: to_str = dts(s)
        Case Is > 1000: to_str = ats(s)
        Case Else: to_str = s
    End Select
End Function
Function ats(ByRef arr As Variant, Optional level As Integer = 0) As String
    For Each e In arr
        ats = ats & ", "
        Select Case VarType(e)
            Case 9: ats = ats & " " & dts(e)
            Case Is > 1000: ats = ats & ats(e)
            Case Else: ats = ats & e
        End Select
    Next
    ats = "[" & Mid(ats, 3) & "]"
End Function
Function dts(ByRef dict As Variant) As String
    For Each k In dict
        dts = dts & ", "
        Select Case VarType(dict(k))
            Case 9: dts = dts & dts(dict(k))
            Case Is > 1000: dts = dts & ats(dict(k))
            Case Else: dts = dts & dict(k)
        End Select
    Next
    dts = "{" & Mid(dts, 3) & "}"
End Function
Function dict_to_string(ByVal inst As Object, Optional level As Integer = 0, _
                        Optional ByRef check As Object = Nothing) As String
    
    If ObjPtr(check) = 0 Then Set check = dict()
    If ObjPtr(inst) = 0 Then Exit Function
        
    For Each Key In inst.keys()
        tmp_str = vbNewLine & String(level, vbTab)
        Select Case VarType(inst.item(Key))
            Case 9 ' is Dictionary
                dict_to_string = dict_to_string & tmp_str & "[" & Key & "] "
                If Not check.exists(ObjPtr(inst.item(Key)) & Key) Then
                    check.item(ObjPtr(inst.item(Key)) & Key) = True
                    dict_to_string = dict_to_string & dict_to_string(inst.item(Key), level + 1, check)
                End If
            Case Is > 800 ' is Array
                dict_to_string = dict_to_string & tmp_str & Key & " : " & Join(inst.item(Key), ", ")
            Case Else
                dict_to_string = dict_to_string & tmp_str & Key & vbTab & inst.item(Key)
        End Select
    Next
    
End Function
Sub dict_to_file(ByVal inst As Object, Optional name_file As String = Empty)
    Dim fso As Object, text_stream As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If name_file = Empty Then name_file = Split(fso.GetTempName(), ".")(0) & ".txt"
    Set text_stream = fso.OpenTextFile(fso.BuildPath(ThisWorkbook.Path, name_file), 2, True)
    text_stream.Write Mid(dict_to_string(inst), 2) & vbNewLine
    text_stream.Close
End Sub









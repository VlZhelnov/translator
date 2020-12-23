Attribute VB_Name = "stack_"
Function stack() As Object
    Set stack = CreateObject("Scripting.Dictionary"): stack.add "top", 0
End Function
Function Stack_pop(ByRef stack As Object)
    With stack
        If .item("top") < 1 Then Exit Function
        If IsObject(.item(.item("top"))) Then
            Set Stack_pop = .item(.item("top"))
        Else
            Stack_pop = .item(.item("top"))
        End If
        .Remove .item("top")
        .item("top") = .item("top") - 1
    End With
End Function
Sub Stack_push(ByRef stack As Object, ByRef value)
    With stack
        .item("top") = .item("top") + 1
        .add .item("top"), value
    End With
End Sub
Function stack_top(ByRef stack As Object)
    stack_top = stack.item(stack.item("top"))
End Function

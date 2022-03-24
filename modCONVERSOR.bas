Attribute VB_Name = "modCONVERSOR"
Option Compare Text

Function SIZE_OF_NATIVE_TYPE(ByVal t As CML_TYPE_VARIABLE) As Long
    Select Case t
        Case var_type_byte
            SIZE_OF_NATIVE_TYPE = 1
        Case var_type_word
            SIZE_OF_NATIVE_TYPE = 2
        Case var_type_dword, var_type_float
            SIZE_OF_NATIVE_TYPE = IIf(x64, 8, 4)
        Case var_type_qword
            SIZE_OF_NATIVE_TYPE = 8
    End Select
End Function

Function SIZE_OF(v As CML_VARIABLE_DEFINITION) As Long
    If v.pointer Then
        SIZE_OF = IIf(x64, 8, 4)
    ElseIf v.datatype = var_type_compose Then
        SIZE_OF = G_GLOBAL_COMPOSES(v.extra).size
    Else
        SIZE_OF = SIZE_OF_NATIVE_TYPE(v.datatype)
    End If
End Function


Function OPERATOR_TO_PRECEDENCE(op As String) As Single
    Select Case op
        Case "!"
            OPERATOR_TO_PRECEDENCE = 10#
        Case "*"
            OPERATOR_TO_PRECEDENCE = 9.9
        Case "/"
            OPERATOR_TO_PRECEDENCE = 9.8
        Case "<<"
            OPERATOR_TO_PRECEDENCE = 9.7
        Case ">>"
            OPERATOR_TO_PRECEDENCE = 9.6
        Case "%"
            OPERATOR_TO_PRECEDENCE = 9.5
        Case "^"
            OPERATOR_TO_PRECEDENCE = 9.4
        Case "+"
            OPERATOR_TO_PRECEDENCE = 8.1
        Case "-"
            OPERATOR_TO_PRECEDENCE = 8#
        Case "|"
            OPERATOR_TO_PRECEDENCE = 7.1
        Case "Xor"
            OPERATOR_TO_PRECEDENCE = 7#
        Case "&"
            OPERATOR_TO_PRECEDENCE = 6#
        Case "=="
            OPERATOR_TO_PRECEDENCE = 5.9
        Case "<>", "!="
            OPERATOR_TO_PRECEDENCE = 5.8
        Case "<="
            OPERATOR_TO_PRECEDENCE = 5.7
        Case "<"
            OPERATOR_TO_PRECEDENCE = 5.6
        Case ">"
            OPERATOR_TO_PRECEDENCE = 5.5
        Case ">="
            OPERATOR_TO_PRECEDENCE = 5.4
        Case "="
            OPERATOR_TO_PRECEDENCE = 1#
    End Select
End Function

Function SCOPE_AND_INDEX_TO_VARIABLE(ByVal idx As Long, ByVal scope As CML_VARIABLE_SCOPE) As CML_VARIABLE_DEFINITION
    Select Case scope
        Case var_scope_local
            
    End Select
End Function

Attribute VB_Name = "modVERIFICATION"
Option Explicit
Option Compare Text

Function IS_I32_4BYTES(ByVal datatype As CML_TYPE_VARIABLE) As Boolean
    IS_I32_4BYTES = datatype = var_type_dword Or datatype = var_type_compose
End Function

Private Sub IS_CORRECT_NUMBER_FOR_DATATYPE(ByVal n As Long, ByVal datatype As CML_TYPE_VARIABLE)
    Select Case datatype
        Case var_type_byte
            IS_CORRECT_NUMBER_FOR_DATATYPE = (n >= -128) And (n <= 255)
        Case var_type_compose
        Case var_type_dword
            IS_CORRECT_NUMBER_FOR_DATATYPE = (n >= -2147483648#) And (n <= 4294967295#)
        Case var_type_float
            
        Case var_type_qword
        Case var_type_word
            IS_CORRECT_NUMBER_FOR_DATATYPE = (n >= -32768) And (n <= 65535)
    End Select
End Sub

Function IS_OPERATOR(char As Byte) As Boolean
    Select Case char
        Case 33, 38, 42, 43, 45, 47, 60 To 62, 92, 94, 124, 126, 246
            IS_OPERATOR = True
    End Select
End Function

Function IS_CHARACTER(char As Byte) As Boolean
    Select Case char
        Case 65 To 90, 97 To 122, 160 To 165, 224, 130
            IS_CHARACTER = True
    End Select
End Function

Function IS_NUMERIC(char As Byte) As Boolean
    Select Case char
        Case 48 To 57
            IS_NUMERIC = True
    End Select
End Function

Function IS_SEPARATOR(char As Byte) As Boolean
    Select Case char
        Case 40, 41, 46, 44, 35, 58, 59, 63, 64, 91, 93, 95, 123, 125, 36, 34
            IS_SEPARATOR = True
    End Select
End Function

Function IS_SPACE(char As Byte) As Boolean
    Select Case char
        Case 32, 9, 10, 13
            IS_SPACE = True
    End Select
End Function

Function IS_INSTRUCTION(text As String, ByRef i As CML_INSTRUCTIONS) As Boolean
    IS_INSTRUCTION = True
    Select Case text
        Case "Public"
            i = i_public
        Case "Private"
            i = i_private
        Case "Import"
            i = i_import
        Case "If"
            i = i_if
        Case "Else"
            i = i_else
        Case "Then"
            i = i_then
        Case "Include"
            i = i_include
        Case "Struct"
            i = i_struct
        Case "Union"
            i = i_union
        Case "End"
            i = i_end
        Case "Proc"
            i = i_proc
        Case "Var"
            i = i_var
        Case "Inherit"
            i = i_inherit
        Case Else
            IS_INSTRUCTION = False
    End Select
End Function

Function IS_NATIVE_DATATYPE(text As String, Optional t As CML_TYPE_VARIABLE) As Boolean
    IS_NATIVE_DATATYPE = True
    Select Case text
        Case "BYTE"
            t = var_type_byte
        Case "WORD"
            t = var_type_word
        Case "DWORD"
            t = var_type_dword
        Case "FLOAT"
            t = var_type_float
        Case "QWORD"
            t = var_type_qword
        Case Else
            IS_NATIVE_DATATYPE = False
    End Select
End Function

Private Function VARIABLE_LIST_CONTAINS_NAME(arr() As CML_VARIABLE_DEFINITION, text As String, Optional idx As Long = 1) As Boolean
    For idx = 1 To UBound(arr)
        If arr(idx).name = text Then
            VARIABLE_LIST_CONTAINS_NAME = True
            Exit Function
        End If
    Next
    idx = 1
End Function

Function IS_COMPOSE(text As String, Optional idx As Long = 1) As Boolean
    Dim f As String
    f = PeekS(source_list)
    For idx = 1 To UBound(G_GLOBAL_COMPOSES)
        If G_GLOBAL_COMPOSES(idx).private And G_GLOBAL_COMPOSES(idx).file = f Then
            IS_COMPOSE = G_GLOBAL_COMPOSES(idx).name = text
        ElseIf G_GLOBAL_COMPOSES(idx).private = False Then
            IS_COMPOSE = G_GLOBAL_COMPOSES(idx).name = text
        End If
        If IS_COMPOSE Then
            Exit Function
        End If
    Next
End Function

Function IS_MEMBER_OF_COMPOSE(ByVal idxc As Long, member As String, Optional idx As Long = 1) As Boolean
    IS_MEMBER_OF_COMPOSE = VARIABLE_LIST_CONTAINS_NAME(G_GLOBAL_COMPOSES(idxc).members, member, idx)
End Function

Function IS_GLOBAL_VARIABLE(text As String, Optional idx As Long = 1) As Boolean
    IS_GLOBAL_VARIABLE = VARIABLE_LIST_CONTAINS_NAME(G_GLOBAL_VARIABLES, text, idx)
End Function

Function IS_LOCAL_VARIABLE(text As String, Optional idx As Long = 1) As Boolean
    Dim idxp As Long
    
    If UBound(G_PROC_IDX_ARRAY) = 0 Then
        Exit Function
    End If
    
    idxp = PeekL(G_PROC_IDX_ARRAY)
    
    IS_LOCAL_VARIABLE = VARIABLE_LIST_CONTAINS_NAME(G_GLOBAL_PROCEDURES(idxp).locals, text, idx)
End Function

Function IS_PARAMETER(text As String, Optional idx As Long = 1) As Boolean
    Dim idxp As Long
    
    If UBound(G_PROC_IDX_ARRAY) = 0 Then
        Exit Function
    End If
    
    idxp = PeekL(G_PROC_IDX_ARRAY)
    
    IS_PARAMETER = VARIABLE_LIST_CONTAINS_NAME(G_GLOBAL_PROCEDURES(idxp).params, text, idx)
End Function

Function IS_PROCEDURE(text As String, Optional idx As Long = 1) As Boolean
    Dim f As String
    f = PeekS(source_list)
    For idx = idx To UBound(G_GLOBAL_PROCEDURES)
        If G_GLOBAL_PROCEDURES(idx).private And G_GLOBAL_PROCEDURES(idx).file = f Then
            IS_PROCEDURE = G_GLOBAL_PROCEDURES(idx).name = text
        ElseIf G_GLOBAL_PROCEDURES(idx).private = False Then
            IS_PROCEDURE = G_GLOBAL_PROCEDURES(idx).name = text
        End If
        If IS_PROCEDURE Then
            Exit Function
        End If
    Next
End Function

Function IS_ANY_VARIABLE(text As String, Optional idx As Long = 1, Optional scope As CML_VARIABLE_SCOPE) As Boolean
    Select Case True
        Case IS_LOCAL_VARIABLE(text, idx)
            scope = var_scope_local
        Case IS_PARAMETER(text, idx)
            scope = var_scope_param
        Case IS_GLOBAL_VARIABLE(text, idx)
            scope = var_scope_global
    End Select
    IS_ANY_VARIABLE = scope
End Function

Function IS_DATATYPE(text As String, Optional t As CML_TYPE_VARIABLE, Optional idx As Long) As Boolean
    IS_DATATYPE = True
    Select Case True
        Case IS_NATIVE_DATATYPE(text, t)
            ' DoNothing
        Case IS_COMPOSE(text, idx)
            t = var_type_compose
        Case Else
            IS_DATATYPE = False
    End Select
End Function

Function IS_VALID_IDENTIFIER(text As String, Optional ByVal globalvariables As Boolean = True, _
                                            Optional ByVal localvariables As Boolean = True, _
                                            Optional ByVal parameters As Boolean = True, _
                                            Optional ByVal procedures As Boolean = True, _
                                            Optional ByVal datatypes As Boolean = True) As Boolean
    IS_VALID_IDENTIFIER = Not (IS_GLOBAL_VARIABLE(text) And globalvariables) Or _
                              (IS_LOCAL_VARIABLE(text) And localvariables) Or _
                              (IS_PARAMETER(text) And parameters) Or _
                              (IS_PROCEDURE(text) And procedures) Or _
                              (IS_DATATYPE(text) And datatypes)
End Function

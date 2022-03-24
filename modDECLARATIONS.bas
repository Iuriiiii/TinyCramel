Attribute VB_Name = "modDECLARATIONS"
Option Explicit

Public main_file As String
Public x64 As Boolean
Public source_list() As String ' Lista de códigos fuentes actualmente siendo procesados

Enum CML_VARIABLE_PARSING_TYPE
    variable_parsing_type_variable
    variable_parsing_type_parameter
    variable_parsing_type_member
End Enum

Enum CML_INSTRUCTIONS
    i_expression = 0
    i_private = 1
    i_public
    i_import
    i_declare
    i_prototype
    i_struct
    i_union
    i_include
    i_var
    i_if
    i_then
    i_else
    i_end
    i_proc
    i_inherit
End Enum

Enum CML_TOKEN_TYPE
    TOKEN_TYPE_NONE = 0
    TOKEN_TYPE_IDENTIFIER
    TOKEN_TYPE_STRING
    TOKEN_TYPE_OPERATOR
    TOKEN_TYPE_SEPARATOR
    TOKEN_TYPE_NUMBER
    TOKEN_TYPE_FLOAT
    TOKEN_TYPE_INSTRUCTION
    TOKEN_TYPE_EOL
    TOKEN_TYPE_EOF
    TOKEN_TYPE_REGISTER
End Enum

Type CML_TOKEN
    t As CML_TOKEN_TYPE
    i As CML_INSTRUCTIONS
    s As String
End Type

Enum CML_TYPE_VARIABLE
    var_type_byte = 1
    var_type_word
    var_type_dword
    var_type_qword
    var_type_float
    var_type_compose
End Enum

Enum CML_TYPE_PROCEDURE
    CALL_STDCALL = 0
    CALL_CDECL
    CALL_FASTCALL
End Enum

Enum CML_VARIABLE_SCOPE
    var_scope_local = 1
    var_scope_param
    var_scope_global
End Enum

Type CML_VARIABLE_DEFINITION
    name As String
    datatype As CML_TYPE_VARIABLE
    pointer As Long
    extra As Long ' Indice del tipo de dato compuesto
    ' Miembros
    private As Boolean
    ' Interno
    uname As String
    offset As Long
    file As String ' Ruta del código fuente dueño
End Type

Type CML_COMPOSE_DEFINITION
    name As String
    members() As CML_VARIABLE_DEFINITION
    size As Long
    private As Boolean
    file As String
End Type

Type CML_EXPRESSION_DEFINITION
    tokens() As CML_TOKEN
End Type

Type CML_LINES_DEFINITION
    general As Boolean
    l As Long
    i As CML_INSTRUCTIONS
    idx As Long
End Type

Type CML_PROCEDURE_DEFINITION
    name As String
    uname As String
    locals() As CML_VARIABLE_DEFINITION
    params() As CML_VARIABLE_DEFINITION
    return As CML_VARIABLE_DEFINITION
    export As Boolean
    library As String
    type As CML_TYPE_PROCEDURE
    instructions() As Long
    private As Boolean
    file As String
End Type

Type CML_END_INSTRUCTION_DEFINITION
    instruction As CML_INSTRUCTIONS
End Type

Public G_GLOBAL_VARIABLES() As CML_VARIABLE_DEFINITION
Public G_GLOBAL_PROCEDURES() As CML_PROCEDURE_DEFINITION
Public G_GLOBAL_COMPOSES() As CML_COMPOSE_DEFINITION
Public G_EXPRESSIONS() As CML_EXPRESSION_DEFINITION
Public G_GLOBAL_LINES() As CML_LINES_DEFINITION
Public G_NESTED() As CML_INSTRUCTIONS
Public G_ENDED() As CML_END_INSTRUCTION_DEFINITION

' Si el contador está en 1, el procedimiento es normal.
' Si el contador es mayor a 1, el procedimiento es una expresión
Public G_PROC_COUNTER As Long
Public G_PROC_IDX_ARRAY() As Long

Sub AddNewLine(ByVal i As CML_INSTRUCTIONS, Optional ByVal index As Long)
    Dim idx As Long
    ReDim Preserve G_GLOBAL_LINES(UBound(G_GLOBAL_LINES) + 1)
    G_GLOBAL_LINES(UBound(G_GLOBAL_LINES)).i = i
    G_GLOBAL_LINES(UBound(G_GLOBAL_LINES)).idx = index
    G_GLOBAL_LINES(UBound(G_GLOBAL_LINES)).general = UBound(G_PROC_IDX_ARRAY) = 0
    If UBound(G_PROC_IDX_ARRAY) Then
        idx = PeekL(G_PROC_IDX_ARRAY)
        PushL G_GLOBAL_PROCEDURES(idx).instructions, UBound(G_GLOBAL_LINES)
    End If
End Sub

Sub INIT_DECLARATIONS()
    ReDim G_GLOBAL_VARIABLES(0)
    ReDim G_GLOBAL_PROCEDURES(0)
    ReDim G_GLOBAL_COMPOSES(0)
    ReDim G_EXPRESSIONS(1)
    ReDim G_EXPRESSIONS(1).tokens(0)
    ReDim G_PROC_IDX_ARRAY(0)
    ReDim source_list(0)
    ReDim G_LINES(0)
    ReDim G_NESTED(0)
    ReDim G_GLOBAL_LINES(0)
    ReDim G_ENDED(0)
End Sub

Sub PushI(arr() As CML_INSTRUCTIONS, ByVal i As CML_INSTRUCTIONS)
    ReDim Preserve arr(UBound(arr) + 1)
    arr(UBound(arr)) = i
End Sub

Function PeekI(arr() As CML_INSTRUCTIONS)
    PeekI = arr(UBound(arr))
End Function

Function PopI(arr() As CML_INSTRUCTIONS) As CML_INSTRUCTIONS
    If UBound(arr) Then
        PopI = arr(UBound(arr))
        ReDim Preserve arr(UBound(arr) - 1)
    End If
End Function

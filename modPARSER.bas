Attribute VB_Name = "modPARSER"
Option Explicit
Private private_section As Boolean

Sub PARSE(Optional file As String)
    Dim tokens() As CML_TOKEN
    Dim idx As Long
    
    If file = Empty Then file = main_file
    
    Assert FILE_EXISTS(file), "El archivo '[1]' no pudo ser abierto.", file
    
    private_section = False
    tokens = TOKENIZATE(FILE_READ_AS_ARRAY(file))
    
    'ShowTokens tokens
    'End
    
    PushS source_list, file
    PARSE_TOKENS tokens, idx
    PopS source_list
End Sub

Private Sub PARSE_VARIABLE_DECLARATIONS(tokens() As CML_TOKEN, idx As Long, ret() As CML_VARIABLE_DEFINITION, ByVal ptype As CML_VARIABLE_PARSING_TYPE)
    Dim initialpos As Long
    Dim offset As Long
    
    initialpos = UBound(ret) + 1
    
    'If initialpos = 1 Then
    ReDim Preserve ret(initialpos)
    'End If
    
    offset = ret(UBound(ret)).offset + SIZE_OF(ret(UBound(ret)))
    
    Do While True
        Select Case tokens(idx).t
            Case TOKEN_TYPE_EOL, TOKEN_TYPE_EOF
                Exit Do
            Case TOKEN_TYPE_SEPARATOR
                Select Case tokens(idx).s
                    Case "["
                        Dim expr() As CML_TOKEN
                        
                        Assert ptype = variable_parsing_type_parameter, "Imposible establecer una dimensión a un parámetro."
                        
                        expr = EXTRACT_FROM_SEPARATORS(tokens, idx)
                    Case ","
                        ReDim Preserve ret(UBound(ret) + 1)
                        
                        ret(UBound(ret)).private = private_section
                        ret(UBound(ret)).file = PeekS(source_list)
                    Case "@"
                        Inc ret(UBound(ret)).pointer
                    Case ":"
                        Dim t As CML_TYPE_VARIABLE
                        Dim idxc As Long
                        
                        Inc idx
                        Assert tokens(idx).t = TOKEN_TYPE_IDENTIFIER, "Se identificador del tipo de dato."
                        Assert IS_DATATYPE(tokens(idx).s, t, idxc), "Tipo de dato inválido: '[1]'.", tokens(idx).s
                        
                        For initialpos = initialpos To UBound(ret)
                        
                            If ptype = variable_parsing_type_parameter And t = var_type_compose Then
                                Assert ret(initialpos).pointer, "Un parámetro de tipo compuesto debe ser un puntero."
                            End If
                        
                            ret(initialpos).datatype = t
                            
                            If t = var_type_compose Then
                                ret(initialpos).extra = idxc
                            End If
                        Next
                        
                        'Inc initialpos
                End Select
            Case TOKEN_TYPE_INSTRUCTION
                Select Case tokens(idx).i
                    Case i_private
                        Assert ptype = variable_parsing_type_member, "Solamente se pueden privatizar miembros."
                        ret(UBound(ret)).private = True
                End Select
            Case TOKEN_TYPE_IDENTIFIER
                Assert ret(UBound(ret)).name = Empty, "La variable, miembro o parámetro ya tiene nombre."
                ret(UBound(ret)).uname = UString
                ret(UBound(ret)).name = tokens(idx).s
            Case Else
                Exit Do
        End Select
        Inc idx
    Loop
End Sub

Private Function EXTRACT_TOKENS(tokens() As CML_TOKEN, ByVal min As Long, ByVal max As Long) As CML_TOKEN()
    Dim ret() As CML_TOKEN
    
    ReDim ret(0)
    
    While min < max
        ReDim Preserve ret(UBound(ret) + 1)
        ret(UBound(ret)) = tokens(min)
        Inc min
    Wend
    
    EXTRACT_TOKENS = ret
End Function

Function EXTRACT_FROM_SEPARATORS(tokens() As CML_TOKEN, idx As Long, Optional min As Long, Optional max As Long) As CML_TOKEN()
    Dim par As Long
    Dim cor As Long
    Dim lla As Long
    Dim ret() As CML_TOKEN
    
    Assert tokens(idx).t = TOKEN_TYPE_SEPARATOR, "Se esperaba separador."
    
    If min = 0 Then
        min = idx + 1
    End If
    
    Do While True
        Select Case tokens(idx).t
            Case TOKEN_TYPE_SEPARATOR
                Select Case tokens(idx).s
                    Case "("
                        Inc par
                    Case ")"
                        Dec par
                    Case "["
                        Inc cor
                    Case "]"
                        Dec cor
                    Case "{"
                        Inc lla
                    Case "}"
                        Dec lla
                End Select
        End Select
        Inc idx
        If par = 0 And cor = 0 And lla = 0 Then
            Exit Do
        End If
    Loop
    
    max = idx - 1
    ret = EXTRACT_TOKENS(tokens, min, max)
    PushEOL ret
    
    EXTRACT_FROM_SEPARATORS = ret
End Function

Private Sub PARSE_TOKENS(tokens() As CML_TOKEN, idx As Long)
    
    Inc idx
    
    Select Case tokens(idx).t
        Case TOKEN_TYPE_INSTRUCTION
            Select Case tokens(idx).i
                Case i_private, i_public
                    PARSE_SECTION tokens, idx, tokens(idx).i = i_private
                Case i_struct, i_union
                    PARSE_COMPOSE tokens, idx, tokens(idx).i = i_union
                Case i_include
                    PARSE_INCLUDE tokens, idx
                Case i_import
                    
                Case i_var
                    PARSE_VAR tokens, idx
                Case i_if
                    
                Case i_then
                    
                Case i_proc
                    PARSE_PROC tokens, idx
                Case i_end
                    PARSE_END tokens, idx
            End Select
        Case Else
            PARSE_EXPRESSION tokens, idx
    End Select
    
    If tokens(idx).t = TOKEN_TYPE_EOF Then
        Exit Sub
    End If
    
    PARSE_TOKENS tokens, idx
End Sub
Private Sub PARSE_EXPRESSION(tokens() As CML_TOKEN, idx As Long)
    Dim n1 As Long
    Dim n2 As Long
    
    n1 = UBound(G_EXPRESSIONS)
    n2 = UBound(G_EXPRESSIONS(n1).tokens) + 1
    ReDim Preserve G_EXPRESSIONS(n1).tokens(n2)
    G_EXPRESSIONS(n1).tokens(n2) = tokens(idx)
    
    If tokens(idx).t = TOKEN_TYPE_EOL Or tokens(idx).t = TOKEN_TYPE_EOF Then
        n1 = UBound(G_EXPRESSIONS)
        AddNewLine i_expression, n1
        ReDim Preserve G_EXPRESSIONS(n1 + 1)
        ReDim Preserve G_EXPRESSIONS(n1 + 1).tokens(0)
    End If
End Sub
Private Sub PARSE_SECTION(tokens() As CML_TOKEN, idx As Long, ByVal isprivate As Boolean)
    Inc idx
    Assert tokens(idx).t = TOKEN_TYPE_SEPARATOR, "Se esperaba separador."
    Assert tokens(idx).s = ":", "Se esperaba ':'."
    
    private_section = isprivate
End Sub

Private Sub PARSE_COMPOSE(tokens() As CML_TOKEN, idx As Long, ByVal isunion As Boolean)
    Dim name As String
    Dim idxc As Long
    
    Inc idx
    Assert tokens(idx).t = TOKEN_TYPE_IDENTIFIER, "Se esperaba identificador."
    Assert IS_VALID_IDENTIFIER(tokens(idx).s), "Identificador inválido."
    
    name = tokens(idx).s
    idxc = UBound(G_GLOBAL_COMPOSES) + 1
    
    Inc idx
    Assert tokens(idx).t = TOKEN_TYPE_SEPARATOR, "Se esperaba separador."
    Assert tokens(idx).s = ",", "Se esperaba ','."
    Inc idx
    
    ReDim Preserve G_GLOBAL_COMPOSES(idxc)
    ReDim G_GLOBAL_COMPOSES(idxc).members(0)
    G_GLOBAL_COMPOSES(idxc).file = PeekS(source_list)
    G_GLOBAL_COMPOSES(idxc).private = private_section
    
    G_GLOBAL_COMPOSES(idxc).name = name
    
    PARSE_VARIABLE_DECLARATIONS tokens, idx, G_GLOBAL_COMPOSES(idxc).members, variable_parsing_type_member
End Sub

Private Sub PARSE_INCLUDE(tokens() As CML_TOKEN, idx As Long)
    Inc idx
    Assert tokens(idx).t = TOKEN_TYPE_STRING, "Se esperaba una cadena."
    ' TODO: COMPLETAR
    ' Añadir variables
    PARSE tokens(idx).s
End Sub

Private Sub PARSE_PROC(tokens() As CML_TOKEN, idx As Long)
    Dim name As String
    Dim idxp As Long
    
    Inc idx
    idxp = UBound(G_GLOBAL_PROCEDURES) + 1
    PushL G_PROC_IDX_ARRAY, idxp

    If UBound(G_PROC_IDX_ARRAY) > 1 Then
        Assert tokens(idx).t = TOKEN_TYPE_SEPARATOR Or tokens(idx).t = TOKEN_TYPE_EOL, "Se esperaba separador o salto de línea."
        name = "Procedure_" + CStr(idxp)
    Else
        Assert tokens(idx).t = TOKEN_TYPE_IDENTIFIER, "Se esperaba identificador."
        Assert IS_VALID_IDENTIFIER(tokens(idx).s), "Identificador inválido."
        name = tokens(idx).s
        Inc idx
    End If
    
    ReDim Preserve G_GLOBAL_PROCEDURES(idxp)
    ReDim G_GLOBAL_PROCEDURES(idxp).locals(0)
    ReDim G_GLOBAL_PROCEDURES(idxp).params(0)
    ReDim G_GLOBAL_PROCEDURES(idxp).instructions(0)
    G_GLOBAL_PROCEDURES(idxp).file = PeekS(source_list)
    G_GLOBAL_PROCEDURES(idxp).name = name
    
    If tokens(idx).t = TOKEN_TYPE_SEPARATOR Then
        Select Case tokens(idx).s
            Case "("
                Dim variables() As CML_TOKEN
                
                variables = EXTRACT_FROM_SEPARATORS(tokens, idx)
                
                If UBound(variables) Then
                    PARSE_VARIABLE_DECLARATIONS variables, 1, G_GLOBAL_PROCEDURES(idxp).params, variable_parsing_type_parameter
                End If
            Case ":"
                Dim t As CML_TYPE_VARIABLE
                Dim idxc As Long
                Dim ptr As Long
                
                Inc idx
                Assert tokens(idx).t = TOKEN_TYPE_IDENTIFIER, "Se identificador del tipo de dato."
                Assert IS_DATATYPE(tokens(idx).s, t, idxc), "Tipo de dato inválido: '[1]'.", tokens(idx).s
                
                If t = var_type_compose Then
                    Inc idx
                    Assert tokens(idx).t = TOKEN_TYPE_SEPARATOR, "Se esperaba '@'."
                    Assert tokens(idx).s = "@", "Se esperaba '@'."
                    Inc G_GLOBAL_PROCEDURES(idxp).return.pointer
                End If
                G_GLOBAL_PROCEDURES(idxp).return.datatype = t
            Case Else
                Assert False, "Se esperaba '(' o ':'."
        End Select
        Inc idx
    End If
    
    Assert tokens(idx).t = TOKEN_TYPE_EOL, "Se esperaba salto de línea."
    AddNewLine i_proc, idxp
End Sub

Private Sub PARSE_VAR(tokens() As CML_TOKEN, idx As Long)
    Dim idxp As Long
    
    Inc idx
    
    Select Case True
        Case UBound(G_PROC_IDX_ARRAY)
            idxp = PeekL(G_PROC_IDX_ARRAY)
            PARSE_VARIABLE_DECLARATIONS tokens, idx, G_GLOBAL_PROCEDURES(idxp).locals, variable_parsing_type_variable
            
        Case Else
            PARSE_VARIABLE_DECLARATIONS tokens, idx, G_GLOBAL_VARIABLES, variable_parsing_type_variable
    End Select
    
    AddNewLine i_var
End Sub

Private Sub PARSE_END(tokens() As CML_TOKEN, idx As Long)
    Dim idxe As Long
    
    Inc idx
    
    Assert tokens(idx).t = TOKEN_TYPE_INSTRUCTION, "Se esperaba instrucción a terminar."
    
    idxe = UBound(G_ENDED) + 1
    ReDim Preserve G_ENDED(idxe)
    
    Select Case tokens(idx).i
        Case i_proc
            Assert UBound(G_PROC_IDX_ARRAY), "Cierre de procedimiento inválido."
            PopL G_PROC_IDX_ARRAY
            G_ENDED(idxe).instruction = i_proc
        Case i_if
            
    End Select
    AddNewLine i_end, idxe
End Sub

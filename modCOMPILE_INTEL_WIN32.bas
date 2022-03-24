Attribute VB_Name = "modCOMPILE_INTEL_WIN32"
Option Explicit

Private Type CML_VALUE_POINTER_DEFINITION
    value As String
    t As CML_TOKEN_TYPE
    name As String
End Type

Private Enum CML_REGISTER
    NONE = 0
    eax
    EBX
    ECX
    EDX
    EDI
    ESI
    AL
    AH
    AX
    BL
    BH
    BX
    CL
    CH
    CX
    DL
    DH
    DX
End Enum

Private Type CML_COMPILED_IDENTIFIER_DEFINITION
    scope As CML_VARIABLE_SCOPE
    text As String
    register As CML_REGISTER
    code As String
    v As CML_VARIABLE_DEFINITION
    waitingmember As Boolean
    pointer As Long
    inmemory As Boolean
    lastchar As String ' Interno
    'ispurepointer as Long
End Type

Private code As String
Private G_PRIVATE_VALUES() As CML_VALUE_POINTER_DEFINITION

Private Function REGISTER_TO_STRING(ByVal regiser As CML_REGISTER) As String
    Select Case register
        Case NONE: Exit Function
        Case eax: REGISTER_TO_STRING = "eax"
        Case EBX: REGISTER_TO_STRING = "ebx"
        Case ECX: REGISTER_TO_STRING = "ecx"
        Case EDX: REGISTER_TO_STRING = "edx"
        Case EDI: REGISTER_TO_STRING = "edi"
        Case ESI: REGISTER_TO_STRING = "esi"
    End Select
End Function

Private Sub WLine(line As String, ParamArray param() As Variant)
    Dim i As Long
    For i = 0 To UBound(param)
        line = Replace$(line, "[" + CStr(i) + "]", CStr(param(i)))
    Next
    Ins code, line
    Ins code, vbCrLf
End Sub

Sub COMPILE_INTEL_WIN32()
    COMPILE_INITIALIZE
    COMPILE_HEADER
    COMPILE_DATA_READEABLE_WRITEABLE_SECTION
    COMPILE_EXECUTABE_CODE_SECTION
    COMPILE_INSTRUCTIONS
End Sub

Private Sub COMPILE_INITIALIZE()
    ReDim G_PRIVATE_VALUES(0)
End Sub

Private Sub COMPILE_HEADER()
    WLine "format PE GUI 4.0"
End Sub

Private Sub COMPILE_DATA_READEABLE_WRITEABLE_SECTION()
    Dim i As Long
    WLine "section '.data' data readable writeable"
    For i = 1 To UBound(G_GLOBAL_VARIABLES)
        WLine "[0] rb [1] dup(?)", G_GLOBAL_VARIABLES(i).uname, SIZE_OF(G_GLOBAL_VARIABLES(i))
    Next
    For i = 1 To UBound(G_PRIVATE_VALUES)
        Select Case G_PRIVATE_VALUES(i).t
            Case TOKEN_TYPE_NUMBER
                WLine "[0] dd [1]", G_PRIVATE_VALUES(i).name, G_PRIVATE_VALUES(i).value
            Case TOKEN_TYPE_STRING
                WLine "[0] db '[1]'", G_PRIVATE_VALUES(i).name, G_PRIVATE_VALUES(i).value
            Case TOKEN_TYPE_FLOAT
                WLine "[0] dq [1]", G_PRIVATE_VALUES(i).name, G_PRIVATE_VALUES(i).value
        End Select
    Next
End Sub

Private Sub COMPILE_EXECUTABE_CODE_SECTION()
    WLine "section '.text' code readable executable"
    WLine "entry start"
    WLine "start:"
End Sub

Private Sub COMPILE_INSTRUCTIONS(Optional ByVal min As Long = 1, Optional ByVal max As Long, Optional ByVal general As Boolean = True)
    
    max = IIf(max, max, UBound(G_GLOBAL_LINES))
    
    For min = min To max
        If G_GLOBAL_LINES(min).general = general Then
            Select Case G_GLOBAL_LINES(min).i
                Case i_proc
                    PushL G_PROC_IDX_ARRAY, G_GLOBAL_LINES(min).idx
                Case i_end
                    Select Case G_ENDED(G_GLOBAL_LINES(min).idx).instruction
                        Case i_proc
                            PopL G_PROC_IDX_ARRAY
                    End Select
                Case i_expression
                    Dim rpn() As SHUNTING_ITEM_DEFINE
                    
                    rpn = SHUNTING_YARD(G_EXPRESSIONS(G_GLOBAL_LINES(min).idx).tokens)
                    
                    If UBound(rpn) Then
                        'ShowShuntingYard rpn
                        COMPILE_EXPRESSION rpn
                    End If
            End Select
        End If
    Next
End Sub

Function GET_COMPILED_CODE() As String
    GET_COMPILED_CODE = code
End Function

Private Sub COMPILE_FOOTER()
    
End Sub

Private Function COMPILE_EXPRESSION(rpn() As SHUNTING_ITEM_DEFINE, Optional register As CML_REGISTER = eax) As SHUNTING_ITEM_DEFINE
    Dim i As Long
    Dim stack() As SHUNTING_ITEM_DEFINE
    
    ReDim stack(0)
    
    For i = 1 To UBound(rpn)
        Select Case rpn(i).tokens(1).t
            Case TOKEN_TYPE_SEPARATOR
                
            Case TOKEN_TYPE_OPERATOR
                COMPILE_EXPRESSION_BY_OPERATOR stack, GetTokenS(rpn(i).tokens(1)), register
            Case Else
                PushSID stack, rpn(i)
        End Select
    Next
    
    COMPILE_EXPRESSION = stack(1)
End Function

Private Sub COMPILE_EXPRESSION_BY_OPERATOR(stack() As SHUNTING_ITEM_DEFINE, op As String, register As CML_REGISTER)
    Dim l As SHUNTING_ITEM_DEFINE
    Dim r As SHUNTING_ITEM_DEFINE
    
    Select Case op
        Case "!"
            Assert UBound(stack), "Se esperaba expresión."
            
        Case Else
            Assert UBound(stack) > 1, "Se esperaba expresión."
            
            r = PopSID(stack)
            l = PopSID(stack)
            
            Select Case True
                Case l.tokens(1).t = TOKEN_TYPE_NUMBER And r.tokens(1).t = TOKEN_TYPE_NUMBER
                    PushSID stack, COMPILE_NUMBER_WITH_NUMBER(l, r, op, register)
                Case l.tokens(1).t = TOKEN_TYPE_NUMBER And r.tokens(1).t = TOKEN_TYPE_STRING
                    
                Case l.tokens(1).t = TOKEN_TYPE_REGISTER And r.tokens(1).t = TOKEN_TYPE_IDENTIFIER
                    PushSID stack, COMPILE_REGISTER_WITH_IDENTIFIER(l, r, op, register)
                Case l.tokens(1).t = TOKEN_TYPE_IDENTIFIER And r.tokens(1).t = TOKEN_TYPE_REGISTER
                    PushSID stack, COMPILE_IDENTIFIER_WITH_REGISTER(l, r, op, register)
                Case l.tokens(1).t = TOKEN_TYPE_IDENTIFIER And r.tokens(1).t = TOKEN_TYPE_NUMBER
                    PushSID stack, COMPILE_IDENTIFIER_WITH_NUMBER(l, r, op, register)
                Case l.tokens(1).t = TOKEN_TYPE_REGISTER And r.tokens(1).t = TOKEN_TYPE_NUMBER
                    PushSID stack, COMPILE_REGISTER_WITH_NUMBER(l, r, op, register)
                Case l.tokens(1).t = TOKEN_TYPE_NUMBER And r.tokens(1).t = TOKEN_TYPE_REGISTER
                    PushSID stack, COMPILE_NUMBER_WITH_REGISTER(l, r, op, register)
                Case l.tokens(1).t = TOKEN_TYPE_NUMBER And r.tokens(1).t = TOKEN_TYPE_IDENTIFIER
                    PushSID stack, COMPILE_NUMBER_WITH_IDENTIFIER(l, r, op, register)
                Case l.tokens(1).t = TOKEN_TYPE_REGISTER And r.tokens(1).t = TOKEN_TYPE_REGISTER
                    PushSID stack, COMPILE_REGISTER_WITH_REGISTER(l, r, op, register)
                Case l.tokens(1).t = TOKEN_TYPE_IDENTIFIER And r.tokens(1).t = TOKEN_TYPE_IDENTIFIER
                    PushSID stack, COMPILE_IDENTIFIER_WITH_IDENTIFIER(l, r, op, register)
            End Select
    End Select
End Sub

Private Sub COMPILE_IDENTIFIER(expr As SHUNTING_ITEM_DEFINE, ret As CML_COMPILED_IDENTIFIER_DEFINITION, register As CML_REGISTER, Optional idx As Long, Optional ByVal varregister As CML_REGISTER = EDI, Optional ByVal nowrite As Boolean)
    Dim idxv As Long
    Dim idxm As Long
    Dim scope As CML_VARIABLE_SCOPE
    Dim tokens() As CML_TOKEN
    Dim r As String
    Dim r2 As String
    
    r = REGISTER_TO_TEXT(varregister)
    tokens = expr.tokens
    Inc idx
    
    If idx > UBound(tokens) Then
        If Not nowrite Then
            WLine ret.code
            ret.code = Empty
        End If
        Exit Sub
    End If
    
    If idx = 1 Then
        ret.code = Empty
        
        Assert IS_ANY_VARIABLE(GetTokenS(tokens(idx)), idxv, ret.scope), "Se esperaba variable."
        
        Select Case ret.scope
            Case var_scope_global
                ret.v = G_GLOBAL_VARIABLES(idxv)
                ret.text = ret.v.uname
            Case var_scope_local
                ret.v = G_GLOBAL_PROCEDURES(PeekL(G_PROC_IDX_ARRAY)).locals(idxv)
                ret.text = "ebp-4-" + CStr(ret.v.offset)
            Case var_scope_param
                ret.v = G_GLOBAL_PROCEDURES(PeekL(G_PROC_IDX_ARRAY)).params(idxv)
                ret.text = "ebp+8+" + CStr(ret.v.offset)
        End Select
        
        ret.register = IIf(ret.v.pointer, varregister, NONE)
        
        If ret.v.pointer Then
            InsLine ret.code, "mov [1],DWORD[[2]]", r, ret.text
            ret.text = r
        End If
        
    Else
        If ret.waitingmember Then
            Assert IS_MEMBER_OF_COMPOSE(ret.v.extra, tokens(idx).s, idxm), "El miembro '[0]' no existe en el tipo de dato compuesto '[1]'.", tokens(idx).s, G_GLOBAL_COMPOSES(ret.v.extra).name
            
            ret.v = G_GLOBAL_COMPOSES(ret.v.extra).members(idxm)
            ret.register = IIf(ret.v.pointer, varregister, NONE)
            
            If ret.register Then
                InsLine ret.code, "mov [1],DWORD[[2]]", r, ret.text
                ret.text = r
            Else
                Ins ret.text, "+" + CStr(ret.v.offset)
            End If
            
            ret.waitingmember = False
        Else
            Assert tokens(idx).t = TOKEN_TYPE_SEPARATOR, "Se esperaba separador."
            
            Select Case tokens(idx).s
                Case "@"
                    Inc ret.pointer
                    
                    
                    
                    'Select Case ret.scope
                    '    Case var_scope_global
                    '        If ret.pointer > 1 Then
                    '            ret.text = sprintf("DWORD[[0]]", ret.text)
                    '            ret.inmemory = True
                    '        End If
                    '    Case var_scope_local, var_scope_param
                    '        If ret.pointer = 1 Then
                                InsLine ret.code, "[1] [2],DWORD[[3]]", IIf(ret.pointer = 1, IIf(ret.register, "mov", "lea"), "mov"), r, ret.text
                                ret.text = r
                                ret.register = varregister
                   '         Else
                   '             ret.register = varregister
                   '             InsLine ret.code, "mov [1],DWORD[[2]]", r, ret.text
                   '             ret.text = r
                   '         End If
                   ' End Select
                Case "["
                    Dim dimension() As CML_TOKEN
                    Dim sid() As SHUNTING_ITEM_DEFINE
                    Dim compiled As SHUNTING_ITEM_DEFINE
                    
                    dimension = EXTRACT_FROM_SEPARATORS(tokens, idx)
                    sid = SHUNTING_YARD(dimension)
                    compiled = COMPILE_EXPRESSION(sid, register + 1)
                    
                    ' TODO: terminar
                    Select Case compiled.tokens(1).t
                        Case TOKEN_TYPE_IDENTIFIER
                            Dim cexpr As CML_COMPILED_IDENTIFIER_DEFINITION
                            COMPILE_IDENTIFIER compiled, cexpr, register, , EDI
                            
                            InsLine ret.code, cexpr.code
                            InsLine ret.code, "add [1],[2]", r, "edi"
                            
                        Case TOKEN_TYPE_STRING
                            
                        Case TOKEN_TYPE_NUMBER
                            InsLine "add [1],[2]", r, CLng(compiled.tokens(1).s)
                        Case TOKEN_TYPE_REGISTER
                            InsLine ret.code, "add [1],[2]", r, REGISTER_TO_TEXT(compiled.extra)
                        Case Else
                            Assert False, "Dimensión inválida."
                    End Select
                
                Case "."
                    Assert ret.v.datatype = var_type_compose, "Se esperaba tipo de dato compuesto."
                    ret.waitingmember = True
                    
                Case "("
                    
            End Select
        End If
    End If
    
    'ret.lastchar = tokens(idx).s
    
    COMPILE_IDENTIFIER expr, ret, register, idx, varregister, nowrite
End Sub

Private Function COMPILE_NUMBER_WITH_NUMBER(l As SHUNTING_ITEM_DEFINE, r As SHUNTING_ITEM_DEFINE, op As String, register As CML_REGISTER) As SHUNTING_ITEM_DEFINE
    Dim ret As SHUNTING_ITEM_DEFINE

    ReDim ret.tokens(1)
    
    ret.tokens(1).t = TOKEN_TYPE_NUMBER
    ret.tokens(1).s = CALC_NUMBER_WITH_NUMBER(CLng(GetTokenS(l.tokens(1))), CLng(GetTokenS(r.tokens(1))), op)
    
    COMPILE_NUMBER_WITH_NUMBER = ret
End Function

Private Function COMPILE_REGISTER_WITH_NUMBER(l As SHUNTING_ITEM_DEFINE, r As SHUNTING_ITEM_DEFINE, op As String, register As CML_REGISTER) As SHUNTING_ITEM_DEFINE
    Dim ret As SHUNTING_ITEM_DEFINE
    
    ReDim ret.tokens(1)
    
    ret.tokens(1).t = TOKEN_TYPE_REGISTER
    ret.extra = l.extra
    
    ret.extra = CALC_REGISTER_WITH_NUMBER(ret.extra, CLng(GetTokenS(r.tokens(1))), op, False)
    
    COMPILE_REGISTER_WITH_NUMBER = ret
End Function

Private Function COMPILE_REGISTER_WITH_IDENTIFIER(l As SHUNTING_ITEM_DEFINE, r As SHUNTING_ITEM_DEFINE, op As String, register As CML_REGISTER) As SHUNTING_ITEM_DEFINE
    Dim ret As SHUNTING_ITEM_DEFINE
    Dim ccid As CML_COMPILED_IDENTIFIER_DEFINITION
    
    ReDim ret.tokens(1)
    ret.tokens(1).t = TOKEN_TYPE_REGISTER
    
    COMPILE_IDENTIFIER r, ccid, register
    
    Select Case True
        Case op = "="
        Case ccid.pointer
            ret.extra = CALC_REGISTER_WITH_REGISTER(register, EDI, op)
        Case ccid.register <> NONE
            CALC_REGISTER_WITH_REGISTER_VALUE register, EDI, op, ccid.v.datatype
        Case ccid.register = NONE
            CALC_REGISTER_WITH_VARIABLE register, ccid.v, op
    End Select
    
    COMPILE_REGISTER_WITH_IDENTIFIER = ret
End Function

Private Function COMPILE_NUMBER_WITH_REGISTER(l As SHUNTING_ITEM_DEFINE, r As SHUNTING_ITEM_DEFINE, op As String, register As CML_REGISTER) As SHUNTING_ITEM_DEFINE
    Dim ret As SHUNTING_ITEM_DEFINE
    
    ReDim ret.tokens(1)
    ret.tokens(1).t = TOKEN_TYPE_REGISTER
    ret.extra = r.extra
    ret.extra = CALC_REGISTER_WITH_NUMBER(ret.extra, CLng(GetTokenS(l.tokens(1))), op, True)
    
    COMPILE_NUMBER_WITH_REGISTER = ret
End Function

Private Function COMPILE_REGISTER_WITH_REGISTER(l As SHUNTING_ITEM_DEFINE, r As SHUNTING_ITEM_DEFINE, op As String, ByVal register As CML_REGISTER) As SHUNTING_ITEM_DEFINE
    Dim ret As SHUNTING_ITEM_DEFINE
    
    ReDim ret.tokens(1)
    
    ret.tokens(1).t = TOKEN_TYPE_REGISTER
    ret.extra = CALC_REGISTER_WITH_REGISTER(l.extra, r.extra)
    
    COMPILE_REGISTER_WITH_REGISTER = ret
End Function

Private Function COMPILE_IDENTIFIER_WITH_REGISTER(l As SHUNTING_ITEM_DEFINE, r As SHUNTING_ITEM_DEFINE, op As String, ByVal register As CML_REGISTER) As SHUNTING_ITEM_DEFINE
    Dim ret As SHUNTING_ITEM_DEFINE
    Dim ccid As CML_COMPILED_IDENTIFIER_DEFINITION
    
    ReDim ret.tokens(1)
    ret.tokens(1).t = TOKEN_TYPE_REGISTER
    
    COMPILE_IDENTIFIER l, ccid, register
    
    If op = "=" Then
        MOVE_REGISTER_TO_POINTER_VARIABLE register, ccid
        ret = l
    ElseIf ccid.register = 0 Then
        'CALC_IDENTIFIER_WITH_REGISTER
        'XCHG register, EBX
        If op = "/" Then
            CALC_REGISTER_WITH_VARIABLE register, ccid.v, op, True
        Else
            CALC_REGISTER_WITH_VARIABLE register, ccid.v, op
        End If
    ElseIf ccid.register Then
        Inc register
        MOV_VARIABLE_TO_REGISTER ccid.v, register
        CALC_REGISTER_WITH_REGISTER register, r.extra, op
        Dec register
        ret.extra = r.extra
    End If
    
    COMPILE_IDENTIFIER_WITH_REGISTER = ret
End Function

Private Function COMPILE_IDENTIFIER_WITH_IDENTIFIER(l As SHUNTING_ITEM_DEFINE, r As SHUNTING_ITEM_DEFINE, op As String, ByVal register As CML_REGISTER) As SHUNTING_ITEM_DEFINE
    Dim ret As SHUNTING_ITEM_DEFINE
    Dim ccidr As CML_COMPILED_IDENTIFIER_DEFINITION
    Dim ccidl As CML_COMPILED_IDENTIFIER_DEFINITION
    
    ReDim ret.tokens(1)
    ret.tokens(1).t = TOKEN_TYPE_REGISTER
    
    COMPILE_IDENTIFIER r, ccidr, register, , EDI
    COMPILE_IDENTIFIER l, ccidl, register, , register, True
    
    ret.extra = register
    
    'Clipboard.SetText ccidl.code
    
    Select Case True
        Case op = "="
            COMPILE_IDENTIFIER l, ccidl, register, , register
            
        
        Case ccidl.pointer And ccidl.register = NONE And ccidr.register = NONE And ccidr.pointer = 0
            MOV_INPUT_TO_REGISTER ccidl.v.uname, register
            CALC_REGISTER_WITH_VARIABLE register, ccidr.v, op
        
        Case ccidl.pointer And ccidl.register = NONE And ccidr.register = NONE And ccidr.pointer
            MOV_INPUT_TO_REGISTER ccidl.v.uname, register
            CALC_REGISTER_WITH_INPUT register, ccidr.text, op
            
        Case ccidl.pointer And ccidl.register = NONE And ccidr.register <> NONE And ccidr.pointer = 0
            MOV_INPUT_TO_REGISTER ccidl.v.uname, register
            CALC_REGISTER_WITH_REGISTER_VALUE register, EDI, op, ccidr.v.datatype
        
        Case ccidl.pointer And ccidl.register = NONE And ccidr.register <> NONE And ccidr.pointer
            MOV_INPUT_TO_REGISTER ccidl.v.uname, register
            CALC_REGISTER_WITH_REGISTER register, EDI, op
        
        
        Case ccidl.pointer And ccidl.register <> NONE And ccidr.register = NONE And ccidr.pointer = 0
            ' Lo recompilamos al registro correcto y esta vez si se escribe en el archivo de salida
            COMPILE_IDENTIFIER l, ccidl, register, , register
            CALC_REGISTER_WITH_VARIABLE register, ccidr.v, op
            
        Case ccidl.pointer And ccidl.register <> NONE And ccidr.register = NONE And ccidr.pointer
            COMPILE_IDENTIFIER l, ccidl, register, , register
            CALC_REGISTER_WITH_INPUT register, ccidr.text, op
            
        Case ccidl.pointer And ccidl.register <> NONE And ccidr.register <> NONE And ccidr.pointer = 0
            COMPILE_IDENTIFIER l, ccidl, register, , register
            CALC_REGISTER_WITH_REGISTER_VALUE register, EDI, op, ccidr.v.datatype
            
        Case ccidl.pointer And ccidl.register <> NONE And ccidr.register <> NONE And ccidr.pointer
            COMPILE_IDENTIFIER l, ccidl, register, , register
            CALC_REGISTER_WITH_REGISTER register, EDI, op
        
        ' La variable de la izquierda no es un puntero ni un registro
        Case ccidl.pointer = 0 And ccidl.register = NONE And ccidr.register = NONE And ccidr.pointer = 0
            MOV_INPUT_TO_REGISTER ccidl.v.uname, register
            CALC_REGISTER_WITH_VARIABLE register, ccidr.v, op
        
        Case ccidl.pointer = 0 And ccidl.register = NONE And ccidr.register = NONE And ccidr.pointer
            MOV_VARIABLE_TO_REGISTER ccidl.v, register
            CALC_REGISTER_WITH_INPUT register, ccidr.text, op
            
        Case ccidl.pointer = 0 And ccidl.register = NONE And ccidr.register <> NONE And ccidr.pointer = 0
            MOV_INPUT_TO_REGISTER ccidl.v.uname, register
            CALC_REGISTER_WITH_REGISTER_VALUE register, EDI, op, ccidr.v.datatype
        
        Case ccidl.pointer = 0 And ccidl.register = NONE And ccidr.register <> NONE And ccidr.pointer
            MOV_INPUT_TO_REGISTER ccidl.v.uname, register
            CALC_REGISTER_WITH_REGISTER register, EDI, op
        ' La variable de la izquierda no es un puntero
        Case ccidl.pointer = 0 And ccidl.register <> NONE And ccidr.register = NONE And ccidr.pointer = 0
            
            COMPILE_IDENTIFIER l, ccidl, register, , register
            CALC_REGISTER_WITH_VARIABLE register, ccidr.v, op
            
        Case ccidl.pointer = 0 And ccidl.register <> NONE And ccidr.register = NONE And ccidr.pointer
            COMPILE_IDENTIFIER l, ccidl, register, , register
            CALC_REGISTER_WITH_INPUT register, ccidr.text, op
            
        Case ccidl.pointer = 0 And ccidl.register <> NONE And ccidr.register <> NONE And ccidr.pointer = 0
            COMPILE_IDENTIFIER l, ccidl, register, , register
            CALC_REGISTER_WITH_REGISTER_VALUE register, EDI, op, ccidr.v.datatype
            
        Case ccidl.pointer = 0 And ccidl.register <> NONE And ccidr.register <> NONE And ccidr.pointer
            COMPILE_IDENTIFIER l, ccidl, register, , register
            CALC_REGISTER_WITH_REGISTER register, EDI, op
            
            
            
            
            
            
            
            
        Case ccidl.register = NONE And ccidr.register <> NONE
            MOV_VARIABLE_TO_REGISTER ccidl.v, register
            CALC_REGISTER_WITH_REGISTER register, ccidr.register, op
            
        Case ccidl.register = NONE And ccidr.register = NONE
            MOV_VARIABLE_TO_REGISTER ccidl.v, register
            CALC_REGISTER_WITH_VARIABLE register, ccidr.v, op
            
        Case ccidl.register <> NONE And ccidr.register <> NONE
            COMPILE_IDENTIFIER l, ccidl, register, , register
            CALC_REGISTER_WITH_REGISTER_VALUE register, ccidr.register, op, ccidr.v.datatype
        
        Case ccidl.register <> NONE And ccidr.register = NONE
            COMPILE_IDENTIFIER l, ccidl, register, , register
            CALC_REGISTER_WITH_INPUT register, ccidr.text, op
        
    End Select

    COMPILE_IDENTIFIER_WITH_IDENTIFIER = ret
End Function

Private Function COMPILE_IDENTIFIER_WITH_NUMBER(l As SHUNTING_ITEM_DEFINE, r As SHUNTING_ITEM_DEFINE, op As String, register As CML_REGISTER) As SHUNTING_ITEM_DEFINE
    Dim ret As SHUNTING_ITEM_DEFINE
    
    ReDim ret.tokens(1)
    
    ret.tokens(1).t = TOKEN_TYPE_REGISTER
    ret.extra = CALC_IDENTIFIER_WITH_NUMBER(l, CLng(r.tokens(1).s), op, register, False)
    
    COMPILE_IDENTIFIER_WITH_NUMBER = ret
End Function

Private Function COMPILE_NUMBER_WITH_IDENTIFIER(l As SHUNTING_ITEM_DEFINE, r As SHUNTING_ITEM_DEFINE, op As String, register As CML_REGISTER) As SHUNTING_ITEM_DEFINE
    Dim ret As SHUNTING_ITEM_DEFINE
    
    ReDim ret.tokens(1)
    ret.tokens(1).t = TOKEN_TYPE_REGISTER
    ret.extra = CALC_IDENTIFIER_WITH_NUMBER(r, CLng(l.tokens(1).s), op, register, True)
    
    COMPILE_NUMBER_WITH_IDENTIFIER = ret
End Function

Private Sub MOV_VARIABLE_TO_REGISTER(v As CML_VARIABLE_DEFINITION, ByVal register As CML_REGISTER)
    Dim r As String
    r = REGISTER_TO_TEXT(register)
    Select Case v.datatype
        Case var_type_byte
            WLine "movzx [0],BYTE[[1]]", r, v.uname
        Case var_type_dword
            WLine "mov [0],DWORD[[1]]", r, v.uname
        Case var_type_float
            WLine "mov [0],DWORD[[1]]", r, v.uname
        Case var_type_qword
            'WLine "movzx [0],BYTE[[1]]", r, v.uname
        Case var_type_word
            WLine "movzx [0],WORD[[1]]", r, v.uname
        Case var_type_compose
            WLine "lea [0],[1]", r, v.uname
    End Select
End Sub

Private Sub MOV_REGISTER_TO_REGISTER(ByVal dst As CML_REGISTER, ByVal src As CML_REGISTER, Optional ByVal datatype As CML_TYPE_VARIABLE = var_type_dword)
    Dim rdst As String
    Dim rsrc As String
    
    rdst = REGISTER_TO_TEXT(dst)
    rsrc = REGISTER_TO_TEXT(src)
    
    Select Case datatype
        Case var_type_byte
            WLine "movzx [0],BYTE[[1]]", rdst, rsrc
        Case var_type_dword
            WLine "mov [0],DWORD[[1]]", rdst, rsrc
        Case var_type_float
            WLine "mov [0],DWORD[[1]]", rdst, rsrc
        Case var_type_qword
            'WLine "movzx [0],BYTE[[1]]", r, v.uname
        Case var_type_word
            WLine "movzx [0],WORD[[1]]", rdst, rsrc
        Case var_type_compose
            WLine "mov [0],[1]", rdst, rsrc
    End Select
End Sub

Private Sub MOV_REGISTER_TO_REGISTER_EX(ByVal dst As CML_REGISTER, ByVal src As CML_REGISTER, v As CML_VARIABLE_DEFINITION)
    Dim rdst As String
    Dim rsrc As String
    rdst = REGISTER_TO_TEXT(dst)
    rsrc = REGISTER_TO_TEXT(src)
    
    If v.pointer Then
        If dst = src Then
            Exit Sub
        End If
        
        WLine "mov [0],[1]", rdst, rsrc
    Else
        Select Case v.datatype
            Case var_type_byte
                WLine "movzx [0],BYTE[[1]]", rdst, rsrc
            Case var_type_dword
                WLine "mov [0],DWORD[[1]]", rdst, rsrc
            Case var_type_float
                WLine "mov [0],DWORD[[1]]", rdst, rsrc
            Case var_type_qword
                'WLine "movzx [0],BYTE[[1]]", r, v.uname
            Case var_type_word
                WLine "movzx [0],WORD[[1]]", rdst, rsrc
            Case var_type_compose
                WLine "mov [0],[1]", rdst, rsrc
        End Select
    End If
End Sub

Private Sub MOV_REGISTER_TO_REGISTER2(ByVal dst As CML_REGISTER, ByVal src As CML_REGISTER)
    Dim r1 As String, r2 As String
    If dst = src Then Exit Sub
    
    r1 = REGISTER_TO_TEXT(dst)
    r2 = REGISTER_TO_TEXT(src)
    
    WLine "mov [0],[1]", r1, r2
End Sub

Private Function REGISTER_TO_TEXT(ByVal register As CML_REGISTER) As String
    Select Case register
        Case eax
            REGISTER_TO_TEXT = "eax"
        Case EBX
            REGISTER_TO_TEXT = "ebx"
        Case ECX
            REGISTER_TO_TEXT = "ecx"
        Case EDX
            REGISTER_TO_TEXT = "edx"
        Case EDI
            REGISTER_TO_TEXT = "edi"
        Case ESI
            REGISTER_TO_TEXT = "esi"
    End Select
End Function

Private Function BYTE_OF_REGISTER(ByVal register As CML_REGISTER) As String
    Select Case register
        Case eax
            BYTE_OF_REGISTER = "al"
        Case EBX
            BYTE_OF_REGISTER = "bl"
        Case ECX
            BYTE_OF_REGISTER = "cl"
        Case EDX
            BYTE_OF_REGISTER = "dl"
        Case EDI
            Assert False, "El registro EDI no posee división de 8 bits."
        Case ESI
            Assert False, "El registro ESI no posee división de 8 bits."
    End Select
End Function

Private Function WORD_OF_REGISTER(ByVal register As CML_REGISTER) As String
    Select Case register
        Case eax
            REGISTER_TO_TEXT = "ax"
        Case EBX
            REGISTER_TO_TEXT = "bx"
        Case ECX
            REGISTER_TO_TEXT = "cx"
        Case EDX
            REGISTER_TO_TEXT = "dx"
        Case EDI
            Assert False, "El registro EDI no posee división de 16 bits."
        Case ESI
            Assert False, "El registro ESI no posee división de 16 bits."
    End Select
End Function

Private Function CALC_REGISTER_WITH_REGISTER_VALUE(ByVal r1 As CML_REGISTER, ByVal r2 As CML_REGISTER, op As String, datatype As CML_TYPE_VARIABLE) As CML_REGISTER
    Dim rl As String
    Dim rr As String
    Dim bsl As String
    Dim ds As String
    
    rl = REGISTER_TO_TEXT(r1)
    rr = REGISTER_TO_TEXT(r2)
    bsl = BYTE_OF_REGISTER(r1)
    CALC_REGISTER_WITH_REGISTER_VALUE = r1
    ds = DATATYPE_TO_I32(datatype)
    
    Select Case op
        Case "*", "/"
            If r1 <> eax Then XCHG eax, r1
            
            Select Case datatype
                Case var_type_byte, var_type_word
                    WLine "movzx ecx,[0][[1]]", ds, rr
                    WLine "[0] ecx", IIf(op = "*", "mul", "div")
                Case var_type_dword, var_type_compose
                    WLine "[0] DWORD[[1]]", IIf(op = "*", "mul", "div"), rr
                Case var_type_float
                    
                Case var_type_qword
                    
            End Select
            
            If r1 <> eax Then XCHG eax, r1
        Case "<<"
            Select Case datatype
                Case var_type_byte, var_type_word
                    WLine "movzx ecx,[0][[1]]", ds, rr
                    WLine "[0] ecx", IIf(op = "*", "mul", "div")
                Case var_type_dword, var_type_compose
                    WLine "mov ecx,DWORD[[0]]", rr
                Case var_type_float
                    
                Case var_type_qword
                    
            End Select
            WLine "shl [0],cl", rl
        Case ">>"
            Select Case datatype
                Case var_type_byte, var_type_word
                    WLine "movzx ecx,[0][[1]]", ds, rr
                    WLine "[0] ecx", IIf(op = "*", "mul", "div")
                Case var_type_dword, var_type_compose
                    WLine "mov ecx,DWORD[[0]]", rr
                Case var_type_float
                    
                Case var_type_qword
                    
            End Select
            WLine "shr [0],cl", rl
        Case "%"
            If r1 <> eax Then XCHG eax, r1
            CALC_REGISTER_WITH_REGISTER_VALUE = EDX
            
            WLine "[0] [1]", IIf(op = "/", "div", "mul"), rr
            
            If r1 <> eax Then XCHG eax, r1
            WLine "; El resultado está en EDX"
        Case "^"
            'OPERATOR_TO_PRECEDENCE = 9.4
        Case "+"
            WLine "add [0],[1][[2]]", rl, ds, rr
        Case "-"
            WLine "sub [0],[1][[2]]", rl, ds, rr
        Case "|"
            WLine "or [0],[1][[2]]", rl, ds, rr
        Case "Xor"
            WLine "xor [0],[1][[2]]", rl, ds, rr
        Case "&"
            WLine "and [0],[1][[2]]", rl, ds, rr
        Case "==", "<>", "!="
             Select Case datatype
                Case var_type_byte, var_type_word
                    WLine "movzx ecx,[0][[1]]", ds, rr
                    WLine "test [0],ecx", rl
                Case var_type_dword, var_type_compose
                    WLine "test [0],[1][[2]]", rl, ds, rr
                Case var_type_float
                    
                Case var_type_qword
                    
            End Select
            
            Select Case op
                Case "=="
                    WLine "sete [0]", bsl
                Case Else
                    WLine "setne [0]", bsl
            End Select
            WLine "movzx [0],[1]", rl, bsl
            
        Case "<=", "<", ">", ">="
             Select Case datatype
                Case var_type_byte, var_type_word
                    WLine "movzx ecx,[0][[1]]", ds, rr
                    WLine "cmp [0],ecx", rl
                Case var_type_dword, var_type_compose
                    WLine "cmp [0],[1][[2]]", rl, ds, rr
                Case var_type_float
                    
                Case var_type_qword
                    
            End Select
            
            Select Case op
                Case "<="
                    WLine "setle [0]", bsl
                Case "<"
                    WLine "setb [0]", bsl
                Case ">"
                    WLine "seta [0]", bsl
                Case ">="
                    WLine "setae [0]", bsl
            End Select
            WLine "movzx [0],[1]", rl, bsl
        Case "="
            Assert False, "Imposible asignar un valor a un número."
    End Select

End Function

Private Function CALC_REGISTER_WITH_REGISTER(ByVal r1 As CML_REGISTER, ByVal r2 As CML_REGISTER, op As String) As CML_REGISTER
    Dim rl As String
    Dim rr As String
    Dim bsl As String
    
    
    rl = REGISTER_TO_TEXT(r1)
    rr = REGISTER_TO_TEXT(r2)
    bsl = BYTE_OF_REGISTER(r1)
    CALC_REGISTER_WITH_REGISTER = r1
    
    Select Case op
        Case "*", "/"
            If r1 <> eax Then XCHG eax, r1
            
            WLine "[0] [1]", IIf(op = "/", "div", "mul"), rr
            
            If r1 <> eax Then XCHG eax, r1
        Case "<<"
            MOV_REGISTER_TO_REGISTER ECX, rr
            WLine "shl [0],cl", rl
        Case ">>"
            MOV_REGISTER_TO_REGISTER ECX, rr
            WLine "shr [0],cl", rl
        Case "%"
            If r1 <> eax Then XCHG eax, r1
            CALC_REGISTER_WITH_REGISTER = EDX
            
            WLine "[0] [1]", IIf(op = "/", "div", "mul"), rr
            
            If r1 <> eax Then XCHG eax, r1
            WLine "; El resultado está en EDX"
        Case "^"
            'OPERATOR_TO_PRECEDENCE = 9.4
        Case "+"
            WLine "add [0],[1]", rl, rr
        Case "-"
            WLine "sub [0],[1]", rl, rr
        Case "|"
            WLine "or [0],[1]", rl, rr
        Case "Xor"
            WLine "xor [0],[1]", rl, rr
        Case "&"
            WLine "and [0],[1]", rl, rr
        Case "==", "<>", "!="
            WLine "test [0],[1]", rl, rr
            
            Select Case op
                Case "=="
                    WLine "sete [0]", bsl
                Case Else
                    WLine "setne [0]", bsl
            End Select
            WLine "movzx [0],[1]", rl, bsl
            
        Case "<=", "<", ">", ">="
            WLine "cmp [0],[1]", rl, rr
            
            Select Case op
                Case "<="
                    WLine "setle [0]", bsl
                Case "<"
                    WLine "setb [0]", bsl
                Case ">"
                    WLine "seta [0]", bsl
                Case ">="
                    WLine "setae [0]", bsl
            End Select
            WLine "movzx [0],[1]", rl, bsl
        Case "="
            Assert False, "Imposible asignar un valor a un número."
    End Select

End Function

Private Function CALC_IDENTIFIER_WITH_NUMBER(l As SHUNTING_ITEM_DEFINE, ByVal n As Long, op As String, register As CML_REGISTER, Optional ByVal invert As Boolean) As CML_REGISTER
    Dim fr As String
    Dim br As String
    Dim cexpr As CML_COMPILED_IDENTIFIER_DEFINITION
    
    fr = REGISTER_TO_TEXT(register)
    br = BYTE_OF_REGISTER(register)
    CALC_IDENTIFIER_WITH_NUMBER = register
    COMPILE_IDENTIFIER l, cexpr, register, , register
    
    If cexpr.pointer And cexpr.register Then
        CALC_IDENTIFIER_WITH_NUMBER = CALC_REGISTER_WITH_NUMBER(cexpr.register, n, op, invert)
    ElseIf cexpr.pointer And cexpr.register = NONE Then
        CALC_IDENTIFIER_WITH_NUMBER = CALC_REGISTER_WITH_INPUT(register, cexpr.text, op, invert)
    ElseIf Not cexpr.pointer And cexpr.register Then
        CALC_IDENTIFIER_WITH_NUMBER = CALC_REGISTER_WITH_NUMBER(cexpr.register, n, op, invert)
    ElseIf Not cexpr.pointer And cexpr.register = NONE Then
        Dim v As CML_VARIABLE_DEFINITION
        v.uname = cexpr.text
        v.datatype = cexpr.v.datatype
        CALC_IDENTIFIER_WITH_NUMBER = CALC_REGISTER_WITH_VARIABLE(register, v, op, invert)
    End If
End Function

Private Function CALC_REGISTER_WITH_INPUT(ByVal register As CML_REGISTER, s As String, op As String, Optional ByVal invert As Boolean) As CML_REGISTER
    Dim fr As String
    Dim br As String
    
    fr = REGISTER_TO_TEXT(register)
    br = BYTE_OF_REGISTER(register)
    CALC_REGISTER_WITH_INPUT = register
    
    Select Case op
        Case "*"
            If register = eax Then
                WLine "mov ecx,[0]", s
                WLine "mul ecx"
            Else
                'WLine "mov ecx,", s
                WLine "imul [0],[1]", fr, s
            End If
        Case "/", "%"
            WLine "mov ecx,[0]", s
            
            If invert Then XCHG register, ECX
            
            If register = eax Then
                WLine "div ecx"
            Else
                XCHG eax, register
                WLine "div ecx"
                XCHG eax, register
            End If
            
            If invert Then XCHG register, ECX
            
            If op = "%" Then CALC_REGISTER_WITH_INPUT = EDX
            
            'If register = EAX Then
            '    WLine "xchg eax,ecx"
            '    WLine "div ecx"
            'Else
            '    WLine "xchg eax,ecx"
            '    WLine "div [0]", fr
            '    WLine "xchg eax,ecx"
            'End If
        Case "<<"
            WLine "mov ecx,[0]", s
            If invert Then XCHG register, ECX
            
            WLine "shl [0],cl", fr
        Case ">>"
            WLine "mov ecx,[0]", s
            If invert Then XCHG register, ECX
            
            WLine "shr [0],cl", fr
        Case "^"
            'OPERATOR_TO_PRECEDENCE = 9.4
        Case "+"
            WLine "add [0],[1]", fr, s
        Case "-"
            WLine "sub [0],[1]", fr, s
        Case "|"
            WLine "or [0],[1]", fr, s
        Case "Xor"
            WLine "xor [0],[1]", fr, s
        Case "&"
            WLine "and [0],[1]", fr, s
        Case "==", "<>", "!="
            WLine "test [0],[1]", fr, s
            Select Case op
                Case "=="
                    WLine "sete [0]", br
                Case Else
                    WLine "setne [0]", br
            End Select
            WLine "movzx [0],[1]", fr, br
        Case "<=", "<", ">", ">="
            WLine "cmp [0],[1]", s, fr
            Select Case op
                Case "<="
                    WLine "setle [0]", br
                Case "<"
                    WLine "setb [0]", br
                Case ">"
                    WLine "seta [0]", br
                Case ">="
                    WLine "setae [0]", br
            End Select
            WLine "movzx [0],[1]", fr, br
        Case "="
            Assert False, "Imposible asignar un valor a un número."
    End Select
End Function

Private Function CALC_REGISTER_WITH_NUMBER(ByVal register As CML_REGISTER, ByVal n As Long, op As String, Optional ByVal invert As Boolean, Optional ByVal datatype As CML_TYPE_VARIABLE = var_type_dword) As CML_REGISTER
    Dim fr As String
    Dim br As String
    
    fr = REGISTER_TO_TEXT(register)
    br = BYTE_OF_REGISTER(register)
    CALC_REGISTER_WITH_NUMBER = register
    
    Select Case op
        Case "*"
            If register = eax Then
                MOV_NUMBER_TO_REGISTER ECX, n
                WLine "mul ecx"
            Else
                WLine "imul [0],[1]", fr, n
            End If
        Case "/", "%"
            MOV_NUMBER_TO_REGISTER EDX, 0
            MOV_NUMBER_TO_REGISTER ECX, n
            
            If invert Then XCHG register, ECX
            
            If register = eax Then
                WLine "div ecx"
            Else
                XCHG eax, register
                WLine "div ecx"
                XCHG eax, register
            End If
            
            If invert Then XCHG register, ECX
            
            If op = "%" Then CALC_REGISTER_WITH_NUMBER = EDX
            
            'If register = EAX Then
            '    WLine "xchg eax,ecx"
            '    WLine "div ecx"
            'Else
            '    WLine "xchg eax,ecx"
            '    WLine "div [0]", fr
            '    WLine "xchg eax,ecx"
            'End If
        Case "<<"
            MOV_NUMBER_TO_REGISTER ECX, n
            If invert Then XCHG register, ECX
            
            WLine "shl [0],cl", fr
        Case ">>"
            MOV_NUMBER_TO_REGISTER ECX, n
            If invert Then XCHG register, ECX
            
            WLine "shr [0],cl", fr
        Case "^"
            'OPERATOR_TO_PRECEDENCE = 9.4
        Case "+"
            ADD_NUMBER_TO_REGISTER register, n
        Case "-"
            SUB_NUMBER_TO_REGISTER register, n
        Case "|"
            WLine "or [0],[1]", fr, n
        Case "Xor"
            WLine "xor [0],[1]", fr, n
        Case "&"
            WLine "and [0],[1]", fr, n
        Case "==", "<>", "!="
            WLine "test [0],[1]", fr, n
            Select Case op
                Case "=="
                    WLine "sete [0]", br
                Case Else
                    WLine "setne [0]", br
            End Select
            WLine "movzx [0],[1]", fr, br
        Case "<=", "<", ">", ">="
            WLine "cmp [0],[1]", n, fr
            Select Case op
                Case "<="
                    WLine "setle [0]", br
                Case "<"
                    WLine "setb [0]", br
                Case ">"
                    WLine "seta [0]", br
                Case ">="
                    WLine "setae [0]", br
            End Select
            WLine "movzx [0],[1]", fr, br
        Case "="
            Select Case datatype
                Case var_type_byte
                    WLine "mov BYTE[[0]],[1]", fr, n
                Case var_type_word
                    WLine "mov WORD[[0]],[1]", fr, n
                Case var_type_dword, var_type_compose
                    WLine "mov DWORD[[0]],[1]", fr, n
                Case var_type_float
                    'WLine "mov BYTE[[0]],[1]", fr, n
                Case var_type_qword
                    'WLine "mov BYTE[[0]],[1]", fr, n
            End Select
            'Assert False, "Imposible asignar un valor a un número."
    End Select
End Function

Private Function CALC_VARIABLE_WITH_VARIABLE(l As CML_VARIABLE_DEFINITION, r As CML_VARIABLE_DEFINITION, ByRef register As CML_REGISTER, op As String, Optional ByVal invert As Boolean) As CML_REGISTER

    MOV_VARIABLE_TO_REGISTER r, register
    
    CALC_VARIABLE_WITH_VARIABLE = CALC_REGISTER_WITH_VARIABLE(register, l, op, invert)
    
End Function

Private Function CALC_REGISTER_WITH_VARIABLE(register As CML_REGISTER, v As CML_VARIABLE_DEFINITION, op As String, Optional ByVal invert As Boolean) As CML_REGISTER
    Dim r As String
    Dim ds As String
    Dim br As String
    
    r = REGISTER_TO_TEXT(register)
    br = BYTE_OF_REGISTER(register)
    ds = DATATYPE_TO_I32(v.datatype)
    
    Select Case op
        Case "*", "/"
            XCHG eax, register
            
            If IS_I32_4BYTES(v.datatype) Then
                WLine "[0] DWORD[[1]]", IIf(op = "/", "div", "mul"), v.uname
            Else
                MOVE_VARIABLE_TO_REGISTER v, ECX
                WLine "[0] ecx", IIf(op = "/", "div", "mul")
            End If
            
            XCHG eax, register
        Case "<<"
            MOVE_VARIABLE_TO_REGISTER v, ECX
            WLine "shl [0],cl", r
        Case ">>"
            MOVE_VARIABLE_TO_REGISTER v, ECX
            WLine "shr [0],cl", r
        Case "%"
            MOV_NUMBER_TO_REGISTER EDX, 0
            MOVE_VARIABLE_TO_REGISTER v, ECX
            register = EDX
            
            If register <> eax Then WLine "xchg eax,[0]", r
            
            If IS_I32_4BYTES(v.datatype) Then
                WLine "div DWORD[[0]]", v.uname
            Else
                MOVE_VARIABLE_TO_REGISTER v, ECX
                WLine "div ecx"
            End If
            
            If register <> eax Then WLine "xchg eax,[0]", r
            WLine "; El resultado está en EDX"
        Case "^"
            'OPERATOR_TO_PRECEDENCE = 9.4
        Case "+"
            WLine "add [0],[1][[2]]", r, ds, v.uname
        Case "-"
            WLine "sub [0],[1][[2]]", r, ds, v.uname
        Case "|"
            WLine "or [0],[1][[2]]", r, ds, v.uname
        Case "Xor"
            WLine "xor [0],[1][[2]]", r, ds, v.uname
        Case "&"
            WLine "and [0],[1][[2]]", r, ds, v.uname
        Case "==", "<>", "!="
        
            If IS_I32_4BYTES(v.datatype) Then
                WLine "test [0],DWORD[[1]]", r, v.uname
            Else
                MOVE_VARIABLE_TO_REGISTER v, ECX
                WLine "test [0],ecx", r
            End If
            
            Select Case op
                Case "=="
                    WLine "sete [0]", br
                Case Else
                    WLine "setne [0]", br
            End Select
            WLine "movzx [0],[1]", r, br
            
        Case "<=", "<", ">", ">="
        
            If IS_I32_4BYTES(v.datatype) Then
                WLine "cmp [0],DWORD[[1]]", r, v.uname
            Else
                MOVE_VARIABLE_TO_REGISTER v, ECX
                WLine "cmp [0],ecx", r
            End If

            Select Case op
                Case "<="
                    WLine "setle [0]", br
                Case "<"
                    WLine "setb [0]", br
                Case ">"
                    WLine "seta [0]", br
                Case ">="
                    WLine "setae [0]", br
            End Select
            WLine "movzx [0],[1]", r, br
        Case "="
            Assert False, "Imposible asignar un valor a un número."
    End Select

End Function

Private Sub SUB_NUMBER_TO_REGISTER(ByVal register As CML_REGISTER, ByVal n As Long)
    Dim r As String
    r = REGISTER_TO_TEXT(register)
    
    Select Case n
        Case -2
            WLine "inc [0]", r
            WLine "inc [0]", r
        Case -1
            WLine "inc [0]", r
        Case 0
            ' DoNothing
        Case 1
            WLine "dec [0]", r
        Case 2
            WLine "dec [0]", r
            WLine "dec [0]", r
        Case Else
            WLine "sub [0],[1]", r, n
    End Select
End Sub

Private Sub ADD_NUMBER_TO_REGISTER(ByVal register As CML_REGISTER, ByVal n As Long)
    Dim r As String
    r = REGISTER_TO_TEXT(register)
    
    Select Case n
        Case -2
            WLine "dec [0]", r
            WLine "dec [0]", r
        Case -1
            WLine "dec [0]", r
        Case 0
            ' DoNothing
        Case 1
            WLine "inc [0]", r
        Case 2
            WLine "inc [0]", r
            WLine "inc [0]", r
        Case Else
            WLine "add [0],[1]", r, n
    End Select
End Sub

Private Sub MOV_INPUT_TO_REGISTER(s As String, ByVal register As CML_REGISTER)
    WLine "mov [0],[1]", REGISTER_TO_TEXT(register), s
End Sub

Private Sub MOV_NUMBER_TO_REGISTER(ByVal register As CML_REGISTER, ByVal n As Long, Optional ByVal datatype As CML_TYPE_VARIABLE)
    Dim r As String
    r = REGISTER_TO_TEXT(register)
    If datatype = 0 Then
        Select Case n
            Case -2
                WLine "xor [0],[0]", r
                WLine "dec [0]", r
                WLine "dec [0]", r
            Case -1
                WLine "xor [0],[0]", r
                WLine "dec [0]", r
            Case 0
                WLine "xor [0],[0]", r
            Case 1
                WLine "xor [0],[0]", r
                WLine "inc [0]", r
            Case 2
                WLine "xor [0],[0]", r
                WLine "inc [0]", r
                WLine "inc [0]", r
            Case Else
                WLine "mov [0],[1]", r, n
        End Select
    Else
        Select Case datatype
            Case var_type_byte
                WLine "mov BYTE[[0]],[1]", r, n
            Case var_type_word
                WLine "mov WORD[[0]],[1]", r, n
            Case var_type_compose
                Assert False, "Imposible asignar un valor a un tipo de dato compuesto."
            Case var_type_float, var_type_dword
                WLine "mov DWORD[[0]],[1]", r, n
            Case var_type_qword
                ' TODO: COMPLETAR
        End Select
    End If
End Sub


Private Sub MOVE_VARIABLE_TO_REGISTER(v As CML_VARIABLE_DEFINITION, ByVal register As CML_REGISTER, Optional ByVal pointer As Long)
    Dim r As String
    Dim ds As String
    
    r = REGISTER_TO_TEXT(register)
    
    If Not pointer Then
        ds = DATATYPE_TO_I32(v.datatype)
        
        Select Case v.datatype
            Case var_type_byte, var_type_word
                WLine "movzx [0],[1][[2]]", r, ds, v.uname
            Case var_type_float, var_type_dword, var_type_compose
                WLine "mov [0],[1][[2]]", r, ds, v.uname
            Case var_type_qword
                ' TODO: COMPLETAR
        End Select
    Else
        WLine "mov [0],DWORD[[1]]", r, v.uname
        Dec pointer
        While pointer
            WLine "mov [0],DWORD[[0]]", r
            Dec pointer
        Wend
    End If
End Sub

Private Function SIZED_REGISTER(ByVal register As CML_REGISTER, ByVal datatype As CML_TYPE_VARIABLE)
    Select Case datatype
        Case var_type_byte
            DATATYPE_TO_I32 = "BYTE"
        Case var_type_word
            DATATYPE_TO_I32 = "WORD"
        Case var_type_compose
            DATATYPE_TO_I32 = "DWORD"
        Case var_type_float, var_type_dword
            DATATYPE_TO_I32 = "DWORD"
        Case var_type_qword
            DATATYPE_TO_I32 = "QWORD"
    End Select
End Function

Private Sub MOVE_REGISTER_TO_POINTER_VARIABLE(ByVal register As CML_REGISTER, ciid As CML_COMPILED_IDENTIFIER_DEFINITION)
    Select Case ciid.v.datatype
        Case var_type_byte
            WLine "mov BYTE[[0]],[1]", REGISTER_TO_TEXT(ciid.register), BYTE_OF_REGISTER(register)
        Case var_type_word
            WLine "mov WORD[[0]],[1]", REGISTER_TO_TEXT(ciid.register), WORD_OF_REGISTER(register)
        Case var_type_dword
            WLine "mov DWORD[[0]],[1]", REGISTER_TO_TEXT(ciid.register), REGISTER_TO_TEXT(register)
        'Case var_type_float,var_type_compose
        '    DATATYPE_TO_I32 = "DWORD"
        'Case var_type_qword
        '    DATATYPE_TO_I32 = "QWORD"
        'TODO: COMPLETAR
    End Select
End Sub

Private Function DATATYPE_TO_I32(ByVal datatype As CML_TYPE_VARIABLE) As String
    Select Case datatype
        Case var_type_byte
            DATATYPE_TO_I32 = "BYTE"
        Case var_type_word
            DATATYPE_TO_I32 = "WORD"
        Case var_type_compose
            DATATYPE_TO_I32 = "DWORD"
        Case var_type_float, var_type_dword
            DATATYPE_TO_I32 = "DWORD"
        Case var_type_qword
            DATATYPE_TO_I32 = "QWORD"
    End Select
End Function

Private Sub XCHG(ByVal r1 As CML_REGISTER, ByVal r2 As CML_REGISTER)
    If r1 = r2 Then Exit Sub
    WLine "xchg [0],[1]", REGISTER_TO_TEXT(r1), REGISTER_TO_TEXT(r2)
End Sub

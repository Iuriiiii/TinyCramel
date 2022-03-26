Attribute VB_Name = "modSHUNTINGYARD"
Option Explicit

Enum OPERATOR_PRECEDENCE
    BELOW = 1
    EQUAL
    ABOVE
End Enum

Private Type SHUNTING_PARAMETER_OR_DIMENSION_DEFINITION
    tokens() As CML_TOKEN
End Type

Type SHUNTING_ITEM_DEFINE
    tokens() As CML_TOKEN
    arguments() As SHUNTING_PARAMETER_OR_DIMENSION_DEFINITION
    dimensions() As SHUNTING_PARAMETER_OR_DIMENSION_DEFINITION
    ptr As Long
    calleable As Boolean
    extra As Long
    v As CML_VARIABLE_DEFINITION
End Type

Private Function VERIFY_PRECEDENCE(op1 As String, op2 As String) As OPERATOR_PRECEDENCE
    Dim s1 As Single
    Dim s2 As Single
    
    s1 = OPERATOR_TO_PRECEDENCE(op1)
    s2 = OPERATOR_TO_PRECEDENCE(op2)
    
    Select Case True
        Case CLng(s1) < CLng(s2)
            VERIFY_PRECEDENCE = BELOW
        Case CLng(s1) = CLng(s2)
            VERIFY_PRECEDENCE = EQUAL
        Case CLng(s1) > CLng(s2)
            VERIFY_PRECEDENCE = ABOVE
    End Select
End Function

Sub PushSID(arr() As SHUNTING_ITEM_DEFINE, item As SHUNTING_ITEM_DEFINE)
    ReDim Preserve arr(UBound(arr) + 1)
    arr(UBound(arr)) = item
End Sub

Private Sub PushStringToSID(arr() As SHUNTING_ITEM_DEFINE, s As String, t As CML_TOKEN_TYPE)
    Dim item1 As CML_TOKEN
    Dim item2 As SHUNTING_ITEM_DEFINE
    item1.s = s
    item1.t = t
    ReDim item2.tokens(1)
    item2.tokens(1) = item1
    PushSID arr, item2
End Sub

Private Sub PushTokenToSID(arr() As SHUNTING_ITEM_DEFINE, token As CML_TOKEN)
    Dim item2 As SHUNTING_ITEM_DEFINE
    Dim n1 As Long
    Dim n2 As Long
    n1 = UBound(arr) + 1
    ReDim Preserve arr(n1)
    ReDim arr(n1).tokens(1)
    arr(n1).tokens(1) = token
End Sub

Private Sub AppendTokenToSID(arr() As SHUNTING_ITEM_DEFINE, token As CML_TOKEN)
    'Dim item2 As SHUNTING_ITEM_DEFINE
    'Dim n1 As Long
    'Dim n2 As Long
    'n1 = UBound(arr)
    'n2 = UBound(arr(n1).tokens) + 1
    'ReDim Preserve arr(n1).tokens(n2)
    'arr(n1).tokens(n2) = token
    PushT arr(UBound(arr)).tokens, token
End Sub

Function PeekSID(arr() As SHUNTING_ITEM_DEFINE) As SHUNTING_ITEM_DEFINE
    Dim sid As SHUNTING_ITEM_DEFINE
    If UBound(arr) = 0 Then
        ReDim sid.tokens(1)
        PeekSID = sid
    Else
        PeekSID = arr(UBound(arr))
    End If
End Function

Function PopSID(arr() As SHUNTING_ITEM_DEFINE) As SHUNTING_ITEM_DEFINE
    If UBound(arr) Then
        PopSID.tokens = arr(UBound(arr)).tokens
        PopSID.extra = arr(UBound(arr)).extra
        ReDim Preserve arr(UBound(arr) - 1)
    End If
End Function

Function GetTokenS(token As CML_TOKEN) As String
    GetTokenS = token.s
End Function

Private Sub POP_OPERATORS_TO_STK(ops() As String, stk() As SHUNTING_ITEM_DEFINE)
    While PeekS(ops) <> "(" And PeekS(ops) <> ""
        PushStringToSID stk, PopS(ops), TOKEN_TYPE_OPERATOR
    Wend
End Sub

Sub ShowShuntingYard(rpn() As SHUNTING_ITEM_DEFINE)
    Dim s As String
    Dim i As Long
    Dim k As Long
    
    For i = 1 To UBound(rpn)
        Ins s, "[" + CStr(i) + "]: "
        For k = 1 To UBound(rpn(i).tokens)
            Ins s, rpn(i).tokens(k).s
        Next
        Ins s, " ("
        Ins s, BinToTokenType(rpn(i).tokens(1).t)
        Ins s, ")"
        Ins s, vbCrLf
    Next
    MsgBox s
End Sub

Function EXPRESSION_CONTAIN_CALLEABLE(rpn() As SHUNTING_ITEM_DEFINE) As Boolean
    Dim i As Long
    
    For i = 1 To UBound(rpn)
        If rpn(i).calleable Then
            EXPRESSION_CONTAIN_CALLEABLE = True
            Exit Function
        End If
    Next
End Function

Function SHUNTING_YARD(tokens() As CML_TOKEN, Optional ByVal idx As Long = 1) As SHUNTING_ITEM_DEFINE()
    Dim stk() As SHUNTING_ITEM_DEFINE
    Dim ops() As String
    Dim pre() As String
    
    ReDim ret(0)
    ReDim ops(0)
    ReDim stk(0)
    ReDim pre(0)
    
    ' De derecha a izquierda
    ' For idx = Ubound(tockens) to 1 Step -1
        ' * Remover TOKEN_TYPE_EOL, TOKEN_TYPE_EOF
    
    ' De izquierda a derecha
    For idx = idx To UBound(tokens)
        Select Case tokens(idx).t
            Case TOKEN_TYPE_EOL, TOKEN_TYPE_EOF
                Exit For
            Case TOKEN_TYPE_IDENTIFIER
                PROCESS_IDENTIFIER tokens, idx, ops, stk
            Case TOKEN_TYPE_SEPARATOR
                PROCESS_SEPARATOR tokens, idx, ops, stk
            Case TOKEN_TYPE_OPERATOR
                PROCESS_OPERATOR tokens, idx, ops, pre, stk
            Case Else
                PushTokenToSID stk, tokens(idx)
        End Select
    Next
    
    While PeekS(ops) <> "(" And PeekS(ops) <> ""
        PushStringToSID stk, PopS(ops), TOKEN_TYPE_OPERATOR
    Wend
    
    While PeekS(pre) <> ""
        PushStringToSID stk, PopS(pre), TOKEN_TYPE_OPERATOR
    Wend
    
    SHUNTING_YARD = stk
End Function

Private Sub PROCESS_IDENTIFIER(tokens() As CML_TOKEN, ByRef idx As Long, ops() As String, stk() As SHUNTING_ITEM_DEFINE)
    Dim par As Long
    Dim cor As Long
    Dim lla As Long
    
    ReDim Preserve stk(UBound(stk) + 1)
    ReDim stk(UBound(stk)).tokens(0)
    
    While tokens(idx).t <> TOKEN_TYPE_OPERATOR And tokens(idx).t <> TOKEN_TYPE_EOL And tokens(idx).t <> TOKEN_TYPE_EOF Or (cor Or lla Or par)
        Select Case tokens(idx).t
            Case TOKEN_TYPE_SEPARATOR
                Select Case tokens(idx).s
                    Case "["
                        Inc cor
                    Case "]"
                        Dec cor
                    Case "{"
                        Inc lla
                    Case "}"
                        Dec lla
                    Case "("
                        Inc par
                    Case ")"
                        Dec par
                End Select
                AppendTokenToSID stk, tokens(idx)
            Case Else
                AppendTokenToSID stk, tokens(idx)
        End Select
        Inc idx
    Wend
    Dec idx
End Sub

Private Sub PROCESS_SEPARATOR(tokens() As CML_TOKEN, ByRef idx As Long, ops() As String, stk() As SHUNTING_ITEM_DEFINE)
    Dim data() As SHUNTING_PARAMETER_OR_DIMENSION_DEFINITION
    
    Select Case GetTokenS(tokens(idx))
        Case "("
            If PeekT(PeekSID(stk).tokens).t <> TOKEN_TYPE_OPERATOR Then
                data = stk(UBound(stk)).arguments
                ReDim Preserve stk(UBound(stk)).arguments(UBound(data) + 1)
                ReDim stk(UBound(stk)).arguments(UBound(data)).tokens(0)
                
                
                While GetTokenS(tokens(idx)) <> ")"
                    'AppendTokenToSID stk, tokens(idx)
                    PushT stk(UBound(stk)).arguments(UBound(data)).tokens, tokens(idx)
                    Inc idx
                Wend
                'AppendTokenToSID stk, tokens(idx)
                PushT stk(UBound(stk)).arguments(UBound(data)).tokens, tokens(idx)
                stk(UBound(stk)).calleable = True
            Else
                PushS ops, "("
            End If
        Case ","
            POP_OPERATORS_TO_STK ops, stk
        Case ")"
            POP_OPERATORS_TO_STK ops, stk
            PopS ops ' Remueve '('
        Case "["
            Assert PeekT(PeekSID(stk).tokens).t <> TOKEN_TYPE_OPERATOR, "Imposible establecer una dimensión a un operador."
            Assert PeekT(PeekSID(stk).tokens).t <> TOKEN_TYPE_NUMBER, "Imposible establecer una dimensión a un número."
            Assert PeekT(PeekSID(stk).tokens).t <> TOKEN_TYPE_FLOAT, "Imposible establecer una dimensión a un número flotante."
            
            data = stk(UBound(stk)).dimensions
            ReDim Preserve stk(UBound(stk)).dimensions(UBound(data) + 1)
            ReDim stk(UBound(stk)).dimensions(UBound(data)).tokens(0)
                
            While GetTokenS(tokens(idx)) <> "]"
                'AppendTokenToSID stk, tokens(idx)
                Inc idx
            Wend
        Case "."
            AppendTokenToSID stk, tokens(idx)
        Case "@"
            Inc stk(UBound(stk)).ptr
    End Select
End Sub

Private Sub PROCESS_OPERATOR(tokens() As CML_TOKEN, ByVal idx As Long, ops() As String, pre() As String, stk() As SHUNTING_ITEM_DEFINE)
    If GetTokenS(tokens(idx)) = "=" Then
        PushS pre, GetTokenS(tokens(idx))
        Exit Sub
    End If
    
    Select Case VERIFY_PRECEDENCE(GetTokenS(tokens(idx)), PeekS(ops))
        Case BELOW
            POP_OPERATORS_TO_STK ops, stk
            'PushS GetTokenS(tokens(idx)), ops
        Case EQUAL
            PushStringToSID stk, PopS(ops), TOKEN_TYPE_OPERATOR
            'PushS GetTokenS(tokens(idx)), ops
        'Case ABOVE
            'PushS GetTokenS(tokens(idx)), ops
    End Select
    
    PushS ops, GetTokenS(tokens(idx))
End Sub

Attribute VB_Name = "modTOKENIZER"
Option Explicit

Private Sub retredim(ret() As CML_TOKEN, ByVal t As CML_TOKEN_TYPE)
    If ret(UBound(ret)).t <> TOKEN_TYPE_NONE And ret(UBound(ret)).t <> t Then
        ReDim Preserve ret(UBound(ret) + 1)
    End If
    
    ret(UBound(ret)).t = t
End Sub

Function BinToTokenType(ByVal t As CML_TOKEN_TYPE) As String
    Select Case t
        Case TOKEN_TYPE_NONE
            BinToTokenType = "None"
        Case TOKEN_TYPE_IDENTIFIER
            BinToTokenType = "Identifier"
        Case TOKEN_TYPE_STRING
            BinToTokenType = "String"
        Case TOKEN_TYPE_OPERATOR
            BinToTokenType = "Operator"
        Case TOKEN_TYPE_SEPARATOR
            BinToTokenType = "Separator"
        Case TOKEN_TYPE_NUMBER
            BinToTokenType = "Number"
        Case TOKEN_TYPE_FLOAT
            BinToTokenType = "Float"
        Case TOKEN_TYPE_INSTRUCTION
            BinToTokenType = "Instruction"
        Case TOKEN_TYPE_EOL
            BinToTokenType = "EOL"
    End Select
End Function

Sub PushT(arr() As CML_TOKEN, t As CML_TOKEN)
    ReDim Preserve arr(UBound(arr) + 1)
    arr(UBound(arr)) = t
End Sub

Function PeekT(arr() As CML_TOKEN) As CML_TOKEN
    PeekT = arr(UBound(arr))
End Function

Function PopT(arr() As CML_TOKEN) As CML_TOKEN
    If UBound(arr) Then
        PopT = arr(UBound(arr))
        ReDim Preserve arr(UBound(arr) - 1)
    End If
End Function

Sub ShowTokens(ret() As CML_TOKEN)
    Dim s As String
    Dim i As Long
    For i = 1 To UBound(ret)
        s = s + "[" + CStr(i) + "]: " + ret(i).s + " (" + BinToTokenType(ret(i).t) + ") " + vbCrLf
    Next
    MsgBox s
End Sub

Private Sub FIX_INSTRUCTIONS(ret() As CML_TOKEN)
    Dim i As Long
    Dim instruction As CML_INSTRUCTIONS
    For i = 1 To UBound(ret)
        If IS_INSTRUCTION(ret(i).s, instruction) Then
            ret(i).t = TOKEN_TYPE_INSTRUCTION
            ret(i).i = instruction
        End If
    Next
End Sub

Function TOKENIZATE(data() As Byte) As CML_TOKEN()
    Dim char As Byte
    Dim ret() As CML_TOKEN
    Dim previous As CML_TOKEN
    Dim actual As CML_TOKEN
    Dim idx As Long
    Dim isstring As Boolean
    Dim iscomment As Boolean
    Dim lastminustype As CML_TOKEN_TYPE
    
    Inc idx
    ReDim ret(1)
    
    For idx = 0 To UBound(data)
        char = data(idx)
        previous = ret(UBound(ret))
        
        If char = 0 Then
            Exit For
        End If
        
        If char = 34 Then
            isstring = Not isstring
            If isstring Then
                retredim ret, TOKEN_TYPE_STRING
            End If
        End If
        
        If isstring Then
            GoTo CONTINUE
        End If
        
        If IS_OPERATOR(char) Then
            retredim ret, TOKEN_TYPE_OPERATOR
            
            Select Case True
                Case char = 43 And data(idx + 1) = 43 Or _
                     char = 45 And data(idx + 1) = 45 Or _
                     char = 60 And data(idx + 1) = 60 Or _
                     char = 60 And data(idx + 1) = 61 Or _
                     char = 60 And data(idx + 1) = 62 Or _
                     char = 62 And data(idx + 1) = 62 Or _
                     char = 62 And data(idx + 1) = 61 Or _
                     char = 33 And data(idx + 1) = 61

                    InsB ret(UBound(ret)).s, char
                    Inc idx
                    char = data(idx)
                Case char = 45
                    lastminustype = previous.t
            End Select
        ElseIf IS_SEPARATOR(char) Then
            If char = 46 And previous.t = TOKEN_TYPE_NUMBER Then
                    ret(UBound(ret)).t = TOKEN_TYPE_FLOAT
            Else
                'retredim ret, TOKEN_TYPE_SEPARATOR
                If previous.t <> TOKEN_TYPE_NONE Then
                    ReDim Preserve ret(UBound(ret) + 1)
                End If
                ret(UBound(ret)).t = TOKEN_TYPE_SEPARATOR
            End If
        ElseIf IS_SPACE(char) Then
            If char = 13 Then
                If data(idx + 1) = 10 Then
                    Inc idx
                    retredim ret, TOKEN_TYPE_EOL
                    char = 0
                End If
            Else
                retredim ret, TOKEN_TYPE_NONE
                char = 0
            End If
        ElseIf IS_NUMERIC(char) Then
            If previous.t = TOKEN_TYPE_IDENTIFIER Or previous.t = TOKEN_TYPE_FLOAT Then
                GoTo CONTINUE
            ElseIf previous.t = TOKEN_TYPE_OPERATOR And previous.s = "-" Then
                If lastminustype = TOKEN_TYPE_OPERATOR Or lastminustype = TOKEN_TYPE_NONE Then
                    ret(UBound(ret)).t = TOKEN_TYPE_NUMBER
                End If
            ElseIf previous.t = TOKEN_TYPE_NONE Then
                ret(UBound(ret)).t = TOKEN_TYPE_NUMBER
            Else
                retredim ret, TOKEN_TYPE_NUMBER
            End If
            
        Else
            retredim ret, TOKEN_TYPE_IDENTIFIER
        End If
        
CONTINUE:
        InsB ret(UBound(ret)).s, char
    Next
    
    Assert isstring = False, "Se detectó cadena sin fin."
    FIX_INSTRUCTIONS ret
    retredim ret, TOKEN_TYPE_EOF
    
    TOKENIZATE = ret
End Function

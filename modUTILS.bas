Attribute VB_Name = "modUTILS"
Option Explicit

Private counter As Long

Function UString() As String
    UString = "Identifier_" + CStr(counter)
    Inc counter
End Function

Sub PushL(arr() As Long, ByVal n As Long)
    ReDim Preserve arr(UBound(arr) + 1)
    arr(UBound(arr)) = n
End Sub

Function PeekL(arr() As Long) As Long
    PeekL = arr(UBound(arr))
End Function

Function PopL(arr() As Long) As Long
    If UBound(arr) Then
        PopL = arr(UBound(arr))
        ReDim Preserve arr(PopL - 1)
    End If
End Function

Sub PushS(arr() As String, s As String)
    ReDim Preserve arr(UBound(arr) + 1)
    arr(UBound(arr)) = s
End Sub

Function PeekS(arr() As String) As String
    PeekS = arr(UBound(arr))
End Function

Function PopS(arr() As String) As String
    If UBound(arr) Then
        PopS = arr(UBound(arr))
        ReDim Preserve arr(UBound(arr) - 1)
    End If
End Function

Sub Inc(i As Long, Optional ByVal n As Long = 1)
    i = i + n
End Sub

Sub Dec(i As Long, Optional ByVal n As Long = 1)
    i = i - n
End Sub

Sub Ins(s As String, c As String)
    s = s + c
End Sub

Sub InsB(s As String, ByVal c As Byte)
    If c Then s = s + Chr$(c)
End Sub

Sub PushEOL(stack() As CML_TOKEN)
    Dim u As Long
    u = UBound(stack) + 1
    ReDim Preserve stack(u)
    stack(u).t = TOKEN_TYPE_EOL
End Sub

Sub InsLine(code As String, ParamArray param() As Variant)
    Dim i As Long
    Dim s As String
    s = param(0)
    For i = 1 To UBound(param)
        s = Replace$(s, "[" + CStr(i) + "]", CStr(param(i)))
    Next
    code = code + s + vbCrLf
End Sub

Sub RemLastLine(code As String)
    Dim arr() As String
    Dim i As Long
    arr = Split(code, vbCrLf)
    code = Empty
    For i = 0 To UBound(arr) - 1
        Ins code, arr(i) + vbCrLf
    Next
End Sub

Function sprintf(s As String, ParamArray params() As Variant) As String
    Dim i As Long
    For i = 0 To UBound(param)
        s = Replace$(s, "[" + CStr(i) + "]", CStr(param(i)))
    Next
    sprintf = s
End Function

Sub Assert(ByVal state As Boolean, ParamArray param() As Variant)
    Dim i As Long
    Dim s As String
    s = param(0)
    For i = 1 To UBound(param)
        s = Replace$(s, "[" + CStr(i) + "]", CStr(param(i)))
    Next
    If Not state Then
        MsgBox s, vbCritical, "TinyCramel"
        End
    End If
End Sub

Sub ShowStrArray(arr() As String)
    Dim s As String
    Dim i As Long
    For i = 1 To UBound(arr)
        s = s + "[" + CStr(i) + "]: " + arr(i) + vbCrLf
    Next
    MsgBox s
End Sub

Function CommandLine() As String()
    Dim ret() As String
    Dim cmdline As String
    Dim i As Long
    Dim char As String
    Dim isstring As Boolean
    
    ReDim ret(1)
    cmdline = Command$
    
    For i = 1 To Len(cmdline)
        char = Mid$(cmdline, i, 1)
        If char = """" Then
            isstring = Not isstring
            char = ""
        End If
        
        If isstring Then
            GoTo CONTINUE
        End If
        
        If char = " " And ret(UBound(ret)) <> Empty Then
            ReDim Preserve ret(UBound(ret) + 1)
            char = ""
        End If
CONTINUE:
        Ins ret(UBound(ret)), char
    Next
    
    CommandLine = ret
End Function

Sub ParseCommandLine()
    Dim cmdline() As String
    Dim waitingforfile As Boolean
    Dim i As Long
    
    cmdline = CommandLine
    
    For i = 1 To UBound(cmdline)
        Select Case cmdline(i)
            Case "-f", "/f"
                waitingforfile = True
            Case "-x64", "/x64"
                x64 = True
            Case Else
                Select Case True
                    Case waitingforfile
                        Assert FILE_EXISTS(cmdline(i)), "El archivo '[1]' no pudo ser abierto.", cmdline(i)
                        main_file = cmdline(i)
                        waitingforfile = False
                End Select
        End Select
    Next
End Sub

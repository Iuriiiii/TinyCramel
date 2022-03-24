Attribute VB_Name = "modARITMETICS"
Option Explicit

Function CALC_NUMBER_WITH_NUMBER(ByVal nl As Long, ByVal nr As Long, op As String) As String
    Select Case op
        Case "*"
            CALC_NUMBER_WITH_NUMBER = CStr(nl * nr)
        Case "/"
            CALC_NUMBER_WITH_NUMBER = CStr(nl / nr)
        Case "<<"
            'CALC_NUMBER_WITH_NUMBER = CStr(nl * nr)
        Case ">>"
            'OPERATOR_TO_PRECEDENCE = 9.6
        Case "%"
            CALC_NUMBER_WITH_NUMBER = CStr(nl Mod nr)
        Case "^"
            'OPERATOR_TO_PRECEDENCE = 9.4
        Case "+"
            CALC_NUMBER_WITH_NUMBER = CStr(nl + nr)
        Case "-"
            CALC_NUMBER_WITH_NUMBER = CStr(nl - nr)
        Case "|"
            CALC_NUMBER_WITH_NUMBER = CStr(nl Or nr)
        Case "Xor"
            CALC_NUMBER_WITH_NUMBER = CStr(nl Xor nr)
        Case "&"
            CALC_NUMBER_WITH_NUMBER = CStr(nl And nr)
        Case "=="
            CALC_NUMBER_WITH_NUMBER = CStr(nl = nr)
        Case "<>", "!="
            CALC_NUMBER_WITH_NUMBER = CStr(nl <> nr)
        Case "<="
            CALC_NUMBER_WITH_NUMBER = CStr(nl <= nr)
        Case "<"
            CALC_NUMBER_WITH_NUMBER = CStr(nl < nr)
        Case ">"
            CALC_NUMBER_WITH_NUMBER = CStr(nl > nr)
        Case ">="
            CALC_NUMBER_WITH_NUMBER = CStr(nl >= nr)
        Case "="
            CALC_NUMBER_WITH_NUMBER = CStr(nl = nr)
    End Select
End Function

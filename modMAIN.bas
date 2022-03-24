Attribute VB_Name = "modMAIN"
Option Explicit

Sub Main()
    INIT_COMPILER
    PARSE
    COMPILE
    MsgBox GET_COMPILED_CODE
End Sub

Private Sub INIT_COMPILER()
    ParseCommandLine
    INIT_DECLARATIONS
End Sub

Private Sub COMPILE()
    COMPILE_INTEL_WIN32
End Sub

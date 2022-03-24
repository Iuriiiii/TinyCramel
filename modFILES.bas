Attribute VB_Name = "modFILES"
Option Explicit

Private Const OF_EXIST         As Long = &H4000
Private Const OFS_MAXPATHNAME  As Long = 128
Private Const HFILE_ERROR      As Long = -1
 
Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type
 
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long

Public Function FILE_EXISTS(ByVal Fname As String) As Boolean
    Dim lRetVal As Long
    Dim OfSt As OFSTRUCT
    lRetVal = OpenFile(Fname, OfSt, OF_EXIST)
    FILE_EXISTS = lRetVal <> HFILE_ERROR
End Function

Function FILE_READ_AS_ARRAY(path As String) As Byte()
    On Error Resume Next
    Dim iFile As Long
    iFile = FreeFile
    Open path For Input As #iFile
    FILE_READ_AS_ARRAY = InputB(LOF(iFile), iFile)
    ReDim Preserve FILE_READ_AS_ARRAY(LOF(iFile) + 10) ' Añadimos carácteres nulos
    Close #iFile
End Function

Sub FILE_WRITE(path As String, content As String)
    On Error Resume Next
    Dim iFile As Long
    iFile = FreeFile
    Open path For Output As #iFile
    Write #iFile, content
    Close #iFile
End Sub

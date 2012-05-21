Attribute VB_Name = "modTempfile"
Option Explicit

Declare Function GetTempPath Lib "kernel32" Alias _
"GetTempPathA" (ByVal nBufferLength As Long, ByVal _
lpBuffer As String) As Long

Public Const MAX_PATH = 260



Public Function GetTmpPath()
Dim strFolder As String
Dim lngResult As Long
strFolder = String(MAX_PATH, 0)
lngResult = GetTempPath(MAX_PATH, strFolder)
If lngResult <> 0 Then
    GetTmpPath = Left(strFolder, InStr(strFolder, _
    Chr(0)) - 1)
Else
    GetTmpPath = ""
End If
End Function






Attribute VB_Name = "modShared"
Option Explicit

'Public Windows As Collection
'Public CompareWindows As Collection
'Public DiffPos As Long
Global Const LANG_US = &H409

'Public Const ChunkSize As Long = 30000
Public ChunkSize As Long

Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
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

Sub SaveMySetting(key, Value)
    SaveSetting "hexed", "settings", key, Value
End Sub

Function GetMySetting(key, def)
    GetMySetting = GetSetting("hexed", "settings", key, def)
End Function

Sub FormPos(fform As Form, Optional andSize As Boolean = False, Optional save_mode As Boolean = False)
    
    On Error Resume Next
    
    Dim f, sz, i, ff, def
    f = Split(",Left,Top,Height,Width", ",")
    
    If fform.WindowState = vbMinimized Then Exit Sub
    If andSize = False Then sz = 2 Else sz = 4
    
    For i = 1 To sz
        If save_mode Then
            ff = CallByName(fform, f(i), VbGet)
            SaveSetting App.EXEName, fform.name & ".FormPos", f(i), ff
        Else
            def = CallByName(fform, f(i), VbGet)
            ff = GetSetting(App.EXEName, fform.name & ".FormPos", f(i), def)
            CallByName fform, f(i), VbLet, ff
        End If
    Next
    
End Sub

Function isIDE() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number <> 0 Then isIDE = True
End Function

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim i As Long
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Sub push(ary, Value) 'this modifies parent ary object
    On Error GoTo init
    Dim x
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = Value
    Exit Sub
init:     ReDim ary(0): ary(0) = Value
End Sub
 
Function shex(ByVal data) As String
    If Len(data) = 1 Then
        shex = "0" & data
    Else
        shex = data
    End If
End Function

Function toHexString(ByVal data As String, doUnicode As Boolean, Optional prefix As String = "\x") As String
    Dim b() As Byte
    Dim i As Long, r As String
    
    If Len(data) = 0 Then Exit Function
    
    If doUnicode Then
        data = StrConv(data, vbUnicode, LANG_US)
    End If
    
    b() = StrConv(data, vbFromUnicode, LANG_US)
    
    For i = 0 To UBound(b)
        r = r & prefix & shex(Hex(b(i)))
    Next
    
    toHexString = r
    
End Function

Function toCharDump(ByVal data As String) As String
    
    If Len(data) = 0 Then Exit Function
    
    Dim b() As Byte
    Dim r As String
    Dim i As Long
    Dim asciiCount As Long
    
    b() = StrConv(data, vbFromUnicode, LANG_US)
    
    For i = 0 To UBound(b)
        If b(i) >= 32 And b(i) <= 127 Then
            r = r & Chr(b(i))
            asciiCount = asciiCount + 1
        Else
            r = r & "."
        End If
    Next
    
    If pcent(Len(data), asciiCount) < 30 Then
        toCharDump = toHexString(data, False, " ")
    Else
        toCharDump = r
    End If
    
End Function


Private Function pcent(max, cnt) As Long
    On Error Resume Next
    pcent = (cnt / max) * 100
End Function

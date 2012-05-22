VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'License:   GPL
'Copyright: 2005 iDefense a Verisign Company
'Site:      http://labs.idefense.com
'
'Author:    David Zimmer <david@idefense.com, dzzie@yahoo.com>
'
'         This program is free software; you can redistribute it and/or modify it
'         under the terms of the GNU General Public License as published by the Free
'         Software Foundation; either version 2 of the License, or (at your option)
'         any later version.
'
'         This program is distributed in the hope that it will be useful, but WITHOUT
'         ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
'         FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
'         more details.
'
'         You should have received a copy of the GNU General Public License along with
'         this program; if not, write to the Free Software Foundation, Inc., 59 Temple
'         Place, Suite 330, Boston, MA 02111-1307 USA

'Event Progress(percentComplete As Long, abort As Boolean)
'Event MatchMade(offset As Long)

Dim d As New RegExp
Dim mc As MatchCollection
Dim m As match
Dim ret() As String

Function FromBytes(searchFor As String, buf() As Byte, Optional doUnicode As Boolean = False, Optional caseInsensitive As Boolean = False) As String()
    On Error Resume Next
    setPattern searchFor, doUnicode, caseInsensitive
    Search buf, 0
    FromBytes = ret
    Erase ret
End Function

Private Function setPattern(searchFor As String, doUnicode As Boolean, caseInsensitive As Boolean)
    
    d.IgnoreCase = caseInsensitive
    d.Global = True
    If doUnicode Then
        d.Pattern = toHexString(searchFor, True)
    Else
         d.Pattern = searchFor
    End If
    
End Function

Function FromString(searchFor As String, buffer As String, Optional doUnicode As Boolean = False, Optional caseInsensitive As Boolean = False) As String()
    On Error Resume Next
    Dim buf()  As Byte
    setPattern searchFor, doUnicode, caseInsensitive
    buf = StrConv(buffer, vbFromUnicode, LANG_US)
    Search buf, 0
    FromString = ret
    Erase ret
End Function

Function FileExists(path As String) As Boolean
    On Error GoTo hell
    Dim tmp As String
    tmp = path
    If Len(tmp) = 0 Then Exit Function
    If Dir(tmp, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
    Exit Function
hell: FileExists = False
End Function


Function FromFile(searchFor As String, fpath As String, Optional doUnicode As Boolean = False, Optional caseInsensitive As Boolean = False) As String()
    On Error GoTo hell
    
    Dim f As Long, pointer As Long
    Dim buf()  As Byte
    Dim x As Long
    
    f = FreeFile
    
    If Not FileExists(fpath) Then
        MsgBox "File not found: " & fpath, vbExclamation
        GoTo done
    End If
    
    setPattern searchFor, doUnicode, caseInsensitive
        
    ReDim buf(9000)
    Open fpath For Binary Access Read As f
    
    Do While pointer < LOF(f)
        pointer = Seek(f)
        x = LOF(f) - pointer
        If x < 1 Then Exit Do
        If x < 9000 Then ReDim buf(x)
        Get f, , buf()
        Search buf, pointer
    Loop
      
    Close f
    
    On Error Resume Next
    FromFile = ret
    Erase ret
   
        
Exit Function
hell:
      MsgBox "Error getting strings: " & Err.Description, vbExclamation
      Close f
done:

End Function

Private Sub Search(buf() As Byte, base As Long)
    Dim b As String
    
    b = StrConv(buf, vbUnicode, LANG_US)
    Set mc = d.Execute(b)
    
    For Each m In mc
        push ret(), (base + m.FirstIndex - 1) & "," & toCharDump(m.Value)
    Next
    
End Sub


'Function LineGrep(sBuf() As String, sMatch As String) As String
'
'    Dim ret() As String
'    Dim i As Long
'
'    'how i really really hate regexp...
'    'd.Pattern = "\n(\w)*" & Trim(sMatch) & "(\w)*\n"
'    'd.Global = True
'    'd.IgnoreCase = True
'    'Set mc = d.Execute(sBuf)
'    '
'    'For Each m In mc
'    '    push ret(), Mid(m.value, 2)
'    'Next
'
'    For i = 0 To UBound(sBuf)
'        If InStr(1, sBuf(i), sMatch, vbTextCompare) > 0 Then
'            push ret, sBuf(i)
'        End If
'    Next
'
'
'    LineGrep = Join(ret, vbCrLf)
'
'End Function

Private Sub push(ary, Value) 'this modifies parent ary object
    On Error GoTo init
    Dim x
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = Value
    Exit Sub
init:     ReDim ary(0): ary(0) = Value
End Sub
 

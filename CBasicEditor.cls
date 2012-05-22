VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBasicEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private myForm As frmEditor

Property Get Editor() As HexEd
    Set Editor = myForm.HexEditor
End Property

'Public Function LoadFile(fpath As String, Optional ViewOnly As Boolean = True) As Boolean
'    LoadFile = myForm.HexEditor.LoadFile(fpath, ViewOnly)
'End Function
'
'Public Function LoadString(data As String, Optional ViewOnly As Boolean = True) As Boolean
'    LoadString = myForm.HexEditor.LoadString(data, ViewOnly)
'End Function
'
'Public Function LoadByteArray(bArray As Variant, Optional ViewOnly As Boolean = True) As Boolean
'    LoadByteArray = myForm.HexEditor.LoadByteArray(bArray, ViewOnly)
'End Function

Private Sub Class_Initialize()
    Set myForm = New frmEditor
End Sub



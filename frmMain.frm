VERSION 5.00
Begin VB.Form frmEditor 
   Caption         =   "Hexeditor v 0.3"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   Visible         =   0   'False
   Begin prjVbHex.HexEd HexEditor 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   14420
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public key As String
Public Filename As String
Public title As String

Private Sub Form_Activate()
    Me.HexEditor.SetFocus
End Sub

Private Sub Form_GotFocus()
    Me.HexEditor.SetFocus
End Sub

Private Sub Form_Load()
    
    Dim b() As Byte
    b() = StrConv(String(50, "A"), vbFromUnicode, &H409)
    
    'if you want to force it to load from memory only set chunksize = datasize + 1
    'else it will create a temp file for you automatically...
    'HexEditor.ReadChunkSize = &H50000
    HexEditor.ForceMemOnlyLoading = True
    Me.HexEditor.LoadString (String(&H50000, "B"))
    
    Me.Visible = True
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim answer As Long
    If Not HexEditor.ReadOnly And Me.HexEditor.IsDirty Then
        answer = MsgBox("Save changes to " & Me.title & " ?", vbYesNoCancel Or vbExclamation, "VB Hexedit")
        If answer = vbYes Then Me.HexEditor.Save
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    HexEditor.Width = Me.ScaleWidth
    HexEditor.Height = Me.ScaleHeight
End Sub

Private Sub HexEditor_Dirty()
    Me.Caption = Me.title & "  *"
End Sub

Private Sub HexEditor_RightClick()
   PopupMenu mnuFile
End Sub

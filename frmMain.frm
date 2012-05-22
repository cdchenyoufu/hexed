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
      Begin VB.Menu mnuSearch 
         Caption         =   "Search"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'todo: copy inserts 0x03..have to kill that..

Private Sub Form_Activate()
    Me.HexEditor.SetFocus
End Sub

Private Sub Form_GotFocus()
    Me.HexEditor.SetFocus
End Sub

Private Sub Form_Load()
   Me.Visible = True
   mnuOpen.Visible = isIDE()
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim answer As Long
    If Not HexEditor.ReadOnly And Me.HexEditor.IsDirty Then
        answer = MsgBox("Save changes?", vbYesNoCancel Or vbExclamation, "VB Hexedit")
        If answer = vbYes Then Me.HexEditor.Save
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    HexEditor.Width = Me.ScaleWidth
    HexEditor.Height = Me.ScaleHeight
End Sub

Private Sub HexEditor_Dirty()
    'Me.Caption = Me.title & "  *"
End Sub

Private Sub HexEditor_RightClick()
    'mnuOpen.Visible = Not HexEditor.ReadOnly
    mnuSave.Visible = Not HexEditor.ReadOnly
    PopupMenu mnuFile
End Sub

Private Sub mnuAbout_Click()
    HexEditor.ShowAbout
End Sub

Private Sub mnuOpen_Click()
    HexEditor.ShowOpen
End Sub

Private Sub mnuSearch_Click()
     HexEditor.ShowFind
End Sub

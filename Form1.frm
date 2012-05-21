VERSION 5.00
Begin VB.Form frmEditor 
   Caption         =   "Hexeditor v 0.3"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
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
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public key As String
Public Filename As String
Public Title As String




Private Sub Form_Activate()
    Set frmMain.ActiveWindow = Me
    frmMain.tbFiles.Tabs(Me.key).Selected = True
    Me.HexEditor.SetFocus
End Sub

Private Sub Form_GotFocus()
    Me.HexEditor.SetFocus
End Sub


Private Sub Form_Load()
    Me.Width = 12000
    Me.Height = 6000
End Sub

Private Sub Form_LostFocus()
    Set frmMain.ActiveWindow = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim answer As Long
    If Me.HexEditor.IsDirty Then
        answer = MsgBox("Save changes to " & Me.Title & " ?", vbYesNoCancel Or vbExclamation, "VB Hexedit")
        Select Case answer
            Case vbYes
                Me.HexEditor.Save
                Call frmMain.UnloadWindow(Me.key)
            Case vbNo
                Call frmMain.UnloadWindow(Me.key)
            Case vbCancel
                Cancel = True
                Exit Sub
        End Select
    Else
        Call frmMain.UnloadWindow(Me.key)
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    HexEditor.Width = Me.ScaleWidth
    HexEditor.Height = Me.ScaleHeight
End Sub

Private Sub HexEditor_Dirty()
    Me.Caption = Me.Title & "  *"
End Sub

Private Sub HexEditor_RightClick()
   ' PopupMenu frmMain.mnuHead(1)
End Sub

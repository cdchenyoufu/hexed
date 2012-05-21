VERSION 5.00
Begin VB.Form frmWindows 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compare Files"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   Icon            =   "frmWindows.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   4320
      Width           =   1575
   End
   Begin VB.ListBox lstWindows 
      Height          =   4110
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Set CompareWindows = New Collection
    For i = 0 To lstWindows.ListCount - 1
        If lstWindows.Selected(i) Then
            CompareWindows.Add Windows(i + 1)
        End If
    Next
    Me.Hide
    DiffPos = 0
    Call frmMain.TileWindowsV
    Call frmMain.FindDifference
    
End Sub

Private Sub Form_Activate()
    lstWindows.Clear
    Dim Form As frmEditor
    For Each Form In Windows
        lstWindows.AddItem Form.Title
    Next
End Sub


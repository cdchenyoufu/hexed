VERSION 5.00
Begin VB.Form frmInsert 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Insert Bytes"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3480
   Icon            =   "frmInsert.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   1470
      Width           =   3285
      Begin VB.TextBox txtTest 
         Height          =   285
         Left            =   630
         TabIndex        =   13
         Top             =   210
         Width           =   615
      End
      Begin VB.TextBox txtKeyCode 
         Height          =   315
         Left            =   2430
         TabIndex        =   12
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Keycode 0x"
         Height          =   225
         Left            =   1470
         TabIndex        =   11
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1860
      TabIndex        =   7
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   810
      Width           =   3255
      Begin VB.CheckBox chkValueHex 
         Caption         =   "Hex"
         Height          =   285
         Left            =   2490
         TabIndex        =   9
         Top             =   180
         Width           =   675
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1770
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "0"
         Top             =   150
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Value to insert:"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   210
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   3255
      Begin VB.CheckBox chkBytesHex 
         Caption         =   "Hex"
         Height          =   285
         Left            =   2490
         TabIndex        =   8
         Top             =   210
         Width           =   675
      End
      Begin VB.TextBox txtCount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Text            =   "8"
         Top             =   210
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Bytes to insert:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ByteCount As Long
Public ByteValue As Byte
Public Cancel As Boolean

Private Sub cmdCancel_Click()
    Cancel = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    On Error GoTo hell
    
    If chkBytesHex.Value = 1 Then
        ByteCount = CLng("&h" & txtCount)
    Else
        ByteCount = CLng(txtCount)
    End If
    
    If chkValueHex.Value = 1 Then
        ByteValue = CByte("&h" & txtValue)
    Else
        ByteValue = CByte(txtValue)
    End If
    
    Cancel = False
    Me.Hide
    
    Exit Sub
hell:
    MsgBox Err.Description
End Sub

Private Sub Form_Activate()
    txtCount.SetFocus
    txtCount.SelStart = 0
    txtCount.SelLength = Len(txtCount.Text)
End Sub

Private Sub txtTest_KeyPress(KeyAscii As Integer)
     On Error Resume Next
     txtKeyCode = Hex(KeyAscii)
End Sub


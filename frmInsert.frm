VERSION 5.00
Begin VB.Form frmInsert 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Insert Bytes"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3540
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   3255
      Begin VB.OptionButton Option4 
         Caption         =   "Dec"
         Height          =   255
         Left            =   2520
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Hex"
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Text            =   "0"
         Top             =   380
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Value to insert:"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   400
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.TextBox txtCount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Text            =   "8"
         Top             =   380
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Hex"
         Height          =   255
         Left            =   2520
         TabIndex        =   2
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Dec"
         Height          =   255
         Left            =   2520
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Bytes to insert:"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   400
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
    Cancel = False
    Me.Hide
End Sub

Private Sub Form_Activate()
    txtCount.SetFocus
    txtCount.SelStart = 0
    txtCount.SelLength = Len(txtCount.Text)
End Sub


VERSION 5.00
Object = "{71532E87-B06E-431D-AC3A-686170A406ED}#2.1#0"; "hexed.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   10470
      TabIndex        =   1
      Top             =   180
      Width           =   1125
   End
   Begin rhexed.HexEd HexEd1 
      Height          =   5385
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   9499
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim edit As New CBasicEditor
    
    edit.GetEditor.LoadFile "c:\_jbig2.data", True
    
    
End Sub

Private Sub Form_Load()
    HexEd1.ReadOnly = True
End Sub

VERSION 5.00
Object = "{9A143468-B450-48DD-930D-925078198E4D}#1.1#0"; "hexed.ocx"
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
   Begin rhexed.HexEd HexEd1 
      Height          =   4875
      Left            =   900
      TabIndex        =   1
      Top             =   585
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   8599
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   10470
      TabIndex        =   0
      Top             =   180
      Width           =   1125
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    HexEd1.LoadString "test", False
    
    Dim edit As New CHexEditor
    
    'edit.GetEditor.LoadFile "c:\_jbig2.data", False
    edit.Editor.AdjustBaseOffset = &H401000
    edit.Editor.LoadString String(1000, "A"), True
    
    
End Sub

 


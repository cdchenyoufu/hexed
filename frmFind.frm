VERSION 5.00
Begin VB.Form frmFind 
   Caption         =   "Find"
   ClientHeight    =   1455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSensitive 
      Caption         =   "Case Sensitive"
      Height          =   345
      Left            =   2820
      TabIndex        =   4
      Top             =   420
      Width           =   1455
   End
   Begin VB.CheckBox chkUnicode 
      Caption         =   "Unicode"
      Height          =   345
      Left            =   1560
      TabIndex        =   3
      Top             =   420
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   375
      Left            =   4530
      TabIndex        =   2
      Top             =   60
      Width           =   1245
   End
   Begin VB.TextBox txtFind 
      Height          =   345
      Left            =   690
      TabIndex        =   1
      Top             =   30
      Width           =   3645
   End
   Begin VB.Label Label3 
      Caption         =   "(does not work with unicode option)"
      Height          =   255
      Left            =   450
      TabIndex        =   6
      Top             =   1110
      Width           =   5055
   End
   Begin VB.Label Label2 
      Caption         =   "Use \x00 style strings for hex character searchs"
      Height          =   255
      Left            =   90
      TabIndex        =   5
      Top             =   870
      Width           =   3435
   End
   Begin VB.Label Label1 
      Caption         =   "Find"
      Height          =   255
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   555
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public owner As HexEd

Private Sub Command1_Click()
    Dim uni As Boolean
    Dim sens As Boolean
    Dim ret() As String
    
    If Len(txtFind) = 0 Then Exit Sub
    
    uni = IIf(chkUnicode.Value = 1, True, False)
    sens = IIf(chkSensitive.Value = 1, False, True)
    
    ret() = owner.Search(txtFind, uni, sens)
    
    If AryIsEmpty(ret) Then
        MsgBox "0 occurances found..", vbInformation
    Else
        frmOffsetList.LoadList owner, ret()
        UnLoad Me
    End If
    
End Sub

Private Sub Form_Load()
    SetTopMost Me
    FormPos Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormPos Me, False, True
End Sub

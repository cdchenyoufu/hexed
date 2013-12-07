VERSION 5.00
Begin VB.Form frmGoto 
   Caption         =   "Goto Offset"
   ClientHeight    =   1485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4845
   Icon            =   "frmGoto.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   " From "
      Height          =   885
      Left            =   750
      TabIndex        =   4
      Top             =   480
      Width           =   2715
      Begin VB.OptionButton optBeginning 
         Caption         =   "Beginning of file"
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1425
      End
      Begin VB.OptionButton OptCurrent 
         Caption         =   "Current Position"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   540
         Width           =   1515
      End
   End
   Begin VB.CheckBox chkIsHex 
      Caption         =   "Hex"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   90
      Width           =   675
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   345
      Left            =   3660
      TabIndex        =   2
      Top             =   90
      Width           =   1095
   End
   Begin VB.TextBox txtOffset 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   780
      TabIndex        =   1
      Text            =   "0"
      Top             =   60
      Width           =   1755
   End
   Begin VB.Label Label1 
      Caption         =   "Offset"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmGoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public owner As HexEd

Public Sub Display(he As HexEd)
    Set owner = he
    Form_Load
End Sub

Private Sub Command1_Click()
    Dim offset As Long
    Dim curOffset As Long
    Dim finalOffset As Long
    
    On Error GoTo hell
    
    If chkIsHex.Value = 1 Then
        offset = CLng("&h" & txtOffset)
    Else
        offset = CLng(txtOffset)
    End If
    
    If optBeginning.Value Then
        finalOffset = offset - owner.AdjustBaseOffset 'absolute address to relative...
    Else 'OptCurrent.Value Then
        finalOffset = owner.SelStart + offset 'can be negative number too...already a relative value..
    End If
    
    If finalOffset < 0 Then
        MsgBox "The specified offset is invalid ( < 0 )"
        Exit Sub
    End If
    
     If finalOffset > owner.FileSize Then
        MsgBox "The specified offset is invalid ( > FileSize )"
        Exit Sub
    End If
    
    owner.scrollTo finalOffset
    UnLoad Me
    
    Exit Sub
hell:
    MsgBox Err.Description
    
End Sub

Private Sub Form_Load()
    On Error Resume Next
    FormPos Me
    SetTopMost Me
    txtOffset = GetSetting("hexed", "settings", "goto_offset", 0)
    chkIsHex.Value = GetSetting("hexed", "settings", "goto_isHex", 1)
    txtOffset.SelStart = 0
    txtOffset.SelLength = Len(txtOffset)
    Me.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormPos Me, False, True
    SaveSetting "hexed", "settings", "goto_offset", txtOffset
    SaveSetting "hexed", "settings", "goto_isHex", chkIsHex.Value
End Sub

Private Sub txtOffset_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Command1_Click
End Sub

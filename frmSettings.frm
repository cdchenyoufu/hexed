VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Disk"
      Height          =   1935
      Left            =   2760
      TabIndex        =   15
      Top             =   2280
      Width           =   2535
      Begin VB.CommandButton cmdChunk 
         Caption         =   "Change"
         Height          =   375
         Left            =   600
         TabIndex        =   25
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Chunk size"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Font"
      Height          =   1935
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton cmdFont 
         Caption         =   "Change"
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblFont 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Colors"
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin MSComDlg.CommonDialog dlg 
         Left            =   1920
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label16 
         Caption         =   "Ascii FG"
         Height          =   255
         Left            =   600
         TabIndex        =   23
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label Label15 
         Caption         =   "Ascii BG"
         Height          =   255
         Left            =   600
         TabIndex        =   22
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label14 
         Caption         =   "Modified"
         Height          =   255
         Left            =   600
         TabIndex        =   21
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "Hex Odd Column"
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "Hex Even Column"
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "Hex BG"
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Margin FG"
         Height          =   255
         Left            =   600
         TabIndex        =   17
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Margin BG"
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblAFG 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label lblABG 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label lblMOD 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label lblHOC 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label lblHEC 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label lblHBG 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lblMFG 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblMBG 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    Call SetSettings
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdFont_Click()
On Error GoTo hell:
    Dlg.FontName = lblFont.FontName
    Dlg.FontSize = lblFont.FontSize
    Dlg.Flags = cdlCFFixedPitchOnly Or cdlCFScreenFonts
    Dlg.ShowFont
    lblFont.FontName = Dlg.FontName
    lblFont.FontSize = Dlg.FontSize
    lblFont.Caption = Dlg.FontName & "   " & Dlg.FontSize
    Exit Sub
hell:
    Err.Clear
End Sub

Private Sub cmdOK_Click()
    Call SetSettings
    Me.Hide
End Sub

Private Sub Form_Activate()
    If Not frmMain.ActiveWindow Is Nothing Then
        Call GetSettings
    End If
End Sub

Private Sub GetSettings()
    Set lblFont.Font = frmMain.ActiveWindow.HexEditor.Font
    lblFont.Caption = frmMain.ActiveWindow.HexEditor.Font.name & "   " & frmMain.ActiveWindow.HexEditor.Font.size
    lblMFG.BackColor = frmMain.ActiveWindow.HexEditor.MarginColor
    lblHEC.BackColor = frmMain.ActiveWindow.HexEditor.EvenColor
    lblHOC.BackColor = frmMain.ActiveWindow.HexEditor.OddColor
    lblMOD.BackColor = frmMain.ActiveWindow.HexEditor.ModColor
    lblAFG.BackColor = frmMain.ActiveWindow.HexEditor.AsciiColor
    
    lblABG.BackColor = frmMain.ActiveWindow.HexEditor.AsciiBGColor
    lblHBG.BackColor = frmMain.ActiveWindow.HexEditor.HexBGColor
    lblMBG.BackColor = frmMain.ActiveWindow.HexEditor.MarginBGColor
    
End Sub

Private Sub SetSettings()
    frmMain.ActiveWindow.HexEditor.MarginColor = lblMFG.BackColor
    frmMain.ActiveWindow.HexEditor.EvenColor = lblHEC.BackColor
    frmMain.ActiveWindow.HexEditor.OddColor = lblHOC.BackColor
    frmMain.ActiveWindow.HexEditor.ModColor = lblMOD.BackColor
    frmMain.ActiveWindow.HexEditor.AsciiColor = lblAFG.BackColor
    
    frmMain.ActiveWindow.HexEditor.AsciiBGColor = lblABG.BackColor
    frmMain.ActiveWindow.HexEditor.HexBGColor = lblHBG.BackColor
    frmMain.ActiveWindow.HexEditor.MarginBGColor = lblMBG.BackColor
    Set frmMain.ActiveWindow.HexEditor.Font = lblFont.Font
End Sub

Private Sub lblABG_Click()
    On Error GoTo hell
    Dlg.Color = lblABG.BackColor
    Dlg.ShowColor
    lblABG.BackColor = Dlg.Color
    Exit Sub
hell:
    Err.Clear
End Sub

Private Sub lblAFG_Click()
    On Error GoTo hell
    Dlg.Color = lblAFG.BackColor
    Dlg.ShowColor
    lblAFG.BackColor = Dlg.Color
    Exit Sub
hell:
    Err.Clear
End Sub

Private Sub lblHBG_Click()
    On Error GoTo hell
    Dlg.Color = lblHBG.BackColor
    Dlg.ShowColor
    lblHBG.BackColor = Dlg.Color
    Exit Sub
hell:
    Err.Clear
End Sub

Private Sub lblHEC_Click()
    On Error GoTo hell
    Dlg.Color = lblHEC.BackColor
    Dlg.ShowColor
    lblHEC.BackColor = Dlg.Color
    Exit Sub
hell:
    Err.Clear

End Sub

Private Sub lblHOC_Click()
    On Error GoTo hell
    Dlg.Color = lblHOC.BackColor
    Dlg.ShowColor
    lblHOC.BackColor = Dlg.Color
    Exit Sub
hell:
    Err.Clear
End Sub

Private Sub lblMBG_Click()
    On Error GoTo hell
    Dlg.Color = lblMBG.BackColor
    Dlg.ShowColor
    lblMBG.BackColor = Dlg.Color
    Exit Sub
hell:
    Err.Clear
End Sub

Private Sub lblMFG_Click()
    On Error GoTo hell
    Dlg.Color = lblMFG.BackColor
    Dlg.ShowColor
    lblMFG.BackColor = Dlg.Color
    Exit Sub
hell:
    Err.Clear
End Sub

Private Sub lblMOD_Click()
    On Error GoTo hell
    Dlg.Color = lblMOD.BackColor
    Dlg.ShowColor
    lblMOD.BackColor = Dlg.Color
    Exit Sub
hell:
    Err.Clear
End Sub

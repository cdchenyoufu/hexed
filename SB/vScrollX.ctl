VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.UserControl vScrollXL 
   Alignable       =   -1  'True
   BackColor       =   &H00C000C0&
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   ToolboxBitmap   =   "vScrollX.ctx":0000
   Begin VB.PictureBox vScroll 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5250
      Left            =   0
      ScaleHeight     =   350
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   0
      Top             =   0
      Width           =   900
      Begin VB.PictureBox vThumb 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   57
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin VB.PictureBox btnDown 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   57
         TabIndex        =   2
         Top             =   3960
         Width           =   855
      End
      Begin VB.PictureBox btnUp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   57
         TabIndex        =   1
         Top             =   0
         Width           =   855
      End
   End
   Begin MSComctlLib.ImageList gfx 
      Left            =   5520
      Top             =   585
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vScrollX.ctx":0314
            Key             =   "up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vScrollX.ctx":0373
            Key             =   "left"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vScrollX.ctx":03D5
            Key             =   "right"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vScrollX.ctx":0437
            Key             =   "down"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "vScrollXL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_LEFT = &H1
Private Const BF_ADJUST = &H2000   ' Calculate the space left over.
Private Const BF_BOTTOM = &H8
Private Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_DIAGONAL = &H10
Private Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Private Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
Private Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Private Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Private Const BF_FLAT = &H4000     ' For flat rather than 3-D borders.
Private Const BF_MIDDLE = &H800    ' Fill in the middle.
Private Const BF_MONO = &H8000     ' For monochrome borders.
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BF_SOFT = &H1000     ' Use for softer buttons.
Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Private Const BDR_INNER = &HC
Private Const BDR_OUTER = &H3
Private Const BDR_RAISED = &H5
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKEN = &HA
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

Private mMin As Currency
Private mMax As Currency
Private mValue As Currency
Private mRange As Currency
Private mPercent As Double
Private mStep As Double
Private mThumbWdh As Long

Private mStart As Long
Private mOffset As Long
Private DoScroll As Boolean

Public Event Change()





Private Sub ClickTop()
    Call SinkCtl(btnUp, "up")
    If mValue > mMin Then
        mValue = mValue - 1
        RaiseEvent Change
    End If
    Call Redraw
End Sub

Private Sub ClickDown()
    Call SinkCtl(btnDown, "down")
    If mValue < mMax Then
        mValue = mValue + 1
        RaiseEvent Change
    End If
    Call Redraw
    

End Sub


Private Sub NoClickright()
    Call RaiseCtl(btnDown, "down")
End Sub

Private Sub NoClickTop()
    Call RaiseCtl(btnUp, "up")
End Sub


Private Sub btnUp_DblClick()
    Call StartClickTop
End Sub

Private Sub btnUp_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call StartClickTop
End Sub

Private Sub btnUp_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    DoScroll = False
    Call NoClickTop
End Sub

Private Sub btnDown_DblClick()
    Call StartClickDown
End Sub

Private Sub btnDown_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call StartClickDown
End Sub

Private Sub btnDown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    DoScroll = False
    Call NoClickright
End Sub

Private Sub vScroll_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Current As Long
    Dim Diff As Long
    Dim Value As Currency
    If Not Button = vbLeftButton Then Exit Sub
    Current = y - vThumb.ScaleHeight / 2
    
    If Current < btnUp.ScaleHeight Then Current = btnUp.ScaleHeight
    If Current + vThumb.ScaleHeight > UserControl.ScaleHeight - btnDown.ScaleHeight Then Current = UserControl.ScaleHeight - btnDown.ScaleHeight - vThumb.ScaleHeight
    vThumb.Top = Current
    Value = Current - btnUp.ScaleHeight
    Value = Round(Value / mStep + mMin)
    If Value > mMax Then Value = mMax
    If Value <> mValue Then
        mValue = Value
        RaiseEvent Change
    End If
    Call Redraw
End Sub

Private Sub vThumb_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mStart = vThumb.Top
    mOffset = y
End Sub

Private Sub vThumb_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Current As Long
    Dim Diff As Long
    Dim Value As Currency
    If Not Button = vbLeftButton Then Exit Sub
    Current = vThumb.Top
    Diff = (Current + y) - (mStart + mOffset)

    Current = mStart + Diff
    If Current < btnUp.ScaleHeight Then Current = btnUp.ScaleHeight
    If Current + vThumb.ScaleHeight > UserControl.ScaleHeight - btnDown.ScaleHeight Then Current = UserControl.ScaleHeight - btnDown.ScaleHeight - vThumb.ScaleHeight
    vThumb.Top = Current
    vScroll.Refresh
    Value = Current - btnUp.ScaleHeight
    Value = Round(Value / mStep + mMin)
    If Value > mMax Then Value = mMax
    If Value <> mValue Then
        mValue = Value
        RaiseEvent Change
    End If

End Sub

Private Sub vThumb_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Redraw
End Sub

Private Sub UserControl_Initialize()
    mMin = 1
    mMax = 10
    mValue = 1
    
    Dim mix As Long
    Dim btnFace As Long
    btnFace = GetSysColor(&HF&)
    
    vScroll.BackColor = ((btnFace \ 2) And &H7F7F7F) + ((vbWhite \ 2) And &H7F7F7F)

End Sub

Private Sub UserControl_Resize()
  '  On Error Resume Next
    
    vScroll.Width = UserControl.ScaleWidth
    btnUp.Width = UserControl.ScaleWidth
    btnDown.Width = UserControl.ScaleWidth
    vThumb.Width = UserControl.ScaleWidth
    btnDown.Top = UserControl.ScaleHeight - btnDown.ScaleHeight
    mRange = UserControl.ScaleHeight - btnDown.ScaleHeight - btnUp.ScaleHeight
    mThumbWdh = mRange / ((mMax + 1) - mMin)
    If mThumbWdh < 20 Then
        mThumbWdh = 20
    End If
    vThumb.Height = mThumbWdh
    
    mRange = UserControl.ScaleHeight - btnUp.ScaleHeight - btnDown.ScaleHeight - mThumbWdh
    
    If (mMax - mMin) = 0 Then Exit Sub
    mStep = mRange / (mMax - mMin)
    
      
    Call Redraw
    Clear vScroll
    RaiseCtl vThumb, ""
    RaiseCtl btnUp, "up"
    RaiseCtl btnDown, "down"
    DrawHandle vThumb
End Sub

Private Sub Redraw()
   ' mPercent = mValue / (Max + 1 - mMin)
    vThumb.Top = (mValue - mMin) * mStep + btnUp.ScaleHeight ' (mRange * mPercent) + btnUp.Scaleheight
    vScroll.Refresh
End Sub


Private Sub RaiseCtl(ctl As Control, image As Variant)
    Dim rc As RECT
    ctl.Cls
    rc.Top = 0
    rc.Left = 0
    rc.Right = ctl.ScaleWidth
    rc.Bottom = ctl.ScaleHeight
    If image <> "" Then
        gfx.ListImages(image).draw ctl.hdc, Int(ctl.ScaleHeight / 2 - gfx.ImageHeight / 2), Int(ctl.ScaleWidth / 2 - gfx.ImageWidth / 2), 1
    End If
    DrawEdge ctl.hdc, rc, EDGE_RAISED, BF_RECT
    ctl.Refresh
End Sub

Private Sub SinkCtl(ctl As Control, image As Variant)
    Dim rc As RECT
    ctl.Cls
    rc.Top = 0
    rc.Left = 0
    rc.Right = ctl.ScaleWidth
    rc.Bottom = ctl.ScaleHeight
    If image <> "" Then
        gfx.ListImages(image).draw ctl.hdc, Int(ctl.ScaleHeight / 2 - gfx.ImageHeight / 2) + 1, Int(ctl.ScaleWidth / 2 - gfx.ImageWidth / 2) + 1, 1
    End If
    DrawEdge ctl.hdc, rc, BDR_RAISEDINNER, BF_RECT
    ctl.Refresh
End Sub

Public Property Let Value(vData As Variant)
    If vData <> mValue Then
        mValue = vData
        Call Redraw
        RaiseEvent Change
    End If
End Property

Public Property Get Value() As Variant
    Value = mValue
End Property


Public Property Let Min(vData As Variant)
    mMin = vData
    Call UserControl_Resize
End Property

Public Property Get Min() As Variant
    Min = mMin
End Property

Public Property Let Max(vData As Variant)
    mMax = vData
    
    Call UserControl_Resize
End Property

Public Property Get Max() As Variant
    Max = mMax
End Property

Private Sub DrawHandle(ctl As Control)
    Dim rc As RECT
    rc.Left = 5
    rc.Right = ctl.ScaleWidth - 5
    
    rc.Top = Int(ctl.ScaleHeight / 2) - 5
    rc.Bottom = Int(ctl.ScaleHeight / 2) - 3
    DrawEdge ctl.hdc, rc, BDR_RAISEDINNER, BF_RECT
    
    rc.Top = Int(ctl.ScaleHeight / 2) - 1
    rc.Bottom = Int(ctl.ScaleHeight / 2) + 1
    DrawEdge ctl.hdc, rc, BDR_RAISEDINNER, BF_RECT
    
    rc.Top = Int(ctl.ScaleHeight / 2) + 3
    rc.Bottom = Int(ctl.ScaleHeight / 2) + 5
    DrawEdge ctl.hdc, rc, BDR_RAISEDINNER, BF_RECT
    
    ctl.Refresh
End Sub

Private Sub Clear(ctl As Control)
'    Dim rc As RECT
'    rc.Top = 20
'    rc.Left = Int(ctl.ScaleWidth / 2)
'    rc.Right = Int(ctl.ScaleWidth / 2) + 2
'    rc.Bottom = ctl.ScaleHeight - 20
'    DrawEdge ctl.hdc, rc, BDR_RAISED, BF_RECT
'    ctl.Refresh
End Sub

Private Sub StartClickDown()
    DoScroll = True
    Call ClickDown
    
    Wait 0.3
    Do While DoScroll
        Call ClickDown
        DoEvents
    Loop
End Sub


Private Sub StartClickTop()
    DoScroll = True
    Call ClickTop
    
    Wait 0.3
    Do While DoScroll
        Call ClickTop
        DoEvents
    Loop
End Sub

Private Sub Wait(t)
    Dim tim As Variant
    tim = Timer
    Do While Timer < tim + t And Not Timer < tim
        DoEvents
    Loop
End Sub


VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.UserControl hScrollXL 
   Alignable       =   -1  'True
   BackColor       =   &H00C000C0&
   ClientHeight    =   5115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6795
   ScaleHeight     =   341
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   453
   ToolboxBitmap   =   "hScrollX.ctx":0000
   Begin MSComctlLib.ImageList gfx 
      Left            =   5520
      Top             =   1440
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
            Picture         =   "hScrollX.ctx":0314
            Key             =   "up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "hScrollX.ctx":0373
            Key             =   "left"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "hScrollX.ctx":03D5
            Key             =   "right"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "hScrollX.ctx":0437
            Key             =   "down"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox hScroll 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   453
      TabIndex        =   0
      Top             =   0
      Width           =   6795
      Begin VB.PictureBox hThumb 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   2400
         ScaleHeight     =   57
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   3
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox btnRight 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   6480
         ScaleHeight     =   57
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   2
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox btnLeft 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   0
         ScaleHeight     =   57
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   1
         Top             =   0
         Width           =   255
      End
   End
End
Attribute VB_Name = "hScrollXL"
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





Private Sub ClickLeft()
    Call SinkCtl(btnLeft, "left")
    If mValue > mMin Then
        mValue = mValue - 1
        RaiseEvent Change
    End If
    Call Redraw
End Sub

Private Sub ClickRight()
    Call SinkCtl(btnRight, "right")
    If mValue < mMax Then
        mValue = mValue + 1
        RaiseEvent Change
    End If
    Call Redraw
End Sub
Private Sub NoClickright()
    Call RaiseCtl(btnRight, "right")
End Sub

Private Sub NoClickLeft()
    Call RaiseCtl(btnLeft, "left")
End Sub


Private Sub btnLeft_DblClick()
    Call StartClickLeft
End Sub

Private Sub btnLeft_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call StartClickLeft
End Sub

Private Sub btnLeft_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    DoScroll = False
    Call NoClickLeft
End Sub

Private Sub btnRight_DblClick()
    Call StartClickRight
End Sub

Private Sub btnRight_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call StartClickRight
End Sub

Private Sub btnRight_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    DoScroll = False
    Call NoClickright
End Sub

Private Sub hScroll_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Current As Long
    Dim Diff As Long
    Dim Value As Currency
    If Not Button = vbLeftButton Then Exit Sub
    Current = x - hThumb.ScaleWidth / 2
    
    If Current < btnLeft.ScaleWidth Then Current = btnLeft.ScaleWidth
    If Current + hThumb.ScaleWidth > UserControl.ScaleWidth - btnRight.ScaleWidth Then Current = UserControl.ScaleWidth - btnRight.ScaleWidth - hThumb.ScaleWidth
    hThumb.Left = Current
    Value = Current - btnLeft.ScaleWidth
    Value = Round(Value / mStep + mMin)
    If Value > mMax Then Value = mMax
    If Value <> mValue Then
        mValue = Value
        RaiseEvent Change
    End If
    Call Redraw
End Sub

Private Sub hThumb_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mStart = hThumb.Left
    mOffset = x
End Sub

Private Sub hThumb_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Current As Long
    Dim Diff As Long
    Dim Value As Currency
    If Not Button = vbLeftButton Then Exit Sub
    Current = hThumb.Left
    Diff = (Current + x) - (mStart + mOffset)

    Current = mStart + Diff
    If Current < btnLeft.ScaleWidth Then Current = btnLeft.ScaleWidth
    If Current + hThumb.ScaleWidth > UserControl.ScaleWidth - btnRight.ScaleWidth Then Current = UserControl.ScaleWidth - btnRight.ScaleWidth - hThumb.ScaleWidth
    hThumb.Left = Current
    hScroll.Refresh
    Value = Current - btnLeft.ScaleWidth
    Value = Round(Value / mStep + mMin)
    If Value > mMax Then Value = mMax
    If Value <> mValue Then
        mValue = Value
        RaiseEvent Change
    End If

End Sub

Private Sub hThumb_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Redraw
End Sub

Private Sub UserControl_Initialize()
    mMin = 1
    mMax = 10
    mValue = 1
    
    Dim mix As Long
    Dim btnFace As Long
    btnFace = GetSysColor(&HF&)
    hScroll.BackColor = ((btnFace \ 2) And &H7F7F7F) + ((vbWhite \ 2) And &H7F7F7F)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    hScroll.Height = UserControl.ScaleHeight
    btnLeft.Height = UserControl.ScaleHeight
    btnRight.Height = UserControl.ScaleHeight
    hThumb.Height = UserControl.ScaleHeight
    btnRight.Left = UserControl.ScaleWidth - btnRight.ScaleWidth
    mRange = UserControl.ScaleWidth - btnRight.ScaleWidth - btnLeft.ScaleWidth
    mThumbWdh = mRange / ((mMax + 1) - mMin)
    If mThumbWdh < 20 Then
        mThumbWdh = 20
    End If
    hThumb.Width = mThumbWdh
    
    mRange = UserControl.ScaleWidth - btnLeft.ScaleWidth - btnRight.ScaleWidth - mThumbWdh
    If (mMax - mMin) = 0 Then Exit Sub
    mStep = mRange / (mMax - mMin)
    
      
    Call Redraw
    
    RaiseCtl hThumb, ""
    RaiseCtl btnLeft, "left"
    RaiseCtl btnRight, "right"
    DrawHandle hThumb
End Sub

Private Sub Redraw()
   ' mPercent = mValue / (Max + 1 - mMin)
    hThumb.Left = (mValue - mMin) * mStep + btnLeft.ScaleWidth ' (mRange * mPercent) + btnLeft.ScaleWidth
    hScroll.Refresh
End Sub


Private Sub RaiseCtl(ctl As Control, image As Variant)
    Dim rc As RECT
    ctl.Cls
    rc.Left = 0
    rc.Top = 0
    rc.Right = ctl.ScaleWidth
    rc.Bottom = ctl.ScaleHeight
    If image <> "" Then
        gfx.ListImages(image).draw ctl.hdc, Int(ctl.ScaleWidth / 2 - gfx.ImageWidth / 2), Int(ctl.ScaleHeight / 2 - gfx.ImageHeight / 2), 1
    End If
    DrawEdge ctl.hdc, rc, EDGE_RAISED, BF_RECT
    ctl.Refresh
End Sub

Private Sub SinkCtl(ctl As Control, image As Variant)
    Dim rc As RECT
    ctl.Cls
    rc.Left = 0
    rc.Top = 0
    rc.Right = ctl.ScaleWidth
    rc.Bottom = ctl.ScaleHeight
    If image <> "" Then
        gfx.ListImages(image).draw ctl.hdc, Int(ctl.ScaleWidth / 2 - gfx.ImageWidth / 2) + 1, Int(ctl.ScaleHeight / 2 - gfx.ImageHeight / 2) + 1, 1
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
    rc.Top = 5
    rc.Bottom = ctl.ScaleHeight - 5
    
    rc.Left = Int(ctl.ScaleWidth / 2) - 5
    rc.Right = Int(ctl.ScaleWidth / 2) - 3
    DrawEdge ctl.hdc, rc, BDR_RAISEDINNER, BF_RECT
    
    rc.Left = Int(ctl.ScaleWidth / 2) - 1
    rc.Right = Int(ctl.ScaleWidth / 2) + 1
    DrawEdge ctl.hdc, rc, BDR_RAISEDINNER, BF_RECT
    
    rc.Left = Int(ctl.ScaleWidth / 2) + 3
    rc.Right = Int(ctl.ScaleWidth / 2) + 5
    DrawEdge ctl.hdc, rc, BDR_RAISEDINNER, BF_RECT
    
    ctl.Refresh
End Sub


Private Sub StartClickRight()
    DoScroll = True
    Call ClickRight
    
    Wait 0.3
    Do While DoScroll
        Call ClickRight
        DoEvents
    Loop
End Sub


Private Sub StartClickLeft()
    DoScroll = True
    Call ClickLeft
    
    Wait 0.3
    Do While DoScroll
        Call ClickLeft
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


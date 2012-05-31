VERSION 5.00
Begin VB.UserControl HexEd 
   Alignable       =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16170
   KeyPreview      =   -1  'True
   ScaleHeight     =   416
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1078
   ToolboxBitmap   =   "UserControl1.ctx":0000
   Begin rhexed.hScrollXL hScrollAscii 
      Height          =   255
      Left            =   8880
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5040
      Width           =   2175
      _ExtentX        =   2990
      _ExtentY        =   450
   End
   Begin rhexed.vScrollXL vScroll 
      Height          =   5055
      Left            =   11040
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   8916
   End
   Begin VB.PictureBox picFiller 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   11040
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5040
      Width           =   255
   End
   Begin VB.PictureBox Canvas 
      Align           =   3  'Align Left
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6240
      Left            =   1215
      ScaleHeight     =   416
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   7680
      Begin rhexed.hScrollXL hScrollCanvas 
         Height          =   255
         Left            =   0
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   5040
         Width           =   7695
         _ExtentX        =   11880
         _ExtentY        =   450
      End
   End
   Begin VB.PictureBox Ascii 
      Align           =   3  'Align Left
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6240
      Left            =   8895
      ScaleHeight     =   416
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   133
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   1995
   End
   Begin VB.PictureBox Margin 
      Align           =   3  'Align Left
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6240
      Left            =   0
      ScaleHeight     =   416
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1215
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuCopy2 
         Caption         =   "Copy (Ctrl+C)"
      End
      Begin VB.Menu mnuCopyHex2 
         Caption         =   "Copy Hex Codes (F4)"
      End
      Begin VB.Menu mnuSearch2 
         Caption         =   "Search (Ctrl+F)"
      End
      Begin VB.Menu mnuStrings 
         Caption         =   "Strings"
      End
      Begin VB.Menu mnuGoto2 
         Caption         =   "Goto (Ctrl+G)"
      End
      Begin VB.Menu mnuShowBookMarks2 
         Caption         =   "Show BookMarks (F3)"
      End
      Begin VB.Menu mnuHelp2 
         Caption         =   "Help (F1)"
      End
   End
End
Attribute VB_Name = "HexEd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetTempFileName Lib "kernel32" _
 Alias "GetTempFileNameA" (ByVal lpszPath As String, _
 ByVal lpPrefixString As String, ByVal wUnique As Long, _
 ByVal lpTempFileName As String) As Long

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long



Dim CharSet(255) As Byte        'charset filter
Dim HexLookup(255) As String
Dim mEditMode As Long           'edit mode , hex / ascii

'----------------------------------------------------------------Carrets / positions / sizes
Dim mPos As Long                'pos in the file
Dim mColumns As Long            'num columns
Dim mHexWidth As Long           'width of a hex block ... eg 00_  or ff_
Dim mAsciiWidth As Long         'width of a ascii block ... eg  A   or  .
Dim mLinenumberSize As Long     'num of digits in the line number
Dim mLineHeight As Long         'height of one line
Dim mSelectedPos As Long        'carret pos
Dim mSelStart As Long           'sel start
Dim mSelEnd As Long             'sel end
Dim mSelectedCursorPos As Long  'carret small pos
Dim mCanvasOffset As Long       'canvas x offset (canvas scrollbar)
Dim mAsciiOffset As Long        'ascuu x offset
Dim mCanvasMaxWidth As Long     'the width that is required to draw a complete line

'----------------------------------------------------------------GFX
Dim dcCanvas As VirtualDC       'dc for drawing
Dim dcAscii As VirtualDC
Dim dcMargin As VirtualDC

Private mModColor As Long
Private mOddColor As Long
Private mEvenColor As Long
Private mAsciiColor As Long


'----------------------------------------------------------------undo
Dim mUndoBuffer As Collection   'undo buffer
'----------------------------------------------------------------bookmarks
Private mBookmarks As Collection
Private mBookmarkPos As Long

'----------------------------------------------------------------Scrolling
Private mAutoScroll As Boolean
Private mScrolling As Boolean
Private mDirection As Integer

'----------------------------------------------------------------Files
Private mFileHandler As File
Private mIsDirty As Boolean
Private DrawCount As Long
Private KeyCount As Long

'dzzie
'------------------
Public ReadOnly As Boolean
Public ForceMemOnlyLoading As Boolean
Public UseDefaultRightClickMenu As Boolean
Public AdjustBaseOffset As Long
Private mVisibleLines As Long
Private mClipboard As New CClipboard
Private Const mClipFormat = "Hexed Binary"
Private mClipFormatID As Long
'-------------------

Public Event Dirty()
Public Event RightClick()
Public Event Loaded()
Public Event Saved()

Private mLoadedFile As String

Property Get CurrentPosition() As Long
    CurrentPosition = mPos
End Property

Property Get VisibleLines() As Long
        VisibleLines = mVisibleLines
End Property

Friend Property Let LoadedFile(v As String)
    mLoadedFile = v
End Property

Property Get LoadedFile() As String
    LoadedFile = mLoadedFile
End Property

Property Get ReadChunkSize() As Long
    ReadChunkSize = ChunkSize
End Property

Property Let ReadChunkSize(v As Long)
    ChunkSize = v + 1
End Property
    
Public Function ShowOpen(Optional initDir As String, Optional ViewOnly As Boolean = False) As Boolean
    Dim dlg As New clsCmnDlg2
    Dim fpath As String
    fpath = dlg.OpenDialog(AllFiles, initDir, "Open file", UserControl.hWnd)
    If Len(fpath) = 0 Then Exit Function
    ShowOpen = Me.LoadFile(fpath, ViewOnly)
End Function

Public Function ShowAbout()
    frmAbout.ShowAbout
End Function

Public Function ShowHelp()
    frmAbout.ShowHelp
End Function


Public Function Search(match As String, Optional isUnicode As Boolean = False, Optional caseSensitive As Boolean = False) As String()
        'this doesnt really do the rest of the library justice just my cheese implementation..dz
        Search = mFileHandler.Search(match, isUnicode, caseSensitive)
End Function

Public Sub ShowFind()
    Set frmFind.owner = Me
    frmFind.Visible = True
End Sub

'------------------

Public Property Let Columns(vData As Long)
    If vData < 1 Then vData = 1
    If vData > 32 Then vData = 32
    
    mAsciiWidth = dcAscii.CharWidth  ' Ascii.TextWidth("0")
    mHexWidth = mAsciiWidth * 3
    mLineHeight = dcAscii.CharHeight  'Ascii.TextHeight("0") ' + 02
    
    mColumns = vData
    hScrollCanvas.Min = 1
    hScrollCanvas.Max = vData
    hScrollAscii.Min = 1
    hScrollAscii.Max = vData
    vScroll.Min = 0
    vScroll.Max = CLng(Me.DataLength \ mColumns) ' - (Canvas.ScaleHeight / mLineHeight - 2)
    
    'draw
    SetvScroll Int(mPos / mColumns)
    
    mCanvasMaxWidth = mColumns * mHexWidth + 12
    Canvas.Width = mCanvasMaxWidth
    Ascii.Width = mColumns * mAsciiWidth + 20
    

    'draw
    Call UserControl_Resize
    
    'draw
    Call draw
End Property

Public Property Get Columns() As Long
    Columns = mColumns
End Property


Private Sub Ascii_GotFocus()
    mEditMode = 1
End Sub





Private Sub Ascii_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Pos As Long
    Dim xx As Long
    Dim yy As Long
    
    mEditMode = 1
    
    
    x = x + mAsciiOffset * mAsciiWidth
    
    xx = (x - mAsciiWidth - 3) / mAsciiWidth
    yy = (y - mLineHeight / 3) / mLineHeight
    
    If xx < 0 Then xx = 0
    If yy < 0 Then yy = 0
    If xx > mColumns - 1 Then xx = mColumns - 1
    If yy > Canvas.ScaleHeight / mLineHeight - 2 Then yy = Canvas.ScaleHeight / mLineHeight - 2
    
    Pos = xx + mColumns * yy + mPos
   ' mHighlightedPos = pos
    If Button = vbLeftButton Then
        
        mSelectedPos = Pos
        mSelectedCursorPos = 0
        mSelStart = mSelectedPos
        mSelEnd = mSelectedPos

    End If
    mAutoScroll = True
    mScrolling = False
    draw
End Sub

Private Sub Ascii_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Pos As Long
    Dim xx As Long
    Dim yy As Long
    
    If Not mAutoScroll Then Exit Sub
    
    If y < 0 Then
        mDirection = -1 'up
        If Not mScrolling Then Call DoAutoscroll: Exit Sub
    ElseIf y > Ascii.ScaleHeight Then
        mDirection = 1 'down
        If Not mScrolling Then Call DoAutoscroll: Exit Sub
    Else
        mDirection = 0
    End If
    
    x = x + mAsciiOffset * mAsciiWidth
    
    xx = (x - mAsciiWidth - 3) / mAsciiWidth
    yy = (y - mLineHeight / 3) / mLineHeight
    
    If xx < 0 Then xx = 0
    If yy < 0 Then yy = 0
    If xx > mColumns Then xx = mColumns
    If yy > Canvas.ScaleHeight / mLineHeight - 2 Then yy = Canvas.ScaleHeight / mLineHeight - 2
    
    Pos = xx + mColumns * yy + mPos
   ' mHighlightedPos = pos
    If Button = vbLeftButton Then
        'mSelectedPos = mHighlightedPos
        mSelectedCursorPos = 0
        mSelEnd = Pos
        If mDirection = 0 Then Call draw
    End If
    
    
End Sub

Private Sub ScrollDown()
    If vScroll.Value < vScroll.Max Then
        
        If mSelEnd <= mFileHandler.size - mColumns Then
            mSelEnd = mSelEnd + mColumns
        End If
        vScroll.Value = vScroll.Value + 1
    End If
End Sub

Private Sub ScrollUp()
    If vScroll.Value > 0 Then
        
        If mSelEnd >= 0 + mColumns Then
            mSelEnd = mSelEnd - mColumns
        End If
        vScroll.Value = vScroll.Value - 1
    End If
End Sub


Private Sub Ascii_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mAutoScroll = False
    mScrolling = False
    
    If Button = vbRightButton Then
        If UseDefaultRightClickMenu Then
            PopupMenu mnuPopup
        Else
            RaiseEvent RightClick
        End If
    End If
End Sub

Private Sub Canvas_GotFocus()
    mEditMode = 0
End Sub

Public Sub scrollTo(ByVal Pos As Long)
    
    Dim lines As Long
    Dim newBottom As Long
    Dim Max As Long
    
    If CurrentPosition < Pos Then 'scrollto will put offset at very bottom of screen..lets adjust it up some..
        
        lines = CLng(VisibleLines / 2) 'shoot for display in middle of screen
        newBottom = Pos + (lines * Columns)
        Max = FileSize
        
        If newBottom > Max Then newBottom = Max
        SetPos newBottom, 0
        SelStart = Pos
        SelLength = 0
    Else
        Call SetPos(Pos, 0)
    End If

End Sub

Private Sub SetPos(Pos As Long, Shift As Integer, Optional CarretPos As Long = 0)
    Dim yTop As Currency
    Dim yNow As Currency
    Dim yLast As Currency
    Dim xTmp As Currency
    
    
    
    If Pos < 0 Then Pos = 0
    If Pos > Me.DataLength + 1 Then Pos = Me.DataLength + 1
    
    'mSelectedCursorPos = 0
    mSelectedCursorPos = CarretPos
    If Shift = 0 Then
        mSelectedPos = Pos
    End If
    
    
    GetXYfromPos mPos, xTmp, yTop
    GetXYfromPos Pos, xTmp, yNow
    yLast = yTop + Int(((Canvas.ScaleHeight - mLineHeight * 2) / mLineHeight))
    
    
    If Shift = 0 Then
        mSelStart = Pos
        mSelEnd = Pos
    Else
        mSelEnd = Pos
    End If
    
    
    If yNow < yTop Then
        mPos = yNow * mColumns
        vScroll.Value = mPos / mColumns
    ElseIf yNow > yLast Then
        mPos = (yTop + (yNow - yLast)) * mColumns
        vScroll.Value = mPos / mColumns
    Else
        Call draw
    End If

End Sub



Private Sub SetvScroll(val As Long)
     vScroll.Value = val
End Sub

Private Sub Canvas_KeyPress(KeyAscii As Integer)

    If ReadOnly Then Exit Sub
    
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
    Dim Value As Byte
    
    If mSelectedPos > Me.DataLength Then Exit Sub
    
    If InStr("0123456789abcdef", Chr(KeyAscii)) Then
        If mSelectedCursorPos = 0 Then
            Value = mFileHandler.data(mSelectedPos)
            Value = Value And &HF
            Value = Value Or (val("&h" & Chr(KeyAscii)) * 16)
            ChangeData Value, mSelectedPos
            SetPos mSelectedPos, 0, 1
        Else
            Value = mFileHandler.data(mSelectedPos)
            Value = Value And &HF0
            Value = Value Or (val("&h" & Chr(KeyAscii)))
            ChangeData Value, mSelectedPos
            SetPos mSelectedPos + 1, 0
        End If
    Else
        KeyAscii = 0
    End If
    
    
End Sub

Private Sub Ascii_KeyPress(KeyAscii As Integer)
    
    If ReadOnly Then Exit Sub
    
    Dim Value As Byte
    If KeyAscii = vbKeyBack Or KeyAscii < &H20 Then 'codes < &h20 can be ctrl+C etc, dont want these to insert dz
        KeyAscii = 0 'cancel keypress..
    Else
         Value = mFileHandler.data(mSelectedPos)
         ChangeData KeyAscii, mSelectedPos
         SetPos mSelectedPos + 1, 0
    End If
End Sub


Private Sub Canvas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim Pos As Long
    Dim xx As Long
    Dim yy As Long
    
    mEditMode = 0
    
    x = x + mCanvasOffset * mHexWidth
    
    xx = (x - mHexWidth / 2) / mHexWidth
    
    'xx = (x - mHexWidth / 2) / mHexWidth
    yy = (y - mLineHeight / 3) / mLineHeight
    
    If xx < 0 Then xx = 0
    If yy < 0 Then yy = 0
    If xx > mColumns - 1 Then xx = mColumns - 1
    If yy > Canvas.ScaleHeight / mLineHeight - 2 Then yy = Canvas.ScaleHeight / mLineHeight - 2
    
    
    Pos = xx + mColumns * yy + mPos
   ' mHighlightedPos = pos
    If Button = vbLeftButton Then
        mSelectedPos = Pos
        mSelectedCursorPos = 0
        mSelEnd = mSelectedPos
        mSelStart = mSelectedPos

    End If
    
    mScrolling = False
    mAutoScroll = True
    draw
End Sub

Private Sub Canvas_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim Pos As Long
    Dim xx As Long
    Dim yy As Long
    
    
    If Not mAutoScroll Then Exit Sub
    
    If y < 0 Then
        mDirection = -1 'up
        If Not mScrolling Then Call DoAutoscroll: Exit Sub
    ElseIf y > Canvas.ScaleHeight Then
        mDirection = 1 'down
        If Not mScrolling Then Call DoAutoscroll: Exit Sub
    Else
        mDirection = 0
    End If
    
    x = x + mCanvasOffset * mHexWidth
    
    'xx = x / mHexWidth
    xx = (x - mHexWidth / 4) / mHexWidth
    yy = (y - mLineHeight / 3) / mLineHeight
    
    If xx < 0 Then xx = 0
    If yy < 0 Then yy = 0
    If xx > mColumns Then xx = mColumns
    If yy > Canvas.ScaleHeight / mLineHeight - 2 Then yy = Canvas.ScaleHeight / mLineHeight - 2
    
    
    Pos = xx + mColumns * yy + mPos
    If Pos > Me.DataLength + 1 Then Pos = Me.DataLength + 1
    If Pos < 0 Then Pos = 0
    If Button = vbLeftButton Then
        mSelectedCursorPos = 0
        mSelEnd = Pos
        If mDirection = 0 Then Call draw
    End If
End Sub


Private Sub Canvas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mScrolling = False
    mAutoScroll = False
    
    If Button = vbRightButton Then
        If UseDefaultRightClickMenu Then
            PopupMenu mnuPopup
        Else
            RaiseEvent RightClick
        End If
    End If
    
End Sub

Private Sub Command1_Click()
    Dim t As Variant
    Dim i As Long
    t = Timer
    
    For i = 0 To 100
        Call draw
    Next
    
    MsgBox Timer - t
    
    'orginal kod ger: 1.57
    'utan modified text : 1.29
    'optimerad-1 : 1.15
    '-"-      -2 : 1.14
    '         -3 : 0.9
    '         -4 : 0.87
    '         -5 : 0.70
    'full render6: 0.70
End Sub

Private Sub hScrollAscii_GotFocus()
Call ExitScrollFocus
End Sub

Private Sub hScrollCanvas_Change()
    Call hScrollCanvas_Scroll
End Sub

Private Sub hScrollCanvas_GotFocus()
Call ExitScrollFocus
End Sub

Private Sub hScrollCanvas_Scroll()
    mCanvasOffset = hScrollCanvas.Value - 1
    draw
End Sub

Private Sub hScrollAscii_Change()
    Call hScrollAscii_Scroll
End Sub

Private Sub hScrollAscii_Scroll()
    mAsciiOffset = hScrollAscii.Value - 1
    draw
End Sub









Private Sub mnuCopy2_Click()
    Me.DoCopy
End Sub

Private Sub mnuCopyHex2_Click()
    Clipboard.Clear
    Clipboard.SetText Me.SelTextAsHexCodes
End Sub

Private Sub mnuGoto2_Click()
    Me.ShowGoto
End Sub

Private Sub mnuHelp2_Click()
    Me.ShowHelp
End Sub

Private Sub mnuSaveAs_Click()
    Me.SaveAs
End Sub

Private Sub mnuSearch2_Click()
    Me.ShowFind
End Sub

Private Sub mnuShowBookMarks2_Click()
    Me.ShowBookMarks
End Sub

Private Sub mnuStrings_Click()
    
    On Error Resume Next
    
    Const minLen = 7
    Dim Ascii() As String
    Dim uni() As String
    
    Ascii() = Search("[\w0-9 /?.\-_=+$\\@!*\(\)#%~`\^&\|\{\}\[\]:;'""<>\,]{" & minLen & ",}")
    uni() = Search("([\w0-9 /?.\-=+$\\@!\*\(\)#%~`\^&\|\{\}\[\]:;'""<>\,][\x00]){" & minLen & ",}")
    
    frmOffsetList.LoadList Me, Ascii
    frmOffsetList.LoadList Me, uni 'this will append the data...
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If Ambient.UserMode = True Then
        InstallSubclass Me
    End If
End Sub

Private Sub vScroll_Change()
    
    mPos = vScroll.Value * mColumns
    draw
    
End Sub

Private Sub DoAutoscroll()
        mScrolling = True
        Do While mAutoScroll
            
            If mDirection = -1 Then
                KeyCount = KeyCount + 1
                ScrollUp
                
            ElseIf mDirection = 1 Then
                KeyCount = KeyCount + 1
                ScrollDown
                
            End If
            DoEvents
        Loop
End Sub

Private Sub UserControl_EnterFocus()
On Error Resume Next
    If mEditMode = 0 Then
        Canvas.SetFocus
    Else
        Ascii.SetFocus
    End If
End Sub



Private Sub UserControl_GotFocus()
On Error Resume Next
    If mEditMode = 0 Then
        Canvas.SetFocus
    Else
        Ascii.SetFocus
    End If

End Sub

Private Sub UserControl_Initialize()
    Dim i As Long
    
    mClipFormatID = mClipboard.AddFormat(mClipFormat)
    UseDefaultRightClickMenu = True
    ChunkSize = 30000
    Call InitCharset

    Set mUndoBuffer = New Collection
    Set mBookmarks = New Collection
    
    Set dcCanvas = New VirtualDC
    Set dcMargin = New VirtualDC
    Set dcAscii = New VirtualDC
    
    Set mFileHandler = New File
    Set mFileHandler.owner = Me
    
    dcCanvas.CreateFromPBOX Canvas
    dcAscii.CreateFromPBOX Ascii
    dcMargin.CreateFromPBOX Margin
    
    Me.AsciiColor = vbBlack
    Me.OddColor = RGB(0, 0, 128)
    Me.EvenColor = vbBlue
    Me.MarginColor = vbBlack
    Me.ModColor = vbRed
    
    mPos = 0
    
    mLinenumberSize = 10
    Me.Columns = 16

End Sub

Private Sub RefreshSettings()
    Columns = Columns
    Call UserControl_Resize
    Columns = Columns
End Sub



Private Sub draw()
    DrawCount = DrawCount + 1
    mFileHandler.ActivateChunk mPos
    
    dcCanvas.Cls
    dcAscii.Cls
    dcMargin.Cls
    
    Call DrawBookmarks
    Call DrawNormal
    Call DrawGuides
    Call DrawSelection
    Call DrawCursors

    Call Redraw
    
End Sub

Private Sub DrawNormal()
    Dim lines As Long               'number of lines to be drawn
    Dim i As Long                   'counter (outer loop)
    Dim j As Long                   'counter (inner loop)
    Dim Pos As Long                 'pos in the file
    Dim data As String              'temp var for hexvalues
    Dim HexLine As String           'templine with even column hex data
    Dim OddHexline As String        'templine with odd column hex data
    Dim AsciiLine As String         'temp line with Ascii data
    Dim CanvasOffset As Long        'Hex canvas hScrollpos
    Dim AsciiOffset As Long         'ascii canvas hScrollpos
    Dim buff() As Byte              'buffer with the data for the entire screen
    Dim StatusBuff() As Byte        'buffer with the STATUS data for the entire screen
    Dim BuffPos As Long             'position in the screen buffer
    Dim BigHexLine As String
    Dim BigOddHexline As String
    Dim BigPos As Long
    Dim BigAsciiLine As String
    Dim BigAsciiPos As Long
    Dim DiffLine As String
    Dim DiffAsciiLine As String
    Dim modLine As Boolean
    Dim yPos As Long
    Static cleartime As Variant
    Static drawtime As Variant
    Static gettime As Variant
    Dim tmpTime As Variant
    
    
    
    
    CanvasOffset = mCanvasOffset * mHexWidth
    AsciiOffset = mAsciiOffset * mAsciiWidth
    
    tmpTime = Timer
'    dcCanvas.Cls
'    dcAscii.Cls
'    dcMargin.Cls
    
    dcMargin.FillArea Margin.ScaleWidth - 2, 0, Margin.ScaleWidth - 1, Margin.ScaleHeight, 0
    dcCanvas.FillArea Canvas.ScaleWidth - 1, 0, Canvas.ScaleWidth, Canvas.ScaleHeight, 0
    cleartime = cleartime + Timer - tmpTime
    
    
    
    lines = Canvas.ScaleHeight / mLineHeight
    Pos = mPos
    mVisibleLines = lines
    
    tmpTime = Timer
    buff() = mFileHandler.DataScreen(Pos, mColumns * (lines + 1))
    StatusBuff() = mFileHandler.StatusScreen(Pos, mColumns * (lines + 1))
    gettime = gettime + Timer - tmpTime
    BuffPos = 0
    
    BigHexLine = Space((mColumns * 3 + 2) * (lines + 1))
    BigOddHexline = BigHexLine
    BigAsciiLine = Space((mColumns + 2) * (lines + 1))
    BigPos = 1
    BigAsciiPos = 1
    
    dcCanvas.ForeColor = mModColor
    dcAscii.ForeColor = mModColor
    tmpTime = Timer
    For i = 0 To lines
        
        If Pos > Me.DataLength Then
            Exit For
        End If
        
        yPos = i * mLineHeight
        
        HexLine = Hex(Pos + AdjustBaseOffset) 'AdjustedBaseOffset is for loading of memory images..
        HexLine = String(mLinenumberSize - Len(HexLine), "0") & HexLine
        
        dcMargin.PrintText HexLine, 5, i * mLineHeight, Margin.ScaleWidth, i * mLineHeight + mLineHeight, 0
        
        HexLine = Space(mColumns * 3)
        OddHexline = Space(mColumns * 3)
        
        DiffLine = Space(mColumns * 3)
        modLine = False
        
        AsciiLine = Space(mColumns)
        DiffAsciiLine = Space(mColumns)
        
        For j = 0 To mColumns - 1
            
            If Pos > Me.DataLength Then
                Exit For
            End If
            
            data = HexLookup((buff(BuffPos)))
            
            If StatusBuff(BuffPos) = 1 Then
                'do if modified
                Mid(DiffLine, j * 3 + 1, 2) = data
                Mid(DiffAsciiLine, j + 1, 1) = Chr(CharSet(buff(BuffPos)))
                modLine = True
                
            Else
                'do if not modified
                If (j + 1) Mod 2 Then
                    Mid(OddHexline, j * 3 + 1, 2) = data
                Else
                    Mid(HexLine, j * 3 + 1, 2) = data
                End If
                Mid(AsciiLine, j + 1, 1) = Chr(CharSet(buff(BuffPos)))
            End If
            
            
            Pos = Pos + 1
            BuffPos = BuffPos + 1
            
        Next
        
        If modLine Then
             dcCanvas.PrintText DiffLine, 8 - CanvasOffset, yPos, Columns * mHexWidth - CanvasOffset + 8, yPos + mLineHeight, 0
             DiffAsciiLine = Replace(DiffAsciiLine, "&", "&&")
             dcAscii.PrintText DiffAsciiLine, 8 - AsciiOffset, yPos, Columns * mAsciiWidth - AsciiOffset + 8, yPos + mLineHeight, 0
        End If
        
        
        Mid(BigOddHexline, BigPos, mColumns * 3 + 2) = OddHexline & vbCrLf
        Mid(BigHexLine, BigPos, mColumns * 3 + 2) = HexLine & vbCrLf
        
        Mid(BigAsciiLine, BigAsciiPos, mColumns + 2) = AsciiLine & vbCrLf
        BigPos = BigPos + mColumns * 3 + 2
        BigAsciiPos = BigAsciiPos + mColumns + 2
    Next
    
    BigAsciiLine = Replace(BigAsciiLine, "&", "&&")
    dcCanvas.ForeColor = mOddColor
    dcCanvas.PrintText BigHexLine, 8 - CanvasOffset, 0, Canvas.ScaleWidth, Canvas.ScaleHeight, 0
    dcCanvas.ForeColor = mEvenColor
    dcCanvas.PrintText BigOddHexline, 8 - CanvasOffset, 0, Canvas.ScaleWidth, Canvas.ScaleHeight, 0
    dcAscii.ForeColor = mAsciiColor
    dcAscii.PrintText BigAsciiLine, 8 - AsciiOffset, 0, Ascii.ScaleWidth, Ascii.ScaleHeight, 0
    drawtime = drawtime + Timer - tmpTime
    
    
    'frmMain.Caption = drawtime & "    " & cleartime & "    " & gettime & " kc" & KeyCount & " dc" & DrawCount
End Sub

Private Sub DrawModified()
    Dim lines As Long
    Dim i As Long
    Dim j As Long
    Dim Pos As Long
    Dim data As String
    Dim HexLine As String
    Dim AsciiLine As String
    Dim CaretX As Long
    
    Dim yPos As Long
    Dim modLine As Boolean
    
    Dim CanvasOffset As Long
    Dim AsciiOffset As Long
    CanvasOffset = mCanvasOffset * mHexWidth
    AsciiOffset = mAsciiOffset * mAsciiWidth
    
    lines = Canvas.ScaleHeight / mLineHeight
    Pos = mPos
    
    dcCanvas.ForeColor = vbRed
    dcAscii.ForeColor = vbRed
    
    For i = 0 To lines
        
        If Pos > Me.DataLength Then
            Exit For
        End If
        
        yPos = i * mLineHeight

        HexLine = ""
        AsciiLine = ""
        modLine = False
        For j = 0 To mColumns - 1
            
            If Pos > Me.DataLength Then
                Exit For
            End If
            
            If mFileHandler.Status(Pos) = 1 Then
                data = HexLookup(mFileHandler.data(Pos))
                'If Len(Data) = 1 Then Data = "0" & Data
                
                HexLine = HexLine & data & " "
                AsciiLine = AsciiLine & Chr(CharSet(mFileHandler.data(Pos)))
                modLine = True
            Else
                HexLine = HexLine & "   "
                AsciiLine = AsciiLine & " "
            End If
            Pos = Pos + 1
        Next
        If modLine Then
            dcCanvas.PrintText HexLine, 8 - CanvasOffset, yPos, Columns * mHexWidth - CanvasOffset + 8, yPos + mLineHeight, 0
            AsciiLine = Replace(AsciiLine, "&", "&&")
            dcAscii.PrintText AsciiLine, 8 - AsciiOffset, yPos, Columns * mAsciiWidth - AsciiOffset + 8, yPos + mLineHeight, 0
        End If
    Next
End Sub

Private Sub DrawCursors()
    Dim x As Currency
    Dim y As Currency
    Dim xx As Long
    Dim Pos As Long
    Dim data As String
    Dim CanvasOffset As Long
    Dim AsciiOffset As Long
        
    If mSelEnd <> mSelStart Then Exit Sub
        
    CanvasOffset = mCanvasOffset * mHexWidth
    AsciiOffset = mAsciiOffset * mAsciiWidth
    
    
    If mSelectedPos >= 0 And mSelectedPos <= Me.DataLength + 1 Then
        Pos = mSelectedPos - mPos
        GetXYfromPos Pos, x, y
        xx = x * mAsciiWidth - AsciiOffset
        x = x * mHexWidth - CanvasOffset
        y = y * mLineHeight
        If y < 0 Then y = 0
        If y > Canvas.ScaleHeight Then y = Canvas.ScaleHeight
        
        If Len(data) = 1 Then data = "0" & data
        
        If mEditMode = 0 Then
            dcCanvas.FillArea x + 8 + mSelectedCursorPos * mAsciiWidth, y, x + 8 + mSelectedCursorPos * mAsciiWidth + 2, y + mLineHeight, vbBlack
            dcAscii.FillArea xx + 8, y + mLineHeight - 1, xx + mAsciiWidth + 8, y + mLineHeight + 1, vbBlack
        Else
            dcCanvas.FillArea x + 8, y + mLineHeight - 1, x + 8 + 2 * mAsciiWidth + 2, y + mLineHeight + 1, vbBlack
            dcAscii.FillArea xx + 8, y - 2, xx + 8 + 2, y + mLineHeight - 2, vbBlack
        End If
        
    End If
End Sub

Private Sub DrawBookmarks()
    Dim x As Currency
    Dim y As Currency
    Dim xx As Long
    Dim Pos As Long
    Dim data As String
    Dim CanvasOffset As Long
    Dim AsciiOffset As Long
    Dim bm As Bookmark
        
    CanvasOffset = mCanvasOffset * mHexWidth
    AsciiOffset = mAsciiOffset * mAsciiWidth
    
    For Each bm In mBookmarks
    
        Pos = bm.Pos - mPos 'mSelectedPos - mPos
        GetXYfromPos Pos, x, y
        xx = x * mAsciiWidth - AsciiOffset
        x = x * mHexWidth - CanvasOffset
        y = y * mLineHeight
        If y < 0 Then y = 0
        If y > Canvas.ScaleHeight Then y = Canvas.ScaleHeight
        
        If Len(data) = 1 Then data = "0" & data
        
        
        dcCanvas.FillArea x + 4, y, x + 4 + mHexWidth, y + mLineHeight, vbBlack
        dcCanvas.FillArea x + 5, y + 1, x + 3 + mHexWidth, y + mLineHeight - 1, bm.Color
        dcAscii.FillArea xx + 8, y, xx + 8 + mAsciiWidth, y + mLineHeight, bm.Color
    
    Next
        
        

End Sub

Private Sub DrawSelection()
    Dim ss As Long
    Dim se As Long
    Dim x1 As Currency
    Dim y1 As Currency
    Dim x2 As Currency
    Dim y2 As Currency
    Dim xx1 As Currency
    Dim xx2 As Currency
    
    Dim CanvasOffset As Long
    Dim AsciiOffset As Long
    
    If mSelEnd = mSelStart Then Exit Sub 'bailout if no selection
    
    CanvasOffset = mCanvasOffset * mHexWidth
    AsciiOffset = mAsciiOffset * mAsciiWidth
    
    If mSelStart <= mSelEnd Then
        ss = mSelStart - mPos
        se = mSelEnd - mPos
    Else
        se = mSelStart - mPos
        ss = mSelEnd - mPos
    End If
    
    GetXYfromPos ss, x1, y1
    GetXYfromPos se, x2, y2
    

    
    xx1 = x1 * mAsciiWidth + 8 - AsciiOffset
    x1 = x1 * mHexWidth + 4 - CanvasOffset

    
    xx2 = x2 * mAsciiWidth + 8 - AsciiOffset
    x2 = x2 * mHexWidth + 4 - CanvasOffset


    'fix limits
    If y1 < -10 Then y1 = -10
    If y2 < -10 Then y2 = -10
    If y1 > Canvas.ScaleHeight + 10 Then y1 = Canvas.ScaleHeight + 10
    If y2 > Canvas.ScaleHeight + 10 Then y2 = Canvas.ScaleHeight + 10
    
    If x1 < 4 Then x1 = 4
    If x2 < 4 Then x2 = 4
    
    If y1 = y2 Then 'selection is one row or less
        dcCanvas.InvertArea x1, y1 * mLineHeight, x2, y1 * mLineHeight + mLineHeight
        dcAscii.InvertArea xx1, y1 * mLineHeight, xx2, y1 * mLineHeight + mLineHeight
    Else
        dcCanvas.InvertArea x1, y1 * mLineHeight, Canvas.ScaleWidth - 1, y1 * mLineHeight + mLineHeight
        
        dcCanvas.InvertArea 4, y2 * mLineHeight, x2, y2 * mLineHeight + mLineHeight
        
        dcAscii.InvertArea xx1, y1 * mLineHeight, Ascii.ScaleWidth - 12, y1 * mLineHeight + mLineHeight
        dcAscii.InvertArea 8, y2 * mLineHeight, xx2, y2 * mLineHeight + mLineHeight
        
        If y2 - y1 > 1 Then
            dcCanvas.InvertArea 4, y1 * mLineHeight + mLineHeight, Canvas.ScaleWidth - 1, y2 * mLineHeight
            dcAscii.InvertArea 8, y1 * mLineHeight + mLineHeight, Ascii.ScaleWidth - 12, y2 * mLineHeight
        End If
    End If
    
    

End Sub



Private Sub DrawGuides()
    Dim CanvasOffset As Long
    Dim AsciiOffset As Long
    Dim HalfAsciiWidth As Long
    Dim yPos As Long
    
    
    HalfAsciiWidth = mAsciiWidth / 2
    CanvasOffset = mCanvasOffset * mHexWidth
    AsciiOffset = mAsciiOffset * mAsciiWidth
    Dim i As Long
    For i = 1 To (mColumns / 8) - 1
        dcCanvas.FillArea (i * 8) * mHexWidth - HalfAsciiWidth - CanvasOffset + 8, 0, (i * 8) * mHexWidth - HalfAsciiWidth - CanvasOffset + 9, Canvas.ScaleHeight, 0
    Next
End Sub


Private Sub Redraw()
    dcMargin.UpdatePBOX
    dcCanvas.UpdatePBOX
    dcAscii.UpdatePBOX
End Sub

'f2 = bookmark, insert = insert..

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmparr() As Byte

KeyCount = KeyCount + 1
    Select Case KeyCode
        Case vbKeyLeft, vbKeyBack
            SetPos mSelEnd - 1, Shift
        Case vbKeyRight
            SetPos mSelEnd + 1, Shift
        Case vbKeyUp
            SetPos mSelEnd - mColumns, Shift
        Case vbKeyDown
            SetPos mSelEnd + mColumns, Shift
        Case vbKeyPageDown
            KeyCode = 0
            SetPos mSelEnd + mColumns * Int(Canvas.ScaleHeight / mLineHeight - 2), Shift
        Case vbKeyPageUp
            KeyCode = 0
            SetPos mSelEnd - mColumns * Int(Canvas.ScaleHeight / mLineHeight - 2), Shift
        Case vbKeyHome
            KeyCode = 0
            SetPos 0, Shift
        Case vbKeyF1
            ShowHelp
        Case vbKeyF2
            If Shift Then
                Call ToggleBookmark(mSelectedPos)
            Else
                Call GotoNextBookmark
            End If
        Case vbKeyF3: ShowBookMarks
        Case vbKeyF4
            Clipboard.Clear
            Clipboard.SetText Me.SelTextAsHexCodes
        Case vbKeyF5: ShowAbout
        Case vbKeyEnd
            KeyCode = 0
            SetPos Me.DataLength + 1, Shift
        Case vbKeyInsert: ShowInsert
        Case vbKeyDelete: DoDelete
        Case vbKeyA: If Shift = 2 Then SelectAll
        Case vbKeyB: If Shift = 2 Then DoPasteOver
        Case vbKeyC: If Shift = 2 Then DoCopy
        Case vbKeyF: If Shift = 2 Then ShowFind
        Case vbKeyG: If Shift = 2 Then ShowGoto
        Case vbKeyS: If Shift = 2 Then Save
        Case vbKeyX: If Shift = 2 Then DoCut
        Case vbKeyV: If Shift = 2 Then DoPaste
        Case vbKeyZ: If Shift = 2 Then DoUndo
        Case vbKeyO:  If Shift = 2 Then ShowOpen , ReadOnly
    End Select
End Sub

Public Sub ShowGoto()
    frmGoto.Display Me
End Sub
Public Sub DoCopy()
    Call CopyData(Me.SelStart, Me.SelLength)
End Sub

Public Sub DoDelete()
    On Error Resume Next
    If ReadOnly Then
        'MsgBox "File was opened in readonly mode"
        Exit Sub
    End If
    Dim length As Long
    Dim start As Long
    If mSelStart < mSelEnd Then
        start = mSelStart
        length = mSelEnd - mSelStart + 1
    Else
        start = mSelEnd
        length = mSelStart - mSelEnd + 1
    End If
    If length > 1 Then length = length - 1
    mIsDirty = True
    DeleteData start, length
End Sub

Public Sub DoPasteOver()
    On Error Resume Next
     If ReadOnly Then
        'MsgBox "File was opened in readonly mode"
        Exit Sub
    End If
    
    'overwrite data at current offset with data in clipboard.
    Dim b() As Byte
    'b() = StrConv(Clipboard.GetText, vbFromUnicode, LANG_US)
    
    If mClipboard.IsDataAvailableForFormat(mClipFormatID) <> 0 Then
        mClipboard.ClipboardOpen Me.hWnd
        If Not mClipboard.GetBinaryData(mClipFormatID, b()) Then
            MsgBox "Failed to get binary data from clipboard!"
        End If
        mClipboard.ClipboardClose
    Else
        b() = StrConv(Clipboard.GetText(vbCFText), vbFromUnicode, LANG_US)
    End If
    
    If AryIsEmpty(b) Then Exit Sub
    OverWriteData Me.SelStart, b()
End Sub

Public Sub DoPaste()
    On Error Resume Next
    If ReadOnly Then
        'MsgBox "File was opened in readonly mode"
        Exit Sub
    End If

    Dim PasteData As String
    Dim b() As Byte
    'PasteData = Clipboard.GetText(vbCFText)
    
    If mClipboard.IsDataAvailableForFormat(mClipFormatID) <> 0 Then
        mClipboard.ClipboardOpen Me.hWnd
        If Not mClipboard.GetBinaryData(mClipFormatID, b()) Then
            MsgBox "Failed to get binary data from clipboard!"
        End If
        mClipboard.ClipboardClose
    Else
        PasteData = Clipboard.GetText(vbCFText)
        b() = StrConv(PasteData, vbFromUnicode, LANG_US)
    End If
 
    mIsDirty = True
    InsertData mSelectedPos, b()
End Sub

Public Sub SelectAll()
     SelStart = 0
     SelLength = FileSize + 1
End Sub

Public Function Strings(Optional minLen As Long = 7, Optional unicode As Boolean = False) As String()
            
    On Error Resume Next
    Dim tmp() As String
    
    If unicode Then
        tmp() = Search("([\w0-9 /?.\-=+$\\@!\*\(\)#%~`\^&\|\{\}\[\]:;'""<>\,][\x00]){" & minLen & ",}")
    Else
        tmp() = Search("[\w0-9 /?.\-_=+$\\@!*\(\)#%~`\^&\|\{\}\[\]:;'""<>\,]{" & minLen & ",}")
    End If
    
    Strings = tmp()
    
End Function

Public Sub DoCut()
    On Error Resume Next
    
    If ReadOnly Then
        'MsgBox "File was opened in readonly mode"
        Exit Sub
    End If
    
    Dim start As Long, length As Long
    
    Call CopyData(Me.SelStart, Me.SelLength)
    
    If mSelStart < mSelEnd Then
        start = mSelStart
        length = mSelEnd - mSelStart + 1
    Else
        start = mSelEnd
        length = mSelStart - mSelEnd + 1
    End If
    
    If length > 1 Then length = length - 1
    mIsDirty = True
    DeleteData start, length
    
End Sub

Private Sub UserControl_Paint()
    Redraw
End Sub

Public Sub Refresh()
    UserControl_Resize
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Dim wdh As Long
    
    Margin.Width = Margin.TextWidth(String(mLinenumberSize, "0")) + 14
    Margin.Width = dcMargin.CharWidth * mLinenumberSize + 14
    
    wdh = UserControl.ScaleWidth - Margin.Width - Ascii.Width - 30
    If wdh > mCanvasMaxWidth Then wdh = mCanvasMaxWidth
    Canvas.Width = wdh
    
    
    vScroll.Top = Canvas.Top
    vScroll.Left = UserControl.ScaleWidth - vScroll.Width
    vScroll.Height = Canvas.Height - hScrollCanvas.Height
    
    hScrollCanvas.Top = Canvas.Height - hScrollCanvas.Height
    hScrollCanvas.Width = Canvas.Width
    
    hScrollAscii.Top = Ascii.Height - hScrollAscii.Height + Ascii.Top
    hScrollAscii.Width = UserControl.ScaleWidth - Margin.Width - Canvas.Width - vScroll.Width
    hScrollAscii.Left = Ascii.Left
    
    picFiller.Width = vScroll.Width
    picFiller.Height = hScrollAscii.Height
    picFiller.Top = vScroll.Top + vScroll.Height
    picFiller.Left = hScrollAscii.Width + hScrollAscii.Left
    
    'vScroll.LargeChange = Canvas.ScaleHeight / mLineHeight
    
    draw
End Sub

Private Sub UserControl_Terminate()
    '// do cleanup
    dcAscii.Destroy
    dcCanvas.Destroy
    dcMargin.Destroy
    
    UnInstallSubclass Me
    
End Sub



Private Sub vscroll_GotFocus()
    Call ExitScrollFocus
End Sub


Private Sub InitCharset()
    Dim i As Long
    For i = 0 To 255
        CharSet(i) = i
    Next
    For i = 0 To 31
        CharSet(i) = Asc(".")
    Next
    For i = &H7F To &H9F
        CharSet(i) = Asc(".")
    Next
    
    For i = 0 To 255
        HexLookup(i) = UCase(Hex(i))
        If Len(HexLookup(i)) = 1 Then HexLookup(i) = "0" & HexLookup(i)
    Next
End Sub

Private Sub AddToUndobuffer(UB As UndoBlock)
    mUndoBuffer.Add UB
    If mUndoBuffer.Count > 500 Then
        mUndoBuffer.Remove 1
    End If
End Sub

Private Sub ChangeData(ByVal Value As Byte, ByVal Pos As Long)
    'store data to undobuffer
    
    Dim UB As New UndoBlock
    UB.Action = undEdit
    UB.Pos = Pos
    UB.Status = mFileHandler.Status(Pos)
    UB.Value = mFileHandler.data(Pos)
    
    AddToUndobuffer UB
    
    'store value and mark the byte as modified
    If mFileHandler.data(Pos) <> Value Then
        'only set as modified if byte differs
        mFileHandler.Status(Pos) = 1
    End If
    mFileHandler.data(Pos) = Value
    If Not mIsDirty Then
        RaiseEvent Dirty
        mIsDirty = True 'we have changed some data , the file is now dirty..
    End If
    
End Sub





Public Sub DoUndo()
    Dim UB As UndoBlock
    
    If mUndoBuffer.Count = 0 Then Exit Sub
    
    Set UB = mUndoBuffer.Item(mUndoBuffer.Count)
    mUndoBuffer.Remove mUndoBuffer.Count
    
    'always activate the affected chunk/s
    mFileHandler.ActivateChunk UB.Pos
    
    Select Case UB.Action
        Case undEdit
            mFileHandler.data(UB.Pos) = UB.Value
            mFileHandler.Status(UB.Pos) = UB.Status
            SetPos UB.Pos, 0
            
        Case undInsert
            mFileHandler.DeleteData UB.Pos, UB.Custom
            SetPos UB.Pos, 0
            
        Case undDelete
            mFileHandler.InsertDataStatus UB.Pos, UB.Value, UB.Status
            SetPos UB.Pos, 0
    End Select
    
    Call draw
    
End Sub

Private Sub GetXYfromPos(ByRef Pos, ByRef x As Currency, ByRef y As Currency)
  y = Pos \ Columns
  x = Pos - Columns * y
End Sub

Public Property Get SelText() As String
    Dim buff() As Byte, Pos As Long, length As Long
    Pos = SelStart
    length = SelLength
    If length = 0 Then Exit Property
    If length < 0 Then
        Pos = Pos + length
        length = -length
    End If
    buff = mFileHandler.DataScreen(Pos, length)
    SelText = StrConv(buff, vbUnicode, LANG_US)
End Property

Public Property Get SelTextAsHexCodes(Optional prefix As String = Empty) As String
    SelTextAsHexCodes = toHexString(SelText, False, prefix)
End Property

Public Sub CopyData(ByVal Pos As Long, ByVal length As Long)
    Clipboard.Clear
    Dim buff() As Byte, sText As String
    If length = 0 Then Exit Sub
    If length < 0 Then
        Pos = Pos + length
        length = -length
    End If
    buff = mFileHandler.DataScreen(Pos, length)
    sText = StrConv(buff, vbUnicode, LANG_US)
    'Clipboard.SetText sText, vbCFText
    
    mClipboard.ClipboardOpen Me.hWnd
    mClipboard.ClearClipboard
    mClipboard.SetBinaryData mClipFormatID, StrConv(sText, vbFromUnicode, LANG_US)
    mClipboard.ClipboardClose

End Sub

Public Sub ShowInsert()
    
    If ReadOnly Then
        'MsgBox "File was opened in readonly mode"
        Exit Sub
    End If
    
    Dim b() As Byte
    Dim i As Long
    
    frmInsert.Show 1
    
    If frmInsert.Cancel Or frmInsert.ByteCount < 1 Then
        UnLoad frmInsert
        Exit Sub
    End If
    
    ReDim b(frmInsert.ByteCount - 1)
    
    If frmInsert.ByteValue <> 0 Then
        For i = 0 To UBound(b)
            b(i) = frmInsert.ByteValue
        Next
    End If
    
    UnLoad frmInsert
    mIsDirty = True
    InsertData mSelectedPos, b
End Sub

Public Sub InsertData(ByVal Pos As Long, data() As Byte)
    
    If ReadOnly Then
        'MsgBox "File was opened in readonly mode"
        Exit Sub
    End If
    
    Dim UB As New UndoBlock
    UB.Action = undInsert
    UB.Pos = Pos
    UB.Status = ""
    UB.Value = ""
    UB.Custom = UBound(data) + 1    'store length
    
    AddToUndobuffer UB
    mFileHandler.InsertData Pos, data
    Columns = Columns
End Sub

Public Sub OverWriteData(ByVal Pos As Long, data() As Byte)
    If ReadOnly Then
        'MsgBox "File was opened in readonly mode"
        Exit Sub
    End If
    LockWindowUpdate Me.hWnd
    DeleteData Pos, UBound(data) + 1
    InsertData Pos, data()
    LockWindowUpdate 0
End Sub

Public Sub DeleteData(ByVal Pos As Long, ByVal length As Long)
    Dim UB As New UndoBlock
        
    If ReadOnly Then
        ''MsgBox "File was opened in readonly mode"
        Exit Sub
    End If
            
    mFileHandler.ActivateChunk Pos
    
    UB.Action = undDelete
    UB.Pos = Pos
    UB.Status = mFileHandler.StatusScreen(Pos, length)
    UB.Value = mFileHandler.DataScreen(Pos, length)
    
    AddToUndobuffer UB
    mFileHandler.DeleteData Pos, length
    SetPos Pos, 0
    Columns = Columns
End Sub

Public Property Get SelStart() As Long
    SelStart = mSelStart
End Property

Public Property Let SelStart(vData As Long)
    If vData > Me.DataLength + 1 Then Exit Property
    mSelStart = vData
    mSelectedPos = vData
    Call draw
End Property

Public Property Get SelLength() As Long
    SelLength = mSelEnd - mSelStart
End Property

Public Property Let SelLength(vData As Long)
    mSelEnd = mSelStart + vData
    Call draw
End Property

Public Property Get DataLength() As Long
    DataLength = mFileHandler.size ' UBound(mData)
End Property

Private Sub ExitScrollFocus()
    If mEditMode = 0 Then
        Canvas.SetFocus
    Else
        Ascii.SetFocus
    End If
End Sub

'dzzie
'------------------
Public Function LoadString(data As String, Optional ViewOnly As Boolean = True) As Boolean
    Dim b() As Byte
    b() = StrConv(data, vbFromUnicode, LANG_US)
    LoadString = LoadByteArray(b(), ViewOnly)
End Function

 
Public Function LoadByteArray(bArray As Variant, Optional ViewOnly As Boolean = True) As Boolean
        
    If TypeName(bArray) <> "Byte()" Then
        MsgBox "LoadByteArray Error: argument is not a Byte() Array ", vbInformation
        Exit Function
    End If
    
    On Error GoTo hell
    
    Dim f As Long
    Dim path As String
    Dim b() As Byte
    Dim ch As Chunk
    
    b() = bArray
    ReadOnly = ViewOnly
    mLoadedFile = Empty
    
    If ForceMemOnlyLoading Then
        ReadOnly = True
        ReadChunkSize = UBound(b) + 1
    End If
    
    If ReadOnly And UBound(b) + 1 < ChunkSize Then
        'just view in memory only no need to create tempfile,
        'user can reset ChunkSize through ReadChunkSize property to force mem only loading...
        Set mFileHandler = Nothing
        Set mFileHandler = New File
        Set mFileHandler.owner = Me
        LockWindowUpdate Me.hWnd
        mFileHandler.LoadFromMemory b()
        vScroll.Value = 1
        Me.Columns = Me.Columns
        Me.SelStart = 0
        Me.SelLength = 0
        UserControl.vScroll.Value = 0
        mIsDirty = False
        mLoadedFile = Empty
        LockWindowUpdate 0
        Set mUndoBuffer = New Collection
        Set mBookmarks = New Collection
        LoadByteArray = True
        RaiseEvent Loaded
        Exit Function
    End If
    
    path = TempFileName(Environ("temp"))
    
    f = FreeFile
    Open path For Binary As f
    Put f, , b()
    Close f
    
    LoadByteArray = LoadFile(path, ViewOnly)
    If LoadByteArray Then RaiseEvent Loaded
    Exit Function
hell:
End Function

Function FolderExists(path As String) As Boolean
    If Len(path) = 0 Then Exit Function
    If Dir(path, vbDirectory) <> "" Then FolderExists = True
End Function

Function FileExists(path As String) As Boolean
    On Error GoTo hell
    Dim tmp As String
    tmp = path
    If Len(tmp) = 0 Then Exit Function
    If Dir(tmp, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
    Exit Function
hell: FileExists = False
End Function

Private Function TempFileName(sPath As String, Optional prefix As String = "hexed_") As String
    Dim lUnique As Long, sTempFileName As String
    If IsEmpty(sPath) Or Not FolderExists(sPath) Then sPath = Environ("temp")
    If Not FolderExists(sPath) Then sPath = Environ("tmp")
    lUnique = 0
    sTempFileName = Space$(100)
    GetTempFileName sPath, prefix, lUnique, sTempFileName
    sTempFileName = Mid$(sTempFileName, 1, InStr(sTempFileName, Chr$(0)) - 1)
    TempFileName = sTempFileName
End Function
'------------------

Public Function LoadFile(fpath As String, Optional ViewOnly As Boolean = True) As Boolean
    On Error GoTo hell
    
    Set mFileHandler = Nothing
    Set mFileHandler = New File
    Set mFileHandler.owner = Me
    
    ReadOnly = ViewOnly
    mFileHandler.Load fpath
    mLoadedFile = fpath
    
    LockWindowUpdate Me.hWnd
    vScroll.Value = 1
    Me.Columns = Me.Columns
    Me.SelStart = 0
    Me.SelLength = 0
    UserControl.vScroll.Value = 0
    mIsDirty = False
    LockWindowUpdate 0
    
    Set mUndoBuffer = New Collection
    Set mBookmarks = New Collection
    LoadFile = True
    RaiseEvent Loaded
    
    Exit Function
hell:
End Function

Public Sub AsciiView()
    Ascii.Visible = True
    Canvas.Visible = False
End Sub

Public Sub HexView()
    Ascii.Visible = False
    Canvas.Visible = True
End Sub
Public Sub FullView()
    Ascii.Visible = True
    Canvas.Visible = True
End Sub

Public Property Get IsDirty() As Boolean
    IsDirty = mIsDirty
End Property

Public Sub SaveAs(Optional fpath As String = Empty, Optional defaultfName As String)
    
    If Len(defaultfName) = 0 And Len(LoadedFile) > 0 Then
        defaultfName = FileNameFromPath(LoadedFile)
    End If
    
    mFileHandler.SaveAs fpath, defaultfName
    mIsDirty = False
    Call draw
    ReadOnly = False
    RaiseEvent Saved
    
End Sub

Public Sub Save()
    
    If Not mFileHandler.isMemLoad Then
        If ReadOnly Then
            'MsgBox "File was opened in readonly mode"
            Exit Sub
        End If
    End If

    mFileHandler.Save
    mIsDirty = False
    Call draw
    RaiseEvent Saved
    
End Sub

Public Property Set Font(vData As StdFont)
    'set font
    Set dcMargin.Font = vData
    Set dcCanvas.Font = vData
    Set dcAscii.Font = vData
    
    Set Margin.Font = dcCanvas.Font
    Set Canvas.Font = dcCanvas.Font
    Set Ascii.Font = dcCanvas.Font

    Call RefreshSettings
End Property

Public Property Get Font() As StdFont
    Set Font = Canvas.Font
End Property


'--------------------------------------------------------------------
Public Property Get ModColor() As Long
    ModColor = mModColor
End Property

Public Property Let ModColor(ByVal vData As Long)
    mModColor = vData
End Property

Public Property Get OddColor() As Long
    OddColor = mOddColor
End Property

Public Property Let OddColor(ByVal vData As Long)
    mOddColor = vData
End Property

Public Property Get EvenColor() As Long
    EvenColor = mEvenColor
End Property

Public Property Let EvenColor(ByVal vData As Long)
    mEvenColor = vData
End Property

Public Property Get AsciiColor() As Long
    AsciiColor = mAsciiColor
End Property

Public Property Let AsciiColor(ByVal vData As Long)
    mAsciiColor = vData
End Property

Public Property Get MarginColor() As Long
    MarginColor = dcMargin.ForeColor
End Property

Public Property Let MarginColor(ByVal vData As Long)
    dcMargin.ForeColor = vData
    Margin.ForeColor = vData
End Property

Public Property Get MarginBGColor() As Long
    MarginBGColor = dcMargin.BackColor
End Property

Public Property Let MarginBGColor(ByVal vData As Long)
    dcMargin.BackColor = vData
    Margin.BackColor = vData
End Property

Public Property Get HexBGColor() As Long
    HexBGColor = dcCanvas.BackColor
End Property

Public Property Let HexBGColor(ByVal vData As Long)
    dcCanvas.BackColor = vData
    Canvas.BackColor = vData
End Property

Public Property Get AsciiBGColor() As Long
    AsciiBGColor = dcAscii.BackColor
End Property

Public Property Let AsciiBGColor(ByVal vData As Long)
    dcAscii.BackColor = vData
    Ascii.BackColor = vData
    UserControl.BackColor = vData
End Property


Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get hWndCanvas() As Long
    hWndCanvas = Canvas.hWnd
End Property

Public Property Get hWndAscii() As Long
    hWndAscii = Ascii.hWnd
End Property


Public Sub Scroll(ByVal Amount As Long)
    If vScroll.Value - Amount <= vScroll.Max And vScroll.Value - Amount >= vScroll.Min Then
        vScroll.Value = vScroll.Value - Amount
       ' Call Draw
    End If
End Sub


Public Function GetData(ByVal Pos As Long) As Byte
    If Pos > mFileHandler.size Then Exit Function
    mFileHandler.ActivateChunk Pos
    GetData = mFileHandler.data(Pos)
End Function

Public Function GetDataChunk(ByVal Pos As Long) As String
    If Pos > mFileHandler.size Then Exit Function
    mFileHandler.ActivateChunk Pos
    Dim s As String
    s = mFileHandler.DataScreen(Pos, ChunkSize)
    GetDataChunk = s
End Function

Public Function FileSize() As Long
    FileSize = mFileHandler.size
End Function

Property Get BookMarks() As Collection
    Set BookMarks = mBookmarks
End Property

Public Sub ShowBookMarks()
    frmOffsetList.LoadBookMarks Me
End Sub

Public Sub ToggleBookmark(ByVal Pos As Long)
    Dim bm As Bookmark
    On Error Resume Next
    Set bm = mBookmarks("abc" & Pos)
    If Err Then
        'add bookmark
        Set bm = New Bookmark
        bm.Pos = Pos
        mBookmarks.Add bm, "abc" & Pos
        Call draw
    Else
        'remove bookmark
        mBookmarks.Remove ("abc" & Pos)
        Call draw
    End If
End Sub

Public Sub GotoNextBookmark()
    Dim bm As Bookmark
    On Error Resume Next
    If mBookmarkPos >= mBookmarks.Count Then mBookmarkPos = 0
    Set bm = mBookmarks(mBookmarkPos + 1)
    SetPos bm.Pos, 0
    mBookmarkPos = mBookmarkPos + 1
    Err.Clear
End Sub

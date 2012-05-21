VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "VbHexed v1.0"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imlToolbar 
      Left            =   1680
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0BC2
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0CD4
            Key             =   "open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0DE6
            Key             =   "save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0EF8
            Key             =   "print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":100A
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":111C
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":122E
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1340
            Key             =   "paste"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlMenu 
      Left            =   1080
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1452
            Key             =   ""
            Object.Tag             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1564
            Key             =   ""
            Object.Tag             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1676
            Key             =   ""
            Object.Tag             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1788
            Key             =   ""
            Object.Tag             =   "Copy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":189A
            Key             =   ""
            Object.Tag             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":19AC
            Key             =   ""
            Object.Tag             =   "Paste"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1ABE
            Key             =   ""
            Object.Tag             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1BD0
            Key             =   ""
            Object.Tag             =   "Undo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1CE2
            Key             =   ""
            Object.Tag             =   "Redo"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1DF4
            Key             =   ""
            Object.Tag             =   "Preferences"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            ImageKey        =   "new"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            ImageKey        =   "open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            ImageKey        =   "save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            ImageKey        =   "print"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageKey        =   "cut"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageKey        =   "copy"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageKey        =   "paste"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "undo"
            ImageKey        =   "undo"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "redo"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.PictureBox pBottom 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   0
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   792
      TabIndex        =   3
      Top             =   7740
      Width           =   11880
      Begin MSComctlLib.TabStrip tbFiles 
         Height          =   495
         Left            =   75
         TabIndex        =   4
         Top             =   -90
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   873
         TabWidthStyle   =   1
         TabFixedWidth   =   2884
         HotTracking     =   -1  'True
         Placement       =   1
         TabMinWidth     =   794
         ImageList       =   "imlIcons"
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox ptop 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   160
      Left            =   0
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   792
      TabIndex        =   2
      Top             =   720
      Width           =   11880
   End
   Begin VB.PictureBox pRight 
      Align           =   4  'Align Right
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   11715
      ScaleHeight     =   457
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   1
      Top             =   885
      Width           =   160
   End
   Begin VB.PictureBox pLeft 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   0
      ScaleHeight     =   457
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   0
      Top             =   885
      Width           =   160
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   2160
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   480
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483644
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1F06
            Key             =   "b8"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1F73
            Key             =   "b24"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1FED
            Key             =   "b32"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2068
            Key             =   "b16"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":20E4
            Key             =   "hex"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2175
            Key             =   "ascii"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2203
            Key             =   "both"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2286
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":22F7
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2367
            Key             =   "binfile"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "b8"
            ImageKey        =   "b8"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "b16"
            ImageKey        =   "b16"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "b24"
            ImageKey        =   "b24"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "b32"
            ImageKey        =   "b32"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "hex"
            ImageKey        =   "hex"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ascii"
            ImageKey        =   "ascii"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "full"
            ImageKey        =   "both"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu mnuHead 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu FileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu FileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu FileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu FileSaveAs 
         Caption         =   "Save As"
      End
   End
   Begin VB.Menu mnuHead 
      Caption         =   "&Edit"
      Index           =   1
      Begin VB.Menu EditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
   End
   Begin VB.Menu mnuHead 
      Caption         =   "&Options"
      Index           =   2
      Begin VB.Menu mnuPref 
         Caption         =   "Settings"
      End
   End
   Begin VB.Menu mnuHead 
      Caption         =   "&Tools"
      Index           =   3
      Begin VB.Menu toolsCompare 
         Caption         =   "Compare"
      End
      Begin VB.Menu toolsCompareNext 
         Caption         =   "Compare Next"
         Shortcut        =   {F4}
      End
      Begin VB.Menu toolsString 
         Caption         =   "String extractor"
      End
   End
   Begin VB.Menu mnuHead 
      Caption         =   "&Window"
      Index           =   4
      Begin VB.Menu mnuWindows 
         Caption         =   "Cascade"
         Index           =   0
      End
      Begin VB.Menu mnuWindows 
         Caption         =   "Tile Horizontal"
         Index           =   1
      End
      Begin VB.Menu mnuWindows 
         Caption         =   "Tile Vertical"
         Index           =   2
      End
      Begin VB.Menu mnuWindows 
         Caption         =   "-"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ActiveWindow As frmEditor




Private Sub EditUndo_Click()
    If Not Me.ActiveWindow Is Nothing Then
        Me.ActiveWindow.HexEditor.DoUndo
    End If
End Sub

Private Sub FileClose_Click()
    If Not Me.ActiveWindow Is Nothing Then
        UnLoad Me.ActiveWindow
        Set Me.ActiveWindow = Nothing
    End If
End Sub

Private Sub FileOpen_Click()
    LoadFile
End Sub

Private Sub LoadFile(Optional ByVal Filename As String)
    On Error GoTo hell
    Dim frm As frmEditor
    Dim key As String
    Static counter As Long
  '  Dlg.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
    If Filename = "" Then
        dlg.ShowOpen
    End If
    counter = counter + 1
    Set frm = New frmEditor
    key = "abcd" & counter
    frm.key = key
    If Filename = "" Then
        frm.Caption = dlg.FileTitle
        frm.HexEditor.Load dlg.Filename
        frm.title = dlg.FileTitle
        frm.Filename = dlg.Filename
        tbFiles.Tabs.Add , key, dlg.FileTitle & " ", "binfile"
    Else
        Dim title As String
        title = Mid(Filename, InStrRev(Filename, "\") + 1)
        frm.Caption = title
        frm.HexEditor.Load Filename
        frm.title = title
        frm.Filename = Filename
        tbFiles.Tabs.Add , key, title & " ", "binfile"
    End If
    
    Windows.Add frm, key
    tbFiles.Tabs(key).Selected = True
    Redraw
    frm.Show
    Call FixWindowsMenu
    
    Exit Sub
hell:
    Err.Clear

End Sub


Private Sub FileSave_Click()
    If Not Me.ActiveWindow Is Nothing Then
        Me.ActiveWindow.HexEditor.Save
    End If
End Sub

Private Sub MDIForm_Load()

    Me.Show
    
    
    tbFiles.Tabs.Clear
    Set Windows = New Collection
    Call Redraw
    
    
    
    If Command <> "" Then
        Dim file As String
        file = Replace(Command$, Chr(34), "")
        LoadFile file
    End If
    
End Sub

Private Sub TabStrip1_Click()
    Me.SetFocus
End Sub


Private Sub MDIForm_Resize()
    Call Redraw
End Sub

Private Sub Redraw()
On Error Resume Next
    pLeft.Cls
    pRight.Cls
    ptop.Cls
    pBottom.Cls
    PaintEdge pLeft.hdc, 4, -10, 2000, 15000
    PaintEdge pRight.hdc, 0, -10, 6, 15000
    PaintEdge ptop.hdc, 4, 4, Me.Width / Screen.TwipsPerPixelX - 13, pLeft.ScaleHeight + 10
    tbFiles.Width = Me.Width / Screen.TwipsPerPixelX - 18
    If tbFiles.Tabs.Count = 0 Then
        tbFiles.Top = -23
    Else
        tbFiles.Top = -3
    End If
    pLeft.Refresh
    pRight.Refresh
    ptop.Refresh
End Sub

Private Sub mnuFile_Click(Index As Integer)

End Sub


Private Sub mnuEdit_Click(Index As Integer)
    
    
    
End Sub

Private Sub mnuPref_Click()
    If Not Me.ActiveWindow Is Nothing Then
        frmSettings.Show 1
        UnLoad frmSettings
    End If
End Sub

Private Sub mnuWindows_Click(Index As Integer)
    Dim i As Long
    Select Case Index
        Case 2
            Call TileWindowsH
        Case 1
            Call TileWindowsV
        Case Is >= 4
            i = Index - 3
            Windows(i).SetFocus
    End Select
End Sub



Private Sub tbFiles_Click()
    Dim key As String
    Dim frm As frmEditor
    key = tbFiles.SelectedItem.key
    Set frm = Windows(key)
    
    frm.SetFocus
    
End Sub


Public Sub UnloadWindow(ByVal key As String)
    tbFiles.Tabs.Remove (key)
    If Windows(key) Is Me.ActiveWindow Then
        Set Me.ActiveWindow = Nothing
    End If
    Windows.Remove (key)
    
    
    'remove from compare list
    On Error Resume Next
    CompareWindows.Remove (key)
    Err.Clear
    On Error GoTo 0
    
    Call Redraw
    Call FixWindowsMenu
    
End Sub

Private Sub FixWindowsMenu()
    Dim i As Long
    Dim wnd As frmEditor
    For i = mnuWindows.UBound To 4 Step -1
        UnLoad mnuWindows(i)
    Next

    For i = 1 To Windows.Count
       Set wnd = Windows(i)
       Load mnuWindows.Item(3 + i)
       mnuWindows.Item(3 + i).Caption = "* " & wnd.title
    Next

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Me.ActiveWindow Is Nothing Then Exit Sub
    Select Case Button.key
        Case "b8"
            Me.ActiveWindow.HexEditor.Columns = 8
        Case "b16"
            Me.ActiveWindow.HexEditor.Columns = 16
        Case "b24"
            Me.ActiveWindow.HexEditor.Columns = 24
        Case "b32"
            Me.ActiveWindow.HexEditor.Columns = 32
        Case "hex"
            'Me.ActiveWindow.HexEditor.HexView
        Case "ascii"
            'Me.ActiveWindow.HexEditor.AsciiView
        Case "full"
           ' Me.ActiveWindow.HexEditor.FullView
    End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.key
        Case "open"
            Call FileOpen_Click
        Case "save"
            Call FileSave_Click
        Case "undo"
            Call EditUndo_Click
            
    End Select
End Sub

Public Sub TileWindowsH()
    Dim i As Long
    Dim w As Long
    Dim h As Long
    Dim wnd As frmEditor
    i = Windows.Count
    If i = 0 Then Exit Sub
    w = Me.ScaleWidth / i
    h = Me.ScaleHeight
    For i = 1 To Windows.Count
        Set wnd = Windows(i)
        wnd.WindowState = 0
        wnd.Move (i - 1) * w, 0, w, h
    Next
End Sub

Public Sub TileWindowsV()
    Dim i As Long
    Dim w As Long
    Dim h As Long
    Dim wnd As frmEditor
    i = Windows.Count
    If i = 0 Then Exit Sub
    w = Me.ScaleWidth
    h = Me.ScaleHeight / i
    For i = 1 To Windows.Count
        Set wnd = Windows(i)
        wnd.WindowState = 0
        wnd.Move 0, (i - 1) * h, w, h
    Next
End Sub

Private Sub WindowCascade_Click()

End Sub

Private Sub ToolsCompare_Click()
    frmWindows.Show vbModal
End Sub

Private Sub toolsCompareNext_Click()
    Call FindDifference
End Sub

Public Sub FindDifference()
    Dim Form1 As frmEditor
    Dim form2 As frmEditor
    Dim Form As frmEditor
    
    Dim i As Long
    Dim Diff As Boolean
    Dim Pos As Long
    Dim size As Long
    Dim title As String
    If CompareWindows Is Nothing Then Set CompareWindows = New Collection
    
    If CompareWindows.Count < 2 Then
        MsgBox "You need atleast 2 files to compare", vbOKOnly Or vbCritical
        Exit Sub
    End If
    
    size = -1
    
    For Each Form In CompareWindows
        If Form.HexEditor.FileSize < size Or size = -1 Then
            size = Form.HexEditor.FileSize
            title = Form.title
        End If
    Next
    
    Diff = False
    Do While Not Diff And DiffPos <= size + 1
        Set Form1 = CompareWindows(1)
        For i = 2 To CompareWindows.Count
            Set form2 = CompareWindows(i)
            If Form1.HexEditor.GetDataChunk(DiffPos) <> form2.HexEditor.GetDataChunk(DiffPos) Then
                Diff = True
                Exit For
            End If
        Next
        If Diff Then Exit Do
        DiffPos = DiffPos + ChunkSize
    Loop
    
    If Diff Then
        
        Diff = False
        Do While Not Diff And DiffPos <= size + 1
            Set Form1 = CompareWindows(1)
            For i = 2 To CompareWindows.Count
                Set form2 = CompareWindows(i)
                If Form1.HexEditor.GetData(DiffPos) <> form2.HexEditor.GetData(DiffPos) Then
                    Diff = True
                    Exit For
                End If
            Next
            DiffPos = DiffPos + 1
        Loop
        
        If Diff Then
        Pos = DiffPos - 1
            For Each Form In CompareWindows
                Form.HexEditor.ScrollTo Pos
                Form.HexEditor.SelStart = Pos
                Form.HexEditor.SelLength = 1
            Next
        End If
    End If
    
    If DiffPos > size Then
        DiffPos = 0
        MsgBox "End of '" & title & "' reached!", vbOKOnly Or vbInformation
    End If
End Sub

Private Sub toolsString_Click()
    Dim chars As String
    Dim IsUnicode As Boolean
    Dim str As String
    Dim Count As Long
    Dim datalen As Long
    Dim buffer() As Byte
    Dim i As Long
    Dim j As Long
    Dim char As String
    Dim StrPos As Long
    Dim last As String
    Dim li As ListItem
    
    If frmStrings.Visible Then Exit Sub
    
    frmStrings.Show
    frmStrings.lw.ListItems.Clear
    
    
    If ActiveWindow Is Nothing Then Exit Sub
    chars = "abcdefghijklmnopqrstuvwxyzåäöü"
    chars = chars & UCase(chars)
    chars = chars & "01234567890 .,;*+-/\'!""#¤%&/()?_=<>"
    
    datalen = ActiveWindow.HexEditor.FileSize
    frmStrings.pb.Max = datalen
    
    
    Do
        buffer = ActiveWindow.HexEditor.GetDataChunk(i)
        For j = 0 To UBound(buffer)
            char = Chr(buffer(j))
            If InStr(chars, char) Then
                str = str & char
                last = char
            Else
                If char = vbNullChar And last <> vbNullChar Then
                    'do nada
                    last = char
                Else
                    'reset string
                    If Len(str) > 3 Then
                        Set li = frmStrings.lw.ListItems.Add(, , Hex(StrPos))
                        li.SubItems(1) = str
                        
                        
                        'frmStrings.lstStrings.ItemData(frmStrings.lstStrings.NewIndex) = StrPos
                       ' MsgBox Str & "  " & StrPos
                    End If
                    last = char
                    str = ""
                    StrPos = i + j + 1
                End If
            End If
            
            If i + j > datalen Then Exit For
            If j Mod 1000 = 0 Then
                DoEvents
                frmStrings.pb.Value = i + j
            End If
        Next
        i = i + ChunkSize
        If i > datalen Then Exit Do
    Loop
    

    
End Sub

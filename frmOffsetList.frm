VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOffsetList 
   Caption         =   "Offsets"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7245
   Icon            =   "frmOffsetList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lv 
      Height          =   3795
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   6694
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Offset"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Data"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy All"
      End
      Begin VB.Menu mnuCopySelected 
         Caption         =   "Copy Selected"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuEditDescription 
         Caption         =   "Edit Description"
      End
      Begin VB.Menu mnuExportList 
         Caption         =   "Export List"
      End
      Begin VB.Menu mnuImportList 
         Caption         =   "Import List"
      End
   End
End
Attribute VB_Name = "frmOffsetList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private owner As HexEd
Private selLi As ListItem

Private Enum viewMode
    searchList = 0
    bookMarkList = 1
End Enum

Private vMode As viewMode

Function LoadBookMarks(he As HexEd)
    
    Dim b As Bookmark
    Dim bmks As Collection
    Dim li As ListItem
    
    vMode = bookMarkList
    On Error Resume Next
    Set owner = he
    
    Set bmks = owner.BookMarks
    
    For Each b In bmks
        Set li = lv.ListItems.Add(, , Hex(b.Pos))
        Set li.Tag = b
        li.SubItems(1) = b.Description
    Next
    
    Me.Visible = True
    Form_Load
End Function


Function LoadList(he As HexEd, data() As String)
    
    Dim tmp
    Dim li As ListItem
    Dim x
    
    On Error Resume Next
    
    vMode = searchList
    Set owner = he
    
    For Each x In data
        If Len(x) > 0 And InStr(x, ",") > 0 Then
            tmp = Split(x, ",")
            Set li = lv.ListItems.Add(, , Hex(tmp(0)))
            li.SubItems(1) = tmp(1)
            li.Tag = CLng(tmp(0))
        End If
    Next
        
    Me.Caption = lv.ListItems.Count & " occurances found.."
    Me.Visible = True
    Form_Load
    
End Function

Private Sub Form_Resize()
    On Error Resume Next
    lv.Width = Me.Width - 250
    lv.Height = Me.Height - 250
    lv.ColumnHeaders(2).Width = lv.Width - lv.ColumnHeaders(2).Left - 230
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    Dim b As Bookmark
    
    If vMode = searchList Then
        owner.scrollTo CLng(Item.Tag)
        If Len(Item.SubItems(1)) > 0 Then
            owner.SelLength = Len(Item.SubItems(1))
        End If
    ElseIf vMode = bookMarkList Then
        Set b = Item.Tag
        owner.scrollTo b.Pos
    End If
        
    Set selLi = Item
    
End Sub

Private Sub Form_Load()
    FormPos Me, True
    SetTopMost Me
    Form_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormPos Me, True, True
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        mnuEditDescription.Visible = (vMode = bookMarkList)
        PopupMenu mnuPopup
    End If
End Sub

Private Sub mnuCopy_Click()
    Dim li As ListItem
    Dim tmp As String
    For Each li In lv.ListItems
        tmp = tmp & li.Text & vbTab & li.SubItems(1) & vbCrLf
    Next
    Clipboard.Clear
    Clipboard.SetText tmp
End Sub

Private Sub mnuCopySelected_Click()
    If selLi Is Nothing Then Exit Sub
    Dim tmp As String
    tmp = selLi.Text & vbTab & selLi.SubItems(1)
    Clipboard.Clear
    Clipboard.SetText tmp
End Sub

Private Sub mnuEditDescription_Click()
    Dim b As Bookmark
    Dim s As String
    
    If selLi Is Nothing Then Exit Sub
    If vMode = bookMarkList Then
        Set b = selLi.Tag
        s = InputBox("Enter description:", , b.Description)
        If Len(s) = 0 Then Exit Sub
        b.Description = s
        selLi.SubItems(1) = s
    End If
        
End Sub

Private Sub mnuExportList_Click()
    Dim dlg As New clsCmnDlg2
    Dim f As String
    On Error Resume Next
    f = dlg.SaveDialog(AllFiles, , , , Me.hWnd, IIf(vMode = bookMarkList, "bookmarks.bml", "results.txt"))
    If Len(f) = 0 Then Exit Sub
    Dim li As ListItem
    Dim tmp As String
    For Each li In lv.ListItems
        tmp = tmp & li.Text & "," & li.SubItems(1) & vbCrLf
    Next
    WriteFile f, tmp
    Set dlg = Nothing
End Sub

Private Sub WriteFile(path As String, it As Variant)
    Dim f As Long
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
End Sub

Private Function ReadFile(filename) As Variant
  Dim f As Long
  Dim temp As Variant
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function

Private Sub ClearCollection(c As Collection)
    While c.Count > 0
        c.Remove 1
    Wend
End Sub

Private Sub mnuImportList_Click()
    Dim dlg As New clsCmnDlg2
    Dim f, x
    Dim cb As Bookmark
    Dim c As Collection
    Dim tmp
    Dim ff() As String
    
    On Error Resume Next
    
    f = dlg.OpenDialog(AllFiles, , , Me.hWnd)
    If Len(f) = 0 Then Exit Sub
    Set dlg = Nothing
    
    vMode = IIf(LCase(Right(f, 4)) = ".bml", bookMarkList, searchList)
    f = ReadFile(f)
    ff = Split(f, vbCrLf)
    
    If vMode = bookMarkList Then
        Set c = owner.BookMarks
        ClearCollection c
        For Each x In ff
            If Len(x) > 0 And InStr(x, ",") > 0 Then
                Err.Clear
                tmp = Split(x, ",")
                Set cb = New Bookmark
                cb.Pos = CLng("&h" & tmp(0))
                cb.Description = tmp(1)
                If Err.Number = 0 Then c.Add cb
            End If
        Next
        LoadBookMarks owner
        owner.Refresh
    Else
        LoadList owner, ff
    End If
            
                
End Sub




Private Sub mnuRemove_Click()
    Dim b As Bookmark
    On Error Resume Next
    
    If selLi Is Nothing Then Exit Sub
    
    If vMode = bookMarkList Then
        Set b = selLi.Tag
        owner.ToggleBookmark b.Pos
        Set b = Nothing
    End If
    
    lv.ListItems.Remove selLi.Index
    Set selLi = Nothing
    
End Sub

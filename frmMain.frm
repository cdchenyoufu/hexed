VERSION 5.00
Begin VB.Form frmEditor 
   Caption         =   "Hexeditor v 0.3"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin rhexed.HexEd HexEditor 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   14420
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open (Ctrl+O)"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save (Ctrl+S)"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuSelectAll 
         Caption         =   "SelectAll (Ctrl+A)"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy (Ctrl+C)"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut (Ctrl+X)"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste (Ctrl+V)"
      End
      Begin VB.Menu mnuOverWrite 
         Caption         =   "OverWrite (Ctrl+B)"
      End
      Begin VB.Menu mnuCopyHexCodes 
         Caption         =   "Copy HexCodes (F4)"
      End
      Begin VB.Menu mnuSpacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete (Del)"
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "Insert (INS)"
      End
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo (CTRL+Z)"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Search (Ctrl+F)"
      End
      Begin VB.Menu mnuStrings 
         Caption         =   "Strings"
      End
      Begin VB.Menu mnuGoto 
         Caption         =   "Goto (Ctrl+G)"
      End
      Begin VB.Menu mnuSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddBookMark 
         Caption         =   "Add BookMark (SHIFT+F2)"
      End
      Begin VB.Menu mnuGotoNextBookMark 
         Caption         =   "Goto Next BookMark (F2)"
      End
      Begin VB.Menu mnuShowBookMarks 
         Caption         =   "Show BookMarks (F3)"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuAbout 
         Caption         =   "About (F5)"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help (F1)"
      End
      Begin VB.Menu mnuSetStringsMinLen 
         Caption         =   "Strings Min Match Length"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy2 
         Caption         =   "Copy (Ctrl+C)"
      End
      Begin VB.Menu mnuCopyHex2 
         Caption         =   "Copy Hex Codes (F4)"
      End
      Begin VB.Menu mnuSearch2 
         Caption         =   "Search (Ctrl+F)"
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
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim minStringLength As Long

Private Sub Form_Activate()
    On Error Resume Next
    Me.HexEditor.SetFocus
End Sub

Private Sub Form_GotFocus()
    On Error Resume Next
    Me.HexEditor.SetFocus
End Sub

Private Sub Form_Load()
   FormPos Me, True
   Me.Visible = True
   On Error Resume Next
   minStringLength = CLng(GetMySetting("minStringLength", 7))
   If minStringLength < 1 Then minStringLength = 7
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim answer As Long
    If Not HexEditor.ReadOnly And Me.HexEditor.IsDirty Then
        answer = MsgBox("Save changes?", vbYesNoCancel Or vbExclamation, "VB Hexedit")
        If answer = vbYes Then Me.HexEditor.Save
    End If
    FormPos Me, True, True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    HexEditor.Width = Me.ScaleWidth
    HexEditor.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveMySetting "minStringLength", minStringLength
End Sub

Private Sub HexEditor_Dirty()
    If Right(Me.Caption, 1) <> "*" Then Me.Caption = Me.Caption & " *"
End Sub

Private Sub HexEditor_Loaded()
    If HexEditor.ReadOnly Then Me.Caption = Me.Caption & " - Read Only - " & HexEditor.LoadedFile
End Sub

Private Sub HexEditor_Saved()
    Me.Caption = "Hexeditor 0.3 - " & HexEditor.LoadedFile
End Sub

Private Sub mnuAbout_Click()
    HexEditor.ShowAbout
End Sub

Private Sub mnuAddBookMark_Click()
    HexEditor.ToggleBookmark HexEditor.SelStart
End Sub

Private Sub mnuCopy_Click()
    HexEditor.DoCopy
End Sub

Private Sub mnuCopyHexCodes_Click()
   Clipboard.Clear
   Clipboard.SetText HexEditor.SelTextAsHexCodes
End Sub

Private Sub mnuCut_Click()
    HexEditor.DoCut
End Sub

Private Sub mnuDelete_Click()
    HexEditor.DoDelete
End Sub

Private Sub mnuGoto_Click()
    HexEditor.ShowGoto
End Sub

Private Sub mnuGotoNextBookMark_Click()
    HexEditor.GotoNextBookmark
End Sub

Private Sub mnuHelp_Click()
    HexEditor.ShowHelp
End Sub

Private Sub mnuInsert_Click()
    HexEditor.ShowInsert
End Sub

Private Sub mnuOpen_Click()
    HexEditor.ShowOpen
End Sub

Private Sub mnuOverWrite_Click()
    HexEditor.DoPasteOver
End Sub

Private Sub mnuPaste_Click()
    HexEditor.DoPaste
End Sub

Private Sub mnuSave_Click()
    HexEditor.Save
End Sub

Private Sub mnuSaveAs_Click()
    HexEditor.SaveAs
End Sub

Private Sub mnuSearch_Click()
     HexEditor.ShowFind
End Sub
 
Private Sub mnuSelectAll_Click()
    HexEditor.SelectAll
End Sub

Private Sub mnuSetStringsMinLen_Click()
    On Error Resume Next
    minStringLength = CLng(InputBox("Set Strings mininimum match length:", , minStringLength))
    If Err.Number <> 0 Or minStringLength < 1 Then minStringLength = 7
End Sub

Private Sub mnuShowBookMarks_Click()
    HexEditor.ShowBookMarks
End Sub

Private Sub mnuStrings_Click()
    
    On Error Resume Next
    
    Dim Ascii() As String
    Dim uni() As String
    
    Ascii() = HexEditor.Strings(minStringLength)
    uni() = HexEditor.Strings(minStringLength)
    
    frmOffsetList.LoadList Me, Ascii
    frmOffsetList.LoadList Me, uni 'this will append the data...
    
End Sub

Private Sub mnuUndo_Click()
    HexEditor.DoUndo
End Sub

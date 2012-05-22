VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   10995
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4035
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   30
      Width           =   10815
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ShowAbout()

    Dim r()
    
    push r, "VB HexEditor by Rang3r\n"
    push r, "buffered file access and screen display (load files upto 2.1 gb)"
    push r, "copy/paste data in both hex/text mode, delete/insert/overwrite"
    push r, "bookmarks, undo, custom scrollbar\n"
    push r, "http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=34729&lngWId=1\n"
    push r, "dzzie mods:\n\tLoadedFromBytes/String, ReadOnly mode, \n\tForceLoadFromMemOnly, Find, Converted to OCX"
    push r, "\tmisc tweaks/rewires, search/bookmark list form\n"
    push r, "Big thanks to Rang3r for releasing this, its a great codebase!"
    Dim tmp
    
    tmp = Join(r, vbCrLf)
    tmp = Replace(tmp, "\n", vbCrLf)
    tmp = Replace(tmp, "\t", vbTab)
    
    Text1 = tmp
    Me.Visible = True
    
End Sub

Sub ShowHelp()

    Dim r()
    
    push r, "Supported commands:\n"
    push r, "Copy (Ctrl+C),\n Paste (Ctrl+V),\n Delete (DEL),\n Insert (INS),\n Write (Ctrl+B)"
    push r, "Open (Ctrl+O),\n Undo (Ctrl+Z),\n Find (Ctrl+F),\n Help (F1)"
    push r, "Toggle BookMark (Shift+F2),\n GoToNextBookMark (F2),\n ShowBookMarks (F3)"
    push r, "Copy Hex Codes (F4),\n About (F5)\n Goto Offset (Ctrl+G)"
    
    push r, "\nYou can copy data from either the hex or char panes"
    
    
    Dim tmp
    
    tmp = Join(r, vbCrLf)
    tmp = Replace(tmp, "\n ", vbCrLf)
     tmp = Replace(tmp, "\n", vbCrLf)
    tmp = Replace(tmp, "\t", vbTab)
    
    Text1 = tmp
    Me.Visible = True
    
End Sub


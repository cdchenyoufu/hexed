VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   30
      Width           =   9105
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Dim r()
    
    push r, "VB HexEditor by Rang3r\n"
    push r, "filebuffer (load files upto 2.1 gb)"
    push r, "copy paste data in both hex/text mode, delete/insert bytes (INS)"
    push r, "custom scrollbar to allow more than 32k lines"
    push r, "bookmarks (F2)\n"
    push r, "http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=34729&lngWId=1\n"
    push r, "dzzie mods:\n\tLoadedFromBytes/String, ReadOnly, \n\tForceLoadFromMemOnly, Find, Converted to OCX\n"
    push r, "Big thanks to Rang3r for releasing this, its a great codebase!"
    Dim tmp
    
    tmp = Join(r, vbCrLf)
    tmp = Replace(tmp, "\n", vbCrLf)
    tmp = Replace(tmp, "\t", vbTab)
    
    Text1 = tmp

End Sub

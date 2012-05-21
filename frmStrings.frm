VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmStrings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Strings"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   Icon            =   "frmStrings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin MSComctlLib.ProgressBar pb 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4455
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   6735
      Begin MSComctlLib.ListView lw 
         Height          =   4095
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   7223
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   11359
         EndProperty
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   5280
      Width           =   1575
   End
End
Attribute VB_Name = "frmStrings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub lw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    frmMain.ActiveWindow.HexEditor.SelStart = val("&h" & Item.Text)
    frmMain.ActiveWindow.HexEditor.ScrollTo val("&h" & Item.Text)
    
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMyItemData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Existe un pour chaque élément de menu OWNERDRAW"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarsMenuText As String 'local copy
Private mvarfType As Long 'local copy
Private mvariButton As Integer 'local copy
Private mvarbMainMenu As Boolean 'local copy

Public Property Let bMainMenu(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bMainMenu = 5
    mvarbMainMenu = vData
End Property

Public Property Get bMainMenu() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bMainMenu
    bMainMenu = mvarbMainMenu
End Property

Public Property Let iButton(ByVal vData As Integer)
Attribute iButton.VB_Description = "Indexe de l'image"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iButton = 5
    mvariButton = vData
End Property

Public Property Get iButton() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iButton
    iButton = mvariButton
End Property

Public Property Let fType(ByVal vData As Long)
Attribute fType.VB_Description = "Type d'élément (item type) original"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fType = 5
    mvarfType = vData
End Property

Public Property Get fType() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fType
    fType = mvarfType
End Property

Public Property Let sMenuText(ByVal vData As String)
Attribute sMenuText.VB_Description = "Le texte du menu"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MenuText = 5
    mvarsMenuText = vData
End Property

Public Property Get sMenuText() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MenuText
    sMenuText = mvarsMenuText
End Property



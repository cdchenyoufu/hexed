Attribute VB_Name = "MenuSwitcher"
'Simple yet useful, this code adds the ability
'to Switch Icons using the code at the following
'URL: http://planet-source-code.com/xq/ASP/txtCodeId.14206/lngWId.1/qx/vb/scripts/ShowCode.htm
'The author of the code I used to make this work is
'Wazerface (Martin McCormick).
'In case you need more info on what this does,
'it adds the Switch ability. You have an imagelist,
'and you put a certain character (default is @)
'as the tag of an icon that you want to be the
'grayed icon of the one with the tag that doesn't
'have the character. For instance, @&Copy would be
'the grayed version of &Copy, so switching just makes
'the icon of the &Copy item gray. Wazerface's code
'sets icons of menus by their tags, as you will see
'in my Image List. By Tommy W. (homeworkkid@msn.com)

'Use Notes: You must add CMWndProc.bas, CMyItemData.cls,
'CMyItemData.cls, CMyItemDatas.cls, and CoolMenu.cls
'as well as this module to your application for the
'switch function to work.

'Diff is the character used in front of the grayed icon
'version, default is @. Images is the ImageList, Menu
'is the Menu Item array to be checked for the tag
'specified so that the text of the menu will be grayed
'as well as the icon (or non-grayed). Default ImageList
'is ImageList1. Form is self-explanatory.

Public Function Switch(Form As Form, Tag As String, Optional Images As ImageList, Optional Diff As String, Optional menu As Object)
On Error GoTo Err
Dim setvar As Integer
Dim setvar2 As Integer
Dim errs As Integer
If Diff = "" Then Diff = "@"
setvar = Images.ListImages.Count
Cont1:
If setvar = 0 Then Set Images = Form.ImageList1 Else GoTo DoFunction
setvar = Images.ListImages.Count
Cont2:
If setvar = 0 Then GoTo Error
DoFunction:
If IsMissing(menu) = False Then setvar2 = 1
Cont3:
Dim i As Integer
Dim i1 As Integer
Dim i2 As Integer
For i = 1 To Images.ListImages.Count
If Images.ListImages(i).Tag = Tag Then i1 = i
If Images.ListImages(i).Tag = Diff & Tag Then i2 = i
Next i
If i1 = 0 Or i2 = 0 Then GoTo Error
Images.ListImages(i1).Tag = Diff & Tag
Images.ListImages(i2).Tag = Tag
Switch = "Success"
GoTo Done
Error:
Switch = "Error in ImageList"
GoTo Done
Err:
If Err.Number = 91 Then
errs = errs + 1
If errs = 1 Then GoTo Cont1
If errs = 2 Then GoTo Cont2
If errs = 3 Then GoTo Cont3
End If
MsgBox Err.Number & ": " & Err.Description, vbCritical, "Error in " & App.Title
Done:
If setvar2 = 1 Then SwitchMenu Tag, menu
Finished:
End Function

Private Function SwitchMenu(Tag As String, menu As Object)
On Error Resume Next
Dim i As Integer
Dim p As Integer
p = -1
For i = menu.LBound To menu.UBound
If menu(i).Caption = Tag Then p = i
Next i
If p <> -1 Then
If menu(p).Enabled = True Then menu(p).Enabled = False Else menu(p).Enabled = True
End If
Done:
End Function

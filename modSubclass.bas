Attribute VB_Name = "modSubclass"
Option Explicit
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private Const GWL_WNDPROC = (-4)

Private Const wm_MouseWheel = &H20A
Private Const wm_KeyPress = &H100

Private Controls As New Collection

Private Function WndProc(ByVal hWnd As Long, ByVal Msg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim high As Long
    Dim low As Long
    Dim lPrevWndProc As Long
    Dim Control As HexEd
    Dim wnd As Long
    
    On Error Resume Next
    
    wnd = GetProp(hWnd, "Wnd")
    lPrevWndProc = GetProp(hWnd, "OldProc")
    
    Select Case wnd
        Case 1 'user ctrl
             Set Control = Controls.Item("hwnd" & hWnd)
             
             If Control Is Nothing Then
                Debug.Print "hwnd: " & hWnd & " was expected to be a Control but wasnt registered? ignoring..."
                Exit Function
             End If
             
             Select Case Msg
                 Case wm_MouseWheel
                     high = (wParam And &HFFFF0000) \ &H10000
                     low = wParam And &HFFFF&
                     high = high / 120
                     Control.Scroll high
                     
                     'Debug.Print "MSWHEEL_ROLLMSG  " & hWnd & "   " & high & "  " & low & "   " & lParam & "      " & wParam
                     Msg = 0
                     WndProc = 10 'CallWindowProc(lPrevWndProc, hWnd, Msg, wParam, lParam)
                     Exit Function
                Case Else
                    'Debug.Print "other  " & Msg & "   " & hWnd & "   " & high & "  " & low & "   " & lParam
            End Select
            
        Case 2, 3 'canvas , ascii
            Select Case Msg
                Case wm_KeyPress
                    Select Case wParam  'subclass keys

                    End Select
            End Select
            
    End Select
   
   WndProc = CallWindowProc(lPrevWndProc, hWnd, Msg, wParam, lParam)
End Function

Public Sub InstallSubclass(Control As HexEd)
    Dim oldproc As Long
    Dim hWnd As Long
    hWnd = Control.hWnd
    Controls.Add Control, "hwnd" & Control.hWnd
    oldproc = GetWindowLong(hWnd, GWL_WNDPROC)
    SetProp hWnd, "OldProc", oldproc
    SetProp hWnd, "Wnd", 1
    SetWindowLong hWnd, GWL_WNDPROC, AddressOf WndProc
    
    hWnd = Control.hWndCanvas
    oldproc = GetWindowLong(hWnd, GWL_WNDPROC)
    SetProp hWnd, "OldProc", oldproc
    SetProp hWnd, "Wnd", 2
    SetWindowLong hWnd, GWL_WNDPROC, AddressOf WndProc
    
    hWnd = Control.hWndAscii
    oldproc = GetWindowLong(hWnd, GWL_WNDPROC)
    SetProp hWnd, "OldProc", oldproc
    SetProp hWnd, "Wnd", 3
    SetWindowLong hWnd, GWL_WNDPROC, AddressOf WndProc

End Sub

Public Sub UnInstallSubclass(Control As HexEd)
    Dim oldproc As Long
    Dim hWnd As Long
    hWnd = Control.hWndCanvas
    oldproc = GetProp(hWnd, "OldProc")
    SetWindowLong hWnd, GWL_WNDPROC, oldproc
    
    hWnd = Control.hWndAscii
    oldproc = GetProp(hWnd, "OldProc")
    SetWindowLong hWnd, GWL_WNDPROC, oldproc
    
    hWnd = Control.hWnd
    oldproc = GetProp(hWnd, "OldProc")
    SetWindowLong hWnd, GWL_WNDPROC, oldproc
End Sub

Attribute VB_Name = "RtbSubclass"
Option Explicit

'http://www.xtremevbtalk.com/showthread.php?t=78306

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
                          (ByVal lpPrevWndFunc As Long, _
                           ByVal hwnd As Long, _
                           ByVal msg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
                          (ByVal hwnd As Long, _
                           ByVal nIndex As Long, _
                           ByVal dwNewLong As Long) As Long

' A pointer to the old window procedure
Public pOldWindPoc As Long

Public Const GWL_WNDPROC& = (-4)

' Our new window procedure
Public Function WndProc(ByVal hwnd As Long, _
                                       ByVal uMsg As Long, _
                                       ByVal wParam As Long, _
                                       ByVal lParam As Long) As Long
    Dim handled As Long
    handled = False
    
    Const WM_KEYDOWN = &H100
    
    Select Case uMsg
        Case WM_KEYDOWN
            Select Case wParam
                Case vbKeyV
                    'just ignore this message
                    handled = True
            End Select
    End Select
    
    If Not handled Then
        WndProc = CallWindowProc(pOldWindPoc, hwnd, uMsg, wParam, lParam)
    End If
End Function


Attribute VB_Name = "modHooking"
Option Explicit

' Windows Constants used by Subclasser.
Public Const WM_CLOSE               As Long = &H10
Public Const WM_SETCURSOR           As Long = &H20
Public Const WM_MOUSEACTIVATE       As Long = &H21
Public Const WM_SYSCOMMAND          As Long = &H112
Public Const WM_MOUSEMOVE           As Long = &H200
Public Const WM_LBUTTONDOWN         As Long = &H201
Public Const WM_LBUTTONUP           As Long = &H202
Public Const WM_LBUTTONDBLCLK       As Long = &H203
Public Const WM_RBUTTONDOWN         As Long = &H204
Public Const WM_RBUTTONUP           As Long = &H205
Public Const WM_RBUTTONDBLCLK       As Long = &H206
Public Const WM_MBUTTONDOWN         As Long = &H207
Public Const WM_MBUTTONUP           As Long = &H208
Public Const WM_MBUTTONDBLCLK       As Long = &H209
Public Const WM_MOUSEWHEEL          As Long = &H20A

Public cMouseEvents As clsMouseEvents
Public lpPrevWndFunc As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Function HookedForm(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo Err_HookedForm:

    Select Case uMsg
    
        Case Is = WM_MOUSEWHEEL ' MouseWheel
            Call cMouseEvents.MouseWheelUsed(wParam > 0)
        
        Case Is = WM_MOUSEMOVE ' Mouse moving on form
            ' If the hwnd under mouse has changed raise mousemove events
            Call cMouseEvents.DetectMouseEvents

        Case Is = WM_SETCURSOR  ' Mouse moving on control
            ' If the hwnd under mouse has changed raise mousemove events
            Call cMouseEvents.DetectMouseEvents

        Case Is = WM_SYSCOMMAND 'Form unloading
            ' if the form is unloading we will unhook to prevent crashes
            If wParam = 61536 Then cMouseEvents.UnhookMouseEvents

    End Select
  
Err_HookedForm:
    ' Placing the window proc after our filter messages with the select
    ' case above allows us to recieve the messages faster than a regular vb event ;P
    HookedForm = CallWindowProc(lpPrevWndFunc, hwnd, uMsg, wParam, lParam)
  
End Function

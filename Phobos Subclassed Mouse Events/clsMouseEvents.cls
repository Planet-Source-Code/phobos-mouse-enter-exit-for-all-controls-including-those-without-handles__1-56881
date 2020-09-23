VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMouseEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Pointers to last reported MouseEvents.
Private ctlLastEntered As Control
Private bHovering As Boolean

' Subclassing Pointers.
Public frmParent As Form
Private Const GWL_WNDPROC = -4

' Type Declarations
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

' Mouse Events.
Public Event MouseEnter(ctlEntered As Control)
Public Event MouseExit(ctlExited As Control)
Public Event MouseWheelUp(lngMoueX As Long, lngMouseY As Long)
Public Event MouseWheelDown(lngMoueX As Long, lngMouseY As Long)

Private Sub Class_Initialize()

    ' Initialise Historic MouseEvent flags and pointers.
    
    Set ctlLastEntered = Nothing    ' We store a pointer to the last control entered here.
    bHovering = False               ' A flag that indicates if we are hovering over a control.
    lpPrevWndFunc = 0               ' A pointer used by subclassing (zero indicates subclassing inactive).

End Sub

Public Sub HookMouseEvents(frmPassed As Form)
    
    ' Start intercepting form events, and process them in the routine "HookedForm".

    If lpPrevWndFunc = 0 Then
        Set frmParent = frmPassed
        Set modHooking.cMouseEvents = Me
        lpPrevWndFunc = SetWindowLong(frmParent.hwnd, GWL_WNDPROC, AddressOf modHooking.HookedForm)
    End If

End Sub

Public Sub UnhookMouseEvents()
   
    ' Terminate the interception of form events and returns control to it's previous state.
    
    If lpPrevWndFunc <> 0 Then
        Call SetWindowLong(frmParent.hwnd, GWL_WNDPROC, lpPrevWndFunc)
        Set modHooking.cMouseEvents = Nothing
    End If

End Sub

Public Sub MouseWheelUsed(bWheelUp As Boolean)
On Error Resume Next
    
    ' Notify the parent application of a MouseWheel event.
    
    Dim Mouse As POINTAPI
    
    ' Find the current mouse pointer position (to be passed on with the event).
    Call GetCursorPos(Mouse)
    
    ' Now raise the appropriate MouseWheel Event.
    If bWheelUp Then
        RaiseEvent MouseWheelUp(Mouse.X, Mouse.Y)
    Else
        RaiseEvent MouseWheelDown(Mouse.X, Mouse.Y)
    End If

End Sub

Public Sub DetectMouseEvents()
On Error Resume Next

    ' Detect and report all new MouseEnter and MouseExit events.
    
    Dim iControl As Integer, ctlSkinned As Object
    
    ' First we check to see if a MouseExit event should be generated.
    If bHovering Then
        ' The pointer has entered a control so now we find out if it is still there.
        If Not IsUnderPointer(ctlLastEntered) Then
            ' The pointer is no longer over the control it entered so we gnerate a MouseExit event.
            RaiseEvent MouseExit(ctlLastEntered)
            Set ctlLastEntered = Nothing
            bHovering = False
        End If
    End If
    
    ' We now search through all skinned controls to see if one of them is now underneath the mouse pointer.
    For iControl = 0 To frmParent.Controls.Count - 1
        Set ctlSkinned = frmParent.Controls(iControl)
        If UCase(Trim$(ctlSkinned.Tag)) = "SKINNED" Then
            
            ' A skinned control has been found so now we find out if it is underneath the pointer.
            If IsUnderPointer(ctlSkinned) Then
                ' The control is below the pointer so we find out if a MouseEnter event is required.
                If Not bHovering Then
                    RaiseEvent MouseEnter(ctlSkinned)
                    Set ctlLastEntered = ctlSkinned
                    bHovering = True
                End If
            End If
        
        End If
    Next iControl

End Sub

Private Function IsUnderPointer(ctlPassed As Control) As Boolean
On Error Resume Next

    ' Detect if the passed control is underneath the mouse pointer.
 
    Dim Mouse As POINTAPI, TopLeft As POINTAPI, BottomRight As POINTAPI
    Dim bTestX As Boolean, bTestY As Boolean, lngHwnd As Long, bHasHwnd As Boolean
    
    ' Identify if the passed object has a windows handle.
    bHasHwnd = HasHwnd(ctlPassed)
    
    ' Now we know where the image is we need to find out if the mouse pointer is in the mask area.
    Call GetCursorPos(Mouse)
    
    ' Find the windows handle for the container/control under the mouse pointer.
    lngHwnd = WindowFromPoint(Mouse.X, Mouse.Y)
    
    ' Identify if the mouse pointer is over the conrol.
    If bHasHwnd Then
        
        ' The passed object has a handle so we compare it with the handle under the mouse pointer.
        IsUnderPointer = (lngHwnd = ctlPassed.hwnd)
    
    Else
        
        ' No windows handle available so compare the pointer position with the container and
        ' position of the required object within it.
        If lngHwnd = ctlPassed.Container.hwnd Then
            
            ' Identify the position (in pixels) of the first available screen coordinate on the form.
            TopLeft.X = 0:   TopLeft.Y = 0
            Call ClientToScreen(lngHwnd, TopLeft)
    
            ' Using this information we calculate the position in pixels of the passed object.
            With ctlPassed
                TopLeft.X = TopLeft.X + .Left / 15
                TopLeft.Y = TopLeft.Y + .Top / 15
                BottomRight.X = TopLeft.X + .Width / 15
                BottomRight.Y = TopLeft.Y + .Height / 15
            End With
    
            ' Set the MouseEvents boolean flag if the mouse is located over the passed control.
            With ctlPassed
                bTestX = ((Mouse.X >= TopLeft.X) And (Mouse.X <= BottomRight.X))
                bTestY = ((Mouse.Y >= TopLeft.Y) And (Mouse.Y <= BottomRight.Y))
            End With
            IsUnderPointer = (bTestX And bTestY)
        Else
            ' The pointer is not in the correct container so the test is failed.
            IsUnderPointer = False
        End If
    
    End If

End Function

Private Function HasHwnd(ctlPassed As Control) As Boolean
On Error GoTo Error_HasHwnd
        
    ' This function will return a boolean that indicates if the passed control has a windows handler.
    
    HasHwnd = (ctlPassed.hwnd > 0)

Exit_HasHwnd:
    Exit Function

Error_HasHwnd:
    HasHwnd = False
    Resume Exit_HasHwnd

End Function

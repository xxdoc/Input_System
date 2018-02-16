Attribute VB_Name = "modMain"
Option Explicit

' Hook
Private Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal Previous As Long, ByVal hWnd As Long, _
    ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal Index As Long, _
    ByVal NewLong As Long) As Long

Private Const GWL_WNDPROC As Long = -4

Dim Hook_hWnd     As Long
Dim Hook_Previous As Long

' Memory
Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, _
    ByVal Length As Long)

' Loop
Dim Running As Boolean

Public Sub Main()
    frmMain.Show
    
    Hook_hWnd = frmMain.hWnd
    Hook_Previous = SetWindowLong(Hook_hWnd, GWL_WNDPROC, Address(AddressOf WindowProc))
    
    Input_Initialize
    
    Running = True
    Cycle
End Sub

Public Sub Terminate()
    Dim Unhook As Long
    
    Running = False
    Unhook = SetWindowLong(Hook_hWnd, GWL_WNDPROC, Hook_Previous)
    
    Unload frmMain
End Sub

Public Sub Cycle()
    Do While Running
        Input_Check
        
        ' Update UI
        frmMain.lblRaw.Caption = "Raw: " & Input_Raw
        frmMain.lblChar.Caption = "Char: " & Input_Char
        frmMain.lblCoord.Caption = "Mouse Coord: " & Input_MouseCoord(True) & ", " & Input_MouseCoord(False)
        frmMain.lblLeft.Caption = "Mouse Left: " & Input_MouseState(MOUSE_LEFT, MOUSE_DOWN) & ", " & _
            Input_MouseState(MOUSE_LEFT, MOUSE_UP) & ", " & Input_MouseState(MOUSE_LEFT, MOUSE_DOUBLE)
        frmMain.lblRight.Caption = "Mouse Right: " & Input_MouseState(MOUSE_RIGHT, MOUSE_DOWN) & ", " & _
            Input_MouseState(MOUSE_RIGHT, MOUSE_UP) & ", " & Input_MouseState(MOUSE_RIGHT, MOUSE_DOUBLE)
        frmMain.lblMiddle.Caption = "Mouse Middle: " & Input_MouseState(MOUSE_MIDDLE, MOUSE_DOWN) & ", " & _
            Input_MouseState(MOUSE_MIDDLE, MOUSE_UP) & ", " & Input_MouseState(MOUSE_MIDDLE, MOUSE_DOUBLE)
        frmMain.lblWheel.Caption = "Mouse Wheel: " & Input_MouseWheel
        
        DoEvents
    Loop
    
    Terminate
End Sub

Public Function WindowProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case Msg
    
        Case WM_INPUT
        Input_HandleRaw lParam
        
        Case WM_CHAR
        Input_Char = Chr$(wParam)
        
        Case WM_MOUSEMOVE
        Input_HandleMouseMove lParam
        
        Case WM_LBUTTONDOWN
        Input_MouseState(MOUSE_LEFT, MOUSE_DOWN) = True
        Input_MouseState(MOUSE_LEFT, MOUSE_UP) = False
        Input_MouseState(MOUSE_LEFT, MOUSE_DOUBLE) = False
        
        Case WM_LBUTTONUP
        Input_MouseState(MOUSE_LEFT, MOUSE_DOWN) = False
        Input_MouseState(MOUSE_LEFT, MOUSE_UP) = True
        Input_MouseState(MOUSE_LEFT, MOUSE_DOUBLE) = False
        
        Case WM_LBUTTONDBLCLK
        Input_MouseState(MOUSE_LEFT, MOUSE_DOWN) = False
        Input_MouseState(MOUSE_LEFT, MOUSE_UP) = False
        Input_MouseState(MOUSE_LEFT, MOUSE_DOUBLE) = True
        
        Case WM_RBUTTONDOWN
        Input_MouseState(MOUSE_RIGHT, MOUSE_DOWN) = True
        Input_MouseState(MOUSE_RIGHT, MOUSE_UP) = False
        Input_MouseState(MOUSE_RIGHT, MOUSE_DOUBLE) = False
        
        Case WM_RBUTTONUP
        Input_MouseState(MOUSE_RIGHT, MOUSE_DOWN) = False
        Input_MouseState(MOUSE_RIGHT, MOUSE_UP) = True
        Input_MouseState(MOUSE_RIGHT, MOUSE_DOUBLE) = False
        
        Case WM_RBUTTONDBLCLK
        Input_MouseState(MOUSE_RIGHT, MOUSE_DOWN) = False
        Input_MouseState(MOUSE_RIGHT, MOUSE_UP) = False
        Input_MouseState(MOUSE_RIGHT, MOUSE_DOUBLE) = True
        
        Case WM_MBUTTONDOWN
        Input_MouseState(MOUSE_MIDDLE, MOUSE_DOWN) = True
        Input_MouseState(MOUSE_MIDDLE, MOUSE_UP) = False
        Input_MouseState(MOUSE_MIDDLE, MOUSE_DOUBLE) = False
        
        Case WM_MBUTTONUP
        Input_MouseState(MOUSE_MIDDLE, MOUSE_DOWN) = False
        Input_MouseState(MOUSE_MIDDLE, MOUSE_UP) = True
        Input_MouseState(MOUSE_MIDDLE, MOUSE_DOUBLE) = False
        
        Case WM_MBUTTONDBLCLK
        Input_MouseState(MOUSE_MIDDLE, MOUSE_DOWN) = False
        Input_MouseState(MOUSE_MIDDLE, MOUSE_UP) = False
        Input_MouseState(MOUSE_MIDDLE, MOUSE_DOUBLE) = True
        
        Case WM_MOUSEWHEEL
        ' https://msdn.microsoft.com/en-us/library/windows/desktop/ms645617(v=vs.85).aspx
        
        Case Else
        WindowProc = CallWindowProc(Hook_Previous, hWnd, Msg, wParam, lParam)
    
    End Select
End Function

Public Function Address(ByRef A As Long) As Long
    Address = A
End Function

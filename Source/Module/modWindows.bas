Attribute VB_Name = "modWindows"
Option Explicit

' WindowProc
Private Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal Previous As Long, ByVal hWnd As Long, _
    ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal Index As Long, _
    ByVal NewLong As Long) As Long

Private Const GWL_WNDPROC As Long = -4

Dim WndProc_hWnd As Long
Dim WndProc_Prev As Long

' Memory
Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, _
    ByVal Length As Long)

' WM_*
Public Const WM_INPUT         As Long = &HFF&
Public Const WM_KEYDOWN       As Long = &H100
Public Const WM_KEYUP         As Long = &H101
Public Const WM_CHAR          As Long = &H102
Public Const WM_SYSKEYDOWN    As Long = &H104
Public Const WM_SYSKEYUP      As Long = &H105
Public Const WM_MOUSEMOVE     As Long = &H200
Public Const WM_LBUTTONDOWN   As Long = &H201
Public Const WM_LBUTTONUP     As Long = &H202
Public Const WM_LBUTTONDBLCLK As Long = &H203
Public Const WM_RBUTTONDOWN   As Long = &H204
Public Const WM_RBUTTONUP     As Long = &H205
Public Const WM_RBUTTONDBLCLK As Long = &H206
Public Const WM_MBUTTONDOWN   As Long = &H207
Public Const WM_MBUTTONUP     As Long = &H208
Public Const WM_MBUTTONDBLCLK As Long = &H209
Public Const WM_MOUSEWHEEL    As Long = &H20A

Public Sub Windows_Initialize(ByVal hWnd As Long)
    WndProc_hWnd = hWnd
    WndProc_Prev = SetWindowLong(WndProc_hWnd, GWL_WNDPROC, Address(AddressOf WindowProc))
End Sub

Public Sub Windows_Terminate()
    SetWindowLong WndProc_hWnd, GWL_WNDPROC, WndProc_Prev
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
        WindowProc = CallWindowProc(WndProc_Prev, hWnd, Msg, wParam, lParam)
    
    End Select
End Function

Public Function Address(ByRef A As Long) As Long
    Address = A
End Function

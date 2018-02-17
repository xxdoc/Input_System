Attribute VB_Name = "modMain"
Option Explicit

' Loop
Dim Running As Boolean

Public Sub Main()
    frmMain.Show
    
    Windows_Initialize frmMain.hWnd
    Input_Initialize frmMain.hWnd
    
    Running = True
    Cycle
End Sub

Public Sub Terminate()
    Running = False
    
    Windows_Terminate
    
    Unload frmMain
End Sub

Public Sub Cycle()
    Do While Running
        Input_Check
        
        ' Update UI
        frmMain.lblRaw.Caption = "Raw: " & Input_Key
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

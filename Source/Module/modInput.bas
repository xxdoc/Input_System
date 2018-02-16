Attribute VB_Name = "modInput"
Option Explicit

' Raw Input
Private Declare Function RegisterRawInputDevices Lib "User32" (ByRef Devices As tRaw_Device, ByVal Number As Long, _
    ByVal Size As Long) As Long
Private Declare Function GetRawInputData Lib "User32" (ByVal RawInput As Long, ByVal Command As Long, ByRef Data As Any, _
    ByRef Size As Long, ByVal SizeHeader As Long) As Long

Private Type tRaw_Device
    UsagePage As Integer
    Usage     As Integer
    Flags     As Long
    hWnd      As Long
End Type

Private Type tRaw_Header
    Type   As Long
    Size   As Long
    Device As Long
    wParam As Long
End Type

Private Type tRaw_Keyboard
    MakeCode As Integer
    Flags    As Integer
    Reserved As Integer
    VKey     As Integer
    Message  As Long
    Extra    As Long
End Type

Private Const RID_INPUT As Long = &H10000003

' tRaw_Header.Type
Private Const RIM_TYPEKEYBOARD As Long = &H1&

' tRaw_Keyboard.Flags
Private Const RI_KEY_MAKE  As Long = &H0& ' Key Down
Private Const RI_KEY_BREAK As Long = &H1& ' Key Up
Private Const RI_KEY_E0    As Long = &H2& ' Left Version Of Key
Private Const RI_KEY_E1    As Long = &H4& ' Right Version Of Key

' MapVirtualKey
Private Declare Function MapVirtualKey Lib "User32" Alias "MapVirtualKeyA" (ByVal Code As Long, ByVal Map As Long) As Long

' MapVirtualKey.Map
Private Const MAPVK_VK_TO_VSC    As Long = 0 ' Virtual Key into Scan Code
Private Const MAPVK_VSC_TO_VK    As Long = 1 ' Scan Code into Virtual Key
Private Const MAPVK_VK_TO_CHAR   As Long = 2 ' Virtual Key into Unshifted Character
Private Const MAPVK_VSC_TO_VK_EX As Long = 3 ' Scan Code into Virtual Key + Left/Right Keys

' Virtual Key
Private Const VK_LBUTTON             As Long = &H1  ' Left mouse button
Private Const VK_RBUTTON             As Long = &H2  ' Right mouse button
Private Const VK_CANCEL              As Long = &H3  ' Control-break processing
Private Const VK_MBUTTON             As Long = &H4  ' Middle mouse button (three-button mouse)
Private Const VK_XBUTTON1            As Long = &H5  ' X1 mouse button
Private Const VK_XBUTTON2            As Long = &H6  ' X2 mouse button
Private Const VK_BACK                As Long = &H8  ' BACKSPACE key
Private Const VK_TAB                 As Long = &H9  ' TAB key
Private Const VK_CLEAR               As Long = &HC  ' CLEAR key
Private Const VK_RETURN              As Long = &HD  ' ENTER key
Private Const VK_SHIFT               As Long = &H10 ' SHIFT key
Private Const VK_CONTROL             As Long = &H11 ' CTRL key
Private Const VK_MENU                As Long = &H12 ' ALT key
Private Const VK_PAUSE               As Long = &H13 ' PAUSE key
Private Const VK_CAPITAL             As Long = &H14 ' CAPS LOCK key
Private Const VK_KANA                As Long = &H15 ' IME Kana mode
Private Const VK_JUNJA               As Long = &H17 ' IME Junja mode
Private Const VK_FINAL               As Long = &H18 ' IME final mode
Private Const VK_HANJA               As Long = &H19 ' IME Hanja mode
Private Const VK_ESCAPE              As Long = &H1B ' ESC key
Private Const VK_CONVERT             As Long = &H1C ' IME convert
Private Const VK_NONCONVERT          As Long = &H1D ' IME nonconvert
Private Const VK_ACCEPT              As Long = &H1E ' IME accept
Private Const VK_MODECHANGE          As Long = &H1F ' IME mode change request
Private Const VK_SPACE               As Long = &H20 ' SPACEBAR
Private Const VK_PRIOR               As Long = &H21 ' PAGE UP key
Private Const VK_NEXT                As Long = &H22 ' PAGE DOWN key
Private Const VK_END                 As Long = &H23 ' END key
Private Const VK_HOME                As Long = &H24 ' HOME key
Private Const VK_LEFT                As Long = &H25 ' LEFT ARROW key
Private Const VK_UP                  As Long = &H26 ' UP ARROW key
Private Const VK_RIGHT               As Long = &H27 ' RIGHT ARROW key
Private Const VK_DOWN                As Long = &H28 ' DOWN ARROW key
Private Const VK_SELECT              As Long = &H29 ' SELECT key
Private Const VK_PRINT               As Long = &H2A ' PRINT key
Private Const VK_EXECUTE             As Long = &H2B ' EXECUTE key
Private Const VK_SNAPSHOT            As Long = &H2C ' PRINT SCREEN key
Private Const VK_INSERT              As Long = &H2D ' INS key
Private Const VK_DELETE              As Long = &H2E ' DEL key
Private Const VK_HELP                As Long = &H2F ' HELP key
Private Const VK_0                   As Long = &H30 ' 0 key
Private Const VK_1                   As Long = &H31 ' 1 key
Private Const VK_2                   As Long = &H32 ' 2 key
Private Const VK_3                   As Long = &H33 ' 3 key
Private Const VK_4                   As Long = &H34 ' 4 key
Private Const VK_5                   As Long = &H35 ' 5 key
Private Const VK_6                   As Long = &H36 ' 6 key
Private Const VK_7                   As Long = &H37 ' 7 key
Private Const VK_8                   As Long = &H38 ' 8 key
Private Const VK_9                   As Long = &H39 ' 9 key
Private Const VK_A                   As Long = &H41 ' A key
Private Const VK_B                   As Long = &H42 ' B key
Private Const VK_C                   As Long = &H43 ' C key
Private Const VK_D                   As Long = &H44 ' D key
Private Const VK_E                   As Long = &H45 ' E key
Private Const VK_F                   As Long = &H46 ' F key
Private Const VK_G                   As Long = &H47 ' G key
Private Const VK_H                   As Long = &H48 ' H key
Private Const VK_I                   As Long = &H49 ' I key
Private Const VK_J                   As Long = &H4A ' J key
Private Const VK_K                   As Long = &H4B ' K key
Private Const VK_L                   As Long = &H4C ' L key
Private Const VK_M                   As Long = &H4D ' M key
Private Const VK_N                   As Long = &H4E ' N key
Private Const VK_O                   As Long = &H4F ' O key
Private Const VK_P                   As Long = &H50 ' P key
Private Const VK_Q                   As Long = &H51 ' Q key
Private Const VK_R                   As Long = &H52 ' R key
Private Const VK_S                   As Long = &H53 ' S key
Private Const VK_T                   As Long = &H54 ' T key
Private Const VK_U                   As Long = &H55 ' U key
Private Const VK_V                   As Long = &H56 ' V key
Private Const VK_W                   As Long = &H57 ' W key
Private Const VK_X                   As Long = &H58 ' X key
Private Const VK_Y                   As Long = &H59 ' Y key
Private Const VK_Z                   As Long = &H5A ' Z key
Private Const VK_LWIN                As Long = &H5B ' Left Windows key (Natural keyboard)
Private Const VK_RWIN                As Long = &H5C ' Right Windows key (Natural keyboard)
Private Const VK_APPS                As Long = &H5D ' Applications key (Natural keyboard)
Private Const VK_SLEEP               As Long = &H5F ' Computer Sleep key
Private Const VK_NUMPAD0             As Long = &H60 ' Numeric keypad 0 key
Private Const VK_NUMPAD1             As Long = &H61 ' Numeric keypad 1 key
Private Const VK_NUMPAD2             As Long = &H62 ' Numeric keypad 2 key
Private Const VK_NUMPAD3             As Long = &H63 ' Numeric keypad 3 key
Private Const VK_NUMPAD4             As Long = &H64 ' Numeric keypad 4 key
Private Const VK_NUMPAD5             As Long = &H65 ' Numeric keypad 5 key
Private Const VK_NUMPAD6             As Long = &H66 ' Numeric keypad 6 key
Private Const VK_NUMPAD7             As Long = &H67 ' Numeric keypad 7 key
Private Const VK_NUMPAD8             As Long = &H68 ' Numeric keypad 8 key
Private Const VK_NUMPAD9             As Long = &H69 ' Numeric keypad 9 key
Private Const VK_MULTIPLY            As Long = &H6A ' Multiply key
Private Const VK_ADD                 As Long = &H6B ' Add key
Private Const VK_SEPARATOR           As Long = &H6C ' Separator key
Private Const VK_SUBTRACT            As Long = &H6D ' Subtract key
Private Const VK_DECIMAL             As Long = &H6E ' Decimal key
Private Const VK_DIVIDE              As Long = &H6F ' Divide key
Private Const VK_F1                  As Long = &H70 ' F1 key
Private Const VK_F2                  As Long = &H71 ' F2 key
Private Const VK_F3                  As Long = &H72 ' F3 key
Private Const VK_F4                  As Long = &H73 ' F4 key
Private Const VK_F5                  As Long = &H74 ' F5 key
Private Const VK_F6                  As Long = &H75 ' F6 key
Private Const VK_F7                  As Long = &H76 ' F7 key
Private Const VK_F8                  As Long = &H77 ' F8 key
Private Const VK_F9                  As Long = &H78 ' F9 key
Private Const VK_F10                 As Long = &H79 ' F10 key
Private Const VK_F11                 As Long = &H7A ' F11 key
Private Const VK_F12                 As Long = &H7B ' F12 key
Private Const VK_F13                 As Long = &H7C ' F13 key
Private Const VK_F14                 As Long = &H7D ' F14 key
Private Const VK_F15                 As Long = &H7E ' F15 key
Private Const VK_F16                 As Long = &H7F ' F16 key
Private Const VK_F17                 As Long = &H80 ' F17 key
Private Const VK_F18                 As Long = &H81 ' F18 key
Private Const VK_F19                 As Long = &H82 ' F19 key
Private Const VK_F20                 As Long = &H83 ' F20 key
Private Const VK_F21                 As Long = &H84 ' F21 key
Private Const VK_F22                 As Long = &H85 ' F22 key
Private Const VK_F23                 As Long = &H86 ' F23 key
Private Const VK_F24                 As Long = &H87 ' F24 key
Private Const VK_NUMLOCK             As Long = &H90 ' NUM LOCK key
Private Const VK_SCROLL              As Long = &H91 ' SCROLL LOCK key
Private Const VK_LSHIFT              As Long = &HA0 ' Left SHIFT key
Private Const VK_RSHIFT              As Long = &HA1 ' Right SHIFT key
Private Const VK_LCONTROL            As Long = &HA2 ' Left CONTROL key
Private Const VK_RCONTROL            As Long = &HA3 ' Right CONTROL key
Private Const VK_LMENU               As Long = &HA4 ' Left MENU key
Private Const VK_RMENU               As Long = &HA5 ' Right MENU key
Private Const VK_BROWSER_BACK        As Long = &HA6 ' Browser Back key
Private Const VK_BROWSER_FORWARD     As Long = &HA7 ' Browser Forward key
Private Const VK_BROWSER_REFRESH     As Long = &HA8 ' Browser Refresh key
Private Const VK_BROWSER_STOP        As Long = &HA9 ' Browser Stop key
Private Const VK_BROWSER_SEARCH      As Long = &HAA ' Browser Search key
Private Const VK_BROWSER_FAVORITES   As Long = &HAB ' Browser Favorites key
Private Const VK_BROWSER_HOME        As Long = &HAC ' Browser Start and Home key
Private Const VK_VOLUME_MUTE         As Long = &HAD ' Volume Mute key
Private Const VK_VOLUME_DOWN         As Long = &HAE ' Volume Down key
Private Const VK_VOLUME_UP           As Long = &HAF ' Volume Up key
Private Const VK_MEDIA_NEXT_TRACK    As Long = &HB0 ' Next Track key
Private Const VK_MEDIA_PREV_TRACK    As Long = &HB1 ' Previous Track key
Private Const VK_MEDIA_STOP          As Long = &HB2 ' Stop Media key
Private Const VK_MEDIA_PLAY_PAUSE    As Long = &HB3 ' Play/Pause Media key
Private Const VK_LAUNCH_MAIL         As Long = &HB4 ' Start Mail key
Private Const VK_LAUNCH_MEDIA_SELECT As Long = &HB5 ' Select Media key
Private Const VK_LAUNCH_APP1         As Long = &HB6 ' Start Application 1 key
Private Const VK_LAUNCH_APP2         As Long = &HB7 ' Start Application 2 key
Private Const VK_OEM_1               As Long = &HBA ' Misc chars vary. For US standard keyboard, the ';:' key
Private Const VK_OEM_PLUS            As Long = &HBB ' For any country/region, the '+' key
Private Const VK_OEM_COMMA           As Long = &HBC ' For any country/region, the ',' key
Private Const VK_OEM_MINUS           As Long = &HBD ' For any country/region, the '-' key
Private Const VK_OEM_PERIOD          As Long = &HBE ' For any country/region, the '.' key
Private Const VK_OEM_2               As Long = &HBF ' Misc chars vary. For the US standard keyboard, the '/?' key
Private Const VK_OEM_3               As Long = &HC0 ' Misc chars vary. For the US standard keyboard, the '~' key
Private Const VK_OEM_4               As Long = &HDB ' Misc chars vary. For the US standard keyboard, the '[{' key
Private Const VK_OEM_5               As Long = &HDC ' Misc chars vary. For the US standard keyboard, the '\|' key
Private Const VK_OEM_6               As Long = &HDD ' Misc chars vary. For the US standard keyboard, the ']}' key
Private Const VK_OEM_7               As Long = &HDE ' Misc chars vary. For the US standard keyboard, the 'single-quote/double-quote' key
Private Const VK_OEM_8               As Long = &HDF ' Misc chars vary.
Private Const VK_OEM_102             As Long = &HE2 ' Angle bracket key or backslash key on the RT 102-key keyboard
Private Const VK_PROCESSKEY          As Long = &HE5 ' IME PROCESS key
Private Const VK_PACKET              As Long = &HE7 ' Pass Unicode chars as if they were keystrokes
Private Const VK_ATTN                As Long = &HF6 ' Attn key
Private Const VK_CRSEL               As Long = &HF7 ' CrSel key
Private Const VK_EXSEL               As Long = &HF8 ' ExSel key
Private Const VK_EREOF               As Long = &HF9 ' Erase EOF key
Private Const VK_PLAY                As Long = &HFA ' Play key
Private Const VK_ZOOM                As Long = &HFB ' Zoom key
Private Const VK_NONAME              As Long = &HFC ' Reserved
Private Const VK_PA1                 As Long = &HFD ' PA1 key
Private Const VK_OEM_CLEAR           As Long = &HFE ' Clear key

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

' Mouse
Public Enum eMouse
    MOUSE_LEFT = 1
    MOUSE_RIGHT = 2
    MOUSE_MIDDLE = 3
    
    MOUSE_DOWN = 4
    MOUSE_UP = 5
    MOUSE_DOUBLE = 6
End Enum

' Keyboard
Private Type tKeyboard
    Header As tRaw_Header
    Data   As tRaw_Keyboard
End Type

Dim Keyboard As tKeyboard

' Input holders
Dim Hold_Raw      As Integer
Dim Hold_Char     As String
Dim Hold_X        As Long
Dim Hold_Y        As Long
Dim Hold_L_Down   As Boolean
Dim Hold_L_Up     As Boolean
Dim Hold_L_Double As Boolean
Dim Hold_R_Down   As Boolean
Dim Hold_R_Up     As Boolean
Dim Hold_R_Double As Boolean
Dim Hold_M_Down   As Boolean
Dim Hold_M_Up     As Boolean
Dim Hold_M_Double As Boolean
Dim Hold_Wheel    As Long

Public Sub Input_Initialize()
    Raw_Initialize
End Sub

Public Sub Input_Check()
    ' Keyboard
    If Keyboard.Data.Message = WM_KEYDOWN Then
        ' Terminate
        If Keyboard.Data.VKey = VK_ESCAPE Then Terminate
    End If
    
    ' Mouse
    If Input_MouseState(MOUSE_LEFT, MOUSE_DOWN) Then frmMain.Caption = "Hold this Down L"
    If Input_MouseState(MOUSE_LEFT, MOUSE_UP) Then frmMain.Caption = "Input System"
    If Input_MouseState(MOUSE_LEFT, MOUSE_DOUBLE) Then frmMain.Caption = "Double Kill!"
End Sub

Public Sub Input_HandleRaw(ByVal lParam As Long)
    Dim Data() As Byte
    Dim Size   As Long
    
    ' Get keyboard data
    GetRawInputData lParam, RID_INPUT, ByVal 0&, Size, Len(Keyboard.Header)
        ReDim Data(Size - 1)
        
    If GetRawInputData(lParam, RID_INPUT, Data(0), Size, Len(Keyboard.Header)) <> Size Then
        ' Some form of logging
        Debug.Print "Input_Raw: (Incorrect size returned [" & lParam & ", " & Size & "])"
    End If
    
    CopyMemory Keyboard, Data(0), Len(Keyboard)
    
    ' Handle keyboard
    If Keyboard.Header.Type = RIM_TYPEKEYBOARD Then
        If Keyboard.Data.Message = WM_KEYDOWN Or Keyboard.Data.Message = WM_SYSKEYDOWN Then
            Raw_Fix
            
            Input_Raw = Keyboard.Data.VKey
        End If
    End If
End Sub

Public Sub Input_HandleMouseMove(ByVal lParam As Long)
    Dim H As String
    
    ' Extract XY
    H = Right$("00000000" & Hex(lParam), 8)
    
    Input_MouseCoord(True) = CLng("&H" & Right$(H, 4))
    Input_MouseCoord(False) = CLng("&H" & Left$(H, 4))
End Sub

Public Property Get Input_Raw() As Integer
    Input_Raw = Hold_Raw
End Property

Public Property Let Input_Raw(ByVal Value As Integer)
    Hold_Raw = Value
End Property

Public Property Get Input_Char() As String
    Input_Char = Hold_Char
End Property

Public Property Let Input_Char(ByVal Value As String)
    Hold_Char = Value
End Property

Public Property Get Input_MouseCoord(ByVal X As Boolean) As Long
    If X Then
        Input_MouseCoord = Hold_X
    Else
        Input_MouseCoord = Hold_Y
    End If
End Property

Public Property Let Input_MouseCoord(ByVal X As Boolean, ByVal Value As Long)
    If X Then
        Hold_X = Value
    Else
        Hold_Y = Value
    End If
End Property

Public Property Get Input_MouseState(ByVal Button As eMouse, ByVal State As eMouse) As Boolean
    Select Case Button

        Case MOUSE_LEFT
        If State = MOUSE_DOWN Then
            Input_MouseState = Hold_L_Down
        ElseIf State = MOUSE_UP Then
            Input_MouseState = Hold_L_Up
        ElseIf State = MOUSE_DOUBLE Then
            Input_MouseState = Hold_L_Double
        End If

        Case MOUSE_RIGHT
        If State = MOUSE_DOWN Then
            Input_MouseState = Hold_R_Down
        ElseIf State = MOUSE_UP Then
            Input_MouseState = Hold_R_Up
        ElseIf State = MOUSE_DOUBLE Then
            Input_MouseState = Hold_R_Double
        End If

        Case MOUSE_MIDDLE
        If State = MOUSE_DOWN Then
            Input_MouseState = Hold_M_Down
        ElseIf State = MOUSE_UP Then
            Input_MouseState = Hold_M_Up
        ElseIf State = MOUSE_DOUBLE Then
            Input_MouseState = Hold_M_Double
        End If

    End Select
End Property

Public Property Let Input_MouseState(ByVal Button As eMouse, ByVal State As eMouse, ByVal Value As Boolean)
    Select Case Button

        Case MOUSE_LEFT
        If State = MOUSE_DOWN Then
            Hold_L_Down = Value
        ElseIf State = MOUSE_UP Then
            Hold_L_Up = Value
        ElseIf State = MOUSE_DOUBLE Then
            Hold_L_Double = Value
        End If

        Case MOUSE_RIGHT
        If State = MOUSE_DOWN Then
            Hold_R_Down = Value
        ElseIf State = MOUSE_UP Then
            Hold_R_Up = Value
        ElseIf State = MOUSE_DOUBLE Then
            Hold_R_Double = Value
        End If

        Case MOUSE_MIDDLE
        If State = MOUSE_DOWN Then
            Hold_M_Down = Value
        ElseIf State = MOUSE_UP Then
            Hold_M_Up = Value
        ElseIf State = MOUSE_DOUBLE Then
            Hold_M_Double = Value
        End If

    End Select
End Property

Public Property Get Input_MouseWheel() As Long
    Input_MouseWheel = Hold_Wheel
End Property

Public Property Let Input_MouseWheel(ByVal Value As Long)
    Hold_Wheel = Value
End Property

Private Sub Raw_Initialize()
    Dim Device(0) As tRaw_Device
    
    ' Set up keyboard
    Device(0).UsagePage = &H1
    Device(0).Usage = &H6
    Device(0).Flags = &H0
    Device(0).hWnd = frmMain.hWnd
    
    If RegisterRawInputDevices(Device(0), 1, Len(Device(0))) = 0 Then
        Err.Raise 513, "Raw_Initialize", "Failed to register device."
        
        Terminate
    End If
End Sub

Private Sub Raw_Fix()
    Dim VKey  As Integer
    Dim Code  As Integer
    Dim E0    As Boolean
    Dim E1    As Boolean
    
    ' Get key
    VKey = Keyboard.Data.VKey
    Code = Keyboard.Data.MakeCode
    
    ' Correct virtual/scan code
    If VKey = 255 Then
        ' Discard
        Exit Sub
    ElseIf VKey = VK_SHIFT Then
        ' Left/Right shift
        VKey = MapVirtualKey(Code, MAPVK_VSC_TO_VK_EX)
    ElseIf VKey = VK_NUMLOCK Then
        ' Pause/Break/Numlock
        Code = (MapVirtualKey(VKey, MAPVK_VK_TO_VSC) Or &H100)
    End If
    
    ' Figure out E0/E1
    If Not (Keyboard.Data.Flags And RI_KEY_E0) = 0 Then E0 = True
    If Not (Keyboard.Data.Flags And RI_KEY_E1) = 0 Then E1 = True
    
    ' For E1, turn VKey into correct Code, map VK_PAUSE by hand
    If E1 Then
        If VKey = VK_PAUSE Then
            Code = &H45
        Else
            Code = MapVirtualKey(VKey, MAPVK_VK_TO_VSC)
        End If
    End If
    
    ' Sort out rest
    Select Case VKey
    
        Case VK_CONTROL
        If E0 Then
            VKey = VK_RCONTROL
        Else
            VKey = VK_LCONTROL
        End If
        
        Case VK_MENU
        If E0 Then
            VKey = VK_RMENU
        Else
            VKey = VK_LMENU
        End If
        
        Case VK_RETURN
        If E0 Then VKey = VK_SEPARATOR
        
        Case VK_INSERT
        If Not E0 Then VKey = VK_NUMPAD0
        
        Case VK_DELETE
        If Not E0 Then VKey = VK_DECIMAL
        
        Case VK_HOME
        If Not E0 Then VKey = VK_NUMPAD7
        
        Case VK_END
        If Not E0 Then VKey = VK_NUMPAD1
        
        Case VK_PRIOR
        If Not E0 Then VKey = VK_NUMPAD9
        
        Case VK_NEXT
        If Not E0 Then VKey = VK_NUMPAD3
        
        Case VK_LEFT
        If Not E0 Then VKey = VK_NUMPAD4
        
        Case VK_RIGHT
        If Not E0 Then VKey = VK_NUMPAD6
        
        Case VK_UP
        If Not E0 Then VKey = VK_NUMPAD8
        
        Case VK_DOWN
        If Not E0 Then VKey = VK_NUMPAD2
        
        Case VK_CLEAR
        If Not E0 Then VKey = VK_NUMPAD5
    
    End Select
    
    ' Set fixed key
    Keyboard.Data.VKey = VKey
    Keyboard.Data.MakeCode = Code
End Sub

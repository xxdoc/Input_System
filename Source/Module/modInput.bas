Attribute VB_Name = "modInput"
Option Explicit

' Raw Input
Private Declare Function RegisterRawInputDevices Lib "User32" (ByRef Devices As tRaw_Device, ByVal Number As Long, _
    ByVal Size As Long) As Long
Private Declare Function GetRawInputData Lib "User32" (ByVal RawInput As Long, ByVal Command As Long, ByRef Data As Any, _
    ByRef Size As Long, ByVal SizeHeader As Long) As Long

' GetRawInputData.Command
Private Const RID_INPUT As Long = &H10000003

' Raw Structs
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
Public Enum eVirtualKey
    VK_LBUTTON = &H1              ' Left mouse button
    VK_RBUTTON = &H2              ' Right mouse button
    VK_CANCEL = &H3               ' Control-break processing
    VK_MBUTTON = &H4              ' Middle mouse button (three-button mouse)
    VK_XBUTTON1 = &H5             ' X1 mouse button
    VK_XBUTTON2 = &H6             ' X2 mouse button
    VK_BACK = &H8                 ' BACKSPACE key
    VK_TAB = &H9                  ' TAB key
    VK_CLEAR = &HC                ' CLEAR key
    VK_RETURN = &HD               ' ENTER key
    VK_SHIFT = &H10               ' SHIFT key
    VK_CONTROL = &H11             ' CTRL key
    VK_MENU = &H12                ' ALT key
    VK_PAUSE = &H13               ' PAUSE key
    VK_CAPITAL = &H14             ' CAPS LOCK key
    VK_KANA = &H15                ' IME Kana mode
    VK_JUNJA = &H17               ' IME Junja mode
    VK_FINAL = &H18               ' IME final mode
    VK_HANJA = &H19               ' IME Hanja mode
    VK_ESCAPE = &H1B              ' ESC key
    VK_CONVERT = &H1C             ' IME convert
    VK_NONCONVERT = &H1D          ' IME nonconvert
    VK_ACCEPT = &H1E              ' IME accept
    VK_MODECHANGE = &H1F          ' IME mode change request
    VK_SPACE = &H20               ' SPACEBAR
    VK_PRIOR = &H21               ' PAGE UP key
    VK_NEXT = &H22                ' PAGE DOWN key
    VK_END = &H23                 ' END key
    VK_HOME = &H24                ' HOME key
    VK_LEFT = &H25                ' LEFT ARROW key
    VK_UP = &H26                  ' UP ARROW key
    VK_RIGHT = &H27               ' RIGHT ARROW key
    VK_DOWN = &H28                ' DOWN ARROW key
    VK_SELECT = &H29              ' SELECT key
    VK_PRINT = &H2A               ' PRINT key
    VK_EXECUTE = &H2B             ' EXECUTE key
    VK_SNAPSHOT = &H2C            ' PRINT SCREEN key
    VK_INSERT = &H2D              ' INS key
    VK_DELETE = &H2E              ' DEL key
    VK_HELP = &H2F                ' HELP key
    VK_0 = &H30                   ' 0 key
    VK_1 = &H31                   ' 1 key
    VK_2 = &H32                   ' 2 key
    VK_3 = &H33                   ' 3 key
    VK_4 = &H34                   ' 4 key
    VK_5 = &H35                   ' 5 key
    VK_6 = &H36                   ' 6 key
    VK_7 = &H37                   ' 7 key
    VK_8 = &H38                   ' 8 key
    VK_9 = &H39                   ' 9 key
    VK_A = &H41                   ' A key
    VK_B = &H42                   ' B key
    VK_C = &H43                   ' C key
    VK_D = &H44                   ' D key
    VK_E = &H45                   ' E key
    VK_F = &H46                   ' F key
    VK_G = &H47                   ' G key
    VK_H = &H48                   ' H key
    VK_I = &H49                   ' I key
    VK_J = &H4A                   ' J key
    VK_K = &H4B                   ' K key
    VK_L = &H4C                   ' L key
    VK_M = &H4D                   ' M key
    VK_N = &H4E                   ' N key
    VK_O = &H4F                   ' O key
    VK_P = &H50                   ' P key
    VK_Q = &H51                   ' Q key
    VK_R = &H52                   ' R key
    VK_S = &H53                   ' S key
    VK_T = &H54                   ' T key
    VK_U = &H55                   ' U key
    VK_V = &H56                   ' V key
    VK_W = &H57                   ' W key
    VK_X = &H58                   ' X key
    VK_Y = &H59                   ' Y key
    VK_Z = &H5A                   ' Z key
    VK_LWIN = &H5B                ' Left Windows key (Natural keyboard)
    VK_RWIN = &H5C                ' Right Windows key (Natural keyboard)
    VK_APPS = &H5D                ' Applications key (Natural keyboard)
    VK_SLEEP = &H5F               ' Computer Sleep key
    VK_NUMPAD0 = &H60             ' Numeric keypad 0 key
    VK_NUMPAD1 = &H61             ' Numeric keypad 1 key
    VK_NUMPAD2 = &H62             ' Numeric keypad 2 key
    VK_NUMPAD3 = &H63             ' Numeric keypad 3 key
    VK_NUMPAD4 = &H64             ' Numeric keypad 4 key
    VK_NUMPAD5 = &H65             ' Numeric keypad 5 key
    VK_NUMPAD6 = &H66             ' Numeric keypad 6 key
    VK_NUMPAD7 = &H67             ' Numeric keypad 7 key
    VK_NUMPAD8 = &H68             ' Numeric keypad 8 key
    VK_NUMPAD9 = &H69             ' Numeric keypad 9 key
    VK_MULTIPLY = &H6A            ' Multiply key
    VK_ADD = &H6B                 ' Add key
    VK_SEPARATOR = &H6C           ' Separator key
    VK_SUBTRACT = &H6D            ' Subtract key
    VK_DECIMAL = &H6E             ' Decimal key
    VK_DIVIDE = &H6F              ' Divide key
    VK_F1 = &H70                  ' F1 key
    VK_F2 = &H71                  ' F2 key
    VK_F3 = &H72                  ' F3 key
    VK_F4 = &H73                  ' F4 key
    VK_F5 = &H74                  ' F5 key
    VK_F6 = &H75                  ' F6 key
    VK_F7 = &H76                  ' F7 key
    VK_F8 = &H77                  ' F8 key
    VK_F9 = &H78                  ' F9 key
    VK_F10 = &H79                 ' F10 key
    VK_F11 = &H7A                 ' F11 key
    VK_F12 = &H7B                 ' F12 key
    VK_F13 = &H7C                 ' F13 key
    VK_F14 = &H7D                 ' F14 key
    VK_F15 = &H7E                 ' F15 key
    VK_F16 = &H7F                 ' F16 key
    VK_F17 = &H80                 ' F17 key
    VK_F18 = &H81                 ' F18 key
    VK_F19 = &H82                 ' F19 key
    VK_F20 = &H83                 ' F20 key
    VK_F21 = &H84                 ' F21 key
    VK_F22 = &H85                 ' F22 key
    VK_F23 = &H86                 ' F23 key
    VK_F24 = &H87                 ' F24 key
    VK_NUMLOCK = &H90             ' NUM LOCK key
    VK_SCROLL = &H91              ' SCROLL LOCK key
    VK_LSHIFT = &HA0              ' Left SHIFT key
    VK_RSHIFT = &HA1              ' Right SHIFT key
    VK_LCONTROL = &HA2            ' Left CONTROL key
    VK_RCONTROL = &HA3            ' Right CONTROL key
    VK_LMENU = &HA4               ' Left MENU key
    VK_RMENU = &HA5               ' Right MENU key
    VK_BROWSER_BACK = &HA6        ' Browser Back key
    VK_BROWSER_FORWARD = &HA7     ' Browser Forward key
    VK_BROWSER_REFRESH = &HA8     ' Browser Refresh key
    VK_BROWSER_STOP = &HA9        ' Browser Stop key
    VK_BROWSER_SEARCH = &HAA      ' Browser Search key
    VK_BROWSER_FAVORITES = &HAB   ' Browser Favorites key
    VK_BROWSER_HOME = &HAC        ' Browser Start and Home key
    VK_VOLUME_MUTE = &HAD         ' Volume Mute key
    VK_VOLUME_DOWN = &HAE         ' Volume Down key
    VK_VOLUME_UP = &HAF           ' Volume Up key
    VK_MEDIA_NEXT_TRACK = &HB0    ' Next Track key
    VK_MEDIA_PREV_TRACK = &HB1    ' Previous Track key
    VK_MEDIA_STOP = &HB2          ' Stop Media key
    VK_MEDIA_PLAY_PAUSE = &HB3    ' Play/Pause Media key
    VK_LAUNCH_MAIL = &HB4         ' Start Mail key
    VK_LAUNCH_MEDIA_SELECT = &HB5 ' Select Media key
    VK_LAUNCH_APP1 = &HB6         ' Start Application 1 key
    VK_LAUNCH_APP2 = &HB7         ' Start Application 2 key
    VK_OEM_1 = &HBA               ' Misc chars vary. For US standard keyboard, the ';:' key
    VK_OEM_PLUS = &HBB            ' For any country/region, the '+' key
    VK_OEM_COMMA = &HBC           ' For any country/region, the ',' key
    VK_OEM_MINUS = &HBD           ' For any country/region, the '-' key
    VK_OEM_PERIOD = &HBE          ' For any country/region, the '.' key
    VK_OEM_2 = &HBF               ' Misc chars vary. For the US standard keyboard, the '/?' key
    VK_OEM_3 = &HC0               ' Misc chars vary. For the US standard keyboard, the '~' key
    VK_OEM_4 = &HDB               ' Misc chars vary. For the US standard keyboard, the '[{' key
    VK_OEM_5 = &HDC               ' Misc chars vary. For the US standard keyboard, the '\|' key
    VK_OEM_6 = &HDD               ' Misc chars vary. For the US standard keyboard, the ']}' key
    VK_OEM_7 = &HDE               ' Misc chars vary. For the US standard keyboard, the 'single-quote/double-quote' key
    VK_OEM_8 = &HDF               ' Misc chars vary.
    VK_OEM_102 = &HE2             ' Angle bracket key or backslash key on the RT 102-key keyboard
    VK_PROCESSKEY = &HE5          ' IME PROCESS key
    VK_PACKET = &HE7              ' Pass Unicode chars as if they were keystrokes
    VK_ATTN = &HF6                ' Attn key
    VK_CRSEL = &HF7               ' CrSel key
    VK_EXSEL = &HF8               ' ExSel key
    VK_EREOF = &HF9               ' Erase EOF key
    VK_PLAY = &HFA                ' Play key
    VK_ZOOM = &HFB                ' Zoom key
    VK_NONAME = &HFC              ' Reserved
    VK_PA1 = &HFD                 ' PA1 key
    VK_OEM_CLEAR = &HFE           ' Clear key
End Enum

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
    
    Key    As Integer
    Char   As String
End Type

Dim Keyboard As tKeyboard

' Mouse
Private Type tMouse_State
    Down   As Boolean
    Up     As Boolean
    Double As Boolean
End Type

Private Type tMouse
    X      As Long
    Y      As Long
    
    Left   As tMouse_State
    Right  As tMouse_State
    Middle As tMouse_State
    
    Wheel  As Long
End Type

Dim Mouse As tMouse

Public Sub Input_Initialize(ByVal hWnd As Long)
    Raw_Initialize hWnd
    
    ' Initial state
    Input_MouseState(MOUSE_LEFT, MOUSE_UP) = True
    Input_MouseState(MOUSE_RIGHT, MOUSE_UP) = True
    Input_MouseState(MOUSE_MIDDLE, MOUSE_UP) = True
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
    
    CopyMemory Keyboard, Data(0), Len(Keyboard.Header) + Len(Keyboard.Data)
    
    ' Handle keyboard
    If Keyboard.Header.Type = RIM_TYPEKEYBOARD Then
        If Keyboard.Data.Message = WM_KEYDOWN Or Keyboard.Data.Message = WM_SYSKEYDOWN Then
            Raw_Fix
            
            Keyboard.Key = Keyboard.Data.VKey
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

Public Property Get Input_Key() As Integer
    Input_Key = Keyboard.Key
End Property

Public Property Let Input_Key(ByVal Value As Integer)
    Keyboard.Key = Value
End Property

Public Property Get Input_Char() As String
    Input_Char = Keyboard.Char
End Property

Public Property Let Input_Char(ByVal Value As String)
    Keyboard.Char = Value
End Property

Public Property Get Input_MouseCoord(ByVal X As Boolean) As Long
    If X Then
        Input_MouseCoord = Mouse.X
    Else
        Input_MouseCoord = Mouse.Y
    End If
End Property

Public Property Let Input_MouseCoord(ByVal X As Boolean, ByVal Value As Long)
    If X Then
        Mouse.X = Value
    Else
        Mouse.Y = Value
    End If
End Property

Public Property Get Input_MouseState(ByVal Button As eMouse, ByVal State As eMouse) As Boolean
    Select Case Button

        Case MOUSE_LEFT
        If State = MOUSE_DOWN Then
            Input_MouseState = Mouse.Left.Down
        ElseIf State = MOUSE_UP Then
            Input_MouseState = Mouse.Left.Up
        ElseIf State = MOUSE_DOUBLE Then
            Input_MouseState = Mouse.Left.Double
        End If
        
        Case MOUSE_RIGHT
        If State = MOUSE_DOWN Then
            Input_MouseState = Mouse.Right.Down
        ElseIf State = MOUSE_UP Then
            Input_MouseState = Mouse.Right.Up
        ElseIf State = MOUSE_DOUBLE Then
            Input_MouseState = Mouse.Right.Double
        End If
        
        Case MOUSE_MIDDLE
        If State = MOUSE_DOWN Then
            Input_MouseState = Mouse.Middle.Down
        ElseIf State = MOUSE_UP Then
            Input_MouseState = Mouse.Middle.Up
        ElseIf State = MOUSE_DOUBLE Then
            Input_MouseState = Mouse.Middle.Double
        End If

    End Select
End Property

Public Property Let Input_MouseState(ByVal Button As eMouse, ByVal State As eMouse, ByVal Value As Boolean)
    Select Case Button

        Case MOUSE_LEFT
        If State = MOUSE_DOWN Then
            Mouse.Left.Down = Value
        ElseIf State = MOUSE_UP Then
            Mouse.Left.Up = Value
        ElseIf State = MOUSE_DOUBLE Then
            Mouse.Left.Double = Value
        End If

        Case MOUSE_RIGHT
        If State = MOUSE_DOWN Then
            Mouse.Right.Down = Value
        ElseIf State = MOUSE_UP Then
            Mouse.Right.Up = Value
        ElseIf State = MOUSE_DOUBLE Then
            Mouse.Right.Double = Value
        End If

        Case MOUSE_MIDDLE
        If State = MOUSE_DOWN Then
            Mouse.Middle.Down = Value
        ElseIf State = MOUSE_UP Then
            Mouse.Middle.Up = Value
        ElseIf State = MOUSE_DOUBLE Then
            Mouse.Middle.Double = Value
        End If

    End Select
End Property

Public Property Get Input_MouseWheel() As Long
    Input_MouseWheel = Mouse.Wheel
End Property

Public Property Let Input_MouseWheel(ByVal Value As Long)
    Mouse.Wheel = Value
End Property

Private Sub Raw_Initialize(ByVal hWnd As Long)
    Dim Device(0) As tRaw_Device
    
    ' Set up keyboard
    Device(0).UsagePage = &H1
    Device(0).Usage = &H6
    Device(0).Flags = &H0
    Device(0).hWnd = hWnd
    
    If RegisterRawInputDevices(Device(0), 1, Len(Device(0))) = 0 Then
        Err.Raise 513, "Raw_Initialize", "Failed to register device."
        
        Terminate
    End If
End Sub

Private Sub Raw_Fix()
    Dim VKey As Integer
    Dim Code As Integer
    Dim E0   As Boolean
    Dim E1   As Boolean
    
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

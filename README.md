# Input System
Input system built upon the Raw Input API and WM_* messages in Visual Basic 6.

**WM_INPUT (Raw Input):** For key presses.

**WM_CHAR:** For text input.

**WM_MOUSEMOVE, WM_BUTTON:** For mouse input. Supports Left, Right & Middle mouse buttons.

**WM_MOUSEWHEEL:** Not fully implemented. See Task List.

## Documentation
### modMain
Used to setup the input system and act as an example.

### modWindows
Contains Windows API functions, constants and the **WindowProc** function. **modWindows** has two functions that need to be called: 

```vb
Public Sub Windows_Initialize(ByVal hWnd As Long)
```
Call on application startup and pass in the hWnd of the main Form.

And

```vb
Public Sub Windows_Terminate()
```
Call when terminating your application.

### modInput
Contains all the input related code and declarations for the Raw Input API. **modInput** has one function that needs to be called:

```vb
Public Sub Input_Initialize(ByVal hWnd As Long)
```
Call on application startup and pass in the hWnd of the main Form.

**modInput** also provides five getter/setters to mouse and keyboard inputs:

```vb
Input_Key() ' Provides access to the virtual key retrieved by WM_INPUT (Raw Input).
Input_Char() ' Provides access to the character string retrieved by WM_CHAR.
Input_MouseCoord(X) ' Provides access to the X, Y position retrieved by WM_MOUSEMOVE. X is a boolean. When True, Input_MouseCoord will return the X coord, if False, the Y coord will be returned.
Input_MouseState(Button, State) ' Provides access to the state of a mouse button retrieved by the WM_*BUTTON* messages. Values for Button and State can be found in the enum, eMouse.
Input_MouseWheel() ' Provides access to the mouse wheel retrieved by WM_MOUSEWHEEL. Incomplete.
```

Those five getter/setters can be used to check for key/mouse presses. Example can be found in **Input_Check (in modInput):**

```vb
' Keyboard
If Keyboard.Data.Message = WM_KEYDOWN Then
    ' Terminate
    If Keyboard.Data.VKey = VK_ESCAPE Then Terminate
End If

' Mouse
If Input_MouseState(MOUSE_LEFT, MOUSE_DOWN) Then frmMain.Caption = "Hold this Down L"
If Input_MouseState(MOUSE_LEFT, MOUSE_UP) Then frmMain.Caption = "Input System"
If Input_MouseState(MOUSE_LEFT, MOUSE_DOUBLE) Then frmMain.Caption = "Double Kill!"
```

## Task List
- [ ] Mouse Wheel functionality (GET_WHEEL_DELTA_WPARAM(wParam))
- [ ] Binding system like in games.

## Credits
Stefan Reinalter's blog post on [Properly handling keyboard input](https://blog.molecular-matters.com/2011/09/05/properly-handling-keyboard-input/) which acted as inspiration for this project and the code included in the blog post makes up the **Raw_Fix** function.

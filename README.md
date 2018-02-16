# Input System
Input system built on Raw Input API and WM_* messages in Visual Basic 6.

**WM_INPUT (Raw Input):** For key presses.

**WM_CHAR:** For text input.

**WM_MOUSEMOVE, WM_BUTTONDOWN/UP/DOUBLE:** For mouse input. Supports Left, Right & Middle mouse buttons.

**WM_MOUSEWHEEL:** Not fully implemented. See Task List.

## Documentation
**Input_Check (in modInput):** Used to check for key/mouse presses.

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

More...

## Task List
- [ ] Actually implement Mouse Wheel functionality (GET_WHEEL_DELTA_WPARAM(wParam))
- [ ] Binding system like in games.

## Credits
Stefan Reinalter's blog post on [Properly handling keyboard input](https://blog.molecular-matters.com/2011/09/05/properly-handling-keyboard-input/) which acted as inspiration for this project and the code included in the blog post makes up the **Raw_Fix** function.

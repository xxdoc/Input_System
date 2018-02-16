# Input System
Input system built on Raw Input API and WM_* messages in Visual Basic 6.

**WM_INPUT (Raw Input):**

For key presses.

**WM_CHAR:**

For text input.

**WM_MOUSEMOVE, WM_BUTTONDOWN, WM_BUTTONUP, WM_BUTTONDOUBLE:**

For mouse input. Supports Left, Right & Middle mouse buttons.

**WM_MOUSEWHEEL:**

Not fully implemented. See Todo list.

## Documentation
**Input_Check (in modInput)**
Can be used to bind/trigger actions for key/mouse presses using if statements.

More...

## Possible Todo
- [ ] Actually implement Mouse Wheel functionality (GET_WHEEL_DELTA_WPARAM(wParam))
- [ ] Binding system like in games.

## Credits
Stefan Reinalter's blog post on [Properly handling keyboard input](https://blog.molecular-matters.com/2011/09/05/properly-handling-keyboard-input/) which acted as inspiration for this project and the code included in the blog post makes up the **Raw_Fix** function.

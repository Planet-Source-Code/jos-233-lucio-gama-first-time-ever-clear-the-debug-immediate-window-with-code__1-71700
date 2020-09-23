<div align="center">

## First time ever\! Clear the Debug/Immediate window with code\!


</div>

### Description

For the first time ever, a code that can clear the debug window from code!

It's actually a cheat, since VB5-VB6 doesn't allow to clear the window without stopping the running project - the program sets the focus on the debug window, go to the last char and start printing a lot of empty lines - effectively getting rid of all clutter visible on the window.
 
### More Info
 
Win32 API knowledge


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jos&\#233; Lucio Gama](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jos-233-lucio-gama.md)
**Level**          |Advanced
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Debugging and Error Handling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/debugging-and-error-handling__1-26.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jos-233-lucio-gama-first-time-ever-clear-the-debug-immediate-window-with-code__1-71700/archive/master.zip)

### API Declarations

FindWindow, SetForegroundWindow and SetFocus


### Source Code

```
Put this on the top of your form:
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Then, paste this on your code:
Private Sub clearDebug()
  'this will try to get the handle to your
  'Immediate window. To make this work on
  'VB4, you can change the string "Immediate"
  'to "Debug"
  parent_hwnd = FindWindow(vbNullString, "Immediate")
  If parent_hwnd = 0 Then Exit Sub
  ' Set the focus on the debug window
  SetFocusAPI parent_hwnd
  'go to the last line / position on the window
  '(same as pressing CTRL + END on your keyboard
  SendKeys "^{END}", True
  'you can adjust the number of lines
  'printed according to your Immediate
  'window size
  For i = 1 To 100
    Debug.Print ""
  Next
  'give the focus back to your program!
  SetForegroundWindow Me.hwnd
End Sub
Then just call clearDebug() anywhere in your code  and that it!
```


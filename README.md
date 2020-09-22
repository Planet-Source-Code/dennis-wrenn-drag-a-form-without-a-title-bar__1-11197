<div align="center">

## Drag a form WITHOUT a title bar\!


</div>

### Description

With this code you can easily drag a form without a titlebar.
 
### More Info
 
If you dont add If Button = 1 etc.. then if you left click, then right click the form will continue to move even though you arent clicking, its like the form is stuck to your mouse


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dennis Wrenn](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dennis-wrenn.md)
**Level**          |Beginner
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dennis-wrenn-drag-a-form-without-a-title-bar__1-11197/archive/master.zip)

### API Declarations

```
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
```


### Source Code

```
'code
Private Sub FormDrag(frm As Form)
  ReleaseCapture
  Call SendMessage(frm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub
'usage:
'put in MouseDown even of almost anything.
'a form a label, a command button, anything will work.
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then Call FormDrag(Me)
End Sub
'If you dont add If Button = 1 etc..
'then if you left click, then right
'click the form will continue to
'move even though you arent clicking,
'its like the form is stuck to your mouse
```


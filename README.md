<div align="center">

## Resize a form based on display settings


</div>

### Description

Resizes a form based on the display settings
 
### More Info
 
How to declare an API in global module

You may have to play with the size to find the right percentage


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[ROBERT JACKSON](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/robert-jackson.md)
**Level**          |Advanced
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/robert-jackson-resize-a-form-based-on-display-settings__1-6723/archive/master.zip)

### API Declarations

```
'Declare this code in your global module
Global Const SM_CXSCREEN = 0
Global Const SM_CYSCREEN = 1
 Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
```


### Source Code

```

'Use this sub in the form you are coding, it could be improved to be a global procedure
'and pass the form as an argument.
Private Sub Resizeall()
Dim Ctl As Control
 Dim X As Integer
   Dim Size As Double
   ScreenX = GetSystemMetrics(SM_CXSCREEN)
 ScreenY = GetSystemMetrics(SM_CYSCREEN)
' this picks out the display settings.
Select Case ScreenX
    Case 640
         'size = 0.67
        Size = 0.64
    Case 800
        Size = 0.72
    Case 1024
        Exit Sub
    Case 1280
      'Exit Sub
      Size = 1.25
    Case Else
      Exit Sub
  End Select
  'Me.Height = Me.Height * size
  'Me.Top = Me.Top * size
  'Me.Width = Me.Width * size
  'Me.Left = Me.Left * size
  For Each Ctl In Me.Controls
   Ctl.Height = Ctl.Height * Size
   Ctl.Width = Ctl.Width * Size
   Ctl.Top = Ctl.Top * Size
   Ctl.Left = Ctl.Left * Size
   If TypeOf Ctl Is TextBox Or TypeOf Ctl Is Label Or TypeOf Ctl Is CommandButton Then
   'Ctl.SizeToFit
   Ctl.FontName = "Arial"
   Ctl.FontSize = 6.7
   If TypeOf Ctl Is CommandButton Then
   Ctl.FontName = "Arial"
   Ctl.FontSize = 5
   End If
   End If
  'SizeToFit
  Next Ctl
End Sub
```


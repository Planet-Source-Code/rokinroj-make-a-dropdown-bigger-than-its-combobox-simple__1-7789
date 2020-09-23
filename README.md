<div align="center">

## Make a dropdown bigger than its combobox \*\*Simple\*\*


</div>

### Description

The only code I could find for this used tons of api and was a bit difficult (at least for me) this code is only 3 lines and works well.It

lets you save room on your form by making your dropdown bigger than its combobox. My use for this was as a state field on a form. The box showed the two letter abbreviation for each state, but if you dropped down the box it showed the full state name. I'm sure you will find your own uses for this
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rokinroj ](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rokinroj.md)
**Level**          |Beginner
**User Rating**    |4.2 (38 globes from 9 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rokinroj-make-a-dropdown-bigger-than-its-combobox-simple__1-7789/archive/master.zip)

### API Declarations

```
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Const CB_SETDROPPEDWIDTH = &H160
```


### Source Code

```
Private Sub Form_Load()
  SendMessage cboState.hwnd, CB_SETDROPPEDWIDTH, 135, 0
'be sure to either carry the line down with a _, or put it all on one line. The complete line should start with SendMessage and end with 0
End Sub
```


<div align="center">

## Locate Entry in List/Combo using API, no looping


</div>

### Description

Locate an entry in combo or listbox using API call, rather than looping through all entries.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mark Compton](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mark-compton.md)
**Level**          |Advanced
**User Rating**    |3.9 (27 globes from 7 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mark-compton-locate-entry-in-list-combo-using-api-no-looping__1-33729/archive/master.zip)





### Source Code

```
'Declarations
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158
Declare Function SendMessageByString& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String)
'Functions
Function InList(sStringToFind as string, lstListBox As ListBox) As Boolean
 InList = SendMessageByString(lstListBox.hwnd, LB_FINDSTRING, -1, sStringToFind) >= 0
End Function
Function InCombo(sStringToFind, cbCombo As ComboBox) As Boolean
 InCombo = SendMessageByString(cbCombo.hwnd, CB_FINDSTRING, -1, sStringToFind) >= 0
End Function
```


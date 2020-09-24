<div align="center">

## Listbox Reorder


</div>

### Description

A function that will allow you to move the selected item in any listbox up or down in the listbox.
 
### More Info
 
MoveListItem(ListBoxName, (0/1))

(See Code Comments)

-1 if nothing is selected

X is the new position of the selected item


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Troy J](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/troy-j.md)
**Level**          |Beginner
**User Rating**    |4.3 (17 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/troy-j-listbox-reorder__1-13213/archive/master.zip)





### Source Code

```
Public Function MoveListItem(LstBox As Object, WhatDir As Integer)
  'WhatDir = 0 up, 1 down
  'Returns -1 if nothing is selected
  'Returns current position otherwise
  Dim CurPos As Integer, CurData As String, NewPos As Integer
  CurPos = LstBox.ListIndex
  If CurPos < 0 Then MoveListItem = -1: Exit Function
  CurData = LstBox.List(CurPos)
  If WhatDir = 0 Then
    'Move Up
    If (CurPos - 1) < 0 Then NewPos = (LstBox.ListCount - 1) Else NewPos = (CurPos - 1)
  Else
    'Move Down
    If (CurPos + 1) > (LstBox.ListCount - 1) Then NewPos = 0 Else NewPos = (CurPos + 1)
  End If
  LstBox.RemoveItem (CurPos)
  LstBox.AddItem CurData, NewPos
  LstBox.Selected(NewPos) = True
  MoveListItem = NewPos
End Function
```


<div align="center">

## Simple text effect


</div>

### Description

very Simple but interesting text effect for the form caption
 
### More Info
 
needs a timer called timer1


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chris Martian](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chris-martian.md)
**Level**          |Beginner
**User Rating**    |4.2 (25 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chris-martian-simple-text-effect__1-34444/archive/master.zip)





### Source Code

```
Dim x As Integer
Private Sub Form_Load()
x = 0
End Sub
Private Sub Timer1_Timer()
Me.Caption = Mid$("Created by Chris Martian", 1, x)
'find how much of the caption should be displayed
x = x + 1
'add 1 to number of letters to be displayed next time
If x > Len("Created by chris martian") Then x = 0
'if it gets to the end then start again
End Sub
```


<div align="center">

## Search a String and Highlight \(SrchString\)


</div>

### Description

This source code will search for a certain string from a TextBox (can be modified to search from any string-bearing control, etc.). Once the string is found, a message box will appear letting you know that the string was found, and then the string will be highlighted.
 
### More Info
 
1) Start a new Standard EXE project.

2) Add two (2) TextBox controls to the Form.

3) Erase the contents of the TextBox controls. Alternately, you can set Text2's Multiline property to True, as it is going to be the TextBox to search for the string in.

4) Add one (1) CommandButton control to the Form.

5) Set the CommandButton Caption property to "&Search"

6) Click the View Code button; copy and paste the source code below.

7) Run the application.

8) Visit my website! http://www.matnet.com/~pyrosoft


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[silverx10](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/silverx10.md)
**Level**          |Beginner
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/silverx10-search-a-string-and-highlight-srchstring__1-13502/archive/master.zip)





### Source Code

```
'Text2 is the TextBox to search for the string in.
Dim I as Integer
Private Sub Command1_Click()
 For I = 1 To Len(Text2)
  If Mid(Text2, I, Len(Text1)) = Text1 Then
   MsgBox "String located and highlighted."
   Text2.SetFocus
   Text2.SelStart = I - 1
   Text2.SelLength = Len(Text1)
  End If
 Next I
End Sub()
```


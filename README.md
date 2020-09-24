<div align="center">

## Read / Write Combo Boxs


</div>

### Description

This code will let you read in values from a file of your choice, probably a .ini or .txt into a combobox. It will also let you save the contents of the combobox to a file of your choice.

Example:

Call WriteCombo(combo1, "C:/example.ini")

or

Call ReadCombo(combo1, "C:/example.ini")
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[T\-Unit](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/t-unit.md)
**Level**          |Intermediate
**User Rating**    |4.0 (12 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/t-unit-read-write-combo-boxs__1-53501/archive/master.zip)





### Source Code

```
Public Sub ReadCombo(combobox As combobox, Filename As String)
  On Error GoTo Err
  Open Filename For Input As #1
  Do While Not EOF(1)
    Input #1, lstinput
    combobox.AddItem lstinput
  Loop
  Close #1
  Exit Sub
Err:
  MsgBox "Error In ReadCombo" & Chr(13) & Chr(13) & Err.Number _
  & " - " & Err.Description, vbCritical, "Error"
  Exit Sub
End Sub
Public Sub WriteCombo(combobox As combobox, Filename As String)
  If combobox.ListCount <= 0 Then
    MsgBox "Combobox is empty - cannot write To file!", vbCritical, "Error"
    End
  End If
  On Error GoTo Err
  Open Filename For Output As #1
  For i = 0 To combobox.ListCount - 1
    Print #1, combobox.List(i)
  Next
  Close #1
  Exit Sub
Err:
  MsgBox "Error In WriteCombo" & Chr(13) & Chr(13) & Err.Number _
  & " - " & Err.Description, vbCritical, "Error"
  Exit Sub
End Sub
```


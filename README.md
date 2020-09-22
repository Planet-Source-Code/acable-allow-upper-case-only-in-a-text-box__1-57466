<div align="center">

## Allow UPPER\-CASE only in a Text Box


</div>

### Description

Easy trick to allow only text and converts it to upper-case. (Backspace is optional)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Acable](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/acable.md)
**Level**          |Beginner
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, VBA MS Access, VBA MS Excel
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/acable-allow-upper-case-only-in-a-text-box__1-57466/archive/master.zip)





### Source Code

```
Private Sub Text1_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8
      'Delete this line and backspace will not be allowed
    Case 65 To 90
      'UPPER-CASE is allowed, do nothing here
    Case 97 To 122
      'This is lower case. By turning off Bit-5 lower-case will converted to UPPER-CASE
      KeyAscii = KeyAscii Xor 32
    Case Else
      'Other keys are not allowed. Set KeyAscii to 0 and nothing will be printed.
      KeyAscii = 0
  End Select
End Sub
```


<div align="center">

## resolve relative link


</div>

### Description

This is a simple function to resolve a relative link. ie turns www.a/b/c/../e.htm into www.a/b/e.htm

Made this for a web crawler I am working on.

Hope you find this usefull.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[RegX](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/regx.md)
**Level**          |Beginner
**User Rating**    |4.3 (17 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/regx-resolve-relative-link__1-33877/archive/master.zip)





### Source Code

```
Public Function cleanurl(url As String)
Dim x As Long
Dim y As Long
check:
x = InStr(1, url, "../", vbBinaryCompare)
If x = 1 Then
url = Mid(url, 4)
ElseIf x > 1 Then
 y = InStrRev(url, "/", x - 2, vbBinaryCompare)
 If y > 0 Then
 url = Mid(url, 1, y) & Mid(url, x + 3)
 Else
  url = Mid(url, 4)
 Debug.Print url
 End If
GoTo check:
End If
cleanurl = url
End Function
```


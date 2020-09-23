<div align="center">

## Find FULL App Path \+ App Extension\!


</div>

### Description

This simple 10 lines of code (including 1 API call) will return the FULL application path INCLUDING the application extension. Very easy to use. Tired of using App.path & "\" & app.exename & ".exe" when you dont know for sure that your extention will be .exe? This will return it all. Votes are welcome
 
### More Info
 
Just call FullAppName instead of putting in App.path & "\" & app.exename & ".exe"

Simply place the API Code and the function in your project. Use FullAppName every time you refer to your self in your code.

A string containing the full application path and extention

No Side Affects


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Steve Sellers](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/steve-sellers.md)
**Level**          |Beginner
**User Rating**    |4.0 (32 globes from 8 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/steve-sellers-find-full-app-path-app-extension__1-45601/archive/master.zip)





### Source Code

```
Private Declare Function GetModuleFileName Lib "kernel32" _
 Alias "GetModuleFileNameA" (ByVal hModule As Long, _
 ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Function FullAppName() As String
 Dim modName As String * 256
 Dim i As Long
 i = GetModuleFileName(App.hInstance, modName, Len(modName))
 FullAppName = Left$(modName, i)
End Function
```


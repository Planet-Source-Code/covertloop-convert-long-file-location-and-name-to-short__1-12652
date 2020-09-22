<div align="center">

## Convert Long File Location And Name To Short


</div>

### Description

This code enables you to convert a long file location and filename to a short file location and filename. Example: C:\Program Files\Test.txt

will become C:\Progra~1\Test.txt
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[CovertLoop](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/covertloop.md)
**Level**          |Beginner
**User Rating**    |4.3 (17 globes from 4 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/covertloop-convert-long-file-location-and-name-to-short__1-12652/archive/master.zip)

### API Declarations

```
ADD TO A MODULE
Declare Function GetShortPathName Lib "kernel32" Alias _
"GetShortPathNameA" (ByVal lpszLongPath As String, ByVal _
lpszShortPath As String, ByVal cchBuffer As Long) As Long
```


### Source Code

```
Add The Declarations Above To A Module. Then put this whenever you want to perform the conversion:
Dim sFile As String, sShortFile As String * 67
Dim lRet As Long
sFile = "C:\Program Files\Test.txt" 'Long File Location/Name
lRet = GetShortPathName(sFile, sShortFile, Len(sShortFile))
sFile = Left(sShortFile, lRet)
Text1.Text = sFile
```


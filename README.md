<div align="center">

## INI Read and INI Write \- Simple and Error Free


</div>

### Description

INI Read and INI Write made simple. I included examples in each function. If you like my codes, then please vote for me. But only if you like them. Thanks =)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jeffrey C\. Tatum](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jeffrey-c-tatum.md)
**Level**          |Beginner
**User Rating**    |4.0 (24 globes from 6 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jeffrey-c-tatum-ini-read-and-ini-write-simple-and-error-free__1-6132/archive/master.zip)

### API Declarations

```
'INI Read and Write
Declare Function GetPrivateProfileString Lib "kernel32" _
Alias "GetPrivateProfileStringA" (ByVal lpApplicationName _
As String, lpKeyName As Any, ByVal lpDefault As String, _
ByVal lpRetunedString As String, ByVal nSize As Long, _
ByVal lpFilename As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" _
Alias "WritePrivateProfileStringA" (ByVal lpApplicationName _
As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
ByVal lplFileName As String) As Long
```


### Source Code

```
Public Function INIRead(iAppName As String, iKeyName As String, iFileName As String) As String
'Example:
  'x = INIRead("boot", "shell", "C:\WINDOWS\system.ini")
  Dim iStr As String
  iStr = String(255, Chr(0))
  INIRead = Left(iStr, GetPrivateProfileString(iAppName, ByVal iKeyName, "", iStr, Len(iStr), iFileName))
End Function
Public Function INIWrite(iAppName As String, iKeyName As String, iKeyString As String, iFileName As String)
'Example:
  'x = INIWrite("boot", "shell", "Explorer.exe", "C:\WINDOWS\system.ini")
r% = WritePrivateProfileString(iAppName, iKeyName, iKeyString, iFileName)
End Function
```


<div align="center">

## AssociateFileType


</div>

### Description

Associate a file type with a program in windows95.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Found on the World Wide Web](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/found-on-the-world-wide-web.md)
**Level**          |Unknown
**User Rating**    |4.0 (12 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/found-on-the-world-wide-web-associatefiletype__1-675/archive/master.zip)

### API Declarations

```
Declare Function RegCreateKey& Lib "advapi32.DLL" Alias "RegCreateKeyA" (ByVal hKey&, ByVal lpszSubKey$, lphKey&)
Declare Function RegSetValue& Lib "advapi32.DLL" Alias "RegSetValueA" (ByVal hKey&, ByVal lpszSubKey$, ByVal fdwType&, ByVal lpszValue$, ByVal dwLength&)
' Return codes from Registration functions.
Public Const ERROR_SUCCESS = 0&
Public Const ERROR_BADDB = 1&
Public Const ERROR_BADKEY = 2&
Public Const ERROR_CANTOPEN = 3&
Public Const ERROR_CANTREAD = 4&
Public Const ERROR_CANTWRITE = 5&
Public Const ERROR_OUTOFMEMORY = 6&
Public Const ERROR_INVALID_PARAMETER = 7&
Public Const ERROR_ACCESS_DENIED = 8&
Global Const HKEY_CLASSES_ROOT = &H80000000
Public Const MAX_PATH = 256&
Public Const REG_SZ = 1
```


### Source Code

```
'make a new project: one form with a commandcontrol
'insert the code on the right places
'make the nessecary changes concerning your application and extension
'look for the * sign!
' Return codes from Registration functions.
Public Const ERROR_SUCCESS = 0&
Public Const ERROR_BADDB = 1&
Public Const ERROR_BADKEY = 2&
Public Const ERROR_CANTOPEN = 3&
Public Const ERROR_CANTREAD = 4&
Public Const ERROR_CANTWRITE = 5&
Public Const ERROR_OUTOFMEMORY = 6&
Public Const ERROR_INVALID_PARAMETER = 7&
Public Const ERROR_ACCESS_DENIED = 8&
Global Const HKEY_CLASSES_ROOT = &H80000000
Public Const MAX_PATH = 256&
Public Const REG_SZ = 1
Private Sub Command1_Click()
  Dim sKeyName As String  'Holds Key Name in registry.
  Dim sKeyValue As String 'Holds Key Value in registry.
  Dim ret&         'Holds error status if any from API calls.
  Dim lphKey&       'Holds created key handle from RegCreateKey.
  'This creates a Root entry called "MyApp".
  sKeyName = "MyApp" '*
  sKeyValue = "My Application" '*
  ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
  ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)
  'This creates a Root entry called .BAR associated with "MyApp".
  sKeyName = ".bar" '*
  sKeyValue = "MyApp" '*
  ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
  ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)
  'This sets the command line for "MyApp".
  sKeyName = "MyApp" '*
  sKeyValue = "notepad.exe %1" '*
  ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
  ret& = RegSetValue&(lphKey&, "shell\open\command", REG_SZ, sKeyValue, MAX_PATH)
End Sub
```


<div align="center">

## Getting the gradient title bar end colors


</div>

### Description

This submission contains a function to get the gradient end colors of that gradient title bars as set in the control panel by the user and two other ones to check if the gradient effect is enabled/supported.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jonas Wolz](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jonas-wolz.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jonas-wolz-getting-the-gradient-title-bar-end-colors__1-9764/archive/master.zip)





### Source Code

```
'Here I have put the whole code (including API
'declarations) to make pasting it into a module easier
'To get the OS version:
Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128  ' Maintenance string for PSS usage
End Type
Public Const VER_PLATFORM_WIN32_NT = 2
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
'To get the color if supported:
Public Const COLOR_GRADIENTACTIVECAPTION = 27
Public Const COLOR_GRADIENTINACTIVECAPTION = 28
Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
'To see if it's enabled:
Public Const SPI_GETGRADIENTCAPTIONS = &H1008
'Changed the declaration a bit (removed the ByVal from lpvParam) to pass a pointer to Long:
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
'Enumeration for GetGradientColor:
Enum eGradientColors
 clrGradientActiveCaption = COLOR_GRADIENTACTIVECAPTION
 clrGradientInactiveCaption = COLOR_GRADIENTINACTIVECAPTION
End Enum
'Gets the system gradient end colors for active and inactive title bars
'Raises error 5 if gradient title bars are not supported (in your app
' it might be useful to return a default color instead)
Function GetGradientColor(ByVal lClrIdx As eGradientColors) As Long
 'Are gradient title bars aupported ?:
 If IsWin98Or2000 Then
  'Supported, call the GetSysColor() API to get the color:
  GetGradientColor = GetSysColor(lClrIdx)
 Else
  'Not supported, raise an error:
  Err.Raise 5, , "Gradient Titlebars not supported by this OS version !"
  'Might be more useful (if you think so):
  ''Return a default color:
  'GetGradientColor = vbCyan
 End If
End Function
'This function returns True if the gradient effect is enabled/supported
'Under Win98/2000/higher it calls the SystemParametersInfo() API to check if it's enabled,
'under Win95/NT 4 it always returns False.
Function IsGradientEnabled() As Boolean
 Dim lEnabled As Long
 If IsWin98Or2000 Then
  lEnabled = 0
  'Call the API to check if it's enabled:
  SystemParametersInfo SPI_GETGRADIENTCAPTIONS, 0, lEnabled, 0
  'Return the value:
  IsGradientEnabled = CBool(lEnabled)
 Else
  'Gradient not supported, return False:
  IsGradientEnabled = False
 End If
End Function
'This function returns True if the OS Version is Win98, 2000 or higher
' (-> a version which has gradient title bars)
Function IsWin98Or2000() As Boolean
 Static bWasInHere As Boolean, bState As Boolean
 'May it speed up a bit when called often:
 If Not bWasInHere Then
  Dim OSV As OSVERSIONINFO
  OSV.dwOSVersionInfoSize = Len(OSV)
  'Get the OS version:
  GetVersionEx OSV
  bState = False
  'Check if platform Win95/98/ME
  If (OSV.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS) Then
   'dwMinorVersion > 0 And dwMajorVersion =4 -> Win98
   If (OSV.dwMajorVersion > 4) Or ((OSV.dwMajorVersion = 4) And (OSV.dwMinorVersion > 0)) Then
    'It's Win98 or higher
    bState = True
   Else
    'It's Win95:
    bState = False
   End If
  'Check if platform NT/Win2000:
  ElseIf (OSV.dwPlatformId = VER_PLATFORM_WIN32_NT) Then
   If (OSV.dwMajorVersion >= 5) Then
    'It's Win2000 or higher:
    bState = True
   Else
    'Is NT4 (or lower):
    bState = False
   End If
  End If
  bWasInHere = True
 End If
 'Return our result:
 IsWin98Or2000 = bState
End Function
```


<div align="center">

## Network Enumerator


</div>

### Description

A self-contained class that provides a simple method of getting a list of all the computers on all the networks visible from the machine running the code. It also provides a list of the domains, printers, shares, network names, server names - basically everything you can get out of the WNet API.

PLEASE NOTE: (Before anyone flames me for plagiarism)

This code has been adapted from several snippets that I have found on my travels. Credit is due to AGP for pointing me in the right direction, and also Mr.X from Google Groups for supplying some code for me to base my class upon.

I hope this condensed version helps someone out.
 
### More Info
 
None required.

Some experience of using objects would be nice, but the demo code in Main() shows how to do everything.

Various strings containing a | delimited list of the items requested (eg GetServerList)

It can take a while to complete when running across a 64k link into a huge corporate network. :)


<span>             |<span>
---                |---
**Submitted On**   |2001-05-30 19:27:14
**By**             |[FrootLimited](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/frootlimited.md)
**Level**          |Beginner
**User Rating**    |4.4 (66 globes from 15 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Network En202835302001\.zip](https://github.com/Planet-Source-Code/frootlimited-network-enumerator__1-23594/archive/master.zip)

### API Declarations

```
Private Const RESOURCE_CONNECTED As Long = &H1&
Private Const RESOURCE_GLOBALNET As Long = &H2&
Private Const RESOURCE_REMEMBERED As Long = &H3&
Private Const RESOURCEDISPLAYTYPE_DIRECTORY& = &H9
Private Const RESOURCEDISPLAYTYPE_DOMAIN& = &H1
Private Const RESOURCEDISPLAYTYPE_FILE& = &H4
Private Const RESOURCEDISPLAYTYPE_GENERIC& = &H0
Private Const RESOURCEDISPLAYTYPE_GROUP& = &H5
Private Const RESOURCEDISPLAYTYPE_NETWORK& = &H6
Private Const RESOURCEDISPLAYTYPE_ROOT& = &H7
Private Const RESOURCEDISPLAYTYPE_SERVER& = &H2
Private Const RESOURCEDISPLAYTYPE_SHARE& = &H3
Private Const RESOURCEDISPLAYTYPE_SHAREADMIN& = &H8
Private Const RESOURCETYPE_ANY As Long = &H0&
Private Const RESOURCETYPE_DISK As Long = &H1&
Private Const RESOURCETYPE_PRINT As Long = &H2&
Private Const RESOURCETYPE_UNKNOWN As Long = &HFFFF&
Private Const RESOURCEUSAGE_ALL As Long = &H0&
Private Const RESOURCEUSAGE_CONNECTABLE As Long = &H1&
Private Const RESOURCEUSAGE_CONTAINER As Long = &H2&
Private Const RESOURCEUSAGE_RESERVED As Long = &H80000000
Private Const NO_ERROR = 0
Private Const ERROR_MORE_DATA = 234
Private Const RESOURCE_ENUM_ALL As Long = &HFFFF
Private Type NETRESOURCE
  dwScope As Long
  dwType As Long
  dwDisplayType As Long
  dwUsage As Long
  pLocalName As Long
  pRemoteName As Long
  pComment As Long
  pProvider As Long
End Type
Private Type NETRESOURCE_EXTENDED
  dwScope As Long
  dwType As Long
  dwDisplayType As Long
  dwUsage As Long
  sLocalName As String
  sRemoteName As String
  sComment As String
  sProvider As String
End Type
Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Any, lphEnum As Long) As Long
Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, lpBuffer As NETRESOURCE, lpBufferSize As Long) As Long
Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long
Private Declare Function VarPtrAny Lib "vb40032.dll" Alias "VarPtr" (lpObject As Any) As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (lpTo As Any, lpFrom As Any, ByVal lLen As Long)
Private Declare Sub CopyMemByPtr Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpTo As Long, ByVal lpFrom As Long, ByVal lLen As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Any) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
```






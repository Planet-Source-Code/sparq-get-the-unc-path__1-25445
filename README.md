<div align="center">

## Get the UNC Path


</div>

### Description

Networked drives. They're an administrative nightmare. In fact, the average user changes his networked drive letters more often than his underwear.

But you can solve this problem of binding an application to a particular drive by using a Universal Naming Convention (UNC) path. This still references a network area, but doesn&#8217;t tie it to any one drive letter.

And you can retrieve and check the UNC of a particular path in code using this neat little function.

To use it, simply call GetUNCPath, passing it your drive letter along with a pre-declared empty string. If a problem occurs, the relevant number is passed back with the function. These can be matched with the possible return code constants.

However if everything goes swimmingly, the function returns a zero (a constant value of NO_ERROR) and places the UNC path into the ByRef-passed variable.

This code works by making a call to the WNetGetConnection function in MPR.DLL. It's essentially a neat wrapper for a regular API call.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sparq](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sparq.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sparq-get-the-unc-path__1-25445/archive/master.zip)

### API Declarations

```
' To be put inside a module
'Possible return codes from the API
Public Const ERROR_BAD_DEVICE      As Long = 1200
Public Const ERROR_CONNECTION_UNAVAIL  As Long = 1201
Public Const ERROR_EXTENDED_ERROR    As Long = 1208
Public Const ERROR_MORE_DATA      As Long = 234
Public Const ERROR_NOT_SUPPORTED    As Long = 50
Public Const ERROR_NO_NET_OR_BAD_PATH  As Long = 1203
Public Const ERROR_NO_NETWORK      As Long = 1222
Public Const ERROR_NOT_CONNECTED    As Long = 2250
Public Const NO_ERROR          As Long = 0
'This API returns a UNC from a drive letter
Declare Function WNetGetConnection Lib "mpr.dll" Alias _
  "WNetGetConnectionA" _
  (ByVal lpszLocalName As String, _
  ByVal lpszRemoteName As String, _
  cbRemoteName As Long) As Long
Function GetUNCPath(ByVal strDriveLetter As String, _
         ByRef strUNCPath As String) As Long
On Local Error GoTo GetUNCPath_Err
  Dim strMsg As String
  Dim lngReturn As Long
  Dim strLocalName As String
  Dim strRemoteName As String
  Dim lngRemoteName As Long
  strLocalName = strDriveLetter
  strRemoteName = String$(255, Chr$(32))
  lngRemoteName = Len(strRemoteName)
  'Attempt to grab UNC
  lngReturn = WNetGetConnection(strLocalName, _
                 strRemoteName, _
                 lngRemoteName)
  If lngReturn = NO_ERROR Then
    'No problems - return the UNC
    'to the passed ByRef string
    GetUNCPath = NO_ERROR
    strUNCPath = Trim$(strRemoteName)
    strUNCPath = Left$(strUNCPath, Len(strUNCPath) - 1)
  Else
    'Problems - so return original
    'drive letter and error number
    GetUNCPath = lngReturn
    strUNCPath = strDriveLetter & "\"
  End If
GetUNCPath_End:
  Exit Function
GetUNCPath_Err:
  GetUNCPath = ERROR_NOT_SUPPORTED
  strUNCPath = strDriveLetter
  Resume GetUNCPath_End
End Function
```


### Source Code

```
Dim strUNC As String
  If GetUNCPath("H:", strUNC) = NO_ERROR Then
    MsgBox "The UNC of the specified drive is " & strUNC
  Else
    MsgBox "There was a problem, sorry!"
  End If
```


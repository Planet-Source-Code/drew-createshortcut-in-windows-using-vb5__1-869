<div align="center">

## CreateShortcut in Windows using VB5


</div>

### Description

Creates a Windows Shortcut to a program executable, and places it on the Desktop, Start Menu, or a Start Menu subfolder.
 
### More Info
 
frm  - A form to hook onto.

strGroupName - The name of the Group where this shortcut

will be placed. By default, this group is

always placed in the 'Start Menu/Programs' folder.

You can pass '..\..\Desktop' to put this on

the Desktop, or '..' to put this on the 'Start Menu'.

strLinkName - Text caption for the Shortcut link.

strLinkPath - Full path to the target of the Shortcut link.

Ex: 'c:\Program Files\My Application\MyApp.exe'

strLinkArguments - Command-line arguments for the Shortcut link.

Ex: '-f -c "c:\Program Files\My Application\MyApp.dat" -q'

The Windows API/Global Declarations section and the code should be placed in the same bas module. Preferable it's own module because there are a lot of functions need to check the platform your using, convert short filenames, create menu folders/groups ect...


<span>             |<span>
---                |---
**Submitted On**   |1999-12-23 16:18:34
**By**             |[Drew](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/drew.md)
**Level**          |Advanced
**User Rating**    |4.7 (52 globes from 11 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD250612231999\.zip](https://github.com/Planet-Source-Code/drew-createshortcut-in-windows-using-vb5__1-869/archive/master.zip)





### Source Code

```
'***********************************************************'
'************* CREATE PROGRAM GROUP FUNCTIONS **************'
'***********************************************************'
' PRIMARY FUNCTION CALL:
'
Public Sub CreateShortcut(ByRef frm As Form, _
             ByVal strGroupName As String, _
             ByVal strLinkName As String, _
             ByVal strLinkPath As String, _
             ByVal strLinkArguments As String)
'************************************************************************************
' PROCEDURE: CreateShortcut
'        First, the procedure creates the Program Group if necessary,
'        Then it calls CreateProgManItem under Windows NT or
'        CreateFolderLink under Windows 95 to validate and create
'        your link shortcuts.
'
' PARAMETERS:
'   frm       - A form to hook onto.
'
'   strGroupName   - The name of the Group where this shortcut
'             will be placed. By default, this group is
'             always placed in the 'Start Menu/Programs' folder.
'             You can pass '..\..\Desktop' to put this on
'             the Desktop, or '..' to put this on the 'Start Menu'.
'
'   strLinkName   - Text caption for the Shortcut link.
'
'   strLinkPath   - Full path to the target of the Shortcut link.
'              Ex: 'c:\Program Files\My Application\MyApp.exe'
'
'   strLinkArguments - Command-line arguments for the Shortcut link.
'              Ex: '-f -c "c:\Program Files\My Application\MyApp.dat" -q'
'
'************************************************************************************
  'CREATE THE PROGRAM GROUP IF NECCESSARY, THEN THE SHORTCUT'
  If fCreateProgGroup(frm, strGroupName) Then
    If TreatAsWin95() Then
      'CREATE WINDOWS 95 SHORTCUT'
      CreateShellLink strLinkPath, strGroupName, strLinkArguments, strLinkName
    Else
      ' DDE will not work properly if you try to send NT the long filename. If it is
      ' in quotes, then the parameters get ignored. If there are no parameters, the
      ' long filename can be used and the following line could be skipped.
      strLinkPath = GetShortPathName(strUnQuoteString(strLinkPath))
      'CREATE WINDOWS NT SHORTCUT'
      CreateProgManItem frm, strGroupName, strLinkPath & " " & strLinkArguments, strLinkName
    End If
  End If
End Sub
Private Sub CreateShellLink(ByVal strLinkPath As String, ByVal strGroupName As String, ByVal strLinkArguments As String, ByVal strLinkName As String)
  'ReplaceDoubleQuotes strLinkName
  strLinkName = strUnQuoteString(strLinkName)
  strLinkPath = strUnQuoteString(strLinkPath)
  Dim fSuccess As Boolean
  fSuccess = OSfCreateShellLink(strGroupName & "", strLinkName, strLinkPath, strLinkArguments & "")
  If Not fSuccess Then
    MsgBox "Create Shortcut Failed!", vbExclamation, "Ouch!"
  End If
End Sub
Private Sub CreateProgManItem(frm As Form, ByVal strGroupName As String, ByVal strCmdLine As String, ByVal strIconTitle As String)
  PerformDDE frm, strGroupName, strCmdLine, strIconTitle, kDDE_AddItem
End Sub
Private Sub PerformDDE(frm As Form, ByVal strGroup As String, ByVal strCmd As String, ByVal strTitle As String, ByVal intDDE As Integer)
  Const strCOMMA$ = ","
  Const strRESTORE$ = ", 1)]"
  Const strACTIVATE$ = ", 5)]"
  Const strENDCMD$ = ")]"
  Const strSHOWGRP$ = "[ShowGroup("
  Const strADDGRP$ = "[CreateGroup("
  Const strREPLITEM$ = "[ReplaceItem("
  Const strADDITEM$ = "[AddItem("
  Dim intIdx As Integer    'loop variable
  Screen.MousePointer = vbHourglass
  Dim intRetry As Integer
  For intRetry = 1 To 20
    On Error Resume Next
    frm.lblDDE.LinkTopic = "PROGMAN|PROGMAN"
    If Err = 0 Then
      Exit For
    End If
    DoEvents
  Next intRetry
  frm.lblDDE.LinkMode = 2
  For intIdx = 1 To 10
   DoEvents
  Next
  frm.lblDDE.LinkTimeout = 100
  On Error Resume Next
  If Err = 0 Then
    Select Case intDDE
      Case kDDE_AddItem
        ' The item will be created in the group titled strGroup
        '
        ' Force the group strGroup to be the active group. Additem only
        ' puts icons in the active group.
        #If 0 Then
          frm.lblDDE.LinkExecute strSHOWGRP & strGroup & strACTIVATE
        #Else
          frm.lblDDE.LinkExecute strADDGRP & strGroup & strENDCMD
        #End If
        frm.lblDDE.LinkExecute strREPLITEM & strTitle & strENDCMD
        Err = 0
        frm.lblDDE.LinkExecute strADDITEM & strCmd & strCOMMA & strTitle & String$(3, strCOMMA) & strENDCMD
      Case kDDE_AddGroup
        frm.lblDDE.LinkExecute strADDGRP & strGroup & strENDCMD
        frm.lblDDE.LinkExecute strSHOWGRP & strGroup & strRESTORE
      'End Case
    End Select
  End If
  '
  'Disconnect DDE Link
  frm.lblDDE.LinkMode = 0
  frm.lblDDE.LinkTopic = ""
  Screen.MousePointer = vbDefault
  Err = 0
End Sub
'
'
'***********************************************************'
'************* CREATE PROGRAM GROUP FUNCTIONS **************'
'***********************************************************'
'
Private Function fCreateProgGroup(frm As Form, sNewGroupName As String) As Boolean
  'DONT VALIDATE OR CREATE THE 'DESKTOP' GROUP,
  '  OR THE 'START MENU GROUP', THEY SHOULD EXIST ALREADY.
  If UCase(Trim(sNewGroupName)) = kDesktopGroup Or sNewGroupName = kStartMenuGroup Then
    fCreateProgGroup = True
    Exit Function
  Else
    'VALIDATE AND CREATE PROGRAM GROUP'
    If TreatAsWin95() Then
      'WINDOWS 95 - VALIDATE'
      If Not fValid95Filename(sNewGroupName) Then
        MsgBox "Error: Could not validate the Program Group name!", vbQuestion, "Error"
        GoTo CGError
      End If
    Else
      'WINDOWS NT - VALIDATE'
      If Not fValidNTGroupName(sNewGroupName) Then
        MsgBox "Error: Could not validate the Program Group name!", vbQuestion, "Error"
        GoTo CGError
      End If
    End If
    'CREATE THE WINDOWS 95 OR NT PROGRAM GROUP'
    If Not fCreateOSProgramGroup(frm, sNewGroupName) Then
      GoTo CGError
    End If
    fCreateProgGroup = True
  End If
Exit Function
CGError:
  fCreateProgGroup = False
End Function
Private Function fCreateShellGroup(ByVal strFolderName As String) As Boolean
  ReplaceDoubleQuotes strFolderName
  If strFolderName = "" Then
    Exit Function
  End If
  Dim fSuccess As Boolean
  fSuccess = OSfCreateShellGroup(strFolderName)
  If fSuccess Then
  Else
    MsgBox "Create Start Menu Group Failed!", vbExclamation, "Ouch!"
  End If
  fCreateShellGroup = fSuccess
End Function
Private Function fValid95Filename(strFilename As String) As Boolean
' This routine verifies that strFileName is a valid file name.
' It checks that its length is less than the max allowed
' and that it doesn't contain any invalid characters..
  Dim iInvalidChar  As Integer
  Dim iFilename    As Integer
  If Not ValidateFilenameLength(strFilename) Then
    ' Name is too long.
    fValid95Filename = False
    Exit Function
  End If
  ' Search through the list of invalid filename characters and make
  ' sure none of them are in the string.
  For iInvalidChar = 1 To Len(kInvalid95GroupNameChars)
    If InStr(strFilename, Mid$(kInvalid95GroupNameChars, iInvalidChar, 1)) <> 0 Then
      fValid95Filename = False
      Exit Function
    End If
  Next iInvalidChar
  fValid95Filename = True
End Function
Public Function fValidNTGroupName(strGroupName) As Boolean
' This routine verifies that strGroupName is a valid group name.
' It checks that its length is less than the max allowed
' and that it doesn't contain any invalid characters.
  If Len(strGroupName) > kMaxGroupNameLength Then
    fValidNTGroupName = False
    Exit Function
  End If
  ' Search through the list of invalid filename characters and make
  ' sure none of them are in the string.
  Dim iInvalidChar As Integer
  Dim iFilename As Integer
  For iInvalidChar = 1 To Len(kInvalidNTGroupNameChars)
    If InStr(strGroupName, Mid$(kInvalidNTGroupNameChars, iInvalidChar, 1)) <> 0 Then
      fValidNTGroupName = False
      Exit Function
    End If
  Next iInvalidChar
  fValidNTGroupName = True
End Function
Private Function fCreateOSProgramGroup(frm As Form, ByVal strFolderName As String) As Boolean
  If TreatAsWin95() Then
    'CREATE WINDOWS 95 PROGRAM GROUP'
    fCreateOSProgramGroup = fCreateShellGroup(strFolderName)
  Else
    'CREATE WINDOWS NT PROGRAM GROUP'
    CreateProgManGroup frm, strFolderName
    fCreateOSProgramGroup = True
  End If
End Function
Private Sub CreateProgManGroup(frm As Form, ByVal strGroupName As String)
  PerformDDE frm, strGroupName, kEmptyString, kEmptyString, kDDE_AddGroup
End Sub
'
'
'***********************************************************'
'********************* OTHER FUNCTIONS *********************'
'***********************************************************'
Private Function TreatAsWin95() As Boolean
  If IsWindows95() Then
    TreatAsWin95 = True
  ElseIf fNTWithShell() Then
    TreatAsWin95 = True
  Else
    TreatAsWin95 = False
  End If
End Function
Private Function IsWindows95() As Boolean
  Const dwMask95 = &H2&
  If GetWinPlatform() And dwMask95 Then
    IsWindows95 = True
  Else
    IsWindows95 = False
  End If
End Function
Private Function strUnQuoteString(ByVal strQuotedString As String)
' This routine tests to see if strQuotedString is wrapped in quotation
' marks, and, if so, remove them.
  strQuotedString = Trim(strQuotedString)
  If Mid$(strQuotedString, 1, 1) = kQuote And Right$(strQuotedString, 1) = kQuote Then
    ' It's quoted. Get rid of the quotes.
    strQuotedString = Mid$(strQuotedString, 2, Len(strQuotedString) - 2)
  End If
  strUnQuoteString = strQuotedString
End Function
Private Function StripTerminator(ByVal strString As String) As String
  Dim intZeroPos As Integer
  intZeroPos = InStr(strString, Chr$(0))
  If intZeroPos > 0 Then
    StripTerminator = Left$(strString, intZeroPos - 1)
  Else
    StripTerminator = strString
  End If
End Function
Private Sub ReplaceDoubleQuotes(str As String)
  Dim i As Integer
  For i = 1 To Len(str)
    If Mid$(str, i, 1) = """" Then
      Mid$(str, i, 1) = "'"
    End If
  Next i
End Sub
Private Function GetShortPathName(ByVal strLongPath As String) As String
  Const cchBuffer = 300
  Dim strShortPath As String
  Dim lResult As Long
  On Error GoTo 0
  strShortPath = String(cchBuffer, Chr$(0))
  lResult = OSGetShortPathName(strLongPath, strShortPath, cchBuffer)
  If lResult = 0 Then
    Error 53 ' File not found
  Else
    GetShortPathName = StripTerminator(strShortPath)
  End If
End Function
Private Function ValidateFilenameLength(strFilename As String) As Boolean
  ValidateFilenameLength = (Len(strFilename) < kMaxPathLength)
End Function
```


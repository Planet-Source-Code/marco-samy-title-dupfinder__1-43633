Attribute VB_Name = "modShorts"
Option Explicit
Option Base 1
Public Declare Function OSfCreateShellLink Lib "vb6stkit.dll" Alias "fCreateShellLink" _
         (ByVal lpstrFolderName As String, _
         ByVal lpstrLinkName As String, _
         ByVal lpstrLinkPath As String, _
         ByVal lpstrLinkArguments As String, _
         ByVal fPrivate As Long, _
         ByVal sParent As String) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Const gstrQUOTE$ = """"

'-----------------------------------------------------------
' SUB: CreateShellLink
'
' Creates (or replaces) a link in either Start>Programs or
' any of its immediate subfolders in the Windows 95 shell.
'
' IN: [strLinkPath] - full path to the target of the link
'                     Ex: 'c:\Program Files\My Application\MyApp.exe"
'     [strLinkArguments] - command-line arguments for the link
'                     Ex: '-f -c "c:\Program Files\My Application\MyApp.dat" -q'
'     [strLinkName] - text caption for the link
'     [fLog] - Whether or not to write to the logfile (default
'                is true if missing)
'
' OUT:
'   The link will be created in the folder strGroupName
'-----------------------------------------------------------
'
Public Sub CreateShellLinkX(ByVal strLinkPath As String, _
         ByVal strGroupName As String, _
         ByVal strLinkArguments As String, _
         ByVal strLinkName As String, _
         ByVal fPrivate As Boolean, _
         sParent As String, _
         Optional ByVal fLog As Boolean = True)
Dim fSuccess As Boolean
Dim intMsgRet As Integer
Dim lREt       As Boolean
   strLinkName = strUnQuoteString(strLinkName)
   strLinkPath = strUnQuoteString(strLinkPath)
   If StrPtr(strLinkArguments) = 0 Then strLinkArguments = ""
   
   lREt = OSfCreateShellLink(strGroupName, strLinkName, strLinkPath, strLinkArguments, _
         fPrivate, sParent)    'the path should never be enclosed in double quotes
End Sub
Public Function strUnQuoteString(ByVal strQuotedString As String)
'
' This routine tests to see if strQuotedString is wrapped in quotation
' marks, and, if so, remove them.
'
    strQuotedString = Trim$(strQuotedString)

    If Mid$(strQuotedString, 1, 1) = gstrQUOTE Then
        If Right$(strQuotedString, 1) = gstrQUOTE Then
            '
            ' It's quoted.  Get rid of the quotes.
            '
            strQuotedString = Mid$(strQuotedString, 2, Len(strQuotedString) - 2)
        End If
    End If
    strUnQuoteString = strQuotedString
End Function
Function GetWinPath()
Dim TmpStr As String
TmpStr = Space(255)
GetWindowsDirectory TmpStr, 255
GetWinPath = Trim$(TermIt(TmpStr))
End Function
Function TermIt(ByVal strString As String) As String
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        TermIt = Left$(strString, intZeroPos - 1)
    Else
        TermIt = strString
    End If
End Function
'here are our final Function To Craete a ShotCut
Function CreateShortCut(sFolder As String, sFile As String, sCommand As String, sTitle As String)
On Error Resume Next 'display no errors
Dim TmpDir
TmpDir = "ShTmpDir"
'We Create a Temp Direcotory in the Windows Path
MkDir (IIf(Right$(GetWinPath, 1) = "\", GetWinPath, GetWinPath & "\") & TmpDir)
'Make The Shell Link in it
CreateShellLinkX sFile, "..\..\" & TmpDir, sCommand, sTitle, True, "$(Programs)"
'wait to be done
DoEvents
'Copy the Link that created in the Win path to the desired folder
FileCopy IIf(Right$(GetWinPath, 1) = "\", GetWinPath, GetWinPath & "\") & TmpDir & "\" & sTitle & ".lnk", IIf(Right$(sFolder, 1) = "\", sFolder, sFolder & "\") & sTitle & ".lnk"
'Reset attributes of Original Link to delete it
SetAttr IIf(Right$(GetWinPath, 1) = "\", GetWinPath, GetWinPath & "\") & TmpDir & "\" & sTitle & ".lnk", vbNormal
'Delete It
Kill IIf(Right$(GetWinPath, 1) = "\", GetWinPath, GetWinPath & "\") & TmpDir & "\" & sTitle & ".lnk"
'Reset Attributes of the temp folder
SetAttr IIf(Right$(GetWinPath, 1) = "\", GetWinPath, GetWinPath & "\") & TmpDir, vbNormal
'Delete it
RmDir IIf(Right$(GetWinPath, 1) = "\", GetWinPath, GetWinPath & "\") & TmpDir
'Now we are done
End Function

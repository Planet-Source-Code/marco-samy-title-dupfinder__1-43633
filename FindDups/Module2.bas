Attribute VB_Name = "Module2"
'/////////////////////////////////////////////////////////////////
'///////////////////////////File Compare Engine Base File
'///////////////////////////By Marco Samy 2003
'/////////////////////////////////////////////////////////////////
'in This Module the compare process takes it's effects
Public Const MAX_PATH = 260
Private Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public PreformTest As Boolean 'Boolean Variable Indicate If we Make the test or not
Public BytesDone As Double, BytesWith As Double 'where the done bytes and compared bytes is saved.
'Find File is a function by Windows API for a very fast find file
'It works so faster than the normal find Using Dir$
Public Sub FindFiles(strRootFolder As String, strFolder As String, strFile As String, colFilesFound As Collection, PicDir As PictureBox)
If fMain.Cancel = True Then Exit Sub
Dim lngSearchHandle As Long 'Handle returned by the function
Dim udtFindData As WIN32_FIND_DATA 'variale contans find information, basic WinFind routine
Dim strTemp As String, lngRet As Long 'Our Returned values
PicDir.Cls: PicDir.Print strRootFolder 'Just Changing information in the main form
'Check that folder name ends with "\"
If Right$(strRootFolder, 1) <> "\" Then strRootFolder = strRootFolder & "\"
'Find first file/folder in current folder
lngSearchHandle = FindFirstFile(strRootFolder & "*", udtFindData)
'Check that we received a valid handle
If lngSearchHandle = INVALID_HANDLE_VALUE Then Exit Sub
lngRet = 1
Do While lngRet <> 0
'Trim nulls from filename
strTemp = TrimNulls(udtFindData.cFileName)
If (udtFindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
'It's a dir - make sure it isn't . or .. dirs
If strTemp <> "." And strTemp <> ".." Then
'It's a normal dir: let's dive straight
'into it...
    Call FindFiles(strRootFolder & strTemp, strFolder, strFile, colFilesFound, PicDir)
End If
Else
'It's a file. First check if the current folder matches
'the folder path in strFolder
    If (strRootFolder Like strFolder) Then
    'Folder matches, what about file?
    'Found one!
                    If (strTemp Like strFile) Then
    '//////////////////////////////////////////
    '//////////////What To Do Here
    '//////////////////////////////////////////
    colFilesFound.Add strRootFolder & strTemp
                    End If
            End If
        End If
'Get next file/folder
lngRet = FindNextFile(lngSearchHandle, udtFindData)
Loop
'Close find handle
Call FindClose(lngSearchHandle)
End Sub
'this is the main function here what's doing every thing in the program
Function FindDublics(Dirs2 As Boolean, Dir1 As String, Dir2 As String, Pattern As String, cPercent As Single, colFound As Collection, PicDir As PictureBox, PicFile As PictureBox, picFound As PictureBox, picNum As PictureBox, Optional MaxPatch As Long = 1047552, Optional MinFound As Integer = 1)
ResetVars 'Reset Variabls Data
BytesDone = 0: BytesWith = 0 'Reseting values
Dir1 = IIf(Right$(Dir1, 1) = "\", Dir1, Dir1 & "\") 'Editing Working Path, Must be a directory path with "\"
If Dirs2 = False Then Dir2 = Dir1 Else Dir2 = IIf(Right$(Dir2, 1) = "\", Dir2, Dir2 & "\") 'Editing Working Path, Must be a directory path with "\"
Dim Dir1Files As New Collection, Dir2Files As New Collection 'Working Collections, Needed for Values
Dim CurFLen As Single, CutOut As Single, Found As Long, Copies As Long 'Needed Variable ---> you will see why
Dim I, X, Founds As Boolean, iStr As String, CurFound As Single 'Loop and Info.
fMain.Picture7.Cls: fMain.Picture7.Print "Searching Dir1, Please Wait.." 'Displaying Some Text
If fMain.Cancel = True Then Exit Function 'Cancel the Process when the user clicks cancel
FindFiles Dir1, "*", Pattern, Dir1Files, PicDir  'Collecting Files using the past function
If fMain.Cancel = True Then Exit Function 'Cancel the Process when the user clicks cancel
fMain.Picture7.Cls: fMain.Picture7.Print "Searching Dir2, Please Wait ..."
If Dirs2 = True Then FindFiles Dir2, "*", Pattern, Dir2Files, PicDir
If Dirs2 = True Then
For I = 1 To Dir1Files.Count
If InLstEnh(Dir1Files(I), colFound) = True Then GoTo NextI2
iStr = Dir1Files(I)
CurFLen = FileLen(iStr)
Founds = False
CurFound = 0
For X = 1 To Dir2Files.Count
If Dir1Files(I) = Dir2Files(X) Then GoTo NextX2
If fMain.Cancel = True Then Exit Function 'Cancel the Process when the user clicks cancel
CutOut = Abs(FileLen(Dir2Files(X)) - CurFLen)
If (CutOut / CurFLen) <= (cPercent / 100) Then
If TestPassed(Dir1Files(I), Dir2Files(X), cPercent, PicFile, MaxPatch) Then
iStr = iStr & "," & Dir1Files(X)
If Founds = False Then
Founds = True
Found = Found + 1
End If
CurFound = CurFound + 1
Copies = Copies + 1
End If
End If
fMain.picByte.Cls: fMain.picByte.Print BytesDone
fMain.picWith.Cls: fMain.picWith.Print BytesWith
fMain.PB1.Value = Val(Format(I / Dir1Files.Count * 100))
picFound.Cls: picFound.Print Found
picNum.Cls: picNum.Print Copies
NextX2:
Next X
If Val(CurFound) >= Val(MinFound) Then colFound.Add iStr
NextI2:
Next I
'////////////Else
Else
For I = 1 To Dir1Files.Count
iStr = Dir1Files(I)
If InLstEnh(Dir1Files(I), colFound) = True Then GoTo NextI
CurFLen = FileLen(iStr)
Founds = False
CurFound = 0
For X = 1 To Dir1Files.Count
If fMain.Cancel = True Then Exit Function
If X = I Then GoTo NextX
CutOut = Abs(FileLen(Dir1Files(X)) - CurFLen)
If (MinOf(CutOut, CurFLen) / BigOf(CurFLen, CutOut)) <= (cPercent / 100) Then
If InLstEnh(Dir1Files(X), colFound) = True Then GoTo NextX
If TestPassed(Dir1Files(I), Dir1Files(X), cPercent, PicFile, MaxPatch) Then
iStr = iStr & "," & Dir1Files(X)
If Founds = False Then
Founds = True
Found = Found + 1
End If
CurFound = CurFound + 1
Copies = Copies + 1
End If
End If
picFound.Cls: picFound.Print Found
picNum.Cls: picNum.Print Copies
NextX:
fMain.picByte.Cls: fMain.picByte.Print BytesDone
fMain.picWith.Cls: fMain.picWith.Print BytesWith
Next X
If Val(CurFound) >= Val(MinFound) Then colFound.Add iStr
NextI:
fMain.PB1.Value = Val(Format(I / Dir1Files.Count * 100))
DoEvents
Next I
End If
fMain.Founds = Founds
fMain.fAll = Copies
End Function
Function ResetVars()
fMain.PB1.Value = 0 'Update status
BytesDone = 0: BytesWith = 0 'Reset Information
End Function
Function BigOf(Value1 As Single, Value2 As Single) As Single 'get big value of 2 values
If Val(Value1) > Val(Value2) Then BigOf = Value1 Else BigOf = Value2
End Function
Function MinOf(Value1 As Single, Value2 As Single) As Single 'get smaller value
If Val(Value1) < Val(Value2) Then MinOf = Value1 Else MinOf = Value2
End Function
Public Function TrimNulls(strString As String) As String
   Dim l As Long
   l = InStr(1, strString, Chr(0))
   If l = 1 Then
      TrimNulls = ""
   ElseIf l > 0 Then
      TrimNulls = Left$(strString, l - 1)
   Else
      TrimNulls = strString
   End If
End Function
Function TestPassed(sFile1 As String, sFile2 As String, sPercent As Single, PicFiles As PictureBox, Optional MaxSize As Long = 1047552) As Boolean
On Error GoTo Err1:
Dim nf, nf2
Dim cUnit As Long, cUnit2 As Long, PatchNum As Long, PatchNum2 As Long, PatchSize As Long, PatchSize2 As Long, LastPatch As Long, LastPatch2 As Long, cPatch As Long, cPatch2 As Long
Dim Marks As Double
Dim Ab() As Byte, Ab2() As Byte
cUnit = FileLen(sFile1)
cUnit2 = FileLen(sFile2)
BytesDone = BytesDone + cUnit
BytesWith = BytesWith + cUnit2
If fMain.Cancel = True Then Exit Function 'Cancel the Process when the user clicks cancel
fMain.Picture7.Cls: fMain.Picture7.Print "Testing Two Files"
PicFiles.Cls: PicFiles.Print sFile1
fMain.Picture1.Cls: fMain.Picture1.Print sFile2
If PreformTest = False Then TestPassed = True: Exit Function
'-----------------------Calc Patches For The First File
PatchSize = MaxSize
PatchNum = cUnit / PatchSize
If Int(PatchNum) < PatchNum Then PatchNum = Int(PatchNum) + 1
If Int(PatchSize) < PatchSize Then PatchSize = Int(PatchSize) + 1
Patches:
LastPatch = cUnit - Abs(((PatchNum - 1) * PatchSize) - 1)
If LastPatch <= 0 Then If PatchNum <= 0 Then PatchNum = PatchNum + 1: GoTo Patches Else PatchNum = PatchNum - 1: GoTo Patches
'-----------/Calc Patches
'-----------------------Calc Patches For The Second File
PatchSize2 = MaxSize
PatchNum2 = cUnit2 / PatchSize2
If Int(PatchNum2) < PatchNum2 Then PatchNum2 = Int(PatchNum2) + 1
If Int(PatchSize2) < PatchSize2 Then PatchSize2 = Int(PatchSize2) + 1
Patches2:
LastPatch2 = cUnit2 - Abs(((PatchNum2 - 1) * PatchSize2) - 1)
If LastPatch2 <= 0 Then If PatchNum2 <= 0 Then PatchNum2 = PatchNum2 + 1: GoTo Patches2 Else PatchNum2 = PatchNum2 - 1: GoTo Patches2
'-----------/Calc Patches
nf = FreeFile
Open sFile1 For Binary As #nf
nf2 = FreeFile
Open sFile2 For Binary As #nf2
For Z = 1 To IIf(PatchNum < PatchNum2, PatchNum, PatchNum2)
If Z = PatchNum Then cPatch = LastPatch Else cPatch = PatchSize ' Select Current Patch
If Z = PatchNum2 Then cPatch2 = LastPatch2 Else cPatch2 = PatchSize2 ' Select Current Patch2
ReDim Ab(cPatch)
ReDim Ab2(cPatch2)
Get #nf, 1 + ((Z - 1) * PatchSize), Ab()
Get #nf2, 1 + ((Z - 1) * PatchSize), Ab2()
For Y = 0 To IIf((UBound(Ab)) > (UBound(Ab2)), UBound(Ab), UBound(Ab2))
On Error Resume Next
Marks = Marks + ((Abs(Val(Ab(Y)) - Val(Ab2(Y)))) / 255)
Next Y
Next Z
Close #nf: Close #nf2
'Add the late (More Charcter May in one or more files)
Dim adm As Single
adm = (0.5 * (Abs(PatchNum - PatchNum2) - 1) * PatchSize)
If adm < 0 Then adm = 0
Marks = Marks + adm
Marks = Marks + (0.5 * Abs(LastPatch - LastPatch2))
If (Marks / (cUnit + cUnit2 / 2)) <= (sPercent / 100) Then TestPassed = True Else TestPassed = False
DoEvents
Exit Function
Err1:
TestPassed = False
Close #nf: Close #nf2
End Function
'If item included in collection
Function InLstEnh(sItem As String, sLst As Collection) As Boolean
InLstEnh = False
Dim TmpCol As New Collection
For I = 1 To sLst.Count
GetAllAB sLst(I), ",", ",", TmpCol
If TmpCol.Count = 0 Then TmpCol.Add sLst(I)
For Z = 1 To TmpCol.Count
If TmpCol(1) = sItem Then InLstEnh = True: Exit Function
TmpCol.Remove (1)
Next Z
Next I
End Function

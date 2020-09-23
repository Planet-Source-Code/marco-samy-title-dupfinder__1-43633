Attribute VB_Name = "DirFile"
'//////////////////////////////////////////////////////////////////////////
'/////////////////File And Folder Control Module For Magic Copy
'/////////////////By Marco Samy 2002
'//////////////////////////////////////////////////////////////////////////
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetLogicalDrives Lib "kernel32" () As Long
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Public FreeNum As Integer
Public mCurPath As String, mCancel As Boolean
Function GetDirs(sPath As String, sColl As Collection)
On Error Resume Next
Dim Dirx As String
'Dir all Directories including hidden and readonly files
Dirx = Dir(NormPath(sPath), vbDirectory + vbHidden + vbReadOnly + vbSystem + vbArchive)
While Not Dirx = "" 'when the files finished they will give us Null String
If Not Left$(Dirx, 1) = "." Then 'the Directories "." and ".." is not needed
If GetAttr(NormPath(sPath) & Dirx) And vbDirectory Then sColl.Add NormPath(sPath) & Dirx
End If
Dirx = Dir 'this function every time give us a newer value, it hold an information in the memory and give us the result as we ask it
Wend 'loop
End Function
'Dir files in a directory
Function GetFiles(sPath As String, sColl As Collection)
On Error Resume Next
Dim Dirx As String
'Dir all files including hidden and readonly files
Dirx = Dir(NormPath(sPath), vbNormal + vbHidden + vbReadOnly + vbSystem + vbArchive)
While Not Dirx = "" 'when the files finished they will give us Null String
If Not Left$(Dirx, 1) = "." Then sColl.Add NormPath(sPath) & Dirx 'the Directories "." and ".." is not needed
Dirx = Dir 'this function every time give us a newer value, it hold an information in the memory and give us the result as we ask it
Wend 'loop
End Function
'Extarcts the normal path from any path ( a path with "\" at the end)
Function NormPath(sPath As String) As String
If Right$(sPath, 1) = "\" Then NormPath = sPath Else NormPath = sPath & "\"
End Function
'Removes the "\" from the end if it found
Function NoNormPath(sPath As String) As String
If Right$(sPath, 1) = "\" Then NoNormPath = Left$(sPath, Len(sPath) - 1) Else NoNormPath = sPath
End Function

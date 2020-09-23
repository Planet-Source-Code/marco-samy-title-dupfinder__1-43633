VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////Class Drive Information
Rem for Magic Copy - By Marco Samy Nasif 2002
'///////////////////////////////////////////////////////////////////////////////////////
'the following code to get some infomation about drives
Option Explicit
Private iName As String
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetLogicalDrives Lib "kernel32" () As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Public Enum DriveTypes
drvCDROM = 5
drvRamdisk = 6
drvRemote = 4
drvRemovable = 2
drvFixed = 1
End Enum
'property name to get set the current working drive
Public Property Get Name() As String
   Name = iName
End Property
Public Property Let Name(ByVal Name As String)
   iName = Left(Name, 1)
End Property
Public Property Get DriveList() As Object 'The Object is a collection
Dim strDrive As String
Dim dwDrives As Long
Dim Drives As New Collection
dwDrives = GetLogicalDrives()
strDrive = "A"
Do While (dwDrives > 0)
If (dwDrives And 1) = 1 Then ' A logical drive with this letter exists
 Drives.Add strDrive
 dwDrives = dwDrives - 1
End If
strDrive = Chr$(Asc(strDrive) + 1)
dwDrives = dwDrives \ 2
Loop
Set DriveList = Drives
End Property
Public Property Get DriveType() As Long
DriveType = GetDriveType(iName & ":\") 'Getting the working drive's type
'take a look about the drives type enumruation, Up
End Property

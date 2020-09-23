VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Folders 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Folder ..."
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7920
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList LV 
      Left            =   2880
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483643
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList TV 
      Left            =   3000
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483643
      _Version        =   393216
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   5415
      Left            =   3240
      TabIndex        =   14
      Top             =   960
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   9551
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      Icons           =   "TV"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView TV1 
      Height          =   5415
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   9551
      _Version        =   393217
      Indentation     =   529
      LineStyle       =   1
      Style           =   7
      ImageList       =   "TV"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Go"
      Height          =   255
      Left            =   7440
      TabIndex        =   11
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Up"
      Height          =   255
      Left            =   4560
      TabIndex        =   9
      ToolTipText     =   "Up One Level"
      Top             =   645
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   7
      Text            =   "C:\"
      Top             =   300
      Width           =   6735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "New Folder"
      Height          =   255
      Left            =   6360
      TabIndex        =   5
      ToolTipText     =   "Create New Folder"
      Top             =   645
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Help"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   6465
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   6465
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   6480
      Width           =   1335
   End
   Begin VB.PictureBox Pict 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3000
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   6
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "Select Folder :"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   60
      Width           =   7695
   End
   Begin VB.Label Label4 
      Caption         =   "By Marco Samy - 1/2003"
      ForeColor       =   &H80000011&
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   6480
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Path:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   300
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Folders Explorer"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   615
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Select Folder"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   615
      Width           =   2415
   End
End
Attribute VB_Name = "Folders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//////////////////////////////////////////////////////////////
'///////////////Duplicates (Copies) Finder
'//////////////////////////////////////////////////////////////
'Full Desgin, Create and Programming By
'             Marco Samy Nasif
'             Marco_s2@hotmail.com
'//////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////Private Const MAX_PATH = 260
Private Const SHGFI_ICON = &H100
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1                      '  get small icon
Private Const ILD_TRANSPARENT = &H1
Private Type SHFILEINFO 'Structure used by SHGetFileInfo
   hIcon As Long
   iIcon As Long
   dwAttributes As Long
   szDisplayName As String * MAX_PATH
   szTypeName As String * 80
End Type
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal I&, ByVal hDCDest&, ByVal X&, ByVal Y&, ByVal Flags&) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private shinfo As SHFILEINFO
'------------------------------
Public CancelSelected As Boolean
Public FolderPath As String
Public CurPath As String
Dim CH As Boolean
Private Sub Command1_Click()
CancelSelected = False
mCancel = CancelSelected
mCurPath = CurPath
If EditFolder(Text1.Text) = True Then Unload Me Else MsgBox "Unable To Create your folder.", vbCritical: Text1.Text = "C:\": Text1_LostFocus
End Sub
Function EditFolder(sDir As String) As Boolean
On Error GoTo Err1:
EditFolder = False
If Dir(sDir, vbDirectory + vbArchive + vbHidden + vbReadOnly + vbSystem) = "" Then
If MsgBox("The Directory Dosen't Exists, Do you want to create it?", vbYesNo + vbInformation) = vbYes Then
Dim TmpCol As New Collection, CurI
GetAllAB sDir, "\", "\", TmpCol
CH = False
For I = 1 To TmpCol.Count
CurI = ""
For X = 1 To I: CurI = CurI & TmpCol.Item(X) & "\": Next X
If Dir(CurI) = "" Then MkDir Left$(CurI, Len(CurI) - 1)
Next I
EditFolder = True: FolderPath = sDir
End If
Else
EditFolder = True
FolderPath = sDir
End If
Err1:
End Function
Private Sub Command2_Click()
CancelSelected = True
mCancel = CancelSelected
mCurPath = CurPath
Unload Me
End Sub
Private Function StartWithPath(sPath) As String
DoEvents
Text1.Text = sPath
Text1_LostFocus
End Function
Function SetPath(sPath)
Text1.Text = sPath
Text1_LostFocus
End Function
Private Sub Command3_Click()
Dim Msg
Msg = "Folder Selector is to select a folder (or root folder) from your computer" & vbCrLf & "to complete your action or process." & vbCrLf & vbCrLf & "all you have to do is :" & vbCrLf & "Select a folder from folders explorer OR" & vbCrLf & "Explore For a folder from folders Explorer on the right OR" & vbCrLf & "Write your new folder to be created in the text box." & vbrclf & vbCrLf & "The Press OK" & vbCrLf & vbCrLf & vbCrLf & "Created By Marco Samy."
MsgBox Msg, vbInformation
End Sub

Private Sub Command4_Click()
On Error GoTo Err1
MkDir NoNormPath(CurPath) & "\" & InputBox("Making new folder " & vbCrLf & "Enter the valid name of your new folder...")
'Force to Make refresh
Dim oPath As String
oPath = CurPath
CurPath = ""
ChangePath (oPath)
Exit Sub
Err1:
MsgBox "Unable To Craete the folder!", vbCritical
End Sub

Private Sub Command7_Click()
If Len(CurPath) <= 3 Then Exit Sub
ChangePath (GetBL("\", IIf(Right$(CurPath, 1) = "\", Left$(CurPath, Len(CurPath) - 1), CurPath)))
For Z = 1 To TV1.Nodes.Count
If (TV1.Nodes.Item(Z).Key = IIf(Right$(CurPath, 1) = "\", Left$(CurPath, Len(CurPath) - 1), CurPath)) Or (TV1.Nodes.Item(Z).Key = CurPath & "\") Then TV1.Nodes.Item(Z).Selected = True: GoTo ExitZ2
Next Z
ExitZ2:

End Sub

Private Sub Form_Load()
On Error Resume Next
CancelSelected = True
mCancel = CancelSelected
CH = True
picT.Width = 16 * Screen.TwipsPerPixelX: picT.Height = 16 * Screen.TwipsPerPixelY
TV.ListImages.Clear
Dim Drivers As clsDrive, Drv As Object, hIcon, himl As Long
Set Drivers = New clsDrive
Set Drv = Drivers.DriveList
For I = 1 To Drv.Count
himl = SHGetFileInfo(Drv.Item(I) & ":\", 0&, shinfo, Len(shinfo), SHGFI_SYSICONINDEX Or SHGFI_SMALLICON)
picT.Cls
ImageList_Draw himl, shinfo.iIcon, picT.hdc, 0, 0, ILD_TRANSPARENT
DestroyIcon shinfo.iIcon
TV.ListImages.Add , , picT.Image
TV1.Nodes.Add , tvwFirst, Drv.Item(I) & ":\", Drv.Item(I) & ":", TV.ListImages.Count
If I = 1 Then TV1.Nodes.Add Drv.Item(I) & ":\", tvwChild, "": GoTo NextI
If HasSub(Drv.Item(I) & ":") = True Then TV1.Nodes.Add Drv.Item(I) & ":\", tvwChild, ""
NextI:
Next I
If Not mCurPath = "" Then StartWithPath (mCurPath)
End Sub

Private Sub LV1_DblClick()
On Error GoTo Err1
If HasSub(IIf(Right$(CurPath, 1) = "\", CurPath, CurPath & "\") & LV1.SelectedItem.Text) Then
GoAt (IIf(Right$(CurPath, 1) = "\", CurPath, CurPath & "\") & LV1.SelectedItem.Text)
Else
ChangePath (IIf(Right$(CurPath, 1) = "\", CurPath, CurPath & "\") & LV1.SelectedItem.Text)
For Z = 1 To TV1.Nodes.Count
If TV1.Nodes.Item(Z).Key = IIf(Right$(CurPath, 1) = "\", Left$(CurPath, Len(CurPath) - 1), CurPath) Then TV1.Nodes.Item(Z).Selected = True: GoTo ExitZ2
Next Z
ExitZ2:
End If
Err1:
End Sub
Function GoAt(sPath As String)
Dim TmpCol As New Collection, CurI
GetAllAB sPath, "\", "\", TmpCol
CH = False
For I = 1 To TmpCol.Count - 1
CurI = ""
For X = 1 To I: CurI = CurI & TmpCol.Item(X) & "\": Next X
For Z = 1 To TV1.Nodes.Count
If TV1.Nodes.Item(Z).Key = CurI Then TV1_Expand TV1.Nodes.Item(Z): GoTo ExitZ
Next Z
ExitZ:
Next I
CH = True
CurI = ""
For X = 1 To I: CurI = CurI & TmpCol.Item(X) & "\": Next X
CurI = Left$(CurI, Len(CurI) - 1)
For Z = 1 To TV1.Nodes.Count
If TV1.Nodes.Item(Z).Key = CurI Then TV1_Expand TV1.Nodes.Item(Z): GoTo ExitZ2
Next Z
ExitZ2:
End Function

Private Sub Text1_LostFocus()
If EditFolder(Text1.Text) = True Then ChangePath (FolderPath)
End Sub

Private Sub TV1_Click()
picT.Width = 16 * Screen.TwipsPerPixelX: picT.Height = 16 * Screen.TwipsPerPixelY
ChangePath (TV1.SelectedItem.FullPath)
End Sub
Function HasSub(sPath As String) As Boolean
On Error GoTo Err1
sPath = IIf(Right$(sPath, 1) = "\", sPath, sPath & "\")
Dim Dx
Dx = Dir(sPath, vbDirectory + vbHidden + vbReadOnly + vbSystem + vbArchive)
BG:
Dx = Dir
If Left$(Dx, 1) = "." Then GoTo BG
If GetAttr(sPath & Dx) And vbDirectory Then GoTo Err1 Else GoTo BG
Err1:
If Dx = "" Then HasSub = False Else HasSub = True
End Function


Private Sub TV1_Expand(ByVal Node As MSComctlLib.Node)
picT.Width = 16 * Screen.TwipsPerPixelX: picT.Height = 16 * Screen.TwipsPerPixelY
Dim m_Path As String, TmpCol As New Collection
m_Path = Node.Key
If Node.Child.Text = "" Then
TV1.Nodes.Remove Node.Child.Index
GetDirs m_Path, TmpCol
For I = 1 To TmpCol.Count
'-Icon
himl = SHGetFileInfo(IIf(Right$(TmpCol.Item(I), 1) = "\", TmpCol.Item(I), TmpCol.Item(I) & "\"), 0&, shinfo, Len(shinfo), SHGFI_SYSICONINDEX Or SHGFI_SMALLICON)
picT.Cls
ImageList_Draw himl, shinfo.iIcon, picT.hdc, 0, 0, ILD_TRANSPARENT
DestroyIcon shinfo.iIcon
TV.ListImages.Add , , picT.Image
'-/Icon
TV1.Nodes.Add m_Path, tvwChild, TmpCol.Item(I), GetAL("\", IIf(Right$(TmpCol.Item(I), 1) = "\", Left$(TmpCol.Item(I), Len(TmpCol.Item(I)) - 1), TmpCol.Item(I))), TV.ListImages.Count
If HasSub(TmpCol.Item(I)) = True Then TV1.Nodes.Add TmpCol.Item(I), tvwChild, ""
Next I
End If
Node.Selected = True
If CH = True Then ChangePath (m_Path) 'Send Event to the list view to change
End Sub
Function ChangePath(ByVal sPath As String)
If sPath = CurPath Then Exit Function
picT.Width = 32 * Screen.TwipsPerPixelX: picT.Height = 32 * Screen.TwipsPerPixelY
LV1.ListItems.Clear
Set LV1.Icons = Nothing
LV.ListImages.Clear
LV.ListImages.Add , , picT.Image
Set LV1.Icons = LV
Dim TmpCol As New Collection
GetDirs sPath, TmpCol
For I = 1 To TmpCol.Count
'-Icon
himl = SHGetFileInfo(IIf(Right$(TmpCol.Item(I), 1) = "\", TmpCol.Item(I), TmpCol.Item(I) & "\"), 0&, shinfo, Len(shinfo), SHGFI_SYSICONINDEX Or SHGFI_LARGEICON)
picT.Cls
ImageList_Draw himl, shinfo.iIcon, picT.hdc, 0, 0, ILD_TRANSPARENT
DestroyIcon shinfo.iIcon
LV.ListImages.Add , , picT.Image
'-/Icon
LV1.ListItems.Add , , GetAL("\", IIf(Right$(TmpCol.Item(I), 1) = "\", Left$(TmpCol.Item(I), Len(TmpCol.Item(I)) - 1), TmpCol.Item(I))), LV.ListImages.Count
Next I
CurPath = sPath
Text1.Text = IIf(Right$(sPath, 1) = "\", sPath, sPath & "\")

End Function
Property Get DlgCaption() As String
DlgCaption = Label5.Caption
End Property
Property Let DlgCaption(ByVal NewCap As String)
Label5.Caption = NewCap
End Property


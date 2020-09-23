VERSION 5.00
Begin VB.Form Menus 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mupop 
      Caption         =   "mnu"
      Begin VB.Menu SelectAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu SelectNone 
         Caption         =   "Select None"
      End
      Begin VB.Menu SelectInv 
         Caption         =   "Invert Selection"
      End
      Begin VB.Menu is111 
         Caption         =   "-"
      End
      Begin VB.Menu CheckAll 
         Caption         =   "Check All"
      End
      Begin VB.Menu CheckNone 
         Caption         =   "Check None"
      End
      Begin VB.Menu CheckInv 
         Caption         =   "Invert Checks"
      End
      Begin VB.Menu is222 
         Caption         =   "-"
      End
      Begin VB.Menu ChO 
         Caption         =   "Choose Operation"
         Begin VB.Menu Op1 
            Caption         =   "Delete"
         End
         Begin VB.Menu Op2 
            Caption         =   "Rename"
         End
         Begin VB.Menu Op3 
            Caption         =   "Create Shortcuts"
         End
         Begin VB.Menu Op4 
            Caption         =   "Add To Report"
         End
      End
   End
End
Attribute VB_Name = "Menus"
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
'//////////////////////////////////////////////////////////////
'/////////Menu Form
Private Sub CheckAll_Click()
For I = 1 To fMain.List1.ListItems.Count
fMain.List1.ListItems(I).Checked = True
Next I
End Sub

Private Sub CheckInv_Click()
For I = 1 To fMain.List1.ListItems.Count
fMain.List1.ListItems(I).Checked = Not fMain.List1.ListItems(I).Checked
Next I
End Sub

Private Sub CheckNone_Click()
For I = 1 To fMain.List1.ListItems.Count
fMain.List1.ListItems(I).Checked = False
Next I
End Sub

Private Sub Op1_Click()
fMain.Op_Click 0
End Sub

Private Sub Op2_Click()
fMain.Op_Click 1

End Sub

Private Sub Op3_Click()
fMain.Op_Click 2
End Sub

Private Sub Op4_Click()
fMain.Op_Click 3
End Sub

Private Sub SelectAll_Click()
For I = 1 To fMain.List1.ListItems.Count
fMain.List1.ListItems(I).Selected = True
Next I
End Sub

Private Sub SelectInv_Click()
For I = 1 To fMain.List1.ListItems.Count
fMain.List1.ListItems(I).Selected = Not fMain.List1.ListItems(I).Selected
Next I
End Sub

Private Sub SelectNone_Click()
For I = 1 To fMain.List1.ListItems.Count
fMain.List1.ListItems(I).Selected = False
Next I
End Sub


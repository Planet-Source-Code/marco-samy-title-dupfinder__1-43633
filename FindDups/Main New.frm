VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fMain 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Find Copies  by Marco Samy"
   ClientHeight    =   7095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
   ControlBox      =   0   'False
   Icon            =   "Main New.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Main New.frx":030A
   ScaleHeight     =   7095
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Con 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      ForeColor       =   &H80000008&
      Height          =   5055
      Index           =   1
      Left            =   2520
      MousePointer    =   99  'Custom
      ScaleHeight     =   5025
      ScaleWidth      =   6705
      TabIndex        =   7
      Top             =   720
      Width           =   6735
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FF00FF&
         Caption         =   "Preform Bytes Test"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   3360
         Value           =   1  'Checked
         Width           =   3855
      End
      Begin FindCopies.Slide Slide1 
         Height          =   375
         Left            =   1680
         TabIndex        =   75
         Top             =   600
         Width           =   2295
         _extentx        =   4048
         _extenty        =   661
         min             =   1
         value           =   5
         backcolor       =   16711935
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   25
         Text            =   "1047552"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         Text            =   "*.*"
         Top             =   120
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   120
         MaxLength       =   2
         TabIndex        =   8
         Text            =   "1"
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Line Line3 
         X1              =   2520
         X2              =   4200
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Bytes"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3360
         TabIndex        =   26
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"Main New.frx":21D58
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1815
         Left            =   4080
         TabIndex        =   24
         Top             =   3120
         Width           =   2535
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Set The Max Size of test Patch of a file, the MOre you type More Faster and More Memory Needed (Def. 1MB)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   975
         Left            =   4080
         TabIndex        =   23
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Similarity Degree the More you Select, The More Files Not 100% Same will be shown"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   855
         Left            =   4080
         TabIndex        =   22
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Test Size"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "File Pattern To Select a fixed Pattern For Searching Files(Def. *.*)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   855
         Left            =   4080
         TabIndex        =   20
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"Main New.frx":21E43
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1215
         Left            =   120
         TabIndex        =   19
         Top             =   3720
         Width           =   3855
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Times"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         TabIndex        =   18
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Dispaly Copies More Than"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   2640
         Width           =   3855
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "5%"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2400
         TabIndex        =   16
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Like"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   15
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ABc"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   3120
         TabIndex        =   14
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ABC"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   1680
         TabIndex        =   13
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Means"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Test Degree"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "File Pattern"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.PictureBox Con 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      ForeColor       =   &H80000008&
      Height          =   5055
      Index           =   3
      Left            =   3240
      MousePointer    =   99  'Custom
      ScaleHeight     =   5025
      ScaleWidth      =   6705
      TabIndex        =   40
      Top             =   -1800
      Width           =   6735
      Begin MSComctlLib.ListView List1 
         Height          =   2415
         Left            =   120
         TabIndex        =   44
         ToolTipText     =   "Check lines you want to apply operations on it, Unchecked will not be applied, see ? for details, Right Click To Edit List."
         Top             =   360
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4260
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         _Version        =   393217
         SmallIcons      =   "Im"
         ForeColor       =   -2147483640
         BackColor       =   16777152
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Copies"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Operation"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Size of copies"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Duplicated with files"
            Object.Width           =   4939
         EndProperty
      End
      Begin VB.Label Num1 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1920
         TabIndex        =   85
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label NumAll 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   5400
         TabIndex        =   50
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "# Of all"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   4320
         TabIndex        =   49
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Op 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Create A report (Save This Founds) TXT File"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   375
         Index           =   3
         Left            =   120
         MouseIcon       =   "Main New.frx":21F28
         MousePointer    =   99  'Custom
         TabIndex        =   48
         ToolTipText     =   "Save The Result into Text Document Report in Order for free Edit"
         Top             =   4560
         Width           =   6495
      End
      Begin VB.Label Op 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Create Shortcut To Original File Instead"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   375
         Index           =   2
         Left            =   120
         MouseIcon       =   "Main New.frx":22232
         MousePointer    =   99  'Custom
         TabIndex        =   47
         ToolTipText     =   "Delete duplicates and put a shortcut to the original file instead"
         Top             =   4080
         Width           =   6495
      End
      Begin VB.Label Op 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Raname To File Name and Copy#"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   375
         Index           =   1
         Left            =   120
         MouseIcon       =   "Main New.frx":2253C
         MousePointer    =   99  'Custom
         TabIndex        =   46
         ToolTipText     =   "Keep the Original File , and Rename the Rest Of files to The to ""Copy Of""  and Original's Name ."
         Top             =   3600
         Width           =   6495
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Select Your Custom Option For Every File Here"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   2295
         Left            =   5520
         TabIndex        =   45
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Op 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   120
         MouseIcon       =   "Main New.frx":22846
         MousePointer    =   99  'Custom
         TabIndex        =   43
         ToolTipText     =   "Keep The First File and Delete the rest of files"
         Top             =   3120
         Width           =   6495
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "What To Do With Selected Items' Copies"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   2760
         Width           =   6495
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Found Copies"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.PictureBox picT 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2520
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   86
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Con 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   5055
      Index           =   0
      Left            =   360
      ScaleHeight     =   5025
      ScaleWidth      =   6705
      TabIndex        =   0
      Top             =   1440
      Width           =   6735
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   3150
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   2550
         Width           =   3975
      End
      Begin VB.Label BTN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   8
         Left            =   5265
         MouseIcon       =   "Main New.frx":22B50
         MousePointer    =   99  'Custom
         TabIndex        =   70
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label BTN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   7
         Left            =   5280
         MouseIcon       =   "Main New.frx":22E5A
         MousePointer    =   99  'Custom
         TabIndex        =   69
         Top             =   2520
         Width           =   1185
      End
      Begin VB.Label CH 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C000C0&
         Caption         =   "Compare files in 2 Directories"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF80FF&
         Height          =   375
         Index           =   1
         Left            =   120
         MouseIcon       =   "Main New.frx":23164
         MousePointer    =   99  'Custom
         TabIndex        =   57
         ToolTipText     =   "Compare Files Between Two Directories"
         Top             =   1920
         Width           =   6495
      End
      Begin VB.Label CH 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C000C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Compare files in 1 Directory"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   120
         MouseIcon       =   "Main New.frx":2346E
         MousePointer    =   99  'Custom
         TabIndex        =   56
         ToolTipText     =   "Compare Files in One Directory"
         Top             =   120
         Width           =   6495
      End
      Begin VB.Label Tip1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"Main New.frx":23778
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   1335
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   6495
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"Main New.frx":2382D
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   1335
         Left            =   120
         TabIndex        =   5
         Top             =   3600
         Width           =   6495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Directory2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   3180
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Directory1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2580
         Width           =   1095
      End
   End
   Begin VB.PictureBox Con 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   5055
      Index           =   4
      Left            =   6600
      MousePointer    =   99  'Custom
      ScaleHeight     =   5025
      ScaleWidth      =   6705
      TabIndex        =   51
      Top             =   240
      Width           =   6735
      Begin MSComctlLib.ProgressBar PB2 
         Height          =   495
         Left            =   120
         TabIndex        =   53
         Top             =   480
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2385
         Left            =   1920
         Picture         =   "Main New.frx":23917
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   2700
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "This Wizard Was Created and Programmed By  Marco Samy Nasif - El-Minia , EGYPT."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   735
         Left            =   120
         TabIndex        =   55
         Top             =   4200
         Width           =   6495
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "All Processes Done. Click Next To Start The Wizard Again, Exit to Go Out the Wizard"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   735
         Left            =   120
         TabIndex        =   54
         Top             =   1080
         Width           =   6495
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "Working Please Wait"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   0
         Width           =   3975
      End
   End
   Begin VB.PictureBox Con 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H80000008&
      Height          =   5055
      Index           =   2
      Left            =   6360
      MousePointer    =   99  'Custom
      ScaleHeight     =   5025
      ScaleWidth      =   6705
      TabIndex        =   27
      Top             =   840
      Width           =   6735
      Begin MSComctlLib.ProgressBar PB1 
         Height          =   375
         Left            =   2400
         TabIndex        =   82
         Top             =   2400
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.PictureBox picByte 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         FillColor       =   &H00004000&
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   2400
         ScaleHeight     =   225
         ScaleWidth      =   1185
         TabIndex        =   80
         Top             =   2040
         Width           =   1215
      End
      Begin VB.PictureBox picWith 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         FillColor       =   &H00004000&
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   5280
         ScaleHeight     =   225
         ScaleWidth      =   1305
         TabIndex        =   79
         Top             =   2040
         Width           =   1335
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         FillColor       =   &H00004000&
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   2040
         ScaleHeight     =   225
         ScaleWidth      =   4545
         TabIndex        =   76
         Top             =   840
         Width           =   4575
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         FillColor       =   &H00004000&
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   2040
         ScaleHeight     =   225
         ScaleWidth      =   4545
         TabIndex        =   37
         Top             =   1200
         Width           =   4575
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         FillColor       =   &H00004000&
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   2040
         ScaleHeight     =   225
         ScaleWidth      =   4545
         TabIndex        =   36
         Top             =   480
         Width           =   4575
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         FillColor       =   &H00004000&
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   2040
         ScaleHeight     =   225
         ScaleWidth      =   4545
         TabIndex        =   35
         Top             =   120
         Width           =   4575
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         FillColor       =   &H00004000&
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   5280
         ScaleHeight     =   225
         ScaleWidth      =   1305
         TabIndex        =   34
         Top             =   1560
         Width           =   1335
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         FillColor       =   &H00004000&
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   2400
         ScaleHeight     =   225
         ScaleWidth      =   1185
         TabIndex        =   33
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   6600
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   "Overall Progress"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   120
         TabIndex        =   83
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "With Bytes"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   3720
         TabIndex        =   81
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Bytes Compared"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   120
         TabIndex        =   78
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6720
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "With File"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   120
         TabIndex        =   77
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   $"Main New.frx":3AA7D
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   120
         TabIndex        =   39
         Top             =   4080
         Width           =   6495
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   $"Main New.frx":3AB13
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   975
         Left            =   120
         TabIndex        =   38
         Top             =   3120
         Width           =   6495
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Now Seaching"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "File"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Action"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Files Has Copies"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "# of Copies"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   3720
         TabIndex        =   28
         Top             =   1440
         Width           =   1575
      End
   End
   Begin MSComctlLib.ImageList Im 
      Left            =   2880
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16711935
      MaskColor       =   16777152
      _Version        =   393216
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Step Text Apears Here"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   480
      TabIndex        =   73
      Top             =   360
      Width           =   6495
   End
   Begin VB.Label lblStep 
      BackStyle       =   0  'Transparent
      Caption         =   "Step1:Begin"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   71
      Top             =   0
      Width           =   6615
   End
   Begin VB.Label BTN 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   3
      Left            =   0
      MouseIcon       =   "Main New.frx":3ABC6
      MousePointer    =   99  'Custom
      TabIndex        =   64
      Top             =   1080
      Width           =   345
   End
   Begin VB.Label BTN 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   2
      Left            =   0
      MouseIcon       =   "Main New.frx":3AED0
      MousePointer    =   99  'Custom
      TabIndex        =   63
      Top             =   720
      Width           =   345
   End
   Begin VB.Label BTN 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   1
      Left            =   0
      MouseIcon       =   "Main New.frx":3B1DA
      MousePointer    =   99  'Custom
      TabIndex        =   62
      Top             =   360
      Width           =   345
   End
   Begin VB.Label BTN 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   6
      Left            =   960
      MouseIcon       =   "Main New.frx":3B4E4
      MousePointer    =   99  'Custom
      TabIndex        =   61
      Top             =   6600
      Width           =   525
   End
   Begin VB.Label BTN 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   4
      Left            =   4200
      MouseIcon       =   "Main New.frx":3B7EE
      MousePointer    =   99  'Custom
      TabIndex        =   60
      Top             =   6600
      Width           =   675
   End
   Begin VB.Label BTN 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   5
      Left            =   5880
      MouseIcon       =   "Main New.frx":3BAF8
      MousePointer    =   99  'Custom
      TabIndex        =   59
      Top             =   6600
      Width           =   585
   End
   Begin VB.Label BTN 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   0
      Left            =   0
      MouseIcon       =   "Main New.frx":3BE02
      MousePointer    =   99  'Custom
      TabIndex        =   58
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   360
      Picture         =   "Main New.frx":3C10C
      Top             =   6480
      Width           =   6735
   End
   Begin VB.Image leftB 
      Height          =   5640
      Left            =   0
      MousePointer    =   5  'Size
      Picture         =   "Main New.frx":3D48C
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label BTNX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   330
      Index           =   10
      Left            =   20
      MouseIcon       =   "Main New.frx":4446E
      MousePointer    =   99  'Custom
      TabIndex        =   68
      Top             =   1110
      Width           =   345
   End
   Begin VB.Label BTNX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   330
      Index           =   9
      Left            =   20
      MouseIcon       =   "Main New.frx":44778
      MousePointer    =   99  'Custom
      TabIndex        =   67
      Top             =   750
      Width           =   345
   End
   Begin VB.Label BTNX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   330
      Index           =   8
      Left            =   20
      MouseIcon       =   "Main New.frx":44A82
      MousePointer    =   99  'Custom
      TabIndex        =   66
      Top             =   380
      Width           =   345
   End
   Begin VB.Label BTNX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   330
      Index           =   7
      Left            =   20
      MousePointer    =   99  'Custom
      TabIndex        =   65
      Top             =   30
      Width           =   345
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Step1:Begin"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   510
      TabIndex        =   72
      Top             =   30
      Width           =   6615
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "Step Text Apears Here"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   460
      TabIndex        =   74
      Top             =   380
      Width           =   6495
   End
End
Attribute VB_Name = "fMain"
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
Private Const Step1 = "Step1:Begin"
Private Const Step2 = "Step2:Options"
Private Const Step3 = "Step3:Searching Files"
Private Const Step4 = "Step4:Choosing Actions"
Private Const Step5 = "Step5:Applying Actions"
Dim Ox, Oy
Dim Iact As Integer
Dim DirType As Boolean
Public CurrentStep As Integer
Public Working As Boolean
Public Cancel As Boolean
Dim aOp As Integer
Const ForeC = &H808000
Const BackC = &HFFFF80
Const Op1 = "Delete"
Const Op2 = "Rename"
Const Op3 = "Shortcut"
Const Op4 = "Save"
Public Founds As Single, fAll As Single
Private Const MAX_PATH = 260
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
Dim Exts As New Collection

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ICheck
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ICheck
End Sub

Private Sub Label26_Change()
Label29.Caption = Label26.Caption
End Sub

Private Sub lblStep_Change()
Label4.Caption = lblStep.Caption
End Sub

Private Sub leftB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Ox = X: Oy = Y
End Sub

Private Sub leftB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ICheck
If Button = 1 Then Move Left + X - Ox, Top + Y - Oy
End Sub

Private Sub bsGradientLabel11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ICheck
End Sub

Private Sub bsGradientLabel2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ICheck
End Sub



Function ICheck(Optional sItem As Integer = -1)
If Not Iact = sItem Then
If Not Iact = -1 Then BTN(Iact).ForeColor = vbWhite
Iact = sItem
If Not Iact = -1 Then BTN(Iact).ForeColor = vbBlue
End If
End Function


Private Sub BTN_Click(Index As Integer)
Select Case Index
Case 0
If SendMsg(Me, "Are you sure you want to exit", , True) = True Then End
Case 6
If Working = True Then Cancel = True: Exit Sub
If SendMsg(Me, "Are you sure you want to exit", , True) = True Then End
Case 2
SendMsg Me, "Find Copies ver. 1.00.237" & vbCrLf & "Created And Programmend By" & vbCrLf & "Marco Samy" & vbCrLf & "marco_s2@hotmail.com" & vbCrLf & "El-Minia, Egypt.", "About"
Case 3
If CurrentStep = 3 Then
SendMsg Me, "Select Some Files (a Multi Select List, Press Ctrl While you click Items) and Choose Action On Selected files Using The 4 available Actions Below.", "Help"
SendMsg Me, "Check Files you want to apply actions on it, items UnChecked It's action will not be applied Means Lines you not check the program will leave it.", "Help"
SendMsg Me, "Execute Selected Actions on Checked files by Clicking Next.", "Help"
Else
SendMsg Me, "Every Thing is Explained in the Tips.", "Help"
End If
Case 1
WindowState = vbMinimized
Case 5
LoadNextStep
Case 4
LoadBackStep
Case 7
mCurPath = Text1.Text
Folders.Show 1, Me
If Not mCancel Then Text1.Text = mCurPath
Case 8
mCurPath = Text2.Text
Folders.Show 1, Me
If Not mCancel Then Text2.Text = mCurPath
End Select
End Sub
Function LoadNextStep()
'Each Case Before Converting to it we must be sure
'that the right information is entried
Select Case CurrentStep
Case 0
If Trim(Text1.Text) = "" Then SendMsg Me, "Select a Valid Directory First.": Exit Function
If (DirType = True) And (Trim(Text2.Text)) = "" Then SendMsg Me, "Select The Second Directory, Or change the Second Directory Type.": Exit Function
If NormPath(Trim$(Text1.Text)) = NormPath(Trim$(Text2.Text)) Then CH_MouseDown 0, 1, 1, 1, 1
Case 4
CurrentStep = -1
End Select
CurrentStep = CurrentStep + 1
Con(CurrentStep).ZOrder 0
'///after chiking validty
'execute current step
DoCurrentStep
End Function
'Searching Files
Function ActionStep3()
'On Error Resume Next
'Empty Collection of extensions
For I = 1 To Exts.Count
Exts.Remove 1
Next I
'Empty List of Icons
Set List1.SmallIcons = Nothing
Im.ListImages.Clear
Im.ListImages.Add , , picT.Image
Set List1.SmallIcons = Im
'Changing Diaplay as the program is working now
Cancel = False
BTN(5).Enabled = False
BTN(4).Enabled = False
BTN(6).Caption = "Cancel"
Working = True
DoEvents
If Check1.Value = 1 Then PreformTest = True Else PreformTest = False
Dim ColfounDs As New Collection
FindDublics DirType, Text1.Text, Text2.Text, Text4.Text, Slide1.Value, ColfounDs, Picture5, Picture6, Picture3, Picture4, Text5.Text, Val(Text3.Text)
Working = False
If Cancel = True Then Cancel = False: SendMsg Me, "Action Cancelled." 'Dispaly Cancel Message
List1.ListItems.Clear
For I = 1 To ColfounDs.Count
List1.ListItems.Add , , GetBF(",", ColfounDs.Item(I), 1), , GetIconNum(GetBF(",", ColfounDs.Item(I), 1))
List1.ListItems(I).SubItems(1) = FindCount(",", ColfounDs(I))
List1.ListItems(I).SubItems(2) = Op1
List1.ListItems(I).SubItems(4) = GetAF(",", ColfounDs(I), 1)
List1.ListItems(I).SubItems(3) = SpaceValue(GetLenOf(List1.ListItems(I).SubItems(4)))
List1.ListItems(I).Checked = True
Next I
Num1 = Founds
NumAll = fAll
'Returning the display back
BTN(5).Enabled = True
BTN(4).Enabled = True
BTN(6).Caption = "Exit"
Working = False
LoadNextStep
'if minimized , we restore it when it done
WindowState = vbNormal
End Function
Function GetIconNum(sFile As String)
If InColl(GetAL(".", sFile), Exts) = 0 Then
Dim hIcon, himl As Long
himl = SHGetFileInfo(sFile, 0&, shinfo, Len(shinfo), SHGFI_SYSICONINDEX Or SHGFI_SMALLICON)
picT.Cls
ImageList_Draw himl, shinfo.iIcon, picT.hdc, 0, 0, ILD_TRANSPARENT
DestroyIcon shinfo.iIcon
Im.ListImages.Add , , picT.Image
Exts.Add GetAL(".", sFile)
GetIconNum = Im.ListImages.Count
Else
GetIconNum = InColl(GetAL(".", sFile), Exts) + 1
End If
End Function
Function InColl(sItem As Variant, sColl As Collection) As Integer
InColl = 0
For I = 1 To sColl.Count
If LCase(sColl.Item(I)) = LCase(sItem) Then InColl = I: Exit Function
Next I
End Function
Function ActionApply()
'Changing Diaplay as the program is working now
Label37.Visible = False: Image1.Visible = False: Label38.Visible = False: Label45.Visible = True: DoEvents
Cancel = False
BTN(5).Enabled = False
BTN(4).Enabled = False
BTN(6).Caption = "Cancel"
Working = True
DoEvents
'Reset Progress Bar Value
PB2.Value = 0
Dim TmpFile As String, nf
Randomize: nf = FreeFile 'set number of nf is a free file to open it
TmpFile = IIf(Right(App.Path, 1) = "\", Left$(App.Path, Len(App.Path) - 1), App.Path) & "\Tmpfilec." & CStr(Int(Rnd * 999) + 1)
Dim TmpCol As New Collection 'we must craete collection contains the other files
For I = 1 To List1.ListItems.Count
If Cancel = True Then GoTo EndIt
If List1.ListItems(I).Checked = False Then GoTo NextI
'Empty Collection(Clear It)
For Z = 1 To TmpCol.Count
TmpCol.Remove (1) 'remove the first One( We Can Use [[[TmpCol.Remove(TmpCol.Count)]]])
Next Z
'/Empty Collection
'Now We Execute Actions
GetAllAB List1.ListItems(I).SubItems(4), ",", ",", TmpCol 'we collect our other files
If TmpCol.Count = 0 Then TmpCol.Add List1.ListItems(I).SubItems(4) 'the collection must contains something
Select Case List1.ListItems(I).SubItems(2)
Case Op1 'delete the other files
For Z = 1 To TmpCol.Count
SetAttr TmpCol.Item(Z), vbNormal
Kill TmpCol.Item(Z)
Next Z
Case Op2 'Rename to Copy#of
For Z = 1 To TmpCol.Count
Name TmpCol.Item(Z) As GetBL("\", TmpCol.Item(Z)) & "\Copy #" & Z & " Of " & GetAL("\", List1.ListItems(I).Text)
Next Z
Case Op3 'Craete Shortcut Instead
For Z = 1 To TmpCol.Count
SetAttr TmpCol.Item(Z), vbNormal
Kill TmpCol.Item(Z) 'Delete it firs
'Now We leave a short cut instead of it
'the sortcut name is it's same name without extension
'shortcut refers to original(first file) path
CreateShellLink GetBL(".", TmpCol.Item(Z)) & ".lnk", List1.ListItems(I).Text, "", "", "", 0, SHOWNORMAL
'we are done
Next Z
Case Op4 'save report info
'we add information to the report file
'so we oppen it as append shared...
Open TmpFile For Append Shared As #nf
'put information into file
Print #nf, "Original is : " & List1.ListItems(I).Text & ", " & List1.ListItems(I).SubItems(1) & " Duplictes Found is : " & List1.ListItems(I).SubItems(4)
Close #nf
End Select
NextI:
'update progress value
PB2.Value = I / List1.ListItems.Count * 100
DoEvents
Next I
'Save report if found
If Not Dir$(TmpFile) = "" Then 'is report created?
Dim ToFile As String 'where will we save it?
ToFile = DialogFile(fMain, "Save Report", App.Path & "\Report.txt", "*.txt", App.Path, "*.txt")
If ToFile = TmpFile Then Kill TmpFile: GoTo EndIt 'no need more if the save file is our temp file
If ToFile = "" Then GoTo EndIt 'no file selected or cancel was selected
FileCopy TmpFile, ToFile 'we copy the temp file as the save file
Kill TmpFile 'delete the temp file
End If
EndIt:
'Returning the display back
If Cancel = True Then Cancel = False: SendMsg Me, "Action Cancelled." 'Dispaly Cancel Message
BTN(5).Enabled = True
BTN(4).Enabled = True
Working = False
BTN(6).Caption = "Exit"
Label37.Visible = True: Image1.Visible = True: Label38.Visible = True: Label45.Visible = False: DoEvents
Label26.Caption = "Congratulations...Every thing is completed now....press Next to restart the Wizard again."
End Function
Function GetLenOf(sFiles As String) As Double
On Error Resume Next
GetLenOf = 0
Dim TmpCOlx As New Collection
GetAllAB sFiles, ",", ",", TmpCOlx
If TmpCOlx.Count = 0 Then TmpCOlx.Add sFiles
For I = 1 To TmpCOlx.Count
GetLenOf = GetLenOf + FileLen(TmpCOlx(I))
Next I
End Function
Function LoadBackStep()
CurrentStep = CurrentStep - 1
Select Case CurrentStep
Case -1
CurrentStep = 0
SendMsg Me, "This is the first screen, please click next not back.": Exit Function
Case 0
lblStep.Caption = Step1
Case 1
lblStep.Caption = Step2
Case 2
CurrentStep = CurrentStep - 1: DoCurrentStep: Exit Function
Case 3
SendMsg Me, "Cannot Go Back, Please Restart Your Wizard.": CurrentStep = CurrentStep + 1: Exit Function
Case 4
lblStep.Caption = Step5
End Select
Con(CurrentStep).ZOrder 0
End Function
Function DoCurrentStep()
Select Case CurrentStep
Case 0
lblStep.Caption = Step1
Label26.Caption = "Welcom in the File Copies Locators, This wizard will help you to free up alot of free space on your system, first of all you must choose the directory that you want to find copies in or the the two directories if you want to compare"
DoEvents
Case 1
lblStep.Caption = Step2
Label26.Caption = "Select Options Here, you will find here the options you customize to get the maximum useful of the program, in Compare Degree I prefer 2% with Check [Preform Bytes Test]"
DoEvents
Case 2
lblStep.Caption = Step3
Label26.Caption = "It's Working Now, Minimize Me and I will tell you when I finished... Remeber that the Engine was tested carefully and it dosen't display file more than one time..I'm sure of that,stay back and relax or Minimize and work on other program."
DoEvents
ActionStep3
Case 3
lblStep.Caption = Step4
Label26.Caption = "Choose Actions, Every One you select choose it's Own action from buttons below, actions are diplayed in the [Operation], Remeber It's a Multi Select List, Press Ctrl for Multi Select to Choose 1 Action for More than 1 Line."
DoEvents
Case 4
lblStep.Caption = Step5
Label26.Caption = "The Options you choosed is beeing processed now, it will take small time to aplly...to cancel press cancel, you are almost done."
DoEvents
ActionApply
End Select
Con(CurrentStep).ZOrder 0
End Function
Private Sub BTN_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
BTN(Index).ForeColor = vbYellow
End Sub

Private Sub BTN_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 0 Then ICheck Index
End Sub

Private Sub BTN_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ICheck
End Sub

Private Sub CH_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
If DirType = False Then Exit Sub Else DirType = False
CH(0).BorderStyle = 1
CH(1).ForeColor = CH(0).ForeColor
CH(0).ForeColor = vbWhite
Tip1.Top = CH(0).Top + CH(0).Height
CH(1).Top = Tip1.Top + Tip1.Height + 120
CH(1).BorderStyle = 0
Tip1.Caption = "Find Copies 1 One Directory Files Means Find Files What's the Same in data. that means the program will scan all files in this directory and will scan every file with the other."
Else
If DirType = True Then Exit Sub Else DirType = True
CH(0).BorderStyle = 0
CH(0).ForeColor = CH(1).ForeColor
CH(1).ForeColor = vbWhite
CH(1).Top = CH(0).Top + CH(0).Height + 120
Tip1.Top = CH(1).Top + CH(1).Height
CH(1).BorderStyle = 1
Tip1.Caption = "Find Copies Between two directories' Files Means Find Files What's the Same in data between two directories. that means the program will scan all files in the first directory  and will scan every file with files in the second directory."
End If
Label2.Visible = DirType
Text2.Visible = DirType
BTN(8).Visible = DirType
End Sub

Private Sub Con_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ICheck -1
End Sub

Private Sub Form_Load()
Label26.Caption = "Welcom in the File Copies Locators, This wizard will help you to free up alot of free space on your system, first of all you must choose the directory that you want to find copies in or the the two directories if you want to compare"
Con(0).ZOrder 0
For I = 1 To Con.UBound
Con(I).Move Con(0).Left, Con(0).Top
Next I
DirType = False
End Sub
Function ChangeOp(iOp As Integer)
On Error Resume Next
If iOp = aOp Then Exit Function 'Must have changes
'remove old Op focus
Op(aOp).BackColor = BackC
Op(aOp).ForeColor = ForeC
'Setting focus to new op
Op(iOp).BackColor = vbWhite
Op(iOp).ForeColor = 0 'black
aOp = iOp
End Function

Private Sub List1_Click()
On Error GoTo Err1
Dim SelOp As Integer
'Selecting current operation to activate
Select Case List1.SelectedItem.SubItems(2)
Case Op1
SelOp = 0
Case Op2
SelOp = 1
Case Op3
SelOp = 2
Case Op4
SelOp = 3
End Select
'activate it
ChangeOp SelOp
Err1:
End Sub

Private Sub List1_DblClick()
PopupMenu Menus.mupop
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Menus.mupop
End Sub

'Remeber Cols
'1 file
'2 Copies
'3 Operation
'4 Size
'Default Operation is Delete
Public Sub Op_Click(Index As Integer)
ChangeOp Index
Dim CurOpration As String 'Current Operation
'Selecting Operation
Select Case Index
Case 0
CurOpration = Op1
Case 1
CurOpration = Op2
Case 2
CurOpration = Op3
Case 3
CurOpration = Op4
End Select
'applying Operation to selected items
For I = 1 To List1.ListItems.Count
If List1.ListItems(I).Selected = True Then
List1.ListItems(I).SubItems(2) = CurOpration
End If
Next I
End Sub


Private Sub Slide1_Seeking()
Label11.Caption = Slide1.Value & "%"
End Sub

Private Sub Text3_Change()
Text3.Text = Trim(Text3.Text)
If Val(Text3.Text) < 0 Then Text3.Text = 0
If Val(Text3.Text) > 99 Then Text3.Text = 99

End Sub

Private Sub Text5_Change()
Text3.Text = Trim(Text3.Text)
If Val(Text3.Text) < 1000 Then Text3.Text = 1000
If Val(Text3.Text) > 10047552 Then Text3.Text = 10047552
End Sub

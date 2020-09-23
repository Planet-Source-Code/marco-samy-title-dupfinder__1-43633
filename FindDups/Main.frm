VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Main 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Find Copies  by Marco Samy"
   ClientHeight    =   7095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin FindCopies.bsGradientLabel BTN 
      Height          =   375
      Index           =   5
      Left            =   5400
      Top             =   6600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Colour1         =   14737632
      Colour2         =   4210752
      Colour4         =   0
      CaptionAlignment=   1
      BorderStyle     =   3
      HighlightColour =   0
      HighlightDKColour=   12632256
      ShadowColour    =   64
      FlatBorderColour=   16711935
      TextShadowColour=   8421504
      TextShadow      =   -1  'True
      TextShadowYOffset=   1
      MousePointer    =   99
      MouseIcon       =   "Main.frx":030A
   End
   Begin VB.PictureBox Con 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   5055
      Index           =   0
      Left            =   7080
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   3150
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   2550
         Width           =   4095
      End
      Begin FindCopies.bsGradientLabel BTN 
         Height          =   375
         Index           =   7
         Left            =   5280
         Top             =   2520
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Browse"
         BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Colour1         =   14737632
         Colour2         =   4210752
         Colour4         =   0
         CaptionAlignment=   1
         BorderStyle     =   3
         HighlightColour =   0
         HighlightDKColour=   12632256
         ShadowColour    =   64
         FlatBorderColour=   16711935
         TextShadowColour=   8421504
         TextShadow      =   -1  'True
         TextShadowYOffset=   1
         MousePointer    =   99
         MouseIcon       =   "Main.frx":0624
      End
      Begin FindCopies.bsGradientLabel BTN 
         Height          =   375
         Index           =   8
         Left            =   5280
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Browse"
         BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Colour1         =   14737632
         Colour2         =   4210752
         Colour4         =   0
         CaptionAlignment=   1
         BorderStyle     =   3
         HighlightColour =   0
         HighlightDKColour=   12632256
         ShadowColour    =   64
         FlatBorderColour=   16711935
         TextShadowColour=   8421504
         TextShadow      =   -1  'True
         TextShadowYOffset=   1
         MousePointer    =   99
         MouseIcon       =   "Main.frx":093E
      End
      Begin FindCopies.bsGradientLabel CH 
         Height          =   375
         Index           =   0
         Left            =   120
         Top             =   120
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   661
         Caption         =   "Compare files in 1 Directory"
         BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Colour1         =   8388736
         Colour2         =   16711935
         Colour4         =   0
         CaptionAlignment=   1
         BorderStyle     =   2
         HighlightColour =   0
         HighlightDKColour=   12632256
         ShadowColour    =   4194368
         ShadowDKColour  =   12583104
         FlatBorderColour=   16711935
         TextShadowColour=   8388736
         TextShadow      =   -1  'True
         TextShadowYOffset=   1
         MousePointer    =   99
         MouseIcon       =   "Main.frx":0C58
      End
      Begin FindCopies.bsGradientLabel CH 
         Height          =   375
         Index           =   1
         Left            =   120
         Top             =   2040
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   661
         Caption         =   "Compare files in 2 Directories"
         BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionColour   =   16744703
         Colour1         =   8388736
         Colour2         =   16711935
         Colour4         =   0
         CaptionAlignment=   1
         BorderStyle     =   3
         HighlightColour =   0
         HighlightDKColour=   12632256
         ShadowColour    =   4194368
         ShadowDKColour  =   12583104
         FlatBorderColour=   16711935
         TextShadowColour=   8388736
         TextShadow      =   -1  'True
         TextShadowYOffset=   1
         MousePointer    =   99
         MouseIcon       =   "Main.frx":0F72
      End
      Begin VB.Label Tip1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"Main.frx":128C
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
         Caption         =   $"Main.frx":1341
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
         Caption         =   "Directory 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   3180
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Directory 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2580
         Width           =   855
      End
   End
   Begin VB.PictureBox Con 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H80000008&
      Height          =   5055
      Index           =   2
      Left            =   7080
      MousePointer    =   99  'Custom
      ScaleHeight     =   5025
      ScaleWidth      =   6705
      TabIndex        =   27
      Top             =   1320
      Width           =   6735
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         FillColor       =   &H00004000&
         ForeColor       =   &H00004000&
         Height          =   495
         Left            =   2040
         ScaleHeight     =   465
         ScaleWidth      =   4545
         TabIndex        =   37
         Top             =   840
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
         Height          =   375
         Left            =   5280
         ScaleHeight     =   345
         ScaleWidth      =   1305
         TabIndex        =   34
         Top             =   1440
         Width           =   1335
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         FillColor       =   &H00004000&
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   2400
         ScaleHeight     =   345
         ScaleWidth      =   1185
         TabIndex        =   33
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   $"Main.frx":1426
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
         Height          =   1695
         Left            =   120
         TabIndex        =   39
         Top             =   3240
         Width           =   6495
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   $"Main.frx":152A
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
         Height          =   1335
         Left            =   120
         TabIndex        =   38
         Top             =   1800
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
         Caption         =   "Searched Files"
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
         Top             =   720
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
   Begin VB.PictureBox Con 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   5055
      Index           =   4
      Left            =   7080
      MousePointer    =   99  'Custom
      ScaleHeight     =   5025
      ScaleWidth      =   6705
      TabIndex        =   52
      Top             =   1320
      Width           =   6735
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   495
         Left            =   120
         TabIndex        =   54
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
         Picture         =   "Main.frx":15DD
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
         TabIndex        =   56
         Top             =   4200
         Width           =   6495
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "All Processes Done. Click Begin To Start The Wizard Again, Exit to Quite the Wizard"
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
         TabIndex        =   53
         Top             =   0
         Width           =   3975
      End
   End
   Begin VB.PictureBox Con 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      ForeColor       =   &H80000008&
      Height          =   5055
      Index           =   3
      Left            =   7080
      MousePointer    =   99  'Custom
      ScaleHeight     =   5025
      ScaleWidth      =   6705
      TabIndex        =   40
      Top             =   1320
      Width           =   6735
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         ItemData        =   "Main.frx":18743
         Left            =   5520
         List            =   "Main.frx":18753
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   2640
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2415
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4260
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777152
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Copies"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Size"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "99999999"
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
         TabIndex        =   51
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
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   43
         Top             =   3120
         Width           =   6495
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "What To Do With Copies"
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
         Width           =   1575
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
         Width           =   1935
      End
   End
   Begin VB.PictureBox Con 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      ForeColor       =   &H80000008&
      Height          =   5055
      Index           =   1
      Left            =   7080
      MousePointer    =   99  'Custom
      ScaleHeight     =   5025
      ScaleWidth      =   6705
      TabIndex        =   7
      Top             =   1200
      Width           =   6735
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1680
         TabIndex        =   25
         Text            =   "1"
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
         TabIndex        =   8
         Text            =   "1"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1215
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
         Caption         =   $"Main.frx":18785
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
         Caption         =   $"Main.frx":18870
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
         Height          =   1575
         Left            =   120
         TabIndex        =   19
         Top             =   3360
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
         Caption         =   "0%"
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
         Left            =   2400
         TabIndex        =   16
         Top             =   840
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
      Begin VB.Line Line1 
         BorderColor     =   &H00800080&
         BorderWidth     =   3
         X1              =   1680
         X2              =   3960
         Y1              =   720
         Y2              =   720
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
   Begin FindCopies.bsGradientLabel BTN 
      Height          =   375
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Colour1         =   14737632
      Colour2         =   33023
      Colour4         =   0
      CaptionAlignment=   1
      BorderStyle     =   3
      HighlightColour =   0
      HighlightDKColour=   12632256
      ShadowColour    =   64
      FlatBorderColour=   16711935
      TextShadowColour=   8421504
      TextShadow      =   -1  'True
      TextShadowYOffset=   1
      MousePointer    =   99
      MouseIcon       =   "Main.frx":18924
   End
   Begin FindCopies.bsGradientLabel BTN 
      Height          =   375
      Index           =   1
      Left            =   0
      Top             =   360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Colour1         =   14737632
      Colour2         =   33023
      Colour4         =   0
      CaptionAlignment=   1
      BorderStyle     =   3
      HighlightColour =   0
      HighlightDKColour=   12632256
      ShadowColour    =   64
      FlatBorderColour=   16711935
      TextShadowColour=   8421504
      TextShadow      =   -1  'True
      TextShadowYOffset=   1
      MousePointer    =   99
      MouseIcon       =   "Main.frx":18C3E
   End
   Begin FindCopies.bsGradientLabel BTN 
      Height          =   375
      Index           =   2
      Left            =   0
      Top             =   720
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Colour1         =   14737632
      Colour2         =   33023
      Colour4         =   0
      CaptionAlignment=   1
      BorderStyle     =   3
      HighlightColour =   0
      HighlightDKColour=   12632256
      ShadowColour    =   64
      FlatBorderColour=   16711935
      TextShadowColour=   8421504
      TextShadow      =   -1  'True
      TextShadowYOffset=   1
      MousePointer    =   99
      MouseIcon       =   "Main.frx":18F58
   End
   Begin FindCopies.bsGradientLabel BTN 
      Height          =   375
      Index           =   3
      Left            =   0
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Colour1         =   14737632
      Colour2         =   33023
      Colour4         =   0
      CaptionAlignment=   1
      BorderStyle     =   3
      HighlightColour =   0
      HighlightDKColour=   12632256
      ShadowColour    =   64
      FlatBorderColour=   16711935
      TextShadowColour=   8421504
      TextShadow      =   -1  'True
      TextShadowYOffset=   1
      MousePointer    =   99
      MouseIcon       =   "Main.frx":19272
   End
   Begin FindCopies.bsGradientLabel BTN 
      Height          =   375
      Index           =   4
      Left            =   3720
      Top             =   6600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Colour1         =   14737632
      Colour2         =   4210752
      Colour4         =   0
      CaptionAlignment=   1
      BorderStyle     =   3
      HighlightColour =   0
      HighlightDKColour=   12632256
      ShadowColour    =   64
      FlatBorderColour=   16711935
      TextShadowColour=   8421504
      TextShadow      =   -1  'True
      TextShadowYOffset=   1
      MousePointer    =   99
      MouseIcon       =   "Main.frx":1958C
   End
   Begin FindCopies.bsGradientLabel bsGradientLabel1 
      Height          =   5655
      Left            =   0
      Top             =   1440
      Width           =   375
      _ExtentX        =   9975
      _ExtentY        =   661
      Caption         =   "Find Copies  by Marco Samy"
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   65535
      Colour1         =   8421631
      Colour2         =   255
      LabelType       =   1
      BorderStyle     =   3
      HighlightColour =   255
      HighlightDKColour=   192
      ShadowColour    =   64
      FlatBorderColour=   16711935
      TextShadowColour=   8421504
      TextShadow      =   -1  'True
      TextShadowYOffset=   1
      MousePointer    =   15
   End
   Begin FindCopies.bsGradientLabel bsGradientLabel2 
      Height          =   1095
      Left            =   360
      Top             =   360
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1931
      GradientType    =   2
      Caption         =   $"Main.frx":198A6
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   16761024
      Colour1         =   128
      Colour2         =   12583104
      Colour3         =   0
      BorderStyle     =   5
      HighlightColour =   12582912
      HighlightDKColour=   4194368
      TextShadowColour=   16711680
      TextShadow      =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin FindCopies.bsGradientLabel lblStep 
      Height          =   375
      Left            =   360
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   661
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   16711680
      Colour1         =   16744576
      BorderStyle     =   2
      HighlightColour =   12582912
      HighlightDKColour=   8388608
      ShadowColour    =   4194304
      FlatBorderColour=   12583104
      TextShadowColour=   16761024
      TextShadow      =   -1  'True
      TextShadowYOffset=   1
   End
   Begin FindCopies.bsGradientLabel BTN 
      Height          =   375
      Index           =   6
      Left            =   480
      Top             =   6600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Colour1         =   14737632
      Colour2         =   255
      Colour4         =   0
      CaptionAlignment=   1
      BorderStyle     =   3
      HighlightColour =   0
      HighlightDKColour=   12632256
      ShadowColour    =   64
      FlatBorderColour=   16711935
      TextShadowColour=   8421504
      TextShadow      =   -1  'True
      TextShadowYOffset=   1
      MousePointer    =   99
      MouseIcon       =   "Main.frx":19992
   End
   Begin FindCopies.bsGradientLabel bsGradientLabel11 
      Height          =   615
      Left            =   360
      Top             =   6480
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1085
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Colour2         =   16744576
      BorderStyle     =   5
      HighlightColour =   16744576
      HighlightDKColour=   16711680
      ShadowColour    =   12582912
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const Step1 = "Step1:Begin"
Private Const Step2 = "Step2:Options"
Private Const Step3 = "Step3:Searching Files"
Private Const Step4 = "Step4:Choosing Actions"
Private Const Step5 = "Step5:Applying Actions"
Dim Ox, Oy
Dim Iact As Integer
Dim DirType As Boolean
Public CurrentStep As Integer
Private Sub bsGradientLabel1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Ox = x: Oy = Y
End Sub

Private Sub bsGradientLabel1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
ICheck
If Button = 1 Then Move Left + x - Ox, Top + Y - Oy
End Sub

Private Sub bsGradientLabel11_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
ICheck
End Sub

Private Sub bsGradientLabel2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
ICheck
End Sub



Function ICheck(Optional sItem As Integer = -1)
If Not Iact = sItem Then
If Not Iact = -1 Then BTN(Iact).CaptionColour = vbWhite
Iact = sItem
If Not Iact = -1 Then BTN(Iact).CaptionColour = vbBlue
End If
End Function


Private Sub BTN_Click(Index As Integer)
Select Case Index
Case 0
If SendMsg(Me, "Are you sure you want to exit", , True) = True Then End
Case 6
If SendMsg(Me, "Are you sure you want to exit", , True) = True Then End
Case 2
SendMsg Me, "Find Copies ver. 1.00.237         Created And Programmend By Marco Samy         marco_s2@hotmail.com                        El-Minia, Egypt. all rights are reserved.", "About"
Case 3
SendMsg Me, "Every Thing is Explained in the Tips.", "Help"
Case 1
WindowState = vbMinimized
Case 5
LoadNextStep
Case 4
LoadPrevousStep
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
If Trim(Text1.Text) = "" Then SendMsg "Select a Valid Directory First.": Exit Sub
If (DirType = True) And (Trim(Text2.Text)) = "" Then SendMsg "Select The Second Directory, Or change the Second Directory Type.": Exit Sub
End Select
Con(CurrentStep).ZOrder 0
'///after chiking validty
'execute current step
End Function
Function LoadBackStep()
CurrentStep = CurrentStep - 1
Select Case CurrentStep
Case 0
lblStep.Caption = Step1
Case 1
lblStep.Caption = Step2
Case 2
lblStep.Caption = Step3
Case 3
lblStep.Caption = Step4
Case 4
lblStep.Caption = Step5
End Select
Con(CurrentStep).ZOrder 0
End Function

Private Sub BTN_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
BTN(Index).BorderStyle = [Sunken 3D]
BTN(Index).CaptionColour = vbYellow
End Sub

Private Sub BTN_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 0 Then ICheck Index
End Sub

Private Sub BTN_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
BTN(Index).BorderStyle = [Raised 3D]
ICheck
End Sub

Private Sub CH_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
If Index = 0 Then
If DirType = False Then Exit Sub Else DirType = False
CH(0).BorderStyle = [Raised Thin]
CH(1).CaptionColour = CH(0).CaptionColour
CH(0).CaptionColour = vbWhite
Tip1.Top = CH(0).Top + CH(0).Height
CH(1).Top = Tip1.Top + Tip1.Height + 120
CH(1).BorderStyle = [Raised 3D]
Tip1.Caption = "Find Copies 1 One Directory Files Means Find Files What's the Same in data. that means the program will scan all files in this directory and will scan every file with the other."
Else
If DirType = True Then Exit Sub Else DirType = True
CH(0).BorderStyle = [Raised 3D]
CH(0).CaptionColour = CH(1).CaptionColour
CH(1).CaptionColour = vbWhite
CH(1).Top = CH(0).Top + CH(0).Height + 120
Tip1.Top = CH(1).Top + CH(1).Height
CH(1).BorderStyle = [Raised Thin]
Tip1.Caption = "Find Copies Between two directories' Files Means Find Files What's the Same in data between two directories. that means the program will scan all files in the first directory  and will scan every file with files in the second directory."
End If
Label2.Visible = DirType
Text2.Visible = DirType
BTN(8).Visible = DirType
End Sub

Private Sub Con_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
ICheck -1
End Sub

Private Sub Form_Load()
Con(0).ZOrder 0
For I = 1 To Con.UBound
Con(I).Move Con(0).Left, Con(0).Top
Next I
DirType = False
End Sub


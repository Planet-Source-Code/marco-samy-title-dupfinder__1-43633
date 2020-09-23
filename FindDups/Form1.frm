VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3225
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Hide1 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3620
      ScaleHeight     =   375
      ScaleWidth      =   1620
      TabIndex        =   4
      Top             =   2760
      Width           =   1620
   End
   Begin VB.Label BTN 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
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
      Left            =   3990
      MouseIcon       =   "Form1.frx":37A7E
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2760
      Width           =   885
   End
   Begin VB.Label BTN 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
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
      Left            =   990
      MouseIcon       =   "Form1.frx":37D88
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2760
      Width           =   465
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2535
      Left            =   500
      TabIndex        =   1
      Top             =   150
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2535
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
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
'//////////////////////////////////////////////////////////////Dim Ox, Oy
Dim Iact As Integer
Private Sub form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Ox = X: Oy = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Move Left + X - Ox, Top + Y - Oy
ICheck
End Sub

Private Sub BTN_Click(Index As Integer)
If Index = 0 Then Accept = True Else Accept = False
Unload Me
End Sub
Function ICheck(Optional sItem As Integer = -1)
If Not Iact = sItem Then
If Not Iact = -1 Then BTN(Iact).ForeColor = vbWhite
Iact = sItem
If Not Iact = -1 Then BTN(Iact).ForeColor = vbBlue
End If
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
Private Sub Form_Load()
BTN(1).Visible = Ask: Hide1.Visible = Not Ask
Label1 = Msg
End Sub

Private Sub Label1_Change()
Label2.Caption = Label1.Caption
End Sub


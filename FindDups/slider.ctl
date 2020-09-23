VERSION 5.00
Begin VB.UserControl Slide 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF00FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleWidth      =   4800
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "slider.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   4800
      Y1              =   120
      Y2              =   120
   End
End
Attribute VB_Name = "Slide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This User Control We Need a 1 Piece of it in our project
'but you can edit it to match your project's manually
'Default Property Values:
Const m_def_IntegeralValue = True
Const m_def_Min = 0
Const m_def_Max = 100
Const m_def_Value = 0
'Property Variables:
Dim m_IntegeralValue As Boolean
Dim m_Min As Single
Dim m_Max As Single
Dim m_Value As Single
Event Changed()
Event Seeking()
Dim Ox, Oy
'Calculate Current Value from the picture's left
Function CalcValue() As Single
Dim MyDom As Single, ValDom As Single
MyDom = Width - Image1.Width
ValDom = m_Max - m_Min
CalcValue = m_Min + (Image1.Left / MyDom * ValDom)
If m_IntegeralValue Then CalcValue = CSng(Val(Format(CalcValue, "#")))
End Function
'set the picture's left by a fixed value
Function SetValue(ByVal sValue As Single)
On Error Resume Next
Dim MyDom As Single, ValDom As Single
MyDom = Width - Image1.Width
ValDom = m_Max - m_Min
Image1.Left = sValue / ValDom * MyDom
End Function
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Ox = x: Oy = Y
End Sub
'Raising Seeking is Here
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
If Val(Image1.Left + x - Ox + Image1.Width) > Width Then Ox = x
If Val(Image1.Left + x - Ox) < 0 Then Ox = x + Image1.Left
Image1.Move Image1.Left + x - Ox
Dim nVal As Single
nVal = CalcValue
If m_Value = nVal Then Exit Sub
m_Value = nVal
PropertyChanged "Value"
RaiseEvent Seeking
End If
End Sub
'Now We Can Raise Vaule Changed, if it really changed
Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim nVal As Single
Static OldValue As Single 'a static variable contains the old value
nVal = CalcValue
If OldValue = nVal Then Exit Sub
m_Value = nVal
PropertyChanged "Value"
OldValue = m_Value
RaiseEvent Changed
End Sub
Private Sub UserControl_Resize()
If Width < Image1.Width * 2 Then Width = Image1.Width * 2
Dim LineWid As Single
LineWid = (Line1.BorderWidth * Screen.TwipsPerPixelY) - 5
Line1.Y1 = (Height - LineWid) / 2
Line1.Y2 = Line1.Y1
Line1.X2 = Width
Image1.Height = Height
SetValue m_Value
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,1
Public Property Get Min() As Single
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Single)
If New_Min > m_Max Then Exit Property
    m_Min = New_Min
    PropertyChanged "Min"
If m_Value < m_Min Then Value = Min
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,100
Public Property Get Max() As Single
    Max = m_Max
End Property
Public Property Let Max(ByVal New_Max As Single)
    m_Max = New_Max
    PropertyChanged "Max"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,1
Public Property Get Value() As Single
    Value = m_Value
End Property
Public Property Let Value(ByVal New_Value As Single)
If New_Value > m_Max Then Exit Property
    m_Value = New_Value
    SetValue New_Value
    PropertyChanged "Value"
    RaiseEvent Changed
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Line1,Line1,-1,BorderColor
Public Property Get LineColor() As Long
Attribute LineColor.VB_Description = "Returns/sets the color of an object's border."
    LineColor = Line1.BorderColor
End Property
Public Property Let LineColor(ByVal New_LineColor As Long)
    Line1.BorderColor() = New_LineColor
    PropertyChanged "LineColor"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Min = m_def_Min
    m_Max = m_def_Max
    m_Value = m_def_Value
    m_IntegeralValue = m_def_IntegeralValue
End Sub
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    Line1.BorderColor = PropBag.ReadProperty("LineColor", -2147483640)
    m_IntegeralValue = PropBag.ReadProperty("IntegeralValue", m_def_IntegeralValue)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("LineColor", Line1.BorderColor, -2147483640)
    Call PropBag.WriteProperty("IntegeralValue", m_IntegeralValue, m_def_IntegeralValue)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000005)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get IntegeralValue() As Boolean
    IntegeralValue = m_IntegeralValue
End Property
Public Property Let IntegeralValue(ByVal New_IntegeralValue As Boolean)
    m_IntegeralValue = New_IntegeralValue
    PropertyChanged "IntegeralValue"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    UserControl.Refresh
    PropertyChanged "BackColor"
End Property


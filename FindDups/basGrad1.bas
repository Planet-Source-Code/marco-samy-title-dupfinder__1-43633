Attribute VB_Name = "basMsg"
' the message box Control Module
'here are
Public Accept As Boolean, Ask As Boolean, Title As String, Msg As String 'variable handles values of result and text
Function SendMsg(sForm As Form, sText As String, Optional sTitle As String = "File Copies Detector", Optional sAsk As Boolean = False) As Boolean
Ask = sAsk: Title = sTitle: Msg = sText
Form1.Show 1, sForm 'Stopping application
SendMsg = Accept 'result will be set by the form Msg and we pass it to the function's result
End Function
'Get a space in a string
Function SpaceValue(ByVal sVal As Double) As String
Dim nVal As Double, Level, xStr
nVal = Fix(sVal)
Level = 0
While nVal > 1024
nVal = nVal / 1024
Level = Level + 1
Wend
'Beautiflize it some...
nVal = Val(Format$(CStr(nVal), "###.##"))
'gnerate add string from Number Degree
Select Case Level
Case 0: xStr = "Byte"
Case 1: xStr = "KB"
Case 2: xStr = "MB"
Case 3: xStr = "GB"
Case 4: xStr = "TB"
Case 5: xStr = "QB"
End Select
'Last, Editing It.
SpaceValue = CStr(nVal) & xStr
End Function

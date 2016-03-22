Attribute VB_Name = "Module2"
Option Explicit

'Function for keys
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Declare Function joyGetDevCaps Lib "winmm.dll" Alias "joyGetDevCapsA" (ByVal id As Long, lpCaps As JOYCAPS, ByVal uSize As Long) As Long
Declare Function joyGetPos Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFO) As Long
Declare Function joySetCapture Lib "winmm.dll" (ByVal hwnd As Long, uJoyID As Long, uPeriod As Long, fChanged As Long) As Long

Type JOYCAPS
wMid As Integer
wPid As Integer
szPname As String
wXmin As Long
wXmax As Long
wYmin As Long
wYmax As Long
wZmin As Long
wZmax As Long
wNumButtons As Long
wPeriodMin As Long
wPeriodMax As Long
End Type

Type JOYINFO
wXpos As Long
wYpos As Long
wZpos As Long
wButtons As Long
End Type

Public Const JOY_BUTTON1 = &H1
Public Const JOY_BUTTON2 = &H2
Public Const JOY_BUTTON3 = &H4
Public Const JOY_BUTTON4 = &H8
Public Const JOY_BUTTON5 = &H10&
Public Const JOY_BUTTON6 = &H20&
Public Const JOY_BUTTON7 = &H40&
Public Const JOY_BUTTON8 = &H80&

'Constants for getkeys
Global Const KEY_TOGGLED As Integer = &H1
Global Const KEY_PRESSED As Integer = &H1000

Global Joy1Information As JOYINFO
Global Joy2Information As JOYINFO
Global joytestinfo As JOYINFO


Private Function Power(ByVal vBase As Long, ByVal vPower As Long) As Long
Dim i As Integer, result As Long

result = 1

For i = 1 To vPower
  result = result * vBase
Next i

Power = result

End Function
Private Function Text2Binary(ByVal BinStr As String) As Long

Dim result As Long
Dim i As Integer
Dim length As Integer

result = 0
length = Len(BinStr)
BinStr = StrReverse(BinStr)

For i = 1 To length
  If Mid(BinStr, i, 1) = "0" Then
    ' do nothing
  ElseIf Mid(BinStr, i, 1) = "1" Then
     result = result + Power(2, i - 1)
  End If
Next i

Text2Binary = result

End Function
Public Sub ParseJoyKeysPressed(ByVal wButtons As Long, ByRef KeysList() As Integer)

Dim i As Integer
Dim zeros As String, d As String

For i = 0 To 30
  zeros = String(i, "0")
  d = "1" + zeros
  If (Text2Binary(d) And wButtons) > 0 Then
    KeysList(i) = 1
  Else
    KeysList(i) = 0
  End If
Next i

  
End Sub

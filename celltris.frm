VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5265
   Icon            =   "celltris.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   245
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   351
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   2
      Left            =   960
      Top             =   3000
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   240
      Top             =   3000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  Const KEYCODE_UP As Integer = 38
  Const KEYCODE_DOWN As Integer = 40
  Const KEYCODE_LEFT As Integer = 37
  Const KEYCODE_RIGHT As Integer = 39
  Const KEYCODE_B As Integer = 66
  Const KEYCODE_D As Integer = 68
  Const KEYCODE_F As Integer = 70
  Const KEYCODE_Q As Integer = 81
  
  Const KEYCODE_0 As Integer = 48
  Const KEYCODE_1 As Integer = 49
  Const KEYCODE_2 As Integer = 50
  Const KEYCODE_3 As Integer = 51
  Const KEYCODE_4 As Integer = 52
  Const KEYCODE_5 As Integer = 53
  Const KEYCODE_6 As Integer = 54
  Const KEYCODE_7 As Integer = 55
  Const KEYCODE_8 As Integer = 56
  Const KEYCODE_9 As Integer = 57
  Const KEYCODE_ENTER As Integer = 13
  Const KEYCODE_SPACEBAR As Integer = 32
  Const KEYCODE_ESC As Integer = 27
  Const KEYCODE_DASH As Integer = 189
  Const KEYCODE_PLUS As Integer = 187
  
  'MsgBox Str(KeyCode)
  
  Select Case KeyCode
     Case KEYCODE_UP:
        RotateRightFlag = 1
     Case KEYCODE_DOWN:
        DropFlag = 1
     Case KEYCODE_LEFT:
        MoveLeftFlag = 1
     Case KEYCODE_RIGHT:
        MoveRightFlag = 1
     Case KEYCODE_B:
       SpecialActionFlag1 = 1
     Case KEYCODE_D:
        RotateLeftFlag = 1
     Case KEYCODE_F:
        RotateRightFlag = 1
     Case KEYCODE_0:
        Option_Flag_List(0) = 1
     Case KEYCODE_1:
        Option_Flag_List(1) = 1
     Case KEYCODE_2:
        Option_Flag_List(2) = 1
     Case KEYCODE_3:
        Option_Flag_List(3) = 1
     Case KEYCODE_4:
        Option_Flag_List(4) = 1
     Case KEYCODE_5:
        Option_Flag_List(5) = 1
     Case KEYCODE_6:
        Option_Flag_List(6) = 1
     Case KEYCODE_7:
        Option_Flag_List(7) = 1
     Case KEYCODE_8:
        Option_Flag_List(8) = 1
     Case KEYCODE_9:
        Option_Flag_List(9) = 1
     Case KEYCODE_ENTER:
        Option_Running_Flag = 1
     Case KEYCODE_SPACEBAR:
        Option_Running_Flag = 1
     Case KEYCODE_Q:
        Option_Paused_Flag = 1
     Case KEYCODE_DASH:
      If CurrentGameState = GAME_STATE_BEGIN Or CurrentGameState = GAME_STATE_GAME_OVER Then
        CurrentSpeed = CurrentSpeed - 1
        If CurrentSpeed < 1 Then CurrentSpeed = 1
      End If
     Case KEYCODE_PLUS:
      If CurrentGameState = GAME_STATE_BEGIN Or CurrentGameState = GAME_STATE_GAME_OVER Then
        CurrentSpeed = CurrentSpeed + 1
        If CurrentSpeed > MAX_FALL_SPEED Then CurrentSpeed = MAX_FALL_SPEED
      End If
     Case KEYCODE_ESC:
        ' *** end the program ***
        AzClose
        Unload Form1
        End
  End Select
  
  KeyPressFlag = 1
  UserInteraction
  UserInteraction2
  
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

  Const KEYCODE_UP As Integer = 38
  Const KEYCODE_DOWN As Integer = 40
  Const KEYCODE_LEFT As Integer = 37
  Const KEYCODE_RIGHT As Integer = 39
  Const KEYCODE_B As Integer = 66
  Const KEYCODE_D As Integer = 68
  Const KEYCODE_F As Integer = 70
  Const KEYCODE_Q As Integer = 81
  
  Const KEYCODE_0 As Integer = 48
  Const KEYCODE_1 As Integer = 49
  Const KEYCODE_2 As Integer = 50
  Const KEYCODE_3 As Integer = 51
  Const KEYCODE_4 As Integer = 52
  Const KEYCODE_5 As Integer = 53
  Const KEYCODE_6 As Integer = 54
  Const KEYCODE_7 As Integer = 55
  Const KEYCODE_8 As Integer = 56
  Const KEYCODE_9 As Integer = 57
  Const KEYCODE_ENTER As Integer = 13
  Const KEYCODE_SPACEBAR As Integer = 32
  Const KEYCODE_ESC As Integer = 27
  Const KEYCODE_DASH As Integer = 189
  Const KEYCODE_PLUS As Integer = 187
  
  Select Case KeyCode
     Case KEYCODE_UP:
        RotateRightFlag = 0
     Case KEYCODE_DOWN:
        DropFlag = 0
     Case KEYCODE_LEFT:
        MoveLeftFlag = 0
     Case KEYCODE_RIGHT:
        MoveRightFlag = 0
     Case KEYCODE_B:
       SpecialActionFlag1 = 0
     Case KEYCODE_D:
        RotateLeftFlag = 0
     Case KEYCODE_F:
        RotateRightFlag = 0
     Case KEYCODE_0:
        Option_Flag_List(0) = 0
     Case KEYCODE_1:
        Option_Flag_List(1) = 0
     Case KEYCODE_2:
        Option_Flag_List(2) = 0
     Case KEYCODE_3:
        Option_Flag_List(3) = 0
     Case KEYCODE_4:
        Option_Flag_List(4) = 0
     Case KEYCODE_5:
        Option_Flag_List(5) = 0
     Case KEYCODE_6:
        Option_Flag_List(6) = 0
     Case KEYCODE_7:
        Option_Flag_List(7) = 0
     Case KEYCODE_8:
        Option_Flag_List(8) = 0
     Case KEYCODE_9:
        Option_Flag_List(9) = 0
     Case KEYCODE_ENTER:
        Option_Running_Flag = 0
     Case KEYCODE_SPACEBAR:
        Option_Running_Flag = 0
     Case KEYCODE_Q:
        Option_Paused_Flag = 0
  End Select
    
  KeyPressFlag = 0
  
End Sub

Private Sub Form_Load()

InitGame Form1.hwnd
InitializeTetradShapes
CurrentGameState = GAME_STATE_BEGIN
joySetCapture Form1.hwnd, 0, 5, False
End Sub

Private Sub Form_Resize()

Dim Xr!, Yr!

   Xr = Form1.Width / Form1.ScaleWidth
   Yr = Form1.Height / Form1.ScaleHeight

   Form1.Width = GameResolutionX * Xr
   Form1.Height = GameResolutionY * Yr
   
   Form1.Top = Screen.Height / 2 - Form1.Height / 2
   Form1.Left = Screen.Width / 2 - Form1.Width / 2


End Sub

Private Sub Timer1_Timer()

Dim RefreshFlag As Boolean
Dim XOld%, YOld%
Dim LinesFormedCount%
RefreshFlag = False

If CurrentGameState = GAME_STATE_BEGIN Then

  If TimeDelay() Then
    DrawWallsStaticPlayField
    CopyStaticPlayField2OutputPlayField
    DisplayGameInfoNextTetrad
    DrawOutputPlayField2BufferText
    DisplayGameInfo
    DisplayGameStateMessage GAME_STATE_BEGIN
    Bitmap_Buffer2Window Form1.hDC
  End If
  
ElseIf CurrentGameState = GAME_STATE_GAME_OVER Then

 If TimeDelay() Then
    DrawWallsStaticPlayField
    CopyStaticPlayField2OutputPlayField
    DisplayGameInfoNextTetrad
    DrawOutputPlayField2BufferText
    DisplayGameInfo
    DisplayGameStateMessage GAME_STATE_GAME_OVER
    
    If CurrentScore > HighScore Then
       HighScore = CurrentScore
       DoHighScore 2
    End If
    
    Bitmap_Buffer2Window Form1.hDC
 End If
 
ElseIf CurrentGameState = GAME_STATE_PAUSED Then

 If TimeDelay() Then
    DrawWallsStaticPlayField
    CopyStaticPlayField2OutputPlayField
    DisplayGameInfoNextTetrad
    DrawTetrad2OutputPlayField CurrentTetrad
    DrawOutputPlayField2BufferText
    DisplayGameInfo
    DisplayGameStateMessage GAME_STATE_PAUSED
    
    ' display game paused message and info window
    ' get input to resume
    Bitmap_Buffer2Window Form1.hDC
 End If
 
ElseIf CurrentGameState = GAME_STATE_RUNNING Then
    
If InstantDropFlag Then
   InstantDrop CurrentTetrad, False
   InstantDropFlag = 0
End If

XOld = CurrentTetrad.x
YOld = CurrentTetrad.y

If DropTetrad(CurrentTetrad) Or KeyPressFlag Then
   RefreshFlag = True
End If

If CheckTetradCollideStaticPlayField(CurrentTetrad) Then
   
   CurrentTetrad.x = XOld
   CurrentTetrad.y = YOld

   PasteTetrad2StaticPlayField
  
   StartRandomTetradFlag = 1
   
   Add2ScoreHitBottom
   
   LinesFormedCount = CheckFormLines
   Add2ScoreLinesClear LinesFormedCount
   IncreaseSpeed
End If

If StartRandomTetradFlag Then
   StartRandomTetradFlag = 0
   StartRandomTetrad
   
   If CheckTetradCollideStaticPlayField(CurrentTetrad) Then
      CurrentGameState = GAME_STATE_GAME_OVER
   End If
   
End If

DrawWallsStaticPlayField
CopyStaticPlayField2OutputPlayField
    
DrawTetrad2OutputPlayField CurrentTetrad

Dim GhostTetradObject As Tetrad
GhostTetradObject = CreateGhostTetrad(CurrentTetrad)

DrawTetrad2OutputPlayField GhostTetradObject

DisplayGameInfoNextTetrad

If RefreshFlag Then

    DrawOutputPlayField2BufferText
    DisplayGameInfo

    Bitmap_Buffer2Window Form1.hDC
           
    LinesFormedCount = CheckFormLines
    Add2ScoreLinesClear LinesFormedCount
    IncreaseSpeed

End If

ClearOutputPlayField

End If

End Sub

Private Sub GamePadInteraction()
Dim myjoyinfo As JOYINFO
Static myjoyinfoprev As JOYINFO
Static DelayList(0 To 30) As Integer
Dim KeyList(0 To 30) As Integer
Dim PressedFlag As Integer
Const KEY_DELAY As Integer = 7


joyGetPos 0, myjoyinfo

'If myjoyinfo.wXpos = myjoyinfoprev.wXpos And _
'   myjoyinfo.wYpos = myjoyinfoprev.wYpos And _
'   myjoyinfo.wZpos = myjoyinfoprev.wZpos And _
'   myjoyinfo.wButtons = myjoyinfoprev.wButtons Then
'   Exit Sub
'End If

' left
If myjoyinfo.wXpos = 0 Then

  DelayList(0) = DelayList(0) + 1
  If DelayList(0) >= KEY_DELAY Then
    DelayList(0) = 0
    MoveLeftFlag = 1
    PressedFlag = 1
  Else
    MoveLeftFlag = 0
  End If
Else
  MoveLeftFlag = 0
End If
' right
If myjoyinfo.wXpos = 65535 Then
  DelayList(1) = DelayList(1) + 1
  If DelayList(1) >= KEY_DELAY Then
    DelayList(1) = 0
    MoveRightFlag = 1
    PressedFlag = 1
  Else
    MoveLeftFlag = 0
  End If
Else
  MoveRightFlag = 0
End If
' up
If myjoyinfo.wYpos = 0 Then
Else
End If
' down
If myjoyinfo.wYpos = 65535 Then
  DelayList(2) = DelayList(2) + 1
  If DelayList(2) >= KEY_DELAY Then
    DelayList(2) = 0
    DropFlag = 1
    PressedFlag = 1
  Else
    DropFlag = 0
  End If
Else
  DropFlag = 0
End If

ParseJoyKeysPressed myjoyinfo.wButtons, KeyList

'rotate left
If KeyList(3) > 0 Then
  DelayList(3) = DelayList(3) + 1
  If DelayList(3) >= KEY_DELAY Then
    DelayList(3) = 0
    RotateLeftFlag = 1
    PressedFlag = 1
  Else
    RotateLeftFlag = 0
  End If
Else
 RotateLeftFlag = 0
End If
'rotate right
If KeyList(2) > 0 Then
  DelayList(4) = DelayList(4) + 1
  If DelayList(4) >= KEY_DELAY Then
    DelayList(4) = 0
    RotateRightFlag = 1
    PressedFlag = 1
  Else
    RotateRightFlag = 0
  End If
Else
  RotateRightFlag = 0
End If

'L1
If KeyList(4) > 0 Then
  CurrentSpeed = CurrentSpeed - 1
  PressedFlag = 1
Else
End If
'R1
If KeyList(5) > 0 Then
  CurrentSpeed = CurrentSpeed + 1
  PressedFlag = 1
Else
End If

If CurrentSpeed < 0 Then CurrentSpeed = 0
If CurrentSpeed > 9 Then CurrentSpeed = 9

'select
If KeyList(8) > 0 Then
  Option_Paused_Flag = 1
  PressedFlag = 1
Else
  Option_Paused_Flag = 0
End If
'start
If KeyList(9) > 0 Then
  Option_Running_Flag = 1
  PressedFlag = 1
Else
  Option_Running_Flag = 0
End If

myjoyinfoprev = myjoyinfo

If PressedFlag > 0 Then
  UserInteraction
  UserInteraction2
End If

End Sub

Private Sub Timer2_Timer()
'GamePadInteraction
End Sub



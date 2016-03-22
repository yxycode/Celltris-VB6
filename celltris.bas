Attribute VB_Name = "Module1"
Option Explicit

Public Declare Sub Bitmap_Init Lib "bitmap.dll" Alias "Bitmap_Init@12" (ByVal hwnd As Long, ByVal nBufferWidth As Long, ByVal nBufferHeight As Long)
Public Declare Sub Bitmap_LoadBitMapFile Lib "bitmap.dll" Alias "Bitmap_LoadBitMapFile@16" (ByVal Index As Long, ByVal FileName As String, ByVal nWidth As Long, ByVal nHeight As Long)
Public Declare Sub Bitmap_LoadBitMapFileScaled Lib "bitmap.dll" Alias "Bitmap_LoadBitMapFileScaled@24" (ByVal Index As Long, ByVal FileName As String, ByVal nWidth As Long, ByVal nHeight As Long, ByVal fScaleX As Single, ByVal fScaleY As Single)
Public Declare Sub Bitmap_InitFont Lib "bitmap.dll" Alias "Bitmap_InitFont@8" (ByVal nHeight As Long, ByVal FontFace As String)
Public Declare Sub Bitmap_SetTextForeColor Lib "bitmap.dll" Alias "Bitmap_SetTextForeColor@12" (ByVal r As Long, ByVal g As Long, ByVal b As Long)
Public Declare Sub Bitmap_SetTextBackColor Lib "bitmap.dll" Alias "Bitmap_SetTextBackColor@12" (ByVal r As Long, ByVal g As Long, ByVal b As Long)
Public Declare Sub Bitmap_DrawCell Lib "bitmap.dll" Alias "Bitmap_DrawCell@12" (ByVal PicIndex As Long, ByVal x As Long, ByVal y As Long)
Public Declare Sub Bitmap_DrawText Lib "bitmap.dll" Alias "Bitmap_DrawText@12" (ByVal Text As String, ByVal x As Long, ByVal y As Long)
'Public Declare Sub Bitmap_DrawTextList Lib "bitmap.dll" Alias "Bitmap_DrawTextList@16" (TextList() As String, XList() As Long, YList() As Long, ByVal Count As Long)
Public Declare Sub Bitmap_Buffer2Window Lib "bitmap.dll" Alias "Bitmap_Buffer2Window@4" (ByVal DestDC As Long)
Public Declare Sub Bitmap_ClearBuffer Lib "bitmap.dll" Alias "Bitmap_ClearBuffer@0" ()
Public Declare Sub Bitmap_SetOption Lib "bitmap.dll" Alias "Bitmap_SetOption@4" (ByVal OptionStr As String)

Public Declare Function AzSoundInit Lib "asound.dll" Alias "_AzSoundInit@4" (ByVal hwnd As Long) As Long
Public Declare Function AzSoundInitX Lib "asound.dll" Alias "_AzSoundInitX@16" (ByVal hwnd, MonoStereo As Long, ByVal Frequency As Long, ByVal Bits As Long) As Long
Public Declare Function AzAddSound Lib "asound.dll" Alias "_AzAddSound@8" (ByVal FileName As String, ByVal Index As Long) As Long
Public Declare Function AzPlaySound Lib "asound.dll" Alias "_AzPlaySound@4" (ByVal Index As Long) As Long
Public Declare Function AzPlaySoundLooping Lib "asound.dll" Alias "_AzPlaySoundLooping@4" (ByVal Index As Long) As Long
Public Declare Function AzIsSoundPlaying Lib "asound.dll" Alias "_AzIsSoundPlaying@4" (ByVal Index As Long) As Long
Public Declare Sub AzStopSound Lib "asound.dll" Alias "_AzStopSound@4" (ByVal Index As Long)
Public Declare Sub AzClose Lib "asound.dll" Alias "_AzClose@0" ()

Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'########################################################################################################

Public Const MAX_SHAPE_WIDTH As Integer = 5
Public Const MAX_SHAPE_HEIGHT As Integer = 5
Public Const MAX_TETRAD_TYPES As Integer = 7

Public Type cColor
  r As Integer
  g As Integer
  b As Integer
End Type

Public ColorList(0 To MAX_TETRAD_TYPES + 1) As cColor
Public GameTextForeColor As cColor

' T, Z, S, J, L, I, O
' original pit dimensions 9x16

Public Const TETRAD_T As Integer = 1
Public Const TETRAD_Z As Integer = 2
Public Const TETRAD_S As Integer = 3
Public Const TETRAD_J As Integer = 4
Public Const TETRAD_L As Integer = 5
Public Const TETRAD_I As Integer = 6
Public Const TETRAD_O As Integer = 7

' custom pit dimensions 9x20
' PLAYFIELD_WIDTH = 53, PLAYFIELD_HEIGHT = 20

Public Const PLAYFIELD_WIDTH As Integer = 35
Public Const PLAYFIELD_HEIGHT As Integer = 17
 
Public Const MAX_LINES_FORMED_COUNT As Integer = 10

Public Type Tetrad
  id As Integer
  x As Integer
  y As Integer
  Cell(0 To MAX_SHAPE_WIDTH - 1, 0 To MAX_SHAPE_HEIGHT - 1) As Integer
End Type

Public TetradList(0 To MAX_TETRAD_TYPES - 1) As Tetrad

Public CurrentTetrad As Tetrad

Public StaticPlayField(0 To PLAYFIELD_WIDTH - 1, 0 To PLAYFIELD_HEIGHT - 1) As Integer
Public OutputPlayField(0 To PLAYFIELD_WIDTH - 1, 0 To PLAYFIELD_HEIGHT - 1) As Integer

' * * * for custom pit dimensions * * *
'Public Const LEFT_WALL_X As Integer = PLAYFIELD_WIDTH / 2 - 6
Public Const LEFT_WALL_X As Integer = PLAYFIELD_WIDTH / 2 - 8
Public Const RIGHT_WALL_X As Integer = LEFT_WALL_X + 10
Public Const WALL_CELL_VALUE As Integer = 8

' * * * for custom pit dimensions * * *
'Public Const TETRAD_START_X As Integer = PLAYFIELD_WIDTH / 2 - MAX_SHAPE_WIDTH / 2 - 1
Public Const TETRAD_START_X As Integer = PLAYFIELD_WIDTH / 2 - MAX_SHAPE_WIDTH / 2 - 2

Dim FallSpeedCounter As Integer
Dim FallSpeedCounterMax As Integer
Dim FallSpeedIncrement As Integer

Public MoveLeftFlag As Integer
Public MoveRightFlag As Integer
Public DropFlag As Integer
Public RotateRightFlag As Integer
Public RotateLeftFlag As Integer
Public SpecialActionFlag1 As Integer
Public KeyPressFlag As Integer

Public Option_Flag_List(0 To 9) As Integer
Public Option_Running_Flag As Integer
Public Option_Paused_Flag As Integer

Public StartRandomTetradFlag As Integer
Public InstantDropFlag As Integer

Public CurrentGameState As Integer
Public Const GAME_STATE_BEGIN As Integer = 0
Public Const GAME_STATE_GAME_OVER As Integer = 1
Public Const GAME_STATE_RUNNING As Integer = 2
Public Const GAME_STATE_PAUSED As Integer = 3

Public TimeDelayCounter As Integer
Public Const TIME_DELAY_COUNTER_MAX As Integer = 20

Public Const HIT_BOTTOM_POINTS As Integer = 2
Public Const CLEAR_LINE_POINTS As Integer = 10
Public Const LINE_COUNT_BONUS_POINTS As Integer = 5
Public Const TETRIS_BONUS_POINTS As Integer = 100

Public HighScore As Long
Public CurrentScore As Long
Public CurrentSpeed As Long
Public CurrentLineCount As Long
Public CurrentPieceIndex As Integer
Public NextPieceIndex As Integer
Public LinesClearPerLevelCount As Long
Public Const LINES_CLEAR_PER_LEVEL As Long = 50

Dim MoveLeftCounter As Integer, MoveLeftCounterMax As Integer
Dim MoveRightCounter As Integer, MoveRightCounterMax As Integer
Dim RotateRightCounter As Integer, RotateRightCounterMax As Integer
Dim RotateLeftCounter As Integer, RotateLeftCounterMax As Integer
Dim DropCounter As Integer, DropCounterMax As Integer

Dim GeneralKeyDelayCounter As Integer
Dim GeneralKeyDelayCounterMax As Integer

Const MAX_PICTURES As Integer = 100
Public TiledBackGroundFlag As Integer
Public ShowBackGroundFlag As Integer
Public BackGroundPictureIndex As Integer

Public PicWidthList(0 To MAX_PICTURES) As Integer
Public PicHeightList(0 To MAX_PICTURES) As Integer
    
Public GameResolutionX As Integer
Public GameResolutionY As Integer

Public CellWidthPixels As Integer
Public CellHeightPixels As Integer

Public CellTextWidthPixels As Integer
Public CellTextHeightPixels As Integer

Public RandomPieceBagMethodFlag As Boolean
Public RemoveLinesBombFlag As Boolean
Public RemoveLinesBombCount As Integer
Public Const INITIAL_REMOVE_LINES_BOMB_COUNT As Integer = 5
Public Const LINES_REMOVE_PER_BOMB = 8
Public Const MAX_GAME_SCORE = 999999
Public Const MAX_FALL_SPEED As Integer = 15
'########################################################################################################
Public Sub InitColors()

'ColorList(0 To MAX_TETRAD_TYPES + 1)

Randomize Timer / 3

Dim i%

'  ColorList(0).r = Rand(0, 50)
'  ColorList(0).g = Rand(0, 50)
'  ColorList(0).b = Rand(0, 50)
  
'For i = 1 To MAX_TETRAD_TYPES + 1
'  ColorList(i).r = Rand(10, 255)
'  ColorList(i).g = Rand(10, 255)
'  ColorList(i).b = Rand(10, 255)
'Next i

'GameTextForeColor.r = Rand(100, 255)
'GameTextForeColor.g = Rand(100, 255)
'GameTextForeColor.b = Rand(100, 255)

ColorList(0).r = 0
ColorList(0).g = 0
ColorList(0).b = 0

ColorList(1).r = 255
ColorList(1).g = 0
ColorList(1).b = 0

ColorList(2).r = 0
ColorList(2).g = 255
ColorList(2).b = 0

ColorList(3).r = 0
ColorList(3).g = 0
ColorList(3).b = 255

ColorList(4).r = 255
ColorList(4).g = 255
ColorList(4).b = 0

ColorList(5).r = 255
ColorList(5).g = 0
ColorList(5).b = 255

ColorList(6).r = 0
ColorList(6).g = 255
ColorList(6).b = 255

ColorList(7).r = 150
ColorList(7).g = 150
ColorList(7).b = 150

ColorList(8).r = 100
ColorList(8).g = 100
ColorList(8).b = 100

GameTextForeColor.r = 200
GameTextForeColor.g = 200
GameTextForeColor.b = 200

End Sub

Public Sub InitializeTetradShapes()
    
Dim x%, y%, T%
Dim Str(0 To MAX_TETRAD_TYPES - 1, 0 To MAX_SHAPE_HEIGHT - 1) As String
    
' T, Z, S, J, L, I, O
        
Str(0, 0) = "....."
Str(0, 1) = "..#.."
Str(0, 2) = ".###."
Str(0, 3) = "....."
Str(0, 4) = "....."

Str(1, 0) = "....."
Str(1, 1) = ".##.."
Str(1, 2) = "..##."
Str(1, 3) = "....."
Str(1, 4) = "....."

Str(2, 0) = "....."
Str(2, 1) = "..##."
Str(2, 2) = ".##.."
Str(2, 3) = "....."
Str(2, 4) = "....."

Str(3, 0) = "....."
Str(3, 1) = "..#.."
Str(3, 2) = "..#.."
Str(3, 3) = ".##.."
Str(3, 4) = "....."

Str(4, 0) = "....."
Str(4, 1) = "..#.."
Str(4, 2) = "..#.."
Str(4, 3) = "..##."
Str(4, 4) = "....."

Str(5, 0) = "....."
Str(5, 1) = "..#.."
Str(5, 2) = "..#.."
Str(5, 3) = "..#.."
Str(5, 4) = "..#.."

Str(6, 0) = "....."
Str(6, 1) = ".##.."
Str(6, 2) = ".##.."
Str(6, 3) = "....."
Str(6, 4) = "....."
        
For T = 0 To MAX_TETRAD_TYPES - 1
 
 TetradList(T).id = T + 1
 TetradList(T).x = 0
 TetradList(T).y = 0
 
 For y = 0 To MAX_SHAPE_HEIGHT - 1
   For x = 0 To MAX_SHAPE_WIDTH - 1
     If Mid(Str(T, y), x + 1, 1) = "#" Then
        TetradList(T).Cell(x, y) = T + 1
     Else
        TetradList(T).Cell(x, y) = 0
     End If
     
   Next x
 Next y
Next T

End Sub

Sub RotateTetradLeft(TheTetrad As Tetrad)

Dim TheTetrad2 As Tetrad
Dim x%, y%, X2%, Y2%

If TheTetrad.id = TETRAD_O Then Exit Sub

TheTetrad2.id = TheTetrad.id
TheTetrad2.x = TheTetrad.x
TheTetrad2.y = TheTetrad.y

    X2 = 0
   For y = 0 To MAX_SHAPE_HEIGHT - 1
       Y2 = MAX_SHAPE_HEIGHT - 1
     For x = 0 To MAX_SHAPE_WIDTH - 1
        TheTetrad2.Cell(X2, Y2) = TheTetrad.Cell(x, y)
        Y2 = Y2 - 1
     Next x
       X2 = X2 + 1
   Next y
   
   TheTetrad = TheTetrad2
   
End Sub

Sub RotateTetradRight(TheTetrad As Tetrad)

Dim TheTetrad2 As Tetrad
Dim x%, y%, X2%, Y2%

If TheTetrad.id = TETRAD_O Then Exit Sub

TheTetrad2.id = TheTetrad.id
TheTetrad2.x = TheTetrad.x
TheTetrad2.y = TheTetrad.y

    X2 = MAX_SHAPE_WIDTH - 1
    
   For y = 0 To MAX_SHAPE_HEIGHT - 1
       Y2 = 0
     For x = 0 To MAX_SHAPE_WIDTH - 1
        TheTetrad2.Cell(X2, Y2) = TheTetrad.Cell(x, y)
        Y2 = Y2 + 1
     Next x
       X2 = X2 - 1
   Next y
   
   TheTetrad = TheTetrad2
   
End Sub
Sub LoadMedia()

    Dim ScaleX!, ScaleY!, FileName$, PicWidth!, PicHeight!, Index%, BackGroundFlag%
        
    Bitmap_SetOption "initialize"
    
    Open "celltris.ini" For Input As #1
    
    While Not EOF(1)
      Input #1, Index, FileName, PicWidth, PicHeight, BackGroundFlag
      
      If InStr(1, FileName, ".bmp") > 0 Then
        
         PicWidthList(Index) = PicWidth
         PicHeightList(Index) = PicHeight
         ScaleX = CellWidthPixels / PicWidth + 0.01
         ScaleY = CellHeightPixels / PicHeight + 0.01
         
         If BackGroundFlag = 1 And TiledBackGroundFlag = 1 Then
            Bitmap_SetOption "1,1,0,"
            Bitmap_LoadBitMapFile Index, FileName, PicWidth, PicHeight
         ElseIf BackGroundFlag = 1 And TiledBackGroundFlag = 0 Then
            Bitmap_SetOption "1,1,0,"
            ScaleX = GameResolutionX / PicWidth
            ScaleY = GameResolutionY / PicHeight
            Bitmap_LoadBitMapFileScaled Index, FileName, PicWidth, PicHeight, ScaleX, ScaleY
         Else
            Bitmap_SetOption "1,1,1,"
            Bitmap_LoadBitMapFileScaled Index, FileName, PicWidth, PicHeight, ScaleX, ScaleY
         End If
         
      ElseIf InStr(1, FileName, ".wav") > 0 Then
         AzAddSound FileName, Index
      End If
    Wend
    
    Close #1
    
End Sub
Function DoHighScore(WhatDo As Integer) As Long

  If WhatDo = 1 Then
    Open "score.dat" For Input As #1
    Input #1, DoHighScore
    Close #1
  ElseIf WhatDo = 2 Then
    Open "score.dat" For Output As #1
    Print #1, HighScore
    Close #1
  End If
  
End Function
Sub InitGame(hwnd As Long)

    Dim ScaleX!, ScaleY!
    
    InitColors
    
    'GameResolutionX = Screen.Width / Screen.TwipsPerPixelX
    'GameResolutionY = Screen.Height / Screen.TwipsPerPixelY
    
    ShowBackGroundFlag = 1
    TiledBackGroundFlag = 1
    BackGroundPictureIndex = Rand(9, 14)
    
    'GameResolutionX = GetSystemMetrics(SM_CXSCREEN)
    'GameResolutionY = GetSystemMetrics(SM_CYSCREEN)
    
    Const CELL_PIXEL_WIDTH As Integer = 29
    Const CELL_PIXEL_HEIGHT As Integer = 29
    
    GameResolutionX = PLAYFIELD_WIDTH * CELL_PIXEL_WIDTH
    GameResolutionY = PLAYFIELD_HEIGHT * CELL_PIXEL_HEIGHT
    
    CellTextHeightPixels = GameResolutionY / PLAYFIELD_HEIGHT
    CellTextWidthPixels = CellTextHeightPixels * (12! / 20!)
    CellWidthPixels = GameResolutionX / PLAYFIELD_WIDTH
    CellHeightPixels = GameResolutionY / PLAYFIELD_HEIGHT
    
    Bitmap_Init hwnd, GameResolutionX, GameResolutionY
    Bitmap_InitFont CellTextHeightPixels * 1.2, "Courier"

    AzSoundInit hwnd
    LoadMedia
    HighScore = DoHighScore(1)
    
    ' * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    Randomize Timer / 3
    
    
     
    CurrentSpeed = 5
    
    FallSpeedCounterMax = 100
    FallSpeedIncrement = CurrentSpeed
    FallSpeedCounter = 0
    
    MoveLeftCounterMax = 1
    MoveRightCounterMax = 1
    RotateRightCounterMax = 1
    RotateLeftCounterMax = 1
    DropCounterMax = 1
    
    GeneralKeyDelayCounterMax = 1
    
    StartRandomTetradFlag = 0
    
    LinesClearPerLevelCount = 0
    CurrentGameState = GAME_STATE_BEGIN
    StartRandomTetrad
    
    RandomPieceBagMethodFlag = True
    RemoveLinesBombFlag = False
    RemoveLinesBombCount = INITIAL_REMOVE_LINES_BOMB_COUNT
    
End Sub

Sub DrawTetrad2OutputPlayField(TheTetrad As Tetrad)
   
  Dim x%, y%, X2%, Y2%
  
  For y = 0 To MAX_SHAPE_HEIGHT - 1
    For x = 0 To MAX_SHAPE_WIDTH - 1
        
        X2 = TheTetrad.x + x
        Y2 = TheTetrad.y + y
        
        If TheTetrad.Cell(x, y) > 0 And 0 <= X2 And X2 < PLAYFIELD_WIDTH And _
           0 <= Y2 And Y2 < PLAYFIELD_HEIGHT Then
           OutputPlayField(X2, Y2) = TheTetrad.Cell(x, y)
           
        End If
        
    Next x
  Next y
  
End Sub

Sub DrawBackGround()

 Dim Index%, PicWidth%, PicHeight%, TileXCount%, TileYCount%, x%, y%
    
 Index = BackGroundPictureIndex
 
 PicWidth = PicWidthList(Index)
 PicHeight = PicHeightList(Index)
 
 'If PicWidth = 0 Then PicWidth = 1
 'If PicHeight = 0 Then PicHeight = 1
         
 TileXCount = GameResolutionX / (PicWidth + 0.00000001) + 1
 TileYCount = GameResolutionY / (PicHeight + 0.00000001) + 1
 
 If TiledBackGroundFlag > 0 Then
   For y = 0 To TileYCount - 1
    For x = 0 To TileXCount - 1
      Bitmap_DrawCell Index, x * PicWidth, y * PicHeight
    Next x
   Next y
 Else
     Bitmap_DrawCell Index, 0, 0
 End If
 
End Sub
Sub DrawOutputPlayField2BufferText()

  Dim x%, y%, Cell%, X2%, Y2%
  
  If ShowBackGroundFlag = 1 Then
    DrawBackGround
  End If
  
  For y = 0 To PLAYFIELD_HEIGHT - 1
    For x = 0 To PLAYFIELD_WIDTH - 1
        
        Cell = OutputPlayField(x, y)
        X2 = x * CellWidthPixels
        Y2 = y * CellHeightPixels
        
        Select Case Cell
           Case 0:
             'Bitmap_SetTextBackColor 0, 0, 0
             'Bitmap_SetTextBackColor ColorList(0).r, ColorList(0).g, ColorList(0).b
           Case 1:
             'Bitmap_SetTextBackColor 255, 0, 0
             'Bitmap_SetTextBackColor ColorList(1).r, ColorList(1).g, ColorList(1).b
           Case 2:
             'Bitmap_SetTextBackColor 0, 255, 0
             'Bitmap_SetTextBackColor ColorList(2).r, ColorList(2).g, ColorList(2).b
           Case 3:
             'Bitmap_SetTextBackColor 0, 0, 255
             'Bitmap_SetTextBackColor ColorList(3).r, ColorList(3).g, ColorList(3).b
           Case 4:
             'Bitmap_SetTextBackColor 255, 255, 0
             'Bitmap_SetTextBackColor ColorList(4).r, ColorList(4).g, ColorList(4).b
           Case 5:
             'Bitmap_SetTextBackColor 255, 0, 255
             'Bitmap_SetTextBackColor ColorList(5).r, ColorList(5).g, ColorList(5).b
           Case 6:
             'Bitmap_SetTextBackColor 0, 255, 255
             'Bitmap_SetTextBackColor ColorList(6).r, ColorList(6).g, ColorList(6).b
           Case 7:
             'Bitmap_SetTextBackColor 128, 128, 128
             'Bitmap_SetTextBackColor ColorList(7).r, ColorList(7).g, ColorList(7).b
           Case 8:
             'Bitmap_SetTextBackColor 50, 50, 50   ' *** wall ***
             'Bitmap_SetTextBackColor ColorList(8).r, ColorList(8).g, ColorList(8).b
         End Select
                          
        'Bitmap_DrawText " ", X2, Y2
    
         If ShowBackGroundFlag = 1 And Cell = 0 Then
         Else
           Bitmap_DrawCell Cell, X2, Y2
         End If
        
    Next x
  Next y
  
End Sub

Sub ClearStaticPlayField()

Dim x%, y%

    For y = 0 To PLAYFIELD_HEIGHT - 1
     For x = 0 To PLAYFIELD_WIDTH - 1
       StaticPlayField(x, y) = 0
     Next x
    Next y

End Sub

Sub ClearOutputPlayField()

Dim x%, y%

    For y = 0 To PLAYFIELD_HEIGHT - 1
     For x = 0 To PLAYFIELD_WIDTH - 1
       OutputPlayField(x, y) = 0
     Next x
    Next y

End Sub

Function DropTetrad(TheTetrad As Tetrad)
    
    DropTetrad = 0
    
    FallSpeedCounter = FallSpeedCounter + FallSpeedIncrement + 1
    
    If FallSpeedCounter > FallSpeedCounterMax Then
       FallSpeedCounter = 0
       TheTetrad.y = TheTetrad.y + 1
       DropTetrad = 1
    End If
    
End Function

Function Rand(LowerBound%, UpperBound%)
   
   Rand = Int((UpperBound - LowerBound + 1) * Rnd()) + LowerBound
   
End Function
Function RandomPieceIndexBagMethod() As Integer

Static NextPieceList(0 To MAX_TETRAD_TYPES - 1) As Integer
Static NextPieceListIndex As Integer

Dim x%, y%, z%, i%

If NextPieceListIndex > MAX_TETRAD_TYPES - 1 Then
  NextPieceListIndex = 0
End If

If NextPieceListIndex = 0 Then
   For i = 0 To MAX_TETRAD_TYPES - 1
     NextPieceList(i) = i
   Next i

   For i = 0 To (MAX_TETRAD_TYPES - 1) * 2
      x = Rand(0, MAX_TETRAD_TYPES - 1)
      y = Rand(0, MAX_TETRAD_TYPES - 1)
      z = NextPieceList(x)
      NextPieceList(x) = NextPieceList(y)
      NextPieceList(y) = z
  Next i
End If

 RandomPieceIndexBagMethod = NextPieceList(NextPieceListIndex)
 
 NextPieceListIndex = NextPieceListIndex + 1

End Function

Sub StartRandomTetrad()
        
If RandomPieceBagMethodFlag Then

    If CurrentPieceIndex = 0 And NextPieceIndex = 0 Then
       CurrentPieceIndex = RandomPieceIndexBagMethod()
       NextPieceIndex = RandomPieceIndexBagMethod()
    Else
       CurrentPieceIndex = NextPieceIndex
       NextPieceIndex = RandomPieceIndexBagMethod()
    End If


Else

    If CurrentPieceIndex = 0 And NextPieceIndex = 0 Then
       CurrentPieceIndex = Rand(0, MAX_TETRAD_TYPES - 1)
       NextPieceIndex = Rand(0, MAX_TETRAD_TYPES - 1)
    Else
       CurrentPieceIndex = NextPieceIndex
       NextPieceIndex = Rand(0, MAX_TETRAD_TYPES - 1)
    End If

End If

    CurrentTetrad = TetradList(CurrentPieceIndex)
    CurrentTetrad.x = TETRAD_START_X
    CurrentTetrad.y = -2
    
End Sub

Sub DrawWallsStaticPlayField()
 
  Dim x%, y%
  
  For y = 0 To PLAYFIELD_HEIGHT - 1
    
     StaticPlayField(LEFT_WALL_X, y) = WALL_CELL_VALUE
     StaticPlayField(RIGHT_WALL_X, y) = WALL_CELL_VALUE
     
  Next y
  
  For x = LEFT_WALL_X To RIGHT_WALL_X
     StaticPlayField(x, PLAYFIELD_HEIGHT - 1) = WALL_CELL_VALUE
  Next x
  
End Sub

Sub CopyStaticPlayField2OutputPlayField()

Dim x%, y%

    For y = 0 To PLAYFIELD_HEIGHT - 1
     For x = 0 To PLAYFIELD_WIDTH - 1
       OutputPlayField(x, y) = StaticPlayField(x, y)
     Next x
    Next y
    
    
End Sub

Function CheckTetradCollideStaticPlayField(TheTetrad As Tetrad)

  Dim x%, y%, X2%, Y2%
  
  CheckTetradCollideStaticPlayField = 0
  
  For y = 0 To MAX_SHAPE_HEIGHT - 1
    For x = 0 To MAX_SHAPE_WIDTH - 1
        
        X2 = TheTetrad.x + x
        Y2 = TheTetrad.y + y
        
        If TheTetrad.Cell(x, y) > 0 And 0 <= X2 And X2 < PLAYFIELD_WIDTH And _
           0 <= Y2 And Y2 < PLAYFIELD_HEIGHT Then
         If StaticPlayField(X2, Y2) > 0 Then
           CheckTetradCollideStaticPlayField = 1
           Exit For
         End If
        End If
        
    Next x
    
    If CheckTetradCollideStaticPlayField Then
       Exit For
    End If
    
  Next y

End Function

Sub PasteTetrad2StaticPlayField()

  Dim x%, y%, X2%, Y2%
  
  For y = 0 To MAX_SHAPE_HEIGHT - 1
    For x = 0 To MAX_SHAPE_WIDTH - 1
        
        X2 = CurrentTetrad.x + x
        Y2 = CurrentTetrad.y + y
        
        If CurrentTetrad.Cell(x, y) > 0 And 0 <= X2 And X2 < PLAYFIELD_WIDTH And _
           0 <= Y2 And Y2 < PLAYFIELD_HEIGHT Then
           StaticPlayField(X2, Y2) = CurrentTetrad.Cell(x, y)
        End If
        
    Next x


  Next y
End Sub

Sub InstantDrop(TheTetrad As Tetrad, GhostFlag As Boolean)

Dim i As Integer
Dim XOld As Integer, YOld As Integer
Dim x As Integer, y As Integer

For y = PLAYFIELD_HEIGHT - 1 To 0 Step -1

     XOld = TheTetrad.x
     YOld = TheTetrad.y

     TheTetrad.y = TheTetrad.y + 1
  
  If CheckTetradCollideStaticPlayField(TheTetrad) Then
   
     TheTetrad.x = XOld
     TheTetrad.y = YOld
   
   If GhostFlag = False Then
     PasteTetrad2StaticPlayField
     
     StartRandomTetradFlag = 1
   End If
   
     Exit For
  End If
Next y

End Sub
Sub UserInteraction()

 Dim TetradBackUp As Tetrad
 
If CurrentGameState = GAME_STATE_RUNNING Then

 TetradBackUp = CurrentTetrad

 If MoveLeftFlag Then
    MoveLeftCounter = MoveLeftCounter + 1
    
    If MoveLeftCounter >= MoveLeftCounterMax Then
       CurrentTetrad.x = CurrentTetrad.x - 1
       MoveLeftCounter = 0
    End If
 Else
    MoveLeftCounter = 0
 End If
 
  If MoveRightFlag Then
    MoveRightCounter = MoveRightCounter + 1
    
    If MoveRightCounter >= MoveRightCounterMax Then
       CurrentTetrad.x = CurrentTetrad.x + 1
       MoveRightCounter = 0
    End If
 Else
    MoveRightCounter = 0
 End If

  If RotateRightFlag Then
    RotateRightCounter = RotateRightCounter + 1
    
    If RotateRightCounter >= RotateRightCounterMax Then
       RotateTetradRight CurrentTetrad
       RotateRightCounter = 0
       AzPlaySound 2
    End If
 Else
    RotateRightCounter = 0
 End If
 
   If RotateLeftFlag Then
    RotateLeftCounter = RotateLeftCounter + 1
    
    If RotateLeftCounter >= RotateLeftCounterMax Then
       RotateTetradLeft CurrentTetrad
       RotateLeftCounter = 0
       AzPlaySound 2
    End If
 Else
    RotateRightCounter = 0
 End If
 
 If DropFlag Then
    DropCounter = DropCounter + 1
    
    If DropCounter >= DropCounterMax Then
       InstantDropFlag = 1
       DropCounter = 0
       AzPlaySound 0
    End If
 Else
    DropCounterMax = 0
 End If
 
 If SpecialActionFlag1 And RemoveLinesBombFlag Then
   If RemoveLinesBombCount > 0 Then
     FillLinesBottomUp LINES_REMOVE_PER_BOMB
     RemoveLinesBombCount = RemoveLinesBombCount - 1
     AzPlaySound 0
   End If
 End If
  
 If CheckTetradCollideStaticPlayField(CurrentTetrad) Then
    CurrentTetrad = TetradBackUp
 End If
 
End If

End Sub

Function CheckFormLines()

Dim LinesY(1 To MAX_LINES_FORMED_COUNT) As Integer
Dim LinesYCount As Integer
Dim x%, y%, i%, CellCount%

    LinesYCount = 0
    
    For y = PLAYFIELD_HEIGHT - 2 To 1 Step -1
    
      CellCount = 0
      
      For x = LEFT_WALL_X + 1 To RIGHT_WALL_X - 1
         If StaticPlayField(x, y) > 0 Then
           CellCount = CellCount + 1
         End If
      Next x
      
      If CellCount >= RIGHT_WALL_X - LEFT_WALL_X - 1 Then
        
         LinesYCount = LinesYCount + 1
         LinesY(LinesYCount) = y
      End If
      
    Next y

    For i = 1 To LinesYCount
      For x = LEFT_WALL_X + 1 To RIGHT_WALL_X - 1
         StaticPlayField(x, LinesY(i)) = 0
      Next x
    Next i
     
          
For i = 1 To LinesYCount
    For y = PLAYFIELD_HEIGHT - 2 To 1 Step -1
    
      CellCount = 0
      
      For x = LEFT_WALL_X + 1 To RIGHT_WALL_X - 1
         If StaticPlayField(x, y) > 0 Then
           CellCount = CellCount + 1
         End If
      Next x
      
      If CellCount = 0 Then
         For x = LEFT_WALL_X + 1 To RIGHT_WALL_X - 1
            StaticPlayField(x, y) = StaticPlayField(x, y - 1)
            StaticPlayField(x, y - 1) = 0
         Next x
      End If
      
    Next y
Next i

      CheckFormLines = LinesYCount
End Function
Public Sub FillLinesBottomUp(LinesFillCount As Integer)

Dim x%, y%, i%
Const SOLID_CELL_1 As Integer = 1

  If LinesFillCount > PLAYFIELD_HEIGHT - 2 Then
    Exit Sub
  End If

  y = PLAYFIELD_HEIGHT - 2
  For i = 1 To LinesFillCount
    For x = LEFT_WALL_X + 1 To RIGHT_WALL_X - 1
       StaticPlayField(x, y) = SOLID_CELL_1
    Next x
    y = y - 1
  Next i

End Sub
Public Function TimeDelay()

  TimeDelay = 0
  
  TimeDelayCounter = TimeDelayCounter + 1
  
  If TimeDelayCounter >= TIME_DELAY_COUNTER_MAX Then
     TimeDelayCounter = 0
     TimeDelay = 1
  End If
  
End Function

Public Sub DisplayGameInfo()

 'Bitmap_SetTextForeColor 100, 100, 100
 'Bitmap_SetTextBackColor 0, 0, 0
 
  Bitmap_SetTextForeColor GameTextForeColor.r, GameTextForeColor.g, GameTextForeColor.b
  Bitmap_SetTextBackColor ColorList(0).r, ColorList(0).g, ColorList(0).b

 Dim W%, H%
 
 W = CellWidthPixels
 H = CellHeightPixels
 
 Bitmap_DrawText "High Score", RIGHT_WALL_X * W + 2 * W, H * 2
 Bitmap_DrawText HighScore, RIGHT_WALL_X * W + 2 * W, H * 3
 
 Bitmap_DrawText "Current Score", RIGHT_WALL_X * W + 2 * W, H * 4
 Bitmap_DrawText CurrentScore, RIGHT_WALL_X * W + 2 * W, H * 5
 
 Bitmap_DrawText "Current Speed", RIGHT_WALL_X * W + 2 * W, H * 6
 Bitmap_DrawText CurrentSpeed, RIGHT_WALL_X * W + 2 * W, H * 7
 
 Bitmap_DrawText "Lines", RIGHT_WALL_X * W + 2 * W, H * 8
 Bitmap_DrawText CurrentLineCount, RIGHT_WALL_X * W + 2 * W, H * 9
 
 Bitmap_DrawText "Next", RIGHT_WALL_X * W + 2 * W, H * 10
 
 If RemoveLinesBombFlag Then
   Bitmap_DrawText "Bombs x " + Str(RemoveLinesBombCount), RIGHT_WALL_X * W + 2 * W, H * 15
 End If
 
End Sub

Public Sub DisplayGameInfoNextTetrad()

 Dim NextTetrad As Tetrad
 Dim W%, H%
 
 If NextPieceIndex + 1 > 0 Then
 
 W = CellTextWidthPixels
 H = CellTextHeightPixels
 
 NextTetrad = TetradList(NextPieceIndex)
 NextTetrad.x = RIGHT_WALL_X + 2
 NextTetrad.y = 11

 DrawTetrad2OutputPlayField NextTetrad
 
 End If
 
End Sub

Public Sub Add2ScoreLinesClear(LineCount%)
 
 If 0 < LineCount And LineCount < 4 Then
   CurrentScore = CurrentScore + LineCount * CLEAR_LINE_POINTS + LineCount * LINE_COUNT_BONUS_POINTS
   AzPlaySound 3
 ElseIf LineCount >= 4 Then
   CurrentScore = CurrentScore + LineCount * CLEAR_LINE_POINTS + LineCount * LINE_COUNT_BONUS_POINTS + TETRIS_BONUS_POINTS
   AzPlaySound 3
 End If
 
 CurrentScore = CurrentScore + LineCount * CurrentSpeed
 
 CurrentLineCount = CurrentLineCount + LineCount
 LinesClearPerLevelCount = LinesClearPerLevelCount + LineCount
 
 If CurrentScore > MAX_GAME_SCORE Then
   CurrentScore = MAX_GAME_SCORE
 End If
 
End Sub

Public Sub Add2ScoreHitBottom()
   
  CurrentScore = CurrentScore + HIT_BOTTOM_POINTS
  
 If CurrentScore > MAX_GAME_SCORE Then
   CurrentScore = MAX_GAME_SCORE
 End If
   
End Sub

Public Sub UserInteraction2()
  Dim i As Integer
  
  GeneralKeyDelayCounter = GeneralKeyDelayCounter + 1
  
  If GeneralKeyDelayCounter >= GeneralKeyDelayCounterMax Then
   If CurrentGameState = GAME_STATE_BEGIN Or CurrentGameState = GAME_STATE_GAME_OVER Then
     GeneralKeyDelayCounter = 0
    
     For i = 0 To 9
       If Option_Flag_List(i) Then
         CurrentSpeed = i
       End If
     Next i
     
     If Option_Running_Flag Then
        CurrentGameState = GAME_STATE_RUNNING
        ClearStaticPlayField
        ClearOutputPlayField
        StartRandomTetradFlag = 1
        CurrentScore = 0
        CurrentLineCount = 0
        LinesClearPerLevelCount = 0
     End If
     
     FallSpeedIncrement = CurrentSpeed
     
     If RemoveLinesBombFlag Then
       RemoveLinesBombCount = INITIAL_REMOVE_LINES_BOMB_COUNT
     End If
     
   ElseIf CurrentGameState = GAME_STATE_RUNNING Then
          
     If Option_Paused_Flag Then
        Option_Paused_Flag = 0
        CurrentGameState = GAME_STATE_PAUSED
     End If
   
   ElseIf CurrentGameState = GAME_STATE_PAUSED Then
   
     If Option_Paused_Flag Then
        Option_Paused_Flag = 0
        CurrentGameState = GAME_STATE_RUNNING
     End If
   End If
   
  End If

End Sub

Public Sub DisplayGameStateMessage(GameState As Integer)

  Bitmap_SetTextForeColor GameTextForeColor.r, GameTextForeColor.g, GameTextForeColor.b
  Bitmap_SetTextBackColor ColorList(0).r, ColorList(0).g, ColorList(0).b

 Dim W%, H%
 
 W = CellWidthPixels
 H = CellHeightPixels
 
  Select Case GameState
  
     Case GAME_STATE_BEGIN:
          Bitmap_DrawText "Welcome", LEFT_WALL_X * W + 2 * W, H * 6
          Bitmap_DrawText "  to   ", LEFT_WALL_X * W + 2 * W, H * 7
          Bitmap_DrawText "Celltris", LEFT_WALL_X * W + 2 * W, H * 8
          
     Case GAME_STATE_PAUSED:
          Bitmap_DrawText " Paused ", LEFT_WALL_X * W + 2 * W, H * 6
          
     Case GAME_STATE_GAME_OVER:
          Bitmap_DrawText "  Game  ", LEFT_WALL_X * W + 2 * W, H * 6
          Bitmap_DrawText "  Over  ", LEFT_WALL_X * W + 2 * W, H * 7
          
  End Select
  
End Sub

Public Sub IncreaseSpeed()

     If LinesClearPerLevelCount >= LINES_CLEAR_PER_LEVEL Then
        CurrentSpeed = CurrentSpeed + 1
        
        If CurrentSpeed > MAX_FALL_SPEED Then
           CurrentSpeed = MAX_FALL_SPEED
        End If
        
        FallSpeedIncrement = CurrentSpeed
        LinesClearPerLevelCount = 0
        BackGroundPictureIndex = Rand(9, 14)
        AzPlaySound 1
     End If
End Sub

Sub CopyTetrad(DestTetrad As Tetrad, SrcTetrad As Tetrad)

 DestTetrad.id = SrcTetrad.id
 DestTetrad.x = SrcTetrad.x
 DestTetrad.y = SrcTetrad.y

 Dim x%, y%

 For y = 0 To MAX_SHAPE_HEIGHT - 1
   For x = 0 To MAX_SHAPE_WIDTH - 1
     DestTetrad.Cell(x, y) = SrcTetrad.Cell(x, y)
   Next x
Next y

End Sub

Function CreateGhostTetrad(TheTetrad As Tetrad) As Tetrad

   Dim GhostTetradObject As Tetrad
   
   CopyTetrad GhostTetradObject, TheTetrad
   InstantDrop GhostTetradObject, True
  
   Dim x%, y%

   For y = 0 To MAX_SHAPE_HEIGHT - 1
     For x = 0 To MAX_SHAPE_WIDTH - 1
       If TheTetrad.Cell(x, y) > 0 Then
        GhostTetradObject.Cell(x, y) = 15
       End If
     Next x
  Next y

  CreateGhostTetrad = GhostTetradObject
End Function


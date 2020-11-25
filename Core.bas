Attribute VB_Name = "CoreModule"
'This module contains this program's core procedures.
Option Explicit

'This enum lists the different shapes used.
Public Enum ShapesE
   I    '"I" shape.
   J    '"J" shape.
   L    '"L" shape.
   O    '"O" shape.
   S    '"S" shape.
   T    '"T" shape.
   Z    '"Z" shape.
End Enum

'This structure defines the game's state.
Public Type GameStateStr
   Dropper As Timer       'Defines the block dropper.
   GameOver As Boolean    'Indicates whether the game has been lost.
   Pit() As Long          'Defines the pit.
   PitBox As PictureBox   'Defines the pit's graphics.
   Score As Long          'Defines the number of rows cleared.
   Window As Form         'Defines the game's window.
End Type

'This structure defines a rectangle.
Public Type RectangleStr
   Height As Long   'Defines a rectangle's height.
   Width As Long    'Defines a rectangle's width.
   x As Long        'Defines the left side of a rectangle's position.
   y As Long        'Defines the top side of a rectangle's position.
End Type

'This structure defines a shape.
Public Type ShapeStr
   Angle As Long                 'Defines a shape's angle. (range 0-3 (x 90 = degrees.))
   Dimensions As RectangleStr    'Defines a shape's dimensions.
   DropRate As Long              'Defines the length of the interval between a shape's drops.
   Map() As Long                 'Defines a shape's map of blocks (colored and empty).
   PitX As Long                  'Defines a shape's horizontal position inside the pit.
   PitY As Long                  'Defines a shape's vertical position inside the pit.
   Shape As ShapesE              'Defines a shape.
End Type

Public Const BLOCK_SCALE As Long = 48   'Defines the scale at which the blocks are drawn.
Public Const PIT_HEIGHT As Long = 16    'Defines the pit's height.
Public Const PIT_WIDTH As Long = 10     'Defines the pit's width.
Private Const SHAPE_SIZE As Long = 4     'Defines the maximum number of blocks a shape can have per side.

Public ActiveShape As ShapeStr     'Contains the active shape.
Public GameState As GameStateStr   'Contains the game's state.

'This procedure returns an indicator of whether a shape with the specified map and position can move in the specified direction.
Public Function CanMove(Map() As Long, x As Long, y As Long, DirectionX As Long, DirectionY As Long) As Boolean
On Error GoTo ErrorTrap
Dim BlockX As Long
Dim BlockY As Long
Dim PitX As Long
Dim PitY As Long

   For BlockY = LBound(Map, 2) To UBound(Map, 2)
      For BlockX = LBound(Map, 1) To UBound(Map, 1)
         If Not Map(BlockX, BlockY) = 0 Then
            PitX = (x + BlockX) + DirectionX
            PitY = (y + BlockY) + DirectionY
            If PitX >= 0 And PitX < PIT_WIDTH And PitY >= 0 And PitY < PIT_HEIGHT Then
               If Not GameState.Pit(PitX, PitY) = 0 Then
                  CanMove = False
                  Exit Function
               End If
            ElseIf PitX < 0 Or PitX >= PIT_WIDTH Or PitY >= PIT_HEIGHT Then
               CanMove = False
               Exit Function
            End If
         End If
      Next BlockX
   Next BlockY
EndRoutine:
   CanMove = True
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure checks the pit for full rows and gives the command to remove any found.
Private Sub CheckForFullRows()
On Error GoTo ErrorTrap
Dim FullRow As Boolean
Dim PitX As Long
Dim PitY As Long

   With GameState
      For PitY = LBound(.Pit, 2) To UBound(.Pit, 2)
         FullRow = True
         For PitX = LBound(.Pit, 1) To UBound(.Pit, 1)
            If .Pit(PitX, PitY) = 0 Then
               FullRow = False
               Exit For
            End If
         Next PitX
         If FullRow Then RemoveRow PitY
      Next PitY
   End With
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure checks whether the game has been lost and gives the commands to draw the pit and to display the status information.
Private Sub CheckGameState()
On Error GoTo ErrorTrap
   GameState.Dropper.Enabled = Not GameState.GameOver
   GameState.GameOver = (ActiveShape.PitY < 0)

   DrawPit GameState.PitBox
   DisplayStatus GameState.Window
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure creates and returns an empty pit.
Public Function CreatePit() As Long()
On Error GoTo ErrorTrap
Dim Pit(0 To PIT_WIDTH - 1, 0 To PIT_HEIGHT - 1) As Long
Dim PitX As Long
Dim PitY As Long

   For PitY = LBound(Pit, 2) To UBound(Pit, 2)
      For PitX = LBound(Pit, 1) To UBound(Pit, 1)
         Pit(PitX, PitY) = 0
      Next PitX
   Next PitY

EndRoutine:
   CreatePit = Pit
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure creates and returns a shape and enables the dropper.
Private Sub CreateShape()
On Error GoTo ErrorTrap

   With ActiveShape
      .Angle = CLng(Rnd() * 3)
      .DropRate = 1000
      .Shape = CLng(Rnd() * ShapesE.Z) + ShapesE.I
      .Map = GetRotatedShapeMap(.Shape, .Angle)
      .Dimensions = GetShapeDimensions(.Map)
      .PitX = CLng(Rnd() * ((PIT_WIDTH - 1) - .Dimensions.Width)) - .Dimensions.x
      .PitY = -(.Dimensions.y + .Dimensions.Height)
   End With

   GameState.Dropper.Interval = ActiveShape.DropRate
EndRoutine:
    Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure displays the game's status on the specified graphical surface.
Public Sub DisplayStatus(Window As Form)
On Error GoTo ErrorTrap
Dim Text As String

   With GameState
      If .GameOver Then
         Text = "Game over - press Escape."
      Else
         Text = "Score: " & .Score
      End If
   End With

   With Window
      .Cls
      .CurrentX = 0
      .CurrentY = 0
   End With

   Window.Print Text
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure draws a block of the specified color at the specified position on the specified graphical surface.
Private Sub DrawBlock(ColorO As Long, PitX As Long, PitY As Long, PitBox As PictureBox)
On Error GoTo ErrorTrap
Dim DrawX As Long
Dim DrawY As Long

   DrawX = PitX * BLOCK_SCALE
   DrawY = PitY * BLOCK_SCALE

   PitBox.Line (DrawX, DrawY)-Step(BLOCK_SCALE, BLOCK_SCALE), QBColor(ColorO), BF
   PitBox.Line (DrawX + CLng(BLOCK_SCALE / 10), DrawY + CLng(BLOCK_SCALE / 10))-Step(BLOCK_SCALE - CLng(BLOCK_SCALE / 5), BLOCK_SCALE - CLng(BLOCK_SCALE / 5)), 0, B
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure draws the pit on the specified graphical surface.
Public Sub DrawPit(PitBox As PictureBox)
On Error GoTo ErrorTrap
Dim PitX As Long
Dim PitY As Long

   With GameState
      For PitY = LBound(.Pit, 2) To UBound(.Pit, 2)
         For PitX = LBound(.Pit, 1) To UBound(.Pit, 1)
            If .GameOver Then
               DrawBlock 4, PitX, PitY, PitBox
            Else
               DrawBlock .Pit(PitX, PitY), PitX, PitY, PitBox
            End If
         Next PitX
      Next PitY
   End With
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure draws/erases the specified shape at the specified position on the specified graphical surface.
Public Sub DrawShape(Optional EraseShape As Boolean = False)
On Error GoTo ErrorTrap
Dim BlockX As Long
Dim BlockY As Long
Dim PitX As Long
Dim PitY As Long

   With ActiveShape
      For BlockX = LBound(.Map, 1) To UBound(.Map, 1)
         For BlockY = LBound(.Map, 2) To UBound(.Map, 2)
            PitX = .PitX + BlockX
            PitY = .PitY + BlockY
            If PitX >= 0 And PitX < PIT_WIDTH And PitY >= 0 And PitY < PIT_HEIGHT Then
               If EraseShape Then
                  DrawBlock GameState.Pit(PitX, PitY), PitX, PitY, GameState.PitBox
               ElseIf Not .Map(BlockX, BlockY) = 0 Then
                  DrawBlock .Map(BlockX, BlockY), PitX, PitY, GameState.PitBox
               End If
            End If
         Next BlockY
      Next BlockX
   End With
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure drops the active shape until it cannot continue dropping.
Public Sub DropShape()
On Error GoTo ErrorTrap

   With ActiveShape
      If CanMove(.Map, .PitX, .PitY, DirectionX:=0, DirectionY:=1) Then
         DrawShape EraseShape:=True
         .PitY = .PitY + 1
         DrawShape
      Else
         SettleActiveShape
         CheckForFullRows
         CheckGameState
   
         If Not GameState.GameOver Then
            CreateShape
            DrawShape
         End If
      End If
   End With
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure returns the specified shape map rotated at the specified angle.
Public Function GetRotatedShapeMap(Shape As ShapesE, Angle As Long) As Long()
On Error GoTo ErrorTrap
Dim BlockX As Long
Dim BlockY As Long
Dim Map() As Long
Dim NewBlockX As Long
Dim NewBlockY As Long
Dim RotatedMap() As Long

   Map = GetShapeMap(Shape)
   ReDim RotatedMap(LBound(Map, 1) To UBound(Map, 1), LBound(Map, 2) To UBound(Map, 2))

   If Angle = 0 Then
      GetRotatedShapeMap = Map
      Exit Function
   Else
      For BlockX = LBound(Map, 1) To UBound(Map, 1)
         For BlockY = LBound(Map, 2) To UBound(Map, 2)
            Select Case Angle
               Case 1
                  NewBlockX = (UBound(Map, 2) - LBound(Map, 2)) - BlockY
                  NewBlockY = BlockX
               Case 2
                  NewBlockX = (UBound(Map, 1) - LBound(Map, 1)) - BlockX
                  NewBlockY = (UBound(Map, 2) - LBound(Map, 2)) - BlockY
               Case 3
                  NewBlockX = BlockY
                  NewBlockY = (UBound(Map, 1) - LBound(Map, 1)) - BlockX
            End Select

            RotatedMap(NewBlockX, NewBlockY) = Map(BlockX, BlockY)
         Next BlockY
      Next BlockX
   End If

EndRoutine:
   GetRotatedShapeMap = RotatedMap
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure checks the specified shape's colored block positions and returns the rectangular area containing colored blocks.
Public Function GetShapeDimensions(Map() As Long) As RectangleStr
On Error GoTo ErrorTrap
Dim BlockX As Long
Dim BlockY As Long
Dim Dimensions As RectangleStr
Dim LowerRightX As Long
Dim LowerRightY As Long
Dim UpperLeftX As Long
Dim UpperLeftY As Long

   LowerRightX = -((UBound(Map, 1) - LBound(Map, 1)) + 1)
   LowerRightY = -((UBound(Map, 2) - LBound(Map, 2)) + 1)
   UpperLeftX = (UBound(Map, 1) - LBound(Map, 1)) + 1
   UpperLeftY = (UBound(Map, 2) - LBound(Map, 2)) + 1

   For BlockY = LBound(Map, 2) To UBound(Map, 2)
      For BlockX = LBound(Map, 1) To UBound(Map, 1)
         If Not Map(BlockX, BlockY) = 0 Then
            If BlockX <= UpperLeftX Then UpperLeftX = BlockX
            If BlockY <= UpperLeftY Then UpperLeftY = BlockY
            If BlockX >= LowerRightX Then LowerRightX = BlockX
            If BlockY >= LowerRightY Then LowerRightY = BlockY
         End If
      Next BlockX
   Next BlockY

   Dimensions.x = UpperLeftX
   Dimensions.y = UpperLeftY
   Dimensions.Width = LowerRightX - UpperLeftX
   Dimensions.Height = LowerRightY - UpperLeftY
EndRoutine:
   GetShapeDimensions = Dimensions
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure returns the specified shape's map of colored and empty blocks.
Private Function GetShapeMap(Shape As ShapesE) As Long()
On Error GoTo ErrorTrap
Dim BlockX As Long
Dim BlockY As Long
Dim Map() As Variant
Dim ShapeMap() As Long

   Select Case Shape
      Case ShapesE.I
         Map = Array(Array(0, 0, 0, 0), Array(11, 11, 11, 11), Array(0, 0, 0, 0), Array(0, 0, 0, 0))
      Case ShapesE.J
         Map = Array(Array(0, 0, 0, 0), Array(9, 9, 9, 0), Array(0, 0, 9, 0), Array(0, 0, 0, 0))
      Case ShapesE.L
         Map = Array(Array(0, 0, 0, 0), Array(6, 6, 6, 0), Array(6, 0, 0, 0), Array(0, 0, 0, 0))
      Case ShapesE.O
         Map = Array(Array(0, 0, 0, 0), Array(0, 14, 14, 0), Array(0, 14, 14, 0), Array(0, 0, 0, 0))
      Case ShapesE.S
         Map = Array(Array(0, 0, 0, 0), Array(0, 2, 2, 0), Array(2, 2, 0, 0), Array(0, 0, 0, 0))
      Case ShapesE.T
         Map = Array(Array(0, 0, 0, 0), Array(5, 5, 5, 0), Array(0, 5, 0, 0), Array(0, 0, 0, 0))
      Case ShapesE.Z
         Map = Array(Array(0, 0, 0, 0), Array(12, 12, 0, 0), Array(0, 12, 12, 0), Array(0, 0, 0, 0))
   End Select

   ReDim ShapeMap(0 To SHAPE_SIZE - 1, 0 To SHAPE_SIZE - 1)

   For BlockX = LBound(ShapeMap, 1) To UBound(ShapeMap, 1)
      For BlockY = LBound(ShapeMap, 2) To UBound(ShapeMap, 2)
         ShapeMap(BlockX, BlockY) = Map(BlockX)(BlockY)
      Next BlockY
   Next BlockX

EndRoutine:
   GetShapeMap = ShapeMap
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure handles any errors that occur.
Public Sub HandleError()
Dim Description As String

   Description = Err.Description
   On Error GoTo ErrorTrap
   If MsgBox(Description, vbOKCancel Or vbExclamation, App.Title) = vbCancel Then End
EndRoutine:
   Exit Sub

ErrorTrap:
   End
   Resume EndRoutine
End Sub

'This procedure initializes the game.
Public Sub InitializeGame(Window As Form, PitBox As PictureBox, Dropper As Timer)
On Error GoTo ErrorTrap

   Randomize

   With GameState
      .GameOver = False
      .Pit = CreatePit()
      .Score = 0
      Set .Dropper = Dropper
      Set .PitBox = PitBox
      Set .Window = Window
      DrawPit .PitBox
      CreateShape
      .Dropper.Enabled = (Not .GameOver)
   End With

EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure removes the specified row.
Private Sub RemoveRow(PitY As Long)
On Error GoTo ErrorTrap
Dim PitX As Long
Dim RemovedRow As Long

   With GameState
      For RemovedRow = PitY To LBound(.Pit, 2) Step -1
         For PitX = LBound(.Pit, 1) To UBound(.Pit, 1)
            If RemovedRow = 0 Then
               .Pit(PitX, RemovedRow) = 0
            Else
               .Pit(PitX, RemovedRow) = .Pit(PitX, RemovedRow - 1)
            End If
         Next PitX
      Next RemovedRow

      .Score = .Score + 1
   End With
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure settles the active shape in the pit.
Private Sub SettleActiveShape()
On Error GoTo ErrorTrap
Dim BlockX As Long
Dim BlockY As Long
Dim PitX As Long
Dim PitY As Long

   With ActiveShape
      For BlockY = LBound(.Map, 2) To UBound(.Map, 2)
         For BlockX = LBound(.Map, 1) To UBound(.Map, 1)
            PitX = .PitX + BlockX
            PitY = .PitY + BlockY
            If PitX >= 0 And PitX < PIT_WIDTH And PitY >= 0 And PitY < PIT_HEIGHT Then
               If Not .Map(BlockX, BlockY) = 0 Then GameState.Pit(PitX, PitY) = .Map(BlockX, BlockY)
            End If
         Next BlockX
      Next BlockY
   End With
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


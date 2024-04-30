'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Environment
Imports System.Linq
Imports System.Windows.Forms

'This module contains this program's core procedures.
Public Module CoreModule
   'This enum lists the different shapes used.
   Public Enum ShapesE As Integer
      I    '"I" shape.
      J    '"J" shape.
      L    '"L" shape.
      O    '"O" shape.
      S    '"S" shape.
      T    '"T" shape.
      Z    '"Z" shape.
   End Enum

   'This structure defines the game's state.
   Public Structure GameStateStr
      Public GameOver As Boolean       'Indicates whether the game has been lost.
      Public Pit(,) As Color           'Defines the pit.
      Public PitCanvas As Graphics     'Defines the pit's graphics.
      Public Score As ULong            'Defines the number of rows cleared.
      Public StateCanvas As Graphics   'Defines the game state's graphics.
   End Structure

   'This structure defines a shape.
   Public Structure ShapeStr
      Public Angle As Integer           'Defines a shape's angle. (range 0-3 (x 90 = degrees.))
      Public Dimensions As Rectangle    'Defines a shape's dimensions.
      Public DropRate As Integer        'Defines the length of the interval between a shape's drops.
      Public Map(,) As Color            'Defines a shape's map of blocks (colored and empty).
      Public PitXY As Point             'Defines a shape's position inside the pit.
      Public Shape As ShapesE           'Defines a shape.
   End Structure

   Public Const BLOCK_SCALE As Integer = 48   'Defines the scale at which the blocks are drawn.
   Public Const PIT_HEIGHT As Integer = 16    'Defines the pit's height.
   Public Const PIT_WIDTH As Integer = 10     'Defines the pit's width.

   Public ActiveShape As ShapeStr = Nothing     'Contains the active shape.
   Public GameState As GameStateStr = Nothing   'Contains the game's state.

   Public WithEvents Dropper As Timer = Nothing  'This drops the active shape at a specific interval.

   'This procedure returns an indicator of whether a shape with the specified map and position can move in the specified direction.
   Public Function CanMove(Map(,) As Color, xy As Point, Direction As Point) As Boolean
      Try
         Dim PitX As New Integer
         Dim PitY As New Integer

         With Map
            For BlockY As Integer = .GetLowerBound(1) To .GetUpperBound(1)
               For BlockX As Integer = .GetLowerBound(0) To .GetUpperBound(0)
                  If Not Map(BlockX, BlockY) = Nothing Then
                     PitX = (xy.X + BlockX) + Direction.X
                     PitY = (xy.Y + BlockY) + Direction.Y
                     If PitX >= 0 AndAlso PitX < PIT_WIDTH AndAlso PitY >= 0 AndAlso PitY < PIT_HEIGHT Then
                        If Not GameState.Pit(PitX, PitY) = Nothing Then Return False
                     ElseIf PitX < 0 OrElse PitX >= PIT_WIDTH OrElse PitY >= PIT_HEIGHT Then
                        Return False
                     End If
                  End If
               Next BlockX
            Next BlockY
         End With

         Return True
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return False
   End Function

   'This procedure checks the pit for full rows and gives the command to remove any found.
   Private Sub CheckForFullRows()
      Try
         Dim FullRow As New Boolean

         With GameState
            For PitY As Integer = .Pit.GetLowerBound(1) To .Pit.GetUpperBound(1)
               FullRow = True
               For PitX As Integer = .Pit.GetLowerBound(0) To .Pit.GetUpperBound(0)
                  If .Pit(PitX, PitY) = Nothing Then
                     FullRow = False
                     Exit For
                  End If
               Next PitX
               If FullRow Then RemoveRow(PitY)
            Next PitY
         End With
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure checks whether the game has been lost and gives the commands to draw the pit and to display the status information.
   Private Sub CheckGameState()
      Try
         GameState.GameOver = (ActiveShape.PitXY.Y < 0)
         Dropper.Enabled = Not GameState.GameOver

         DrawPit(GameState.PitCanvas)
         DisplayStatus(GameState.StateCanvas)
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure creates and returns an empty pit.
   Private Function CreatePit() As Color(,)
      Try
         Dim Pit(0 To PIT_WIDTH - 1, 0 To PIT_HEIGHT - 1) As Color

         With Pit
            For PitY As Integer = .GetLowerBound(1) To .GetUpperBound(1)
               For PitX As Integer = .GetLowerBound(0) To .GetUpperBound(0)
                  Pit(PitX, PitY) = Nothing
               Next PitX
            Next PitY
         End With

         Return Pit
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return {{}}
   End Function

   'This procedure creates and returns a shape and enables the dropper.
   Private Sub CreateShape()
      Try
         Static RandomO As New Random

         With ActiveShape
            .Angle = RandomO.Next(0, 3)
            .DropRate = 1000
            .Shape = DirectCast(RandomO.Next(ShapesE.I, ShapesE.Z), ShapesE)
            .Map = GetRotatedShapeMap(.Shape, .Angle)
            .Dimensions = GetShapeDimensions(.Map)
            .PitXY.X = RandomO.Next(- .Dimensions.X, (PIT_WIDTH - 1) - .Dimensions.Width)
            .PitXY.Y = -(.Dimensions.Y + .Dimensions.Height)
         End With

         Dropper = New Timer With {.Enabled = True, .Interval = ActiveShape.DropRate}
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure displays the game's status on the specified graphical surface.
   Public Sub DisplayStatus(Canvas As Graphics)
      Try
         Dim Text As String = Nothing

         With GameState
            Text = If(.GameOver, "Game over - press Escape.", $"Score: { .Score}")

            Canvas.Clear(Color.Black)
            Canvas.DrawString(Text, New Font("Comic Sans MS", 16, FontStyle.Bold), Brushes.Red, New Point(0, 0))
         End With
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure draws a block of the specified color at the specified position on the specified graphical surface.
   Private Sub DrawBlock(ColorO As Color, PitXY As Point, Canvas As Graphics)
      Try
         Dim DrawX As Integer = PitXY.X * BLOCK_SCALE
         Dim DrawY As Integer = PitXY.Y * BLOCK_SCALE

         Canvas.FillRectangle(New SolidBrush(ColorO), DrawX, DrawY, BLOCK_SCALE, BLOCK_SCALE)
         Canvas.DrawRectangle(Pens.Black, DrawX + CInt(BLOCK_SCALE / 10), DrawY + CInt(BLOCK_SCALE / 10), BLOCK_SCALE - CInt(BLOCK_SCALE / 5), BLOCK_SCALE - CInt(BLOCK_SCALE / 5))
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure draws the pit on the specified graphical surface.
   Public Sub DrawPit(Canvas As Graphics)
      Try
         With GameState
            For PitY As Integer = .Pit.GetLowerBound(1) To .Pit.GetUpperBound(1)
               For Pitx As Integer = .Pit.GetLowerBound(0) To .Pit.GetUpperBound(0)
                  DrawBlock(If(.Pit(Pitx, PitY) = Nothing, Color.Black, If(.GameOver, Color.Red, .Pit(Pitx, PitY))), New Point(Pitx, PitY), Canvas)
               Next Pitx
            Next PitY
         End With
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure draws/erases the active shape.
   Public Sub DrawShape(Optional EraseShape As Boolean = False)
      Try
         Dim PitX As New Integer
         Dim PitY As New Integer

         With ActiveShape
            For BlockX As Integer = .Map.GetLowerBound(0) To .Map.GetUpperBound(0)
               For BlockY As Integer = .Map.GetLowerBound(1) To .Map.GetUpperBound(1)
                  PitX = .PitXY.X + BlockX
                  PitY = .PitXY.Y + BlockY
                  If PitX >= 0 AndAlso PitX < PIT_WIDTH AndAlso PitY >= 0 AndAlso PitY < PIT_HEIGHT Then
                     DrawBlock(If(EraseShape, If(GameState.Pit(PitX, PitY) = Nothing, Color.Black, GameState.Pit(PitX, PitY)), .Map(BlockX, BlockY)), New Point(PitX, PitY), GameState.PitCanvas)
                  End If
               Next BlockY
            Next BlockX
         End With
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure drops the active shape until it cannot continue dropping.
   Private Sub Dropper_Tick(sender As Object, e As EventArgs) Handles Dropper.Tick
      Try
         With ActiveShape
            If CanMove(.Map, .PitXY, New Point(0, 1)) Then
               DrawShape(EraseShape:=True)
               .PitXY.Y += 1
               DrawShape()
            Else
               SettleActiveShape()
               CheckForFullRows()
               CheckGameState()

               If Not GameState.GameOver Then
                  CreateShape()
                  DrawShape()
               End If
            End If
         End With
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure returns the specified shape map rotated at the specified angle.
   Public Function GetRotatedShapeMap(Shape As ShapesE, Angle As Integer) As Color(,)
      Try
         Dim Map(,) As Color = GetShapeMap(Shape)
         Dim NewBlockX As New Integer
         Dim NewBlockY As New Integer
         Dim RotatedMap(0 To Map.GetUpperBound(0), 0 To Map.GetUpperBound(1)) As Color

         If Angle = 0 Then
            Return Map
         Else
            With Map
               For BlockX As Integer = .GetLowerBound(0) To .GetUpperBound(0)
                  For BlockY As Integer = .GetLowerBound(1) To .GetUpperBound(1)
                     Select Case Angle
                        Case 1
                           NewBlockX = (.GetUpperBound(1) - .GetLowerBound(1)) - BlockY
                           NewBlockY = BlockX
                        Case 2
                           NewBlockX = (.GetUpperBound(0) - .GetLowerBound(0)) - BlockX
                           NewBlockY = (.GetUpperBound(1) - .GetLowerBound(1)) - BlockY
                        Case 3
                           NewBlockX = BlockY
                           NewBlockY = (.GetUpperBound(0) - .GetLowerBound(0)) - BlockX
                     End Select

                     RotatedMap(NewBlockX, NewBlockY) = Map(BlockX, BlockY)
                  Next BlockY
               Next BlockX
            End With
         End If

         Return RotatedMap
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return {{}}
   End Function

   'This procedure checks the specified shape's colored block positions and returns the rectangular area containing colored blocks.
   Public Function GetShapeDimensions(Map(,) As Color) As Rectangle
      Try
         Dim LowerRight As New Point(Int32.MinValue, Int32.MinValue)
         Dim UpperLeft As New Point(Int32.MaxValue, Int32.MaxValue)

         With Map
            For BlockY As Integer = .GetLowerBound(1) To .GetUpperBound(1)
               For BlockX As Integer = .GetLowerBound(0) To .GetUpperBound(0)
                  If Not Map(BlockX, BlockY) = Nothing Then
                     If BlockX <= UpperLeft.X Then UpperLeft.X = BlockX
                     If BlockY <= UpperLeft.Y Then UpperLeft.Y = BlockY
                     If BlockX >= LowerRight.X Then LowerRight.X = BlockX
                     If BlockY >= LowerRight.Y Then LowerRight.Y = BlockY
                  End If
               Next BlockX
            Next BlockY
         End With

         Return New Rectangle(UpperLeft.X, UpperLeft.Y, LowerRight.X - UpperLeft.X, LowerRight.Y - UpperLeft.Y)
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns the specified shape's map of colored and empty blocks.
   Private Function GetShapeMap(Shape As ShapesE) As Color(,)
      Try
         Static Maps As New List(Of Color(,))

         With Maps
            If Not .Any Then
               .Add({{Nothing, Nothing, Nothing, Nothing}, {Color.Cyan, Color.Cyan, Color.Cyan, Color.Cyan}, {Nothing, Nothing, Nothing, Nothing}, {Nothing, Nothing, Nothing, Nothing}})
               .Add({{Nothing, Nothing, Nothing, Nothing}, {Color.Blue, Color.Blue, Color.Blue, Nothing}, {Nothing, Nothing, Color.Blue, Nothing}, {Nothing, Nothing, Nothing, Nothing}})
               .Add({{Nothing, Nothing, Nothing, Nothing}, {Color.Orange, Color.Orange, Color.Orange, Nothing}, {Color.Orange, Nothing, Nothing, Nothing}, {Nothing, Nothing, Nothing, Nothing}})
               .Add({{Nothing, Nothing, Nothing, Nothing}, {Nothing, Color.Yellow, Color.Yellow, Nothing}, {Nothing, Color.Yellow, Color.Yellow, Nothing}, {Nothing, Nothing, Nothing, Nothing}})
               .Add({{Nothing, Nothing, Nothing, Nothing}, {Nothing, Color.Green, Color.Green, Nothing}, {Color.Green, Color.Green, Nothing, Nothing}, {Nothing, Nothing, Nothing, Nothing}})
               .Add({{Nothing, Nothing, Nothing, Nothing}, {Color.Purple, Color.Purple, Color.Purple, Nothing}, {Nothing, Color.Purple, Nothing, Nothing}, {Nothing, Nothing, Nothing, Nothing}})
               .Add({{Nothing, Nothing, Nothing, Nothing}, {Color.Red, Color.Red, Nothing, Nothing}, {Nothing, Color.Red, Color.Red, Nothing}, {Nothing, Nothing, Nothing, Nothing}})
            End If
         End With

         Return Maps(Shape)
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return {{}}
   End Function

   'This procedure handles any errors that occur.
   Public Sub HandleError(ExceptionO As Exception)
      Try
         MessageBox.Show(ExceptionO.Message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
      Catch
         [Exit](0)
      End Try
   End Sub

   'This procedure initializes the game.
   Public Sub InitializeGame(Window As Form, PitBox As PictureBox)
      Try
         CreateShape()
         GameState = New GameStateStr With {.GameOver = False, .Pit = CreatePit(), .PitCanvas = PitBox.CreateGraphics, .Score = 0, .StateCanvas = Window.CreateGraphics}
         Window.Invalidate()
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure removes the specified row.
   Private Sub RemoveRow(PitY As Integer)
      Try
         With GameState
            For RemovedRow As Integer = PitY To .Pit.GetLowerBound(1) Step -1
               For PitX As Integer = .Pit.GetLowerBound(0) To .Pit.GetUpperBound(0)
                  .Pit(PitX, RemovedRow) = If(RemovedRow = 0, Nothing, .Pit(PitX, RemovedRow - 1))
               Next PitX
            Next RemovedRow

            .Score += CULng(1)
         End With
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure settles the active shape in the pit.
   Private Sub SettleActiveShape()
      Try
         Dim PitX As New Integer
         Dim PitY As New Integer

         With ActiveShape
            For BlockY As Integer = .Map.GetLowerBound(1) To .Map.GetUpperBound(1)
               For BlockX As Integer = .Map.GetLowerBound(0) To .Map.GetUpperBound(0)
                  PitX = .PitXY.X + BlockX
                  PitY = .PitXY.Y + BlockY
                  If PitX >= 0 AndAlso PitX < PIT_WIDTH AndAlso PitY >= 0 AndAlso PitY < PIT_HEIGHT Then
                     If Not .Map(BlockX, BlockY) = Nothing Then GameState.Pit(PitX, PitY) = .Map(BlockX, BlockY)
                  End If
               Next BlockX
            Next BlockY
         End With
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub
End Module

'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Drawing
Imports System.Windows.Forms

'This class contains this program's main interface window.
Public Class InterfaceWindow
   'This procedure initializes this window and gives the command to initialize the game.
   Public Sub New()
      Try
         InitializeComponent()

         PitBox.Location = New Point(0, 48)
         PitBox.Size = New Size(PIT_WIDTH * BLOCK_SCALE, PIT_HEIGHT * BLOCK_SCALE)

         With My.Application.Info
            Me.Text = $"{ .Title} v{ .Version} - by: { .CompanyName}"
         End With

         Me.Width += (PitBox.Width - Me.ClientSize.Width)
         Me.Height += ((PitBox.Height + PitBox.Top) - Me.ClientSize.Height)

         InitializeGame(Me, PitBox)
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure handles the user's keystrokes.
   Private Sub Form_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
      Try
         Dim NewAngle As New Integer
         Dim RotatedMap(,) As Color = {{}}

         If GameState.GameOver Then
            If e.KeyCode = Keys.Escape Then InitializeGame(Me, PitBox)
         Else
            With ActiveShape
               Select Case e.KeyCode
                  Case Keys.A
                     DrawShape(EraseShape:=True)
                     If .Angle = 3 Then NewAngle = 0 Else NewAngle = .Angle + 1
                     RotatedMap = GetRotatedShapeMap(.Shape, NewAngle)
                     If CanMove(RotatedMap, .PitXY, New Point(0, 0)) Then
                        .Angle = NewAngle
                        .Map = RotatedMap
                        .Dimensions = GetShapeDimensions(.Map)
                     End If
                     DrawShape()
                  Case Keys.Left
                     DrawShape(EraseShape:=True)
                     If CanMove(.Map, .PitXY, New Point(-1, 0)) Then .PitXY.X -= 1
                     DrawShape()
                  Case Keys.Right
                     DrawShape(EraseShape:=True)
                     If CanMove(.Map, .PitXY, New Point(1, 0)) Then .PitXY.X += 1
                     DrawShape()
                  Case Keys.Space
                     Dropper.Interval = 1
               End Select
            End With
         End If
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure gives the command to redraw the game's graphics when this window is redrawn.
   Private Sub Form_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
      Try
         DrawPit(GameState.PitCanvas)
         DisplayStatus(GameState.StateCanvas)
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub
End Class

VERSION 5.00
Begin VB.Form InterfaceWindow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Dropper 
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox PitBox 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   1800
      ScaleHeight     =   121
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   0
      Top             =   1320
      Width           =   2895
   End
End
Attribute VB_Name = "InterfaceWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This class contains this program's main interface window.
Option Explicit

'This procedure gives the command to drop the active shape.
Private Sub Dropper_Timer()
On Error GoTo ErrorTrap
   DropShape
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to display the status when this window is first shown.
Private Sub Form_Activate()
On Error GoTo ErrorTrap
   DisplayStatus Me
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure initializes this window and gives the command to initialize the game.
Private Sub Form_Initialize()
On Error GoTo ErrorTrap
   PitBox.Left = 0
   PitBox.Top = 48
   PitBox.Width = PIT_WIDTH * BLOCK_SCALE
   PitBox.Height = PIT_HEIGHT * BLOCK_SCALE
   
   With App
      Me.Caption = .Title & " v" & .Major & "." & .Minor & .Revision & " - by: " & .CompanyName
   End With
   
   Me.Width = Me.Width + ((PitBox.ScaleWidth - Me.ScaleWidth) * Screen.TwipsPerPixelX)
   Me.Height = Me.Height + (((PitBox.ScaleHeight + PitBox.Top) - Me.ScaleHeight) * Screen.TwipsPerPixelY)
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure handles the user's keystrokes.
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorTrap
Dim NewAngle As Long
Dim RotatedMap() As Long

   If GameState.GameOver Then
      If KeyCode = vbKeyEscape Then
         InitializeGame Me, PitBox, Dropper
         DisplayStatus Me
      End If
   Else
      With ActiveShape
         Select Case KeyCode
            Case vbKeyA
               DrawShape EraseShape:=True
               If .Angle = 3 Then NewAngle = 0 Else NewAngle = .Angle + 1
               RotatedMap = GetRotatedShapeMap(.Shape, NewAngle)
               If CanMove(RotatedMap, .PitX, .PitY, DirectionX:=0, DirectionY:=0) Then
                  .Angle = NewAngle
                  .Map = RotatedMap
                  .Dimensions = GetShapeDimensions(.Map)
               End If
               DrawShape
            Case vbKeyLeft
               DrawShape EraseShape:=True
               If CanMove(.Map, .PitX, .PitY, DirectionX:=-1, DirectionY:=0) Then .PitX = .PitX - 1
               DrawShape
            Case vbKeyRight
               DrawShape EraseShape:=True
               If CanMove(.Map, .PitX, .PitY, DirectionX:=1, DirectionY:=0) Then .PitX = .PitX + 1
               DrawShape
            Case vbKeySpace
               Dropper.Enabled = False
               Dropper.Interval = 1
               Dropper.Enabled = True
         End Select
      End With
   End If
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure gives the command to initialize the game.
Private Sub Form_Load()
On Error GoTo ErrorTrap
   InitializeGame Me, PitBox, Dropper
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub



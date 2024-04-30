<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class InterfaceWindow
   Inherits System.Windows.Forms.Form

   'Form overrides dispose to clean up the component list.
   <System.Diagnostics.DebuggerNonUserCode()> _
   Protected Overrides Sub Dispose(ByVal disposing As Boolean)
      Try
         If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
         End If
      Finally
         MyBase.Dispose(disposing)
      End Try
   End Sub

   'Required by the Windows Form Designer
   Private components As System.ComponentModel.IContainer

   'NOTE: The following procedure is required by the Windows Form Designer
   'It can be modified using the Windows Form Designer.  
   'Do not modify it using the code editor.
   <System.Diagnostics.DebuggerStepThrough()> _
   Private Sub InitializeComponent()
      Me.PitBox = New System.Windows.Forms.PictureBox()
      CType(Me.PitBox, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.SuspendLayout()
      '
      'PitBox
      '
      Me.PitBox.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
      Me.PitBox.Location = New System.Drawing.Point(2, 37)
      Me.PitBox.Name = "PitBox"
      Me.PitBox.Size = New System.Drawing.Size(277, 216)
      Me.PitBox.TabIndex = 0
      Me.PitBox.TabStop = False
      '
      'InterfaceWindow
      '
      Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
      Me.BackColor = System.Drawing.Color.Black
      Me.ClientSize = New System.Drawing.Size(282, 255)
      Me.Controls.Add(Me.PitBox)
      Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
      Me.KeyPreview = True
      Me.MaximizeBox = False
      Me.Name = "InterfaceWindow"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      CType(Me.PitBox, System.ComponentModel.ISupportInitialize).EndInit()
      Me.ResumeLayout(False)

   End Sub

   Friend WithEvents PitBox As System.Windows.Forms.PictureBox
End Class

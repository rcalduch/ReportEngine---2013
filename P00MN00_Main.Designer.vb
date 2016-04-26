<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class P00MN00_Main
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(P00MN00_Main))
    Me.cmdSortir = New System.Windows.Forms.Button()
    Me.lblTitle = New System.Windows.Forms.Label()
    Me.lblDebug = New System.Windows.Forms.Label()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.cmdTest = New System.Windows.Forms.Button()
    Me.lblPTM = New System.Windows.Forms.Label()
    Me.SuspendLayout()
    '
    'cmdSortir
    '
    Me.cmdSortir.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdSortir.Location = New System.Drawing.Point(342, 55)
    Me.cmdSortir.Name = "cmdSortir"
    Me.cmdSortir.Size = New System.Drawing.Size(75, 23)
    Me.cmdSortir.TabIndex = 1
    Me.cmdSortir.Text = "Sortir"
    Me.cmdSortir.UseVisualStyleBackColor = True
    '
    'lblTitle
    '
    Me.lblTitle.BackColor = System.Drawing.SystemColors.ButtonHighlight
    Me.lblTitle.Dock = System.Windows.Forms.DockStyle.Top
    Me.lblTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblTitle.Location = New System.Drawing.Point(0, 0)
    Me.lblTitle.Name = "lblTitle"
    Me.lblTitle.Size = New System.Drawing.Size(428, 40)
    Me.lblTitle.TabIndex = 2
    Me.lblTitle.Text = "Gestor de llistats CUSTOM"
    Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'lblDebug
    '
    Me.lblDebug.AutoSize = True
    Me.lblDebug.Location = New System.Drawing.Point(12, 70)
    Me.lblDebug.Name = "lblDebug"
    Me.lblDebug.Size = New System.Drawing.Size(59, 13)
    Me.lblDebug.TabIndex = 3
    Me.lblDebug.Text = "Debugging"
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.BackColor = System.Drawing.SystemColors.ButtonHighlight
    Me.Label1.Location = New System.Drawing.Point(396, 20)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(28, 13)
    Me.Label1.TabIndex = 4
    Me.Label1.Text = "1.03"
    '
    'cmdTest
    '
    Me.cmdTest.Location = New System.Drawing.Point(261, 55)
    Me.cmdTest.Name = "cmdTest"
    Me.cmdTest.Size = New System.Drawing.Size(75, 23)
    Me.cmdTest.TabIndex = 5
    Me.cmdTest.Text = "Test"
    Me.cmdTest.UseVisualStyleBackColor = True
    Me.cmdTest.Visible = False
    '
    'lblPTM
    '
    Me.lblPTM.AutoSize = True
    Me.lblPTM.Location = New System.Drawing.Point(12, 50)
    Me.lblPTM.Name = "lblPTM"
    Me.lblPTM.Size = New System.Drawing.Size(70, 13)
    Me.lblPTM.TabIndex = 6
    Me.lblPTM.Text = "Monitoritzant:"
    '
    'P00MN00_Main
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(428, 92)
    Me.Controls.Add(Me.lblPTM)
    Me.Controls.Add(Me.cmdTest)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.lblDebug)
    Me.Controls.Add(Me.lblTitle)
    Me.Controls.Add(Me.cmdSortir)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "P00MN00_Main"
    Me.Text = "DOS Report Engine"
    Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents cmdSortir As System.Windows.Forms.Button
  Friend WithEvents lblTitle As System.Windows.Forms.Label
  Friend WithEvents lblDebug As System.Windows.Forms.Label
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents cmdTest As System.Windows.Forms.Button
  Friend WithEvents lblPTM As System.Windows.Forms.Label

End Class

Public Class frmProgress
  Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

  Public Sub New()
    MyBase.New()

    'This call is required by the Windows Form Designer.
    InitializeComponent()

    'Add any initialization after the InitializeComponent() call

  End Sub

  'Form overrides dispose to clean up the component list.
  Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
    If disposing Then
      If Not (components Is Nothing) Then
        components.Dispose()
      End If
    End If
    MyBase.Dispose(disposing)
  End Sub

  'Required by the Windows Form Designer
  Private components As System.ComponentModel.IContainer

  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.  
  'Do not modify it using the code editor.
  Friend WithEvents lblMainTask As System.Windows.Forms.Label
  Friend WithEvents lblItemTask As System.Windows.Forms.Label
  Friend WithEvents ProgressBar As System.Windows.Forms.ProgressBar
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Me.lblMainTask = New System.Windows.Forms.Label
    Me.lblItemTask = New System.Windows.Forms.Label
    Me.ProgressBar = New System.Windows.Forms.ProgressBar
    Me.cmdCancel = New System.Windows.Forms.Button
    Me.SuspendLayout()
    '
    'lblMainTask
    '
    Me.lblMainTask.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblMainTask.ImeMode = System.Windows.Forms.ImeMode.NoControl
    Me.lblMainTask.Location = New System.Drawing.Point(8, 8)
    Me.lblMainTask.Name = "lblMainTask"
    Me.lblMainTask.Size = New System.Drawing.Size(376, 32)
    Me.lblMainTask.TabIndex = 0
    Me.lblMainTask.Text = "MainTask"
    Me.lblMainTask.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'lblItemTask
    '
    Me.lblItemTask.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblItemTask.ImeMode = System.Windows.Forms.ImeMode.NoControl
    Me.lblItemTask.Location = New System.Drawing.Point(32, 48)
    Me.lblItemTask.Name = "lblItemTask"
    Me.lblItemTask.Size = New System.Drawing.Size(336, 16)
    Me.lblItemTask.TabIndex = 1
    Me.lblItemTask.Text = "ItemTask"
    Me.lblItemTask.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'ProgressBar
    '
    Me.ProgressBar.Location = New System.Drawing.Point(32, 64)
    Me.ProgressBar.Name = "ProgressBar"
    Me.ProgressBar.Size = New System.Drawing.Size(328, 24)
    Me.ProgressBar.TabIndex = 2
    '
    'cmdCancel
    '
    Me.cmdCancel.FlatStyle = System.Windows.Forms.FlatStyle.System
    Me.cmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.cmdCancel.Location = New System.Drawing.Point(168, 96)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.TabIndex = 3
    Me.cmdCancel.Text = "Cancelar"
    '
    'frmProgress
    '
    Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
    Me.ClientSize = New System.Drawing.Size(392, 134)
    Me.ControlBox = False
    Me.Controls.Add(Me.cmdCancel)
    Me.Controls.Add(Me.ProgressBar)
    Me.Controls.Add(Me.lblItemTask)
    Me.Controls.Add(Me.lblMainTask)
    Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmProgress"
    Me.ShowInTaskbar = False
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Informació"
    Me.TopMost = True
    Me.ResumeLayout(False)

  End Sub

#End Region

  ''' <summary>
  ''' Inicialitza el formulari i el progressbar.
  ''' </summary>
  ''' <param name="MainTask"></param>
  ''' <param name="Title"></param>
  ''' <param name="MinValue"></param>
  ''' <param name="MaxValue"></param>
  ''' <param name="StepValue"></param>
  ''' <param name="ShowCancelButton"></param>
  ''' <remarks></remarks>
  Public Sub Display(ByVal MainTask As String, ByVal Title As String, ByVal MinValue As Integer, ByVal MaxValue As Integer, ByVal StepValue As Integer, ByVal ShowCancelButton As Boolean)
    lblMainTask.Text = MainTask
    lblItemTask.Text = ""
    Me.Text = Title
    ProgressBar.Minimum = MinValue
    ProgressBar.Maximum = MaxValue
    ProgressBar.Step = StepValue
    cmdCancel.Visible = ShowCancelButton
    Me.Show()
    'System.Threading.Thread.CurrentThread.Sleep(500)
    System.Windows.Forms.Application.DoEvents()
  End Sub

  Public Sub PerformStep()
    If ProgressBar.Value = ProgressBar.Maximum Then
      ProgressBar.Value = ProgressBar.Minimum
    End If
    ProgressBar.PerformStep()
    System.Windows.Forms.Application.DoEvents()
  End Sub

  Public WriteOnly Property ItemTask() As String
    Set(ByVal Value As String)
      lblItemTask.Text = Value
      System.Windows.Forms.Application.DoEvents()
    End Set
  End Property

  Public WriteOnly Property Maximum() As Integer
    Set(ByVal Value As Integer)
      ProgressBar.Maximum = Value
    End Set
  End Property

  Private _Cancel As Boolean = False
  Public ReadOnly Property Cancel() As Boolean
    Get
      Return _Cancel
    End Get
  End Property

  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    _Cancel = True
  End Sub

  Public WriteOnly Property CurrentValue() As Integer
    Set(ByVal Value As Integer)
      If Value > ProgressBar.Maximum Then
        Value = ProgressBar.Maximum
      End If
      ProgressBar.Value = Value
      System.Windows.Forms.Application.DoEvents()
    End Set
  End Property

End Class

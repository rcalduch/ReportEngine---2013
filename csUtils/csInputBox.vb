' GetDotNetCode Replacement for VB InputBox Function VB.NET 1.0

Public Class csInputBox
  Inherits System.Windows.Forms.Form

  Private m_OnlyNumbers As Boolean

#Region " Windows Form Designer generated code "

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
  Friend WithEvents PromptLabel As System.Windows.Forms.Label
  Friend WithEvents InputTextBox As System.Windows.Forms.TextBox
  Friend WithEvents CancelDialogButton As System.Windows.Forms.Button
  Friend WithEvents OKDialogButton As System.Windows.Forms.Button
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Me.PromptLabel = New System.Windows.Forms.Label
    Me.InputTextBox = New System.Windows.Forms.TextBox
    Me.OKDialogButton = New System.Windows.Forms.Button
    Me.CancelDialogButton = New System.Windows.Forms.Button
    Me.SuspendLayout()
    '
    'PromptLabel
    '
    Me.PromptLabel.Location = New System.Drawing.Point(12, 12)
    Me.PromptLabel.Name = "PromptLabel"
    Me.PromptLabel.Size = New System.Drawing.Size(288, 48)
    Me.PromptLabel.TabIndex = 0
    '
    'InputTextBox
    '
    Me.InputTextBox.Location = New System.Drawing.Point(12, 72)
    Me.InputTextBox.Name = "InputTextBox"
    Me.InputTextBox.Size = New System.Drawing.Size(384, 20)
    Me.InputTextBox.TabIndex = 1
    Me.InputTextBox.Text = ""
    '
    'OKDialogButton
    '
    Me.OKDialogButton.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.OKDialogButton.Location = New System.Drawing.Point(324, 7)
    Me.OKDialogButton.Name = "OKDialogButton"
    Me.OKDialogButton.TabIndex = 2
    Me.OKDialogButton.Text = "Acceptar"
    '
    'CancelDialogButton
    '
    Me.CancelDialogButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.CancelDialogButton.Location = New System.Drawing.Point(324, 36)
    Me.CancelDialogButton.Name = "CancelDialogButton"
    Me.CancelDialogButton.TabIndex = 3
    Me.CancelDialogButton.Text = "Cancelar"
    '
    'csInputBox
    '
    Me.AcceptButton = Me.OKDialogButton
    Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
    Me.ClientSize = New System.Drawing.Size(412, 110)
    Me.Controls.Add(Me.CancelDialogButton)
    Me.Controls.Add(Me.OKDialogButton)
    Me.Controls.Add(Me.InputTextBox)
    Me.Controls.Add(Me.PromptLabel)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "csInputBox"
    Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
    Me.Text = "Form1"
    Me.ResumeLayout(False)

  End Sub

#End Region

  Public Result As String

  ' Advanced GdncInputBox constructor.
  Public Sub New(ByVal title As String, ByVal prompt As String, _
    ByVal defaultResponse As String, ByVal formBorderStyle As Windows.Forms.FormBorderStyle, _
    ByVal OnlyNumbers As Boolean)
    MyBase.New()
    'This call is required by the Windows Form Designer.
    InitializeComponent()

    Me.Text = title
    Me.PromptLabel.Text = prompt
    Me.InputTextBox.Text = defaultResponse
    Me.FormBorderStyle = formBorderStyle
    Me.Icon = Nothing
    Me.StartPosition = Windows.Forms.FormStartPosition.CenterScreen
    Me.m_OnlyNumbers = OnlyNumbers
  End Sub

  Public Sub New(ByVal title As String, ByVal prompt As String, _
ByVal defaultResponse As String, ByVal XPosition As Integer, _
ByVal YPosition As Integer, ByVal formBorderStyle As Windows.Forms.FormBorderStyle, _
ByVal formBackColor As System.Drawing.Color)
    MyBase.New()
    'This call is required by the Windows Form Designer.
    InitializeComponent()

    Me.Text = title
    Me.PromptLabel.Text = prompt
    Me.InputTextBox.Text = defaultResponse
    Dim dialogPosition As System.Drawing.Point
    dialogPosition.X = XPosition
    dialogPosition.Y = YPosition
    Me.Location = dialogPosition
    Me.FormBorderStyle = formBorderStyle
    Me.BackColor = formBackColor
    Me.Icon = Nothing
    Me.m_OnlyNumbers = False
  End Sub


  ' Simple GdncInputBox constructor.
  Public Sub New(ByVal title As String, ByVal prompt As String, _
  ByVal defaultResponse As String)
    MyBase.New()
    'This call is required by the Windows Form Designer.
    InitializeComponent()
    Me.Text = title
    Me.PromptLabel.Text = prompt
    Me.InputTextBox.Text = defaultResponse
    Me.StartPosition = Windows.Forms.FormStartPosition.CenterScreen
    Me.Icon = Nothing
  End Sub

  ' If user clicks Cancel button clear InputTextBox text.
  Private Sub CancelDialogButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CancelDialogButton.Click
    Me.InputTextBox.Text = ""
  End Sub

  Private Sub OKDialogButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OKDialogButton.Click
    Result = Me.InputTextBox.Text
  End Sub

  Private Sub InputTextBox_KeyPress1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles InputTextBox.KeyPress
    If m_OnlyNumbers Then
      If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) Then
        e.Handled = True
      End If
    End If
  End Sub

 
End Class


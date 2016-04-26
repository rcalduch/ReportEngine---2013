'Copyright (C) 2002 Microsoft Corporation
'All rights reserved.
'THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER 
'EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF 
'MERCHANTIBILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.

'Requires the Trial or Release version of Visual Studio .NET Professional (or greater).

Option Strict On

Public Class frmStatus
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
  Friend WithEvents lblStatus As System.Windows.Forms.Label
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Me.lblStatus = New System.Windows.Forms.Label
    Me.SuspendLayout()
    '
    'lblStatus
    '
    Me.lblStatus.BackColor = System.Drawing.SystemColors.ControlLightLight
    Me.lblStatus.Dock = System.Windows.Forms.DockStyle.Fill
    Me.lblStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblStatus.ImeMode = System.Windows.Forms.ImeMode.NoControl
    Me.lblStatus.Location = New System.Drawing.Point(0, 0)
    Me.lblStatus.Name = "lblStatus"
    Me.lblStatus.Size = New System.Drawing.Size(336, 94)
    Me.lblStatus.TabIndex = 0
    Me.lblStatus.Text = "Label1"
    Me.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
    '
    'frmStatus
    '
    Me.AutoScaleBaseSize = New System.Drawing.Size(6, 16)
    Me.ClientSize = New System.Drawing.Size(336, 94)
    Me.ControlBox = False
    Me.Controls.Add(Me.lblStatus)
    Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmStatus"
    Me.ShowInTaskbar = False
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Informació"
    Me.TopMost = True
    Me.ResumeLayout(False)

  End Sub

#End Region

  ' This routine shows the Form with a message.
  Public Overloads Sub Show(ByVal Message As String)
    lblStatus.Text = Message
    Me.Show()
    'System.Threading.Thread.CurrentThread.Sleep(500)
    System.Windows.Forms.Application.DoEvents()
  End Sub

  Public Overloads Sub Show(ByVal Message As String, ByVal Title As String)
    lblStatus.Text = Message
    Me.Text = Title
    Me.Show()
    'System.Threading.Thread.CurrentThread.Sleep(500)
    System.Windows.Forms.Application.DoEvents()
  End Sub

End Class

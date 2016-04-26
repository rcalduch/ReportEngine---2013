<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class csButton
    Inherits System.Windows.Forms.UserControl

    'UserControl reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
  <System.Diagnostics.DebuggerStepThrough()> _
  Private Sub InitializeComponent()
    Me.components = New System.ComponentModel.Container
    Me.cxmPrint = New System.Windows.Forms.ContextMenuStrip(Me.components)
    Me.cxiSeleccionarImpressora = New System.Windows.Forms.ToolStripMenuItem
    Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator
    Me.cxiImprimir = New System.Windows.Forms.ToolStripMenuItem
    Me.cxiVistaPrevia = New System.Windows.Forms.ToolStripMenuItem
    Me.cxiPDF = New System.Windows.Forms.ToolStripMenuItem
    Me.cxiFax = New System.Windows.Forms.ToolStripMenuItem
    Me.cxiEmail = New System.Windows.Forms.ToolStripMenuItem
    Me.cxiExcel = New System.Windows.Forms.ToolStripMenuItem
    Me.picDropDown = New System.Windows.Forms.PictureBox
    Me.cmdPrint = New System.Windows.Forms.Button
    Me.cxmPrint.SuspendLayout()
    CType(Me.picDropDown, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SuspendLayout()
    '
    'cxmPrint
    '
    Me.cxmPrint.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.cxiSeleccionarImpressora, Me.ToolStripSeparator1, Me.cxiImprimir, Me.cxiVistaPrevia, Me.cxiPDF, Me.cxiFax, Me.cxiEmail, Me.cxiExcel})
    Me.cxmPrint.Name = "ContextMenuStrip1"
    Me.cxmPrint.ShowCheckMargin = True
    Me.cxmPrint.ShowImageMargin = False
    Me.cxmPrint.Size = New System.Drawing.Size(201, 164)
    '
    'cxiSeleccionarImpressora
    '
    Me.cxiSeleccionarImpressora.CheckOnClick = True
    Me.cxiSeleccionarImpressora.Name = "cxiSeleccionarImpressora"
    Me.cxiSeleccionarImpressora.Size = New System.Drawing.Size(200, 22)
    Me.cxiSeleccionarImpressora.Text = "Sel.leccionar impressora"
    '
    'ToolStripSeparator1
    '
    Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
    Me.ToolStripSeparator1.Size = New System.Drawing.Size(197, 6)
    '
    'cxiImprimir
    '
    Me.cxiImprimir.CheckOnClick = True
    Me.cxiImprimir.Name = "cxiImprimir"
    Me.cxiImprimir.Size = New System.Drawing.Size(200, 22)
    Me.cxiImprimir.Text = "Imprimir"
    '
    'cxiVistaPrevia
    '
    Me.cxiVistaPrevia.CheckOnClick = True
    Me.cxiVistaPrevia.Name = "cxiVistaPrevia"
    Me.cxiVistaPrevia.Size = New System.Drawing.Size(200, 22)
    Me.cxiVistaPrevia.Text = "Vista prèvia"
    '
    'cxiPDF
    '
    Me.cxiPDF.CheckOnClick = True
    Me.cxiPDF.Name = "cxiPDF"
    Me.cxiPDF.Size = New System.Drawing.Size(200, 22)
    Me.cxiPDF.Text = "Adobe PDF"
    '
    'cxiFax
    '
    Me.cxiFax.CheckOnClick = True
    Me.cxiFax.Name = "cxiFax"
    Me.cxiFax.Size = New System.Drawing.Size(200, 22)
    Me.cxiFax.Text = "Fax"
    '
    'cxiEmail
    '
    Me.cxiEmail.CheckOnClick = True
    Me.cxiEmail.Name = "cxiEmail"
    Me.cxiEmail.Size = New System.Drawing.Size(200, 22)
    Me.cxiEmail.Text = "e-mail"
    '
    'cxiExcel
    '
    Me.cxiExcel.CheckOnClick = True
    Me.cxiExcel.Name = "cxiExcel"
    Me.cxiExcel.Size = New System.Drawing.Size(200, 22)
    Me.cxiExcel.Text = "Excel"
    '
    'picDropDown
    '
    Me.picDropDown.ContextMenuStrip = Me.cxmPrint
    Me.picDropDown.Image = Global.csUtils.My.Resources.Resources.flecha
    Me.picDropDown.Location = New System.Drawing.Point(54, 9)
    Me.picDropDown.Name = "picDropDown"
    Me.picDropDown.Size = New System.Drawing.Size(9, 8)
    Me.picDropDown.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
    Me.picDropDown.TabIndex = 29
    Me.picDropDown.TabStop = False
    '
    'cmdPrint
    '
    Me.cmdPrint.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
    Me.cmdPrint.ContextMenuStrip = Me.cxmPrint
    Me.cmdPrint.Dock = System.Windows.Forms.DockStyle.Fill
    Me.cmdPrint.FlatStyle = System.Windows.Forms.FlatStyle.System
    Me.cmdPrint.ImageAlign = System.Drawing.ContentAlignment.BottomRight
    Me.cmdPrint.Location = New System.Drawing.Point(0, 0)
    Me.cmdPrint.Name = "cmdPrint"
    Me.cmdPrint.Size = New System.Drawing.Size(72, 23)
    Me.cmdPrint.TabIndex = 27
    Me.cmdPrint.Text = "   &Imprimir"
    Me.cmdPrint.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    Me.cmdPrint.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
    '
    'csButton
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.Controls.Add(Me.picDropDown)
    Me.Controls.Add(Me.cmdPrint)
    Me.Name = "csButton"
    Me.Size = New System.Drawing.Size(72, 23)
    Me.cxmPrint.ResumeLayout(False)
    CType(Me.picDropDown, System.ComponentModel.ISupportInitialize).EndInit()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
  Public WithEvents cxiSeleccionarImpressora As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents cxiExcel As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents picDropDown As System.Windows.Forms.PictureBox
  Friend WithEvents cmdPrint As System.Windows.Forms.Button

End Class

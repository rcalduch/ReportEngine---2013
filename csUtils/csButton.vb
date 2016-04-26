Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms
Imports csUtils.csRpt
Imports csUtils.csTabularRpt

<System.ComponentModel.DefaultEventAttribute("ButtonClick")> _
Public Class csButton

#Region "Declaració e inicialització de variables "
  Private dtTable As DataTable
  Public SetPrinter As Boolean = True
  Public SelectedTipusImpresio As String
  Public SeleccionarImpressora As Boolean
#End Region

  Private WithEvents cxmPrint As System.Windows.Forms.ContextMenuStrip
  Private WithEvents cxiImprimir As System.Windows.Forms.ToolStripMenuItem
  Private WithEvents cxiVistaPrevia As System.Windows.Forms.ToolStripMenuItem
  Private WithEvents cxiPDF As System.Windows.Forms.ToolStripMenuItem
  Private WithEvents cxiFax As System.Windows.Forms.ToolStripMenuItem
  Private WithEvents cxiEmail As System.Windows.Forms.ToolStripMenuItem
  Public Event ButtonClick(ByVal e As csButtonPrinterEventargs)


  Private mImprimir As Boolean = True
  Private mVistaPrevia As Boolean = True
  Private mPDF As Boolean = True
  Private mFax As Boolean = True
  Private mEmail As Boolean = True
  Private mExcel As Boolean = True
  Private mPDFAvaiable As Boolean
  Private mText As String = "   &Imprimir"


#Region "Events"

  Private Sub csButton_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Me.SelectedTipusImpresio = "I"
    cxiImprimir.Checked = True
    Me.SeleccionarImpressora = True
    cxiSeleccionarImpressora.Checked = True


    Dim key As Microsoft.Win32.RegistryKey
    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\FinePrint Software\pdfFactory3\FinePrinters\pdfFactory", True)
    mPDFAvaiable = Not IsNothing(key)
    PDF = mPDF
    Fax = mFax
    Email = mEmail
  End Sub

  Private Sub cmdPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdPrint.Click
    RaiseEvent ButtonClick(New csButtonPrinterEventargs(SelectedTipusImpresio, SeleccionarImpressora))
  End Sub

  Private Sub cxiEmail_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cxiEmail.Click
    SelectedTipusImpresio = "E"
    cxiEmail.Checked = True
    cxiFax.Checked = False
    cxiImprimir.Checked = False
    cxiPDF.Checked = False
    cxiVistaPrevia.Checked = False
    cxiExcel.Checked = False
  End Sub

  Private Sub cxiFax_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cxiFax.Click
    SelectedTipusImpresio = "F"
    cxiEmail.Checked = False
    cxiFax.Checked = True
    cxiImprimir.Checked = False
    cxiPDF.Checked = False
    cxiVistaPrevia.Checked = False
    cxiExcel.Checked = False
  End Sub

  Private Sub cxiImprimir_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cxiImprimir.Click
    SelectedTipusImpresio = "I"
    cxiEmail.Checked = False
    cxiFax.Checked = False
    cxiImprimir.Checked = True
    cxiPDF.Checked = False
    cxiVistaPrevia.Checked = False
    cxiExcel.Checked = False
  End Sub

  Private Sub cxiPDF_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cxiPDF.Click
    SelectedTipusImpresio = "P"
    cxiEmail.Checked = False
    cxiFax.Checked = False
    cxiImprimir.Checked = False
    cxiPDF.Checked = True
    cxiVistaPrevia.Checked = False
    cxiExcel.Checked = False
  End Sub

  Private Sub cxiVistaPrevia_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cxiVistaPrevia.Click
    SelectedTipusImpresio = "V"
    cxiEmail.Checked = False
    cxiFax.Checked = False
    cxiImprimir.Checked = False
    cxiPDF.Checked = False
    cxiVistaPrevia.Checked = True
    cxiExcel.Checked = False
  End Sub

  Private Sub cxiExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cxiExcel.Click
    SelectedTipusImpresio = "X"
    cxiEmail.Checked = False
    cxiFax.Checked = False
    cxiImprimir.Checked = False
    cxiPDF.Checked = False
    cxiVistaPrevia.Checked = False
    cxiExcel.Checked = True
  End Sub

  Private Sub cxiSeleccionarImpressora_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cxiSeleccionarImpressora.Click
    SeleccionarImpressora = Not SeleccionarImpressora
  End Sub

#End Region

  <Description("Texto"), Category("Opcions visibles"), DefaultValue("   &Imprimir")> _
 Public Property Texto() As String
    Get
      Return mText
    End Get
    Set(ByVal value As String)
      Me.cmdPrint.Text = value
      mText = value
    End Set
  End Property

  <Description("Opció imprimir"), Category("Opcions visibles"), DefaultValue(True)> _
  Public Property Imprimir() As Boolean
    Get
      Return mImprimir
    End Get
    Set(ByVal value As Boolean)
      cxiImprimir.Visible = value
      mImprimir = value
    End Set
  End Property

  <Description("Opció vista previa"), Category("Opcions visibles"), DefaultValue(True)> _
  Public Property VistaPrevia() As Boolean
    Get
      Return mVistaPrevia
    End Get
    Set(ByVal value As Boolean)
      cxiVistaPrevia.Visible = value
      mVistaPrevia = value
    End Set
  End Property

  <Description("Opció PDF"), Category("Opcions visibles"), DefaultValue(True)> _
  Public Property PDF() As Boolean
    Get
      Return mPDF
    End Get
    Set(ByVal value As Boolean)
      cxiPDF.Visible = value And mPDFAvaiable
      mPDF = value
    End Set
  End Property

  <Description("Opció Fax"), Category("Opcions visibles"), DefaultValue(True)> _
  Public Property Fax() As Boolean
    Get
      Return mFax
    End Get
    Set(ByVal value As Boolean)
      cxiFax.Visible = value And mPDFAvaiable
      mFax = value
    End Set
  End Property

  <Description("Opció e-mail"), Category("Opcions visibles"), DefaultValue(True)> _
  Public Property Email() As Boolean
    Get
      Return mEmail
    End Get
    Set(ByVal value As Boolean)
      cxiEmail.Visible = value And mPDFAvaiable
      mEmail = value
    End Set
  End Property

  <Description("Opció excel"), Category("Opcions visibles"), DefaultValue(True)> _
  Public Property Excel() As Boolean
    Get
      Return mExcel
    End Get
    Set(ByVal value As Boolean)
      cxiExcel.Visible = value
      mExcel = value
    End Set
  End Property

  Private Sub picDropDown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles picDropDown.Click
    cmdPrint.PerformClick()
  End Sub
End Class



Public Class csButtonPrinterEventargs
  Inherits System.EventArgs
  Public Destination As String
  Public ShowPrinterDialog As Boolean

  Public Sub New(ByVal SelectedTipusImpresio As String, ByVal SeleccionarImpressora As Boolean)
    MyBase.New()
    Me.Destination = SelectedTipusImpresio
    Me.ShowPrinterDialog = SeleccionarImpressora
  End Sub

End Class
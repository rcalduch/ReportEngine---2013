Public Structure InfoEMailServerStruct
  Public MailUsuari As String
  Public NomUsuari As String
  Public ServerSMTP As String
  Public UsuariSMPT As String
  Public PasswordSMTP As String
  Public MailSMTP As String
  Public MailFeedback As String
  Public MailFaxAccount As String
End Structure

Public Structure DocumentCartaAgrupats
  Dim Capçalera As String
  Dim PeuCarta As String
  Dim Factura As String
  Dim DataFactura As String
  Dim DataVenciment As String
  Dim Import As String
End Structure

Public Structure ReportWorkInfo
  Dim WorkID As String
  Dim ReportID As String
  Dim PrinterID As String
  Dim WorkToProcess As Boolean
  Dim Succes As Boolean
End Structure
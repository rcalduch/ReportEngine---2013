Public Enum MonthEnum
  Gener = 1
  Febrer = 2
  Març = 3
  Abril = 4
  Maig = 5
  Juny = 6
  Juliol = 7
  Agost = 8
  Setembre = 9
  Octubre = 10
  Novembre = 11
  Decembre = 12
End Enum

Public Enum EmpresesEnum
  Custom = 0
End Enum

Public Enum SystemNodeLevel
  General = 1
  Comptadors = 2
  Comptabilitat = 3
End Enum

Public Enum EstatsIncidenciesEnum
  Obert = 0
  Vist = 1
  Pendent = 2
  Resolt = 3
  Tancat = 4
End Enum

Public Enum enFormStatus
  fsClosing = 1
  fsReadOnly = 2
  fsBrowse = 3
  fsEditing = 4
  fsAddingOne = 5
  fsAddingMany = 6
  fsCanceling = 7
End Enum

Public Enum enAccesLevel
  alDenegat
  alLectura
  alEscritura
  alHeredat
End Enum

Public Enum TipusAdresaEnum
  Fiscal = 1
  Postal = 2
  Enviament = 3
End Enum

Public Enum TipusRegistreEnum
  ErrorAuto = 0
  Incidencia = 1
  Millora = 2
  Modificacio = 3
  AComentar = 4
End Enum

Public Enum TipusEntitatEnum
  Altres = -1
  Client = 0
  Proveidor = 1
  AgentComercial = 2
  EntitatBancaria = 3
  Transportiste = 4
  Empresa = 5
  Autovenda = 6
End Enum

Public Enum TipusImpressoraEnum
  Sense_Especificar = 0
  Etiqueta_Gran = 1
  Etiqueta_Mitjana = 2
  Etiqueta_Petita = 3
  DIN_A4_BiN = 4
  DIN_A4_Color = 5
  DIN_A3_BiN = 6
  DIN_A3_Color = 7
End Enum

Public Enum TipusAvisEnum
  STK_TrencamentStock = 1
  OFE_ActivacioOferta = 2
  CAR_DevolucioEfecte = 3
  TAR_ActivacioNovaTarifa = 4
  TAR_ErrorActivacioTarifa = 5
  STK_ModificacioDisponibilitat = 6
  CAR_TancamentRebutGT10 = 7
  CAR_RebutDescomptat = 8
  CTB_ModificacioFacturaAssentada = 9
  DBG_Debug = 10
End Enum

Public Enum TipusEnviamentDocumentEnum
  NA = 0
  Impressora = 1
  VistaPrevia = 2
  Pdf = 3
  Email = 4
  Fax = 5
  Excel = 6
End Enum

Public Enum TipusDocumentEnum
  No_Document = 0
  Albara_A_Client = 1
  Entrada_Compres = 2
  Comanda_A_Proveidor = 3
  Ordre_De_Treball = 4
  Factura_A_Client = 5
  Produccio_Bonsai = 6
  Moviment_Magatzem_En_Transit = 7
  Sessio_De_Treball = 8
  Proveidor_ID = 9
  Data_Document = 10
  Comanda_Custom_A_Proveidor = 11
  Rebut_A_Client = 12
  Embalum_Gefco = 22
End Enum

Public Enum TipusPrioritatEnum
  Baixa = 0
  Mitjana = 1
  Alta = 2
End Enum

Public Enum TipusSubjecteEnum
  TipusSubjecteRol = 0
  TipusSubjecteUsuari = 1
End Enum

Public Enum TipusRelacioLiquidacionsEnum
  Nota
  Liquidacio
  Relacio
End Enum



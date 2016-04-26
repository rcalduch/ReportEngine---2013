'------------------------------------------------------------------------------
' Clase para manejar ficheros de configuración                      (15/Nov/05)
'
' Las secciones siempre estarán dentro de <configuration>
' al menos así lo guardará esta clase, aunque permite leer pares key / value.
' Para que se sepa que se lee de configuration,
' en el código se indica explícitamente.
'
' Pero para usarla de forma independiente de ConfigurationSettings
'
' Revisado para poder guardar automáticamente                       (21/Feb/06)
' Poder leer todas las secciones y las claves de una sección        (21/Feb/06)
'
'------------------------------------------------------------------------------
Option Explicit On
Option Strict On

Imports Microsoft.VisualBasic
Imports System

Imports System.Collections
Imports System.Collections.Generic
Imports System.Configuration
Imports System.Xml
Imports System.IO


Public Class ConfigXml

  '----------------------------------------------------------------------
  ' Los campos y métodos privados
  '----------------------------------------------------------------------
  Private mGuardarAlAsignar As Boolean = True
  Private Const configuration As String = "configuration/"
  Private ficConfig As String = ""
  Private configXml As New XmlDocument
  '
  ''' <summary>
  ''' Indica si se se guardarán los datos cuando se añadan nuevos.
  ''' </summary>
  ''' <value>Indica si se se guardarán los datos cuando se añadan nuevos.</value>
  ''' <returns>Un valor verdadero o falso según el valor de la propiedad</returns>
  ''' <remarks></remarks>
  Public Property GuardarAlAsignar() As Boolean
    Get
      Return mGuardarAlAsignar
    End Get
    Set(ByVal value As Boolean)
      mGuardarAlAsignar = value
    End Set
  End Property
  '
  ''' <summary>
  ''' Obtiene un valor de tipo cadena de la sección y clave indicadas.
  ''' </summary>
  ''' <param name="seccion">La sección de la que queremos obtener el valor</param>
  ''' <param name="clave">La clave de la que queremos recuperar el valor</param>
  ''' <returns>Un valor de tipo cadena con el valor de la sección y clave indicadas</returns>
  ''' <remarks>
  ''' Existe otra sobrecarga para indicar un valor predeterminado.
  ''' Tanbién hay otras dos sobrecargas para valores enteros y boolean.
  ''' </remarks>
  Public Function GetValue(ByVal seccion As String, ByVal clave As String) As String
    Return GetValue(seccion, clave, "")
  End Function
  ''' <summary>
  ''' Obtiene un valor de tipo cadena de la sección y clave indicadas.
  ''' </summary>
  ''' <param name="seccion">La sección de la que queremos obtener el valor</param>
  ''' <param name="clave">La clave de la que queremos recuperar el valor</param>
  ''' <param name="predeterminado">El valor predeterminado para cuando no exista.</param>
  ''' <returns>Un valor de tipo cadena con el valor de la sección y clave indicadas</returns>
  ''' <remarks></remarks>
  Public Function GetValue(ByVal seccion As String, ByVal clave As String, ByVal predeterminado As String) As String
    Return cfgGetValue(seccion, clave, predeterminado)
  End Function
  ''' <summary>
  ''' Obtiene un valor de tipo entero de la sección y clave indicadas.
  ''' </summary>
  ''' <param name="seccion">La sección de la que queremos obtener el valor</param>
  ''' <param name="clave">La clave de la que queremos recuperar el valor</param>
  ''' <param name="predeterminado">El valor predeterminado para cuando no exista.</param>
  ''' <returns>Un valor de tipo entero con el valor de la sección y clave indicadas</returns>
  ''' <remarks></remarks>
  Public Function GetValue(ByVal seccion As String, ByVal clave As String, ByVal predeterminado As Integer) As Integer
    Return CInt(cfgGetValue(seccion, clave, predeterminado.ToString))
  End Function
  ''' <summary>
  ''' Obtiene un valor de tipo boolean de la sección y clave indicadas.
  ''' </summary>
  ''' <param name="seccion">La sección de la que queremos obtener el valor</param>
  ''' <param name="clave">La clave de la que queremos recuperar el valor</param>
  ''' <param name="predeterminado">El valor predeterminado para cuando no exista.</param>
  ''' <returns>Un valor de tipo boolean con el valor de la sección y clave indicadas</returns>
  ''' <remarks>Internamente el valor se guarda con un cero para False y uno para True</remarks>
  Public Function GetValue(ByVal seccion As String, ByVal clave As String, ByVal predeterminado As Boolean) As Boolean
    Dim def As String = "0"
    If predeterminado Then def = "1"
    def = cfgGetValue(seccion, clave, def)
    If def = "1" Then
      Return True
    Else
      Return False
    End If
  End Function

  ''' <summary>
  ''' Asignar un valor de tipo cadena en la sección y clave indicadas.
  ''' </summary>
  ''' <param name="seccion">La sección de la que queremos obtener el valor</param>
  ''' <param name="clave">La clave de la que queremos recuperar el valor</param>
  ''' <param name="valor">El valor a asignar</param>
  ''' <remarks>
  ''' El valor se guardar como un elemento de la sección indicada.
  ''' <seealso cref="SetKeyValue" />
  ''' </remarks>
  Public Sub SetValue(ByVal seccion As String, ByVal clave As String, ByVal valor As String)
    cfgSetValue(seccion, clave, valor)
  End Sub

  ''' <summary>
  ''' Asignar un valor de tipo entero en la sección y clave indicadas.
  ''' </summary>
  ''' <param name="seccion">La sección de la que queremos obtener el valor</param>
  ''' <param name="clave">La clave de la que queremos recuperar el valor</param>
  ''' <param name="valor">El valor a asignar</param>
  ''' <remarks>
  ''' El valor se guardar como un elemento de la sección indicada.
  ''' El valor siempre se guarda como un valor de cadena.
  ''' <seealso cref="SetKeyValue" />
  ''' </remarks>
  Public Sub SetValue(ByVal seccion As String, ByVal clave As String, ByVal valor As Integer)
    cfgSetValue(seccion, clave, valor.ToString)
  End Sub

  ''' <summary>
  ''' Asignar un valor de tipo boolean en la sección y clave indicadas.
  ''' </summary>
  ''' <param name="seccion">La sección de la que queremos obtener el valor</param>
  ''' <param name="clave">La clave de la que queremos recuperar el valor</param>
  ''' <param name="valor">El valor a asignar</param>
  ''' <remarks>
  ''' El valor se guardar como un elemento de la sección indicada.
  ''' El valor siempre se guarda como un valor de cadena, siendo un 1 para True y 0 para False.
  ''' <seealso cref="SetKeyValue" />
  ''' </remarks>
  Public Sub SetValue(ByVal seccion As String, ByVal clave As String, ByVal valor As Boolean)
    If valor Then
      cfgSetValue(seccion, clave, "1")
    Else
      cfgSetValue(seccion, clave, "0")
    End If
  End Sub

  ''' <summary>
  ''' Asigna un valor de tipo cadena en la sección y clave indicadas.
  ''' </summary>
  ''' <param name="seccion">La sección de la que queremos obtener el valor</param>
  ''' <param name="clave">La clave de la que queremos recuperar el valor</param>
  ''' <param name="valor">El valor a asignar</param>
  ''' <remarks>
  ''' El valor se guarda como un atributo de la sección indicada.
  ''' La clave se guarda con el atributo key y el valor con el atributo value.
  ''' <seealso cref="SetValue" />
  ''' </remarks>
  Public Sub SetKeyValue(ByVal seccion As String, ByVal clave As String, ByVal valor As String)
    cfgSetKeyValue(seccion, clave, valor)
  End Sub

  ''' <summary>
  ''' Asigna un valor de tipo entero en la sección y clave indicadas.
  ''' </summary>
  ''' <param name="seccion">La sección de la que queremos obtener el valor</param>
  ''' <param name="clave">La clave de la que queremos recuperar el valor</param>
  ''' <param name="valor">El valor a asignar</param>
  ''' <remarks>
  ''' El valor se guarda como un atributo de la sección indicada.
  ''' La clave se guarda con el atributo key y el valor con el atributo value.
  ''' El valor siempre se guarda como un valor de cadena.
  ''' <seealso cref="SetValue" />
  ''' </remarks>
  Public Sub SetKeyValue(ByVal seccion As String, ByVal clave As String, ByVal valor As Integer)
    cfgSetKeyValue(seccion, clave, valor.ToString)
  End Sub

  ''' <summary>
  ''' Asigna un valor de tipo boolean en la sección y clave indicadas.
  ''' </summary>
  ''' <param name="seccion">La sección de la que queremos obtener el valor</param>
  ''' <param name="clave">La clave de la que queremos recuperar el valor</param>
  ''' <param name="valor">El valor a asignar</param>
  ''' <remarks>
  ''' El valor se guarda como un atributo de la sección indicada.
  ''' La clave se guarda con el atributo key y el valor con el atributo value.
  ''' El valor siempre se guarda como un valor de cadena, siendo un 1 para True y 0 para False.
  ''' <seealso cref="SetValue" />
  ''' </remarks>
  Public Sub SetKeyValue(ByVal seccion As String, ByVal clave As String, ByVal valor As Boolean)
    If valor Then
      cfgSetKeyValue(seccion, clave, "1")
    Else
      cfgSetKeyValue(seccion, clave, "0")
    End If
  End Sub

  ''' <summary>
  ''' Elimina la sección indicada, aunque en realidad la deja vacía.
  ''' </summary>
  ''' <param name="seccion">La sección a eliminar.</param>
  ''' <remarks></remarks>
  Public Sub RemoveSection(ByVal seccion As String)
    Dim n As XmlNode
    n = configXml.SelectSingleNode(configuration & seccion)
    If Not n Is Nothing Then
      n.RemoveAll()
      If mGuardarAlAsignar Then
        Me.Save()
      End If
    End If
  End Sub

  ' Guardar el fichero de configuración
  ' 
  ''' <summary>
  ''' Guardar el fichero de configuración.
  ''' </summary>
  ''' <remarks>
  ''' Si no se llama a este método, no se guardará de forma permanente.
  ''' Para guardar automáticamente al asignar,
  ''' asignar un valor verdadero a la propiedad <see cref="GuardarAlAsignar">GuardarAlAsignar</see>
  ''' </remarks>
  Public Sub Save()
    configXml.Save(ficConfig)
  End Sub

  ''' <summary>
  ''' Lee el fichero de configuración.
  ''' </summary>
  ''' <remarks>
  ''' Si no existe, se crea uno nuevo con los valores predeterminados.
  ''' </remarks>
  Public Sub Read()
    Dim fic As String = ficConfig
    Const revDate As String = "Tue, 21 Feb 2006 19:45:00 GMT"
    If File.Exists(fic) Then
      configXml.Load(fic)
      ' Actualizar los datos de la información de esta clase
      Dim b As Boolean = mGuardarAlAsignar
      mGuardarAlAsignar = False
      Me.SetValue("configXml_Info", "info", "Generado con Config para Visual Basic 2005")
      Me.SetValue("configXml_Info", "revision", revDate)
      Me.SetValue("configXml_Info", "formatoUTF8", "El formato de este fichero debe ser UTF-8")
      mGuardarAlAsignar = b
      Me.Save()
    Else
      ' Crear el XML de configuración con la sección General
      Dim sb As New System.Text.StringBuilder
      sb.Append("<?xml version=""1.0"" encoding=""utf-8"" ?>")
      sb.Append("<configuration>")
      ' Por si es un fichero appSetting
      sb.Append("<configSections>")
      sb.Append("<section name=""General"" type=""System.Configuration.DictionarySectionHandler"" />")
      sb.Append("</configSections>")
      sb.Append("<General>")
      sb.Append("<!-- Los valores irán dentro del elemento indicado por la clave -->")
      sb.Append("<!-- Aunque también se podrán indicar como pares key / value -->")
      sb.AppendFormat("<add key=""Revision"" value=""{0}"" />", revDate)
      sb.Append("<!-- La clase siempre los añade como un elemento -->")
      sb.Append("<Copyright>Custom Software, scp, 2005-2006</Copyright>")
      sb.Append("</General>")
      '
      sb.AppendFormat("<configXml_Info>{0}", vbCrLf)
      sb.AppendFormat("<info>Generado con Config para Visual Basic 2005</info>{0}", vbCrLf)
      sb.AppendFormat("<Copyright>Custom Software, scp, 2005-2006</Copyright>{0}", vbCrLf)
      sb.AppendFormat("<revision>{0}</revision>{1}", revDate, vbCrLf)
      sb.AppendFormat("<formatoUTF8>El formato de este fichero debe ser UTF-8</formatoUTF8>{0}", vbCrLf)
      sb.AppendFormat("</configXml_Info>{0}", vbCrLf)
      '
      sb.Append("</configuration>")
      ' Asignamos la cadena al objeto
      configXml.LoadXml(sb.ToString)
      '
      ' Guardamos el contenido de configXml y creamos el fichero
      configXml.Save(ficConfig)
    End If
  End Sub

  ''' <summary>
  ''' El nombre del fichero de configuración.
  ''' </summary>
  ''' <value>El path completo con el nombre del fichero de configuración.</value>
  ''' <returns>Una cadena con el fichero de configuración.</returns>
  ''' <remarks>El nombre del fichero se debe indicar en el constructor.</remarks>
  Public Property FileName() As String
    Get
      Return ficConfig
    End Get
    Set(ByVal value As String)
      ' Al asignarlo, NO leemos el contenido del fichero
      ficConfig = value
      'LeerFile()
    End Set
  End Property

  ''' <summary>
  ''' Constructor en el que indicamos el nombre del fichero de configuración.
  ''' </summary>
  ''' <param name="fic">El fichero a usar para guardar los datos de configuración.</param>
  ''' <remarks>
  ''' Si no existe, se creará.
  ''' Al usar este constructor, por defecto se guardarán los valores al asignarlos.
  ''' </remarks>
  Public Sub New(ByVal fic As String)
    ficConfig = fic
    ' Por defecto se guarda al asignar los valores
    mGuardarAlAsignar = True
    Read()
  End Sub

  ' Con este constructor podemos decidir si guardamos o no automáticamente
  ''' <summary>
  ''' Constructor en el que indicamos el nombre del fichero de configuración.
  ''' </summary>
  ''' <param name="fic">El fichero a usar para guardar los datos de configuración.</param>
  ''' <param name="guardarAlAsignar">
  ''' Un valor verdadero o falso para indicar
  ''' si se guardan los datos automáticamente al asignarlos.</param>
  ''' <remarks></remarks>
  Public Sub New(ByVal fic As String, ByVal guardarAlAsignar As Boolean)
    ficConfig = fic
    mGuardarAlAsignar = guardarAlAsignar
    Read()
  End Sub
  '
  ''' <summary>
  ''' Devuelve una colección de tipo List con las secciones del fichero de configuración.
  ''' </summary>
  ''' <returns>Una colección de tipo List(Of String) con las secciones del fichero de configuración.</returns>
  ''' <remarks>Este método solo se puede usar en la versión 2.0 o superior.</remarks>
  Public Function Secciones() As List(Of String)
    Dim d As New List(Of String)
    Dim root As XmlNode
    Dim s As String = "configuration"
    root = configXml.SelectSingleNode(s)
    If root IsNot Nothing Then
      For Each n As XmlNode In root.ChildNodes
        d.Add(n.Name)
      Next
    End If
    Return d
  End Function

  ''' <summary>
  ''' Devuelve una colección de tipo Dictionary con las claves y valores de la sección indicada.
  ''' </summary>
  ''' <param name="seccion">La sección de la que queremos obtener las claves y valores.</param>
  ''' <returns>Una colección de tipo Dictionary(Of String, String) con las claves y valores.</returns>
  ''' <remarks></remarks>
  Public Function Claves(ByVal seccion As String) As Dictionary(Of String, String)
    Dim d As New Dictionary(Of String, String)
    Dim root As XmlNode
    seccion = seccion.Replace(" ", "_")
    root = configXml.SelectSingleNode(configuration & seccion)
    If root IsNot Nothing Then
      For Each n As XmlNode In root.ChildNodes
        If d.ContainsKey(n.Name) = False Then
          d.Add(n.Name, n.InnerText)
        End If
      Next
    End If
    Return d
  End Function
  '
  '----------------------------------------------------------------------
  ' Los métodos privados
  '----------------------------------------------------------------------
  '
  ' El método interno para guardar los valores
  ' Este método siempre guardará en el formato <seccion><clave>valor</clave></seccion>
  Private Sub cfgSetValue( _
                  ByVal seccion As String, _
                  ByVal clave As String, _
                  ByVal valor As String)
    '
    Dim n As XmlNode
    '
    ' Filtrar los caracteres no válidos
    ' en principio solo comprobamos el espacio
    seccion = seccion.Replace(" ", "_")
    clave = clave.Replace(" ", "_")

    ' Se comprueba si es un elemento de la sección:
    '   <seccion><clave>valor</clave></seccion>
    n = configXml.SelectSingleNode(configuration & seccion & "/" & clave)
    If n IsNot Nothing Then
      n.InnerText = valor
    Else
      Dim root As XmlNode
      Dim elem As XmlElement
      root = configXml.SelectSingleNode(configuration & seccion)
      If root Is Nothing Then
        ' Si no existe el elemento principal,
        ' lo añadimos a <configuration>
        elem = configXml.CreateElement(seccion)
        configXml.DocumentElement.AppendChild(elem)
        root = configXml.SelectSingleNode(configuration & seccion)
      End If
      If root IsNot Nothing Then
        ' Crear el elemento
        elem = configXml.CreateElement(clave)
        elem.InnerText = valor
        ' Añadirlo al nodo indicado
        root.AppendChild(elem)
      End If
    End If
    '
    If mGuardarAlAsignar Then
      Me.Save()
    End If
  End Sub

  ' Asigna un atributo a una sección
  ' Por ejemplo: <Seccion clave=valor>...</Seccion>
  ' También se usará para el formato de appSettings: <add key=clave value=valor />
  '   Aunque en este caso, debe existir el elemento a asignar.
  Private Sub cfgSetKeyValue( _
                  ByVal seccion As String, _
                  ByVal clave As String, _
                  ByVal valor As String)
    '
    Dim n As XmlNode
    '
    ' Filtrar los caracteres no válidos
    ' en principio solo comprobamos el espacio
    seccion = seccion.Replace(" ", "_")
    clave = clave.Replace(" ", "_")

    n = configXml.SelectSingleNode(configuration & seccion & "/add[@key=""" & clave & """]")
    If n IsNot Nothing Then
      n.Attributes("value").InnerText = valor
    Else
      Dim root As XmlNode
      Dim elem As XmlElement
      root = configXml.SelectSingleNode(configuration & seccion)
      If root Is Nothing Then
        ' Si no existe el elemento principal,
        ' lo añadimos a <configuration>
        elem = configXml.CreateElement(seccion)
        configXml.DocumentElement.AppendChild(elem)
        root = configXml.SelectSingleNode(configuration & seccion)
      End If
      If root IsNot Nothing Then
        Dim a As XmlAttribute = CType(configXml.CreateNode(XmlNodeType.Attribute, clave, Nothing), XmlAttribute)
        a.InnerText = valor
        root.Attributes.Append(a)
      End If
    End If
    If mGuardarAlAsignar Then
      Me.Save()
    End If
  End Sub

  ' Devolver el valor de la clave indicada
  Private Function cfgGetValue( _
                  ByVal seccion As String, _
                  ByVal clave As String, _
                  ByVal valor As String _
                  ) As String
    '
    Dim n As XmlNode
    '
    ' Filtrar los caracteres no válidos
    ' en principio solo comprobamos el espacio
    seccion = seccion.Replace(" ", "_")
    clave = clave.Replace(" ", "_")
    ' Primero comprobar si están el formato de appSettings: <add key = clave value = valor />
    n = configXml.SelectSingleNode(configuration & seccion & "/add[@key=""" & clave & """]")
    If n IsNot Nothing Then
      Return n.Attributes("value").InnerText
    End If
    '
    ' Después se comprueba si está en el formato <Seccion clave = valor>
    n = configXml.SelectSingleNode(configuration & seccion)
    If n IsNot Nothing Then
      Dim a As XmlAttribute = n.Attributes(clave)
      If a IsNot Nothing Then
        Return a.InnerText
      End If
    End If
    '
    ' Por último se comprueba si es un elemento de seccion:
    '   <seccion><clave>valor</clave></seccion>
    n = configXml.SelectSingleNode(configuration & seccion & "/" & clave)
    If n IsNot Nothing Then
      Return n.InnerText
    End If
    '
    ' Si no existe, se devuelve el valor predeterminado
    Return valor
  End Function
End Class

Imports System.Text
Imports System.Xml
Imports System.Net
Imports System.Security
Imports System.Security.Cryptography.X509Certificates
Imports System.IO


''' <summary>
''' Clase para crear objetos Login Tickets
''' </summary>
''' <remarks>
''' Ver documentacion: 
'''    Especificacion Tecnica del Webservice de Autenticacion y Autorizacion
'''    Version 1.0
'''    Departamento de Seguridad Informatica - AFIP
''' </remarks>
Public Class LoginTicket
    Public Property CertFirmante As X509Certificate2

    Public UniqueId As UInt32 ' Entero de 32 bits sin signo que identifica el requerimiento
    Public GenerationTime As DateTime ' Momento en que fue generado el requerimiento
    Public ExpirationTime As DateTime ' Momento en el que expira la solicitud
    Public Sign As String ' Firma de seguridad recibida en la respuesta
    Public Token As String ' Token de seguridad recibido en la respuesta
    Public XmlLoginTicketResponse As XmlDocument = Nothing
    Private mstrPasswordSecureString As New SecureString
    Private mstrProxy As String = Nothing
    Private mstrProxyUser As String = Nothing
    Private mstrProxyPassword As String = ""
    Private mstrServicio As String = "wsfe"
    Private mstrRutaCertX509Firmante As String = "C:\SiCoFa_Server\CFESICOFA20210362712.pfx"
    Private mstrRutaTA As String = "C:\SiCoFa_Server\TA.txt"
    Private XmlLoginTicketRequest As XmlDocument = Nothing
    Private ReadOnly URLWsaa As String = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms" 'URL WSAA Homologacion
    'Private ReadOnly URLWsaa As String = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms?WSDL" 'URL WSAA Produccion
    Private ReadOnly XmlStrLoginTicketRequestTemplate As String = "<loginTicketRequest><header><uniqueId></uniqueId><generationTime></generationTime><expirationTime></expirationTime></header><service></service></loginTicketRequest>"
    Private Shared _globalUniqueID As UInt32 = 0 ' OJO! NO ES THREAD-SAFE

    Public Sub New()
        Const ID_FNC As String = "[New LoginTicket]"

        'Estos datos todavia no entiendo
        '-----------------------------------------------
        '-----------------------------------------------

        '-----------------------------------------------
        '-----------------------------------------------

        Dim strTicketRespuesta As String

        Try
            If VerificarTA() = False Then
                strTicketRespuesta = ObtenerLoginTicketResponse()
            End If

        Catch excepcionAlObtenerTicket As Exception
            MsgBox(ID_FNC + "***EXCEPCION AL OBTENER TICKET: " + excepcionAlObtenerTicket.Message)

        End Try

    End Sub

    Private Function VerificarTA() As Boolean

        Const ID_FNC As String = "[VerificarTA]"

        Try

            If File.Exists(mstrRutaTA) = False Then
                Return False
            End If

            Dim TA As String
            Dim xmlTA As New XmlDocument
            Dim ET As DateTime 'ExpirationTime del TA 

            TA = File.ReadAllText(mstrRutaTA)
            xmlTA.LoadXml(TA)

            ET = DateTime.Parse(xmlTA.SelectSingleNode("//expirationTime").InnerText)

            If ET > Now Then
                Me.UniqueId = UInt32.Parse(xmlTA.SelectSingleNode("//uniqueId").InnerText)
                Me.GenerationTime = DateTime.Parse(xmlTA.SelectSingleNode("//generationTime").InnerText)
                Me.ExpirationTime = DateTime.Parse(xmlTA.SelectSingleNode("//expirationTime").InnerText)
                Me.Sign = xmlTA.SelectSingleNode("//sign").InnerText
                Me.Token = xmlTA.SelectSingleNode("//token").InnerText
                CertFirmante = CertificadosX509Lib.ObtieneCertificadoDesdeArchivo(mstrRutaCertX509Firmante, mstrPasswordSecureString)
                Return True
            Else
                Return False
            End If

        Catch excepcionAlVerificarTA As Exception
            Throw New Exception(ID_FNC + "***Error Verificando TA: " + excepcionAlVerificarTA.Message)
            Return False
        End Try

    End Function

    ''' <summary>
    ''' Construye un Login Ticket obtenido del WSAA
    ''' </summary>
    ''' <remarks>
    ''' Función que solicita un Ticket de autorización al WSAA 
    ''' </remarks>


    Private Function ObtenerLoginTicketResponse() As Integer

        Const ID_FNC As String = "[ObtenerLoginTicketResponse]"

        Dim cmsFirmadoBase64 As String
        Dim loginTicketResponse As String
        Dim xmlNodoUniqueId As XmlNode
        Dim xmlNodoGenerationTime As XmlNode
        Dim xmlNodoExpirationTime As XmlNode
        Dim xmlNodoService As XmlNode

        ' PASO 1: Genero el Login Ticket Request
        Try
            _globalUniqueID += 1

            XmlLoginTicketRequest = New XmlDocument()
            XmlLoginTicketRequest.LoadXml(XmlStrLoginTicketRequestTemplate)

            xmlNodoUniqueId = XmlLoginTicketRequest.SelectSingleNode("//uniqueId")
            xmlNodoGenerationTime = XmlLoginTicketRequest.SelectSingleNode("//generationTime")
            xmlNodoExpirationTime = XmlLoginTicketRequest.SelectSingleNode("//expirationTime")
            xmlNodoService = XmlLoginTicketRequest.SelectSingleNode("//service")
            xmlNodoGenerationTime.InnerText = DateTime.Now.AddMinutes(-10).ToString("s")
            xmlNodoExpirationTime.InnerText = DateTime.Now.AddMinutes(+10).ToString("s")
            xmlNodoUniqueId.InnerText = CStr(_globalUniqueID)
            xmlNodoService.InnerText = mstrServicio

        Catch excepcionAlGenerarLoginTicketRequest As Exception
            Throw New Exception(ID_FNC + "***Error GENERANDO el LoginTicketRequest : " + excepcionAlGenerarLoginTicketRequest.Message + excepcionAlGenerarLoginTicketRequest.StackTrace)
        End Try

        ' PASO 2: Firmo el Login Ticket Request
        Try

            CertFirmante = CertificadosX509Lib.ObtieneCertificadoDesdeArchivo(mstrRutaCertX509Firmante, mstrPasswordSecureString)

            ' Convierto el login ticket request a bytes, firmo el msg y convierto a Base64
            Dim EncodedMsg As Encoding = Encoding.UTF8
            Dim msgBytes As Byte() = EncodedMsg.GetBytes(XmlLoginTicketRequest.OuterXml)
            Dim encodedSignedCms As Byte() = CertificadosX509Lib.FirmaBytesMensaje(msgBytes, CertFirmante)
            cmsFirmadoBase64 = Convert.ToBase64String(encodedSignedCms)

        Catch excepcionAlFirmar As Exception
            Throw New Exception(ID_FNC + "***Error FIRMANDO el LoginTicketRequest: " + excepcionAlFirmar.Message)
        End Try

        ' PASO 3: Invoco al WSAA para obtener el Login Ticket Response
        Try

            Dim servicioWsaa As New WSAA.LoginCMSService()
            servicioWsaa.Url = URLWsaa

            ' Veo si hay que salir a traves de un proxy
            If mstrProxy IsNot Nothing Then
                servicioWsaa.Proxy = New WebProxy(mstrProxy, True)
                If mstrProxyUser IsNot Nothing Then
                    Dim Credentials As New NetworkCredential(mstrProxyUser, mstrProxyPassword)
                    servicioWsaa.Proxy.Credentials = Credentials
                End If
            End If

            loginTicketResponse = servicioWsaa.loginCms(cmsFirmadoBase64)
            File.WriteAllText(mstrRutaTA, loginTicketResponse)

        Catch excepcionAlInvocarWsaa As Exception
            Throw New Exception(ID_FNC + "***Error INVOCANDO al servicio WSAA: " + excepcionAlInvocarWsaa.Message)

        End Try

        ' PASO 4: Analizo el Login Ticket Response recibido del WSAA
        Try
            XmlLoginTicketResponse = New XmlDocument()
            XmlLoginTicketResponse.LoadXml(loginTicketResponse)

            Me.UniqueId = UInt32.Parse(XmlLoginTicketResponse.SelectSingleNode("//uniqueId").InnerText)
            Me.GenerationTime = DateTime.Parse(XmlLoginTicketResponse.SelectSingleNode("//generationTime").InnerText)
            Me.ExpirationTime = DateTime.Parse(XmlLoginTicketResponse.SelectSingleNode("//expirationTime").InnerText)
            Me.Sign = XmlLoginTicketResponse.SelectSingleNode("//sign").InnerText
            Me.Token = XmlLoginTicketResponse.SelectSingleNode("//token").InnerText

        Catch excepcionAlAnalizarLoginTicketResponse As Exception

            Throw New Exception(ID_FNC + "***Error ANALIZANDO el LoginTicketResponse: " + excepcionAlAnalizarLoginTicketResponse.Message)
        End Try

        Return 0

    End Function

End Class



'

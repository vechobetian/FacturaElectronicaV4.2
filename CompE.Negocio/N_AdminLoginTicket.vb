Imports System.IO
Imports System.Net
Imports System.Security
Imports System.Security.Cryptography.Pkcs
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports System.Xml
Imports SiCoFa.Entidades
Public Class N_AdminLoginTicket
    Property CertFirmante As X509Certificate2
    Property TicketAcceso As LoginTicket

    Private mstrPathServidor As String = ParamTerminal.ReadINI("C:\SiCoFa_Cliente\config.ini", "SiCoFa", "strPathS")
    Private XmlLoginTicketResponse As XmlDocument = Nothing
    Private mstrPasswordSecureString As New SecureString
    Private mstrProxy As String = Nothing
    Private mstrProxyUser As String = Nothing
    Private mstrProxyPassword As String = ""
    Private mstrServicio As String = "wsfe"
    Private mstrRutaCertX509Firmante As String
    Private mstrRutaTA As String = mstrPathServidor & "\SiCoFa_Server\CompE\TA.txt"
    Private XmlLoginTicketRequest As XmlDocument = Nothing
    Private URLWsaa As String
    Private ReadOnly XmlStrLoginTicketRequestTemplate As String = "<loginTicketRequest><header><uniqueId></uniqueId><generationTime></generationTime><expirationTime></expirationTime></header><service></service></loginTicketRequest>"
    Private Shared _globalUniqueID As UInt32 = 0 ' OJO! NO ES THREAD-SAFE
    Public Function AccesoAlWSN(ByVal argEmpCUIT As String) As Boolean

        If argEmpCUIT = "20210362712" Then
            URLWsaa = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms" 'URL WSAA Homologacion      
        Else
            URLWsaa = "https://wsaa.afip.gov.ar/ws/services/LoginCms?WSDL" 'URL WSAA Produccion 
        End If

        If File.Exists(mstrPathServidor & "SiCoFa_Server\CompE\CFESICOFA" & Replace(argEmpCUIT, "-", "") & ".pfx") Then
            mstrRutaCertX509Firmante = mstrPathServidor & "SiCoFa_Server\CompE\CFESICOFA" & Replace(argEmpCUIT, "-", "") & ".pfx"
        Else
            Throw New Exception(vecho.MensajeError(Me.ToString, "AccesoAlWSN", "No existe certificado digital"))
        End If

        Try
            If VerificarTA() = False Then
                If ObtenerLoginTicket() = False Then
                    Return False
                End If
            End If
            Return True

        Catch ex As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "AccesoALWSN", ex.Message))
            Return False
        End Try

    End Function
    Private Function VerificarTA() As Boolean

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

            If ET > Now Then 'el tiket no esta vencido y leo el archivo
                TicketAcceso = New LoginTicket
                With TicketAcceso
                    .UniqueId = UInt32.Parse(xmlTA.SelectSingleNode("//uniqueId").InnerText)
                    .GenerationTime = DateTime.Parse(xmlTA.SelectSingleNode("//generationTime").InnerText)
                    .ExpirationTime = DateTime.Parse(xmlTA.SelectSingleNode("//expirationTime").InnerText)
                    .Sign = xmlTA.SelectSingleNode("//sign").InnerText
                    .Token = xmlTA.SelectSingleNode("//token").InnerText
                End With
                CertFirmante = ObtenerCertificadoDesdeArchivo(mstrRutaCertX509Firmante, mstrPasswordSecureString)
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "VerificarTA", ex.Message))
            Return False
        End Try

    End Function
    Private Function ObtenerLoginTicket() As Boolean

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
            Throw New Exception(vecho.MensajeError(Me.ToString, "ObtenerLoginTicket PASO 1", excepcionAlGenerarLoginTicketRequest.Message))
        End Try

        ' PASO 2: Firmo el Login Ticket Request
        Try

            CertFirmante = ObtenerCertificadoDesdeArchivo(mstrRutaCertX509Firmante, mstrPasswordSecureString)

            ' Convierto el login ticket request a bytes, firmo el msg y convierto a Base64
            Dim EncodedMsg As Encoding = Encoding.UTF8
            Dim msgBytes As Byte() = EncodedMsg.GetBytes(XmlLoginTicketRequest.OuterXml)
            Dim encodedSignedCms As Byte() = FirmarBytesMensaje(msgBytes, CertFirmante)
            cmsFirmadoBase64 = Convert.ToBase64String(encodedSignedCms)

        Catch excepcionAlFirmar As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "ObtenerLoginTicket PASO 2", excepcionAlFirmar.Message))
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
            Throw New Exception(vecho.MensajeError(Me.ToString, "ObtenerLoginTicket PASO 3", excepcionAlInvocarWsaa.Message))

        End Try

        ' PASO 4: Analizo el Login Ticket Response recibido del WSAA
        Try
            XmlLoginTicketResponse = New XmlDocument()
            XmlLoginTicketResponse.LoadXml(loginTicketResponse)

            TicketAcceso = New LoginTicket
            With TicketAcceso
                .UniqueId = UInt32.Parse(XmlLoginTicketResponse.SelectSingleNode("//uniqueId").InnerText)
                .GenerationTime = DateTime.Parse(XmlLoginTicketResponse.SelectSingleNode("//generationTime").InnerText)
                .ExpirationTime = DateTime.Parse(XmlLoginTicketResponse.SelectSingleNode("//expirationTime").InnerText)
                .Sign = XmlLoginTicketResponse.SelectSingleNode("//sign").InnerText
                .Token = XmlLoginTicketResponse.SelectSingleNode("//token").InnerText
            End With

        Catch excepcionAlAnalizarLoginTicketResponse As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "ObtenerLoginTicket PASO 4", excepcionAlAnalizarLoginTicketResponse.Message))

        End Try

        Return True

    End Function
    Private Function FirmarBytesMensaje(ByVal argBytesMsg As Byte(), ByVal argCertFirmante As X509Certificate2) As Byte()

        Try
            ' Pongo el mensaje en un objeto ContentInfo (requerido para construir el obj SignedCms)
            Dim infoContenido As New ContentInfo(argBytesMsg)
            Dim cmsFirmado As New SignedCms(infoContenido)

            ' Creo objeto CmsSigner que tiene las caracteristicas del firmante
            Dim cmsFirmante As New CmsSigner(argCertFirmante)
            cmsFirmante.IncludeOption = X509IncludeOption.EndCertOnly

            ' Firmo el mensaje PKCS #7
            cmsFirmado.ComputeSignature(cmsFirmante)

            ' Encodeo el mensaje PKCS #7.
            Return cmsFirmado.Encode()
        Catch excepcionAlFirmar As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "FirmarBytesMensaje", excepcionAlFirmar.Message))
            Return Nothing
        End Try
    End Function
    Private Function ObtenerCertificadoDesdeArchivo(ByVal argArchivo As String, ByVal argPassword As SecureString) As X509Certificate2

        Try
            Dim objCert As New X509Certificate2
            If argPassword.IsReadOnly Then
                objCert.Import(File.ReadAllBytes(argArchivo), argPassword, X509KeyStorageFlags.PersistKeySet)
            Else
                objCert.Import(File.ReadAllBytes(argArchivo))
            End If
            Return objCert
        Catch excepcionAlImportarCertificado As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "ObtenerCertificadoDesdeArchivo", excepcionAlImportarCertificado.Message))
            Return Nothing
        End Try
    End Function

End Class
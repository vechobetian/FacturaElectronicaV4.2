Imports System.Security
Imports System.Globalization
Imports SiCoFa_CompE.WSN

Public Class CAE

    'Popiedades de la clase CAE
    Property Observaciones As String
    Property Errores As String
    Property Eventos As String
    Property NroCompAutorizado As Integer
    Property CAE As String
    Property VtoCAE As Date

    'Variables de Modulo
    Private ReadOnly URLWsn As String = "https://wswhomo.afip.gov.ar/wsfev1/service.asmx?WSDL" 'URL WSN Homologacion
    'Private ReadOnly URLWsn as String = "https://servicios1.afip.gov.ar/wsfev1/service.asmx?WSDL" 'URL WSN Produccion
    Private mvarLogin As LoginTicket
    Private mvarAuthRequest As FEAuthRequest
    Private mstrEstadoCab As String
    Private mstrEstadoDet As String

    ' Valores por defecto, globales en esta clase
    Const DEFAULT_PROXY As String = Nothing
    Const DEFAULT_PROXY_USER As String = Nothing
    Const DEFAULT_PROXY_PASSWORD As String = ""

    ''' <summary>
    ''' Funcion ObtenerTA
    ''' </summary>
    ''' <returns>0 si terminó bien, valores negativos si hubieron errores</returns>
    ''' <remarks></remarks>
    Public Function ObtenerTA(
    ByVal argServicio As String,
    ByVal argRutaTA As String,
    ByVal argRutaCertX509Firmante As String
    ) As Integer
        Const ID_FNC As String = "[ObtenerTA]"

        Dim strPasswordSecureString As New SecureString
        Dim strProxy As String = DEFAULT_PROXY
        Dim strProxyUser As String = DEFAULT_PROXY_USER
        Dim strProxyPassword As String = DEFAULT_PROXY_PASSWORD

        ' Argumentos OK, entonces procesar normalmente...
        Dim objTicketRespuesta As LoginTicket
        Dim strTicketRespuesta As String

        Try
            objTicketRespuesta = New LoginTicket

            If objTicketRespuesta.VerificarTA(argRutaTA, argRutaCertX509Firmante, strPasswordSecureString) = False Then
                strTicketRespuesta = objTicketRespuesta.ObtenerLoginTicketResponse(argServicio, argRutaTA, argRutaCertX509Firmante, strPasswordSecureString, strProxy, strProxyUser, strProxyPassword)
            End If

        Catch excepcionAlObtenerTicket As Exception
            MsgBox(ID_FNC + "***EXCEPCION AL OBTENER TICKET: " + excepcionAlObtenerTicket.Message)
            Return -10

        End Try
        mvarLogin = objTicketRespuesta
        Return 0
    End Function

    Public Function SolicitarCAE(
    ByVal argMiCUIT As String,
    ByVal argConcepto As Integer,
    ByVal argPtoVenta As Integer,
    ByVal argCbteTipo As Integer,
    ByVal argDocTipo As Integer,
    ByVal argDocNro As String,
    ByVal argImpEx As Double,
    ByVal argImpGrav As Double
    ) As String

        Const ID_FNC As String = "[SolicitarCAE]"
        Dim strRetorno As String = "A"

        Try
            mvarAuthRequest = New FEAuthRequest()
            mvarAuthRequest.Cuit = argMiCUIT
            mvarAuthRequest.Sign = mvarLogin.Sign
            mvarAuthRequest.Token = mvarLogin.Token

            Dim service As WSN.Service = getServicio()
            service.ClientCertificates.Add(mvarLogin.CertFirmante)

            Dim req As New FECAERequest
            Dim cab As New FECAECabRequest
            Dim det As New FECAEDetRequest

            cab.CantReg = 1
            cab.PtoVta = argPtoVenta
            cab.CbteTipo = argCbteTipo
            req.FeCabReq = cab

            With det
                .Concepto = argConcepto
                .DocTipo = argDocTipo
                .DocNro = argDocNro

                Dim lastRes As FERecuperaLastCbteResponse = service.FECompUltimoAutorizado(mvarAuthRequest, argPtoVenta, argCbteTipo)
                Dim last As Integer = lastRes.CbteNro

                NroCompAutorizado = last + 1

                .CbteDesde = last + 1
                .CbteHasta = last + 1
                .CbteFch = Now.ToString("yyyyMMdd")

                If argCbteTipo = 11 Then
                    .ImpTotConc = 0
                    .ImpNeto = argImpEx + argImpGrav
                    .ImpIVA = 0
                    .ImpTotal = argImpEx + argImpGrav
                Else
                    Dim ImpG As Decimal = Math.Round(argImpGrav / 1.21, 2, MidpointRounding.ToEven)
                    Dim IVA As Decimal = Math.Round(ImpG * 21 / 100, 2, MidpointRounding.ToEven)
                    Dim alicuota As New AlicIva

                    .ImpTotConc = argImpEx
                    .ImpNeto = ImpG
                    .ImpIVA = IVA
                    .ImpTotal = ImpG + IVA + argImpEx

                    alicuota.Id = 5
                    alicuota.BaseImp = ImpG
                    alicuota.Importe = IVA
                    .Iva = {alicuota}
                End If

                .ImpOpEx = 0
                .ImpTrib = 0
                .MonId = "PES"
                .MonCotiz = 1
            End With

            req.FeDetReq = {det}

            Dim r = service.FECAESolicitar(mvarAuthRequest, req)

            mstrEstadoCab = r.FeCabResp.Resultado
            mstrEstadoDet = r.FeDetResp(0).Resultado

            If mstrEstadoCab = "A" And mstrEstadoDet = "A" Then
                CAE = r.FeDetResp(0).CAE
                VtoCAE = DateTime.ParseExact(r.FeDetResp(0).CAEFchVto, "yyyyMMdd", CultureInfo.InvariantCulture)
            Else
                strRetorno = "R"
            End If

            If r.FeDetResp(0).Observaciones IsNot Nothing Then
                For Each o In r.FeDetResp(0).Observaciones
                    Observaciones &= String.Format("Obs: {0} ({1})", o.Msg, o.Code) & vbCrLf
                Next
            End If

            If r.Errors IsNot Nothing Then
                strRetorno = "E"
                For Each er In r.Errors
                    Errores &= String.Format("Er: {0}: {1}", er.Code, er.Msg) & vbCrLf
                Next

            End If

            If r.Events IsNot Nothing Then
                For Each ev In r.Events
                    Eventos &= String.Format("Ev: {0}: {1}", ev.Code, ev.Msg) & vbCrLf
                Next
            End If

        Catch excepcionAlSolicitarCAE As Exception
            Throw New Exception(ID_FNC + "***Error Solicitando CAE: " + excepcionAlSolicitarCAE.Message)
            Return False

        End Try

        Return strRetorno

    End Function
    Private Function getServicio() As Service
        Dim s As New Service
        s.Url = URLWsn
        Return s
    End Function

End Class

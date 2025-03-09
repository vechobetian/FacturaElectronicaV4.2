Imports System.Globalization
Imports SiCoFa.Entidades
Imports SiCoFa.Negocio.WSN
Public Class N_AdminCAE
    Property Observaciones As String
    Property Errores As String
    Property Eventos As String
    Private URLWsn As String
    Private mvarLogin As LoginTicket
    Private mvarAuthRequest As FEAuthRequest
    Private mstrEstadoCab As String
    Private mstrEstadoDet As String
    Public Function ObtenerCAE(ByVal argComprobante As Comprobante) As CAE

        If argComprobante.Empresa.CUIT = "20210362712" Then
            URLWsn = "https://wswhomo.afip.gov.ar/wsfev1/service.asmx?WSDL" 'URL WSN Homologacion

        Else
            URLWsn = "https://servicios1.afip.gov.ar/wsfev1/service.asmx?WSDL" 'URL WSN Produccion

        End If

        Dim objN_AdminLT As New N_AdminLoginTicket
        Dim objCAE As CAE = Nothing
        Dim NroCbteAutorizado As Long

        Try

            If objN_AdminLT.AccesoAlWSN(argComprobante.Empresa.CUIT) = False Then
                Me.Observaciones = "No se pudo Obtener Ticket de Acceso"
                Return Nothing
            End If

            mvarLogin = objN_AdminLT.TicketAcceso
            mvarAuthRequest = New FEAuthRequest()
            mvarAuthRequest.Cuit = Replace(argComprobante.Empresa.CUIT, "-", "")
            mvarAuthRequest.Sign = mvarLogin.Sign
            mvarAuthRequest.Token = mvarLogin.Token

            Dim service As WSN.Service = getServicio()
            service.ClientCertificates.Add(objN_AdminLT.CertFirmante)

            Dim req As New FECAERequest
            Dim cab As New FECAECabRequest
            Dim det As New FECAEDetRequest

            cab.CantReg = 1
            cab.PtoVta = argComprobante.PuntoVta
            cab.CbteTipo = argComprobante.TipoComprobante.CodiTC_AFIP

            req.FeCabReq = cab

            With det
                .Concepto = 1
                .DocTipo = argComprobante.Cliente.CodiTD_AFIP
                .DocNro = argComprobante.Cliente.NumDoc

                Dim lastRes As FERecuperaLastCbteResponse = service.FECompUltimoAutorizado(mvarAuthRequest, CInt(argComprobante.PuntoVta), argComprobante.TipoComprobante.CodiTC_AFIP)
                Dim last As Integer = lastRes.CbteNro

                NroCbteAutorizado = last + 1

                .CbteDesde = last + 1
                .CbteHasta = last + 1
                .CbteFch = Now.ToString("yyyyMMdd")

                If argComprobante.TipoComprobante.TipoComprobante = "NOTA DE CREDITO" Then
                    Dim cbteAsoc As New CbteAsoc
                    With cbteAsoc
                        .Tipo = argComprobante.CompAsoc.TipoComprobante.CodiTC_AFIP
                        .PtoVta = argComprobante.CompAsoc.PuntoVta
                        .Nro = argComprobante.CompAsoc.NumComp
                        '.Cuit = Replace(cbteOrigen.Empresa.CUIT, "-", "") por ahora no requerido
                        '.CbteFch = cbteOrigen.FechaComp por ahora no requerido
                    End With
                    .CbtesAsoc = {cbteAsoc}
                End If

                If argComprobante.TipoComprobante.CodiTC_AFIP = 11 Or argComprobante.TipoComprobante.CodiTC_AFIP = 13 Then
                    .ImpTotConc = 0
                    .ImpOpEx = 0
                    .ImpNeto = argComprobante.ImpEx + argComprobante.ImpGrav
                    .ImpIVA = 0
                    .ImpTotal = argComprobante.ImpEx + argComprobante.ImpGrav
                Else
                    .ImpTotConc = 0
                    .ImpOpEx = argComprobante.ImpEx
                    .ImpNeto = argComprobante.ImpNeto
                    .ImpIVA = argComprobante.IVA
                    .ImpTotal = argComprobante.ImpNeto + argComprobante.IVA + argComprobante.ImpEx

                    Dim alicuota As New AlicIva

                    If argComprobante.IVA > 0 Then
                        alicuota.Id = 5
                        alicuota.BaseImp = argComprobante.ImpNeto
                        alicuota.Importe = argComprobante.IVA
                        .Iva = {alicuota}
                    End If

                End If

                Dim IdIvaRec = Me.IdIVAReceptor(argComprobante.Cliente.IVA)

                .ImpTrib = 0
                .MonId = "PES"
                '.CanMisMonExt = "S"
                .MonCotizSpecified = True
                .MonCotiz = 1
                .CondicionIVAReceptorId = IdIvaRec
            End With

            req.FeDetReq = {det}
            service.Timeout = 30000
            Dim r = service.FECAESolicitar(mvarAuthRequest, req)

            mstrEstadoCab = r.FeCabResp.Resultado
            mstrEstadoDet = r.FeDetResp(0).Resultado

            If mstrEstadoCab = "A" And mstrEstadoDet = "A" Then
                objCAE = New CAE(
                                NroCbteAutorizado,
                                r.FeDetResp(0).CAE,
                                DateTime.ParseExact(r.FeDetResp(0).CAEFchVto, "yyyyMMdd", CultureInfo.InvariantCulture)
                                )

                Return objCAE
                Exit Function

            End If

            If r.FeDetResp(0).Observaciones IsNot Nothing Then
                For Each o In r.FeDetResp(0).Observaciones
                    Observaciones &= String.Format("{0} ({1})", o.Msg, o.Code) & vbCrLf
                Next
                'MsgBox(Observaciones, vbInformation, "CompE")
                Throw New Exception(Vecho.MensajeError(Me.ToString, "ObtenerCAE", Observaciones))

            End If

            If r.Errors IsNot Nothing Then

                For Each er In r.Errors
                    Errores &= String.Format("{0}: {1}", er.Code, er.Msg) & vbCrLf
                Next
                Throw New Exception(Vecho.MensajeError(Me.ToString, "ObtenerCAE", Errores))

            End If

            If r.Events IsNot Nothing Then
                For Each ev In r.Events
                    Eventos &= String.Format("{0}: {1}", ev.Code, ev.Msg) & vbCrLf
                Next
                Throw New Exception(Vecho.MensajeError(Me.ToString, "ObtenerCAE", Eventos))

            End If

        Catch ex As Exception
            If ex.HResult <> -2146233079 Then
                Throw New Exception(vecho.MensajeError(Me.ToString, "ObtenerCAE", ex.Message))
            End If
        End Try

    End Function
    Private Function getServicio() As Service
        Dim s As New Service
        s.Url = URLWsn
        Return s
    End Function
    Private Function IdIVAReceptor(IVA As String) As Integer
        Select Case IVA
            Case "RI"
                Return 1
            Case "MT"
                Return 6
            Case "EX"
                Return 4
            Case "CF"
                Return 5
            Case "SI" 'sin identificar
                Return 15
        End Select


        '"Id": 1,
        '"Desc": "IVA Responsable Inscripto",
        '"Cmp_Clase": "A/M/C"    

        '"Id": 6,
        '"Desc": "Responsable Monotributo",
        '"Cmp_Clase": "A/M/C"    

        '"Id": 13,
        '"Desc": "Monotributista Social",
        '"Cmp_Clase": "A/M/C"    

        '"Id": 16,
        '"Desc": "Monotributo Trabajador Independiente Promovido",
        '"Cmp_Clase": "A/M/C"    

        '"Id": 4,
        '"Desc": "IVA Sujeto Exento",
        '"Cmp_Clase": "B/C"    

        '"Id": 5,
        '"Desc": "Consumidor Final",
        '"Cmp_Clase": "B/C"    

        '"Id": 7,
        '"Desc": "Sujeto No Categorizado",
        '"Cmp_Clase": "B/C"    

        '"Id": 8,
        '"Desc": "Proveedor del Exterior",
        '"Cmp_Clase": "B/C"    

        '"Id": 9,
        '"Desc": "Cliente del Exterior",
        '"Cmp_Clase": "B/C"    

        '"Id": 10,
        '"Desc": "IVA Liberado – Ley N° 19.640",
        '"Cmp_Clase": "B/C"    

        '"Id": 15,
        '"Desc": "IVA No Alcanzado",
        '"Cmp_Clase": "B/C"

    End Function
End Class
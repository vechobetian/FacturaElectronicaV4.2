Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Web.Script.Serialization
Imports SiCoFa.Entidades

Public Class N_AdminIMED

    ' Endpoints oficiales para Homologación y Producción
    Private Const URL_HOMO_TOKEN As String = "https://homo.imed-sigedad.com/token"
    Private Const URL_PROD_TOKEN As String = "https://ingest.imed-sigedad.com/token"

    Private Const URL_HOMO_INGEST As String = "https://homo.imed-sigedad.com/v2/ingest"
    Private Const URL_PROD_INGEST As String = "https://ingest.imed-sigedad.com/v2/ingest"

    ''' <summary>
    ''' Obtiene el Token JWT desde la API de IMED usando el GLN y la Clave de la farmacia.
    ''' </summary>
    Private Function ObtenerTokenJwt(ByVal gln As String, ByVal clave As String, ByVal esHomologacion As Boolean) As String
        Dim urlToken As String = If(esHomologacion, URL_HOMO_TOKEN, URL_PROD_TOKEN)

        Using client As New HttpClient()
            client.DefaultRequestHeaders.Clear()
            client.DefaultRequestHeaders.Add("x-farma-id", gln)
            client.DefaultRequestHeaders.Add("x-farma-key", clave)
            client.Timeout = TimeSpan.FromSeconds(15)

            Dim response = client.GetAsync(urlToken).Result

            If response.IsSuccessStatusCode Then
                Dim tokenJwt As String = response.Content.ReadAsStringAsync().Result
                Return tokenJwt.Trim().Replace("""", "") ' Limpiamos comillas si el JSON viene plano
            Else
                Throw New ErrorImedException("Error al autenticar en IMED (Código: " & CInt(response.StatusCode) & ")")
            End If
        End Using
    End Function

    Public Function ProcesarDeclaracionDispensa(ByVal objOpera As Operacion, ByVal cbte As Comprobante) As Boolean
        Try
            Dim estaHabilitado As Boolean = True
            If Not estaHabilitado Then Return True

            ' Definir entorno (True para Homologación, False para Producción)
            Dim esHomologacion As Boolean = True
            Dim urlImed As String = If(esHomologacion, URL_HOMO_INGEST, URL_PROD_INGEST)

            ' 1. Obtener credenciales de la farmacia desde el objeto o configuración
            Dim glnFarmacia As String = cbte.Empresa.GLN.ToString() ' O ajusta según tu entidad
            Dim claveFarmacia As String = "CLAVE_DE_FARMACIA" ' Cárgalo desde tu tabla de parámetros de la farmacia

            ' 2. Solicitar el Token JWT dinámicamente
            Dim tokenImed As String = ObtenerTokenJwt(glnFarmacia, claveFarmacia, esHomologacion)

            If String.IsNullOrEmpty(tokenImed) Then
                Throw New Exception("No se pudo obtener un Token JWT válido de IMED.")
            End If

            ' 3. Formatear fechas 
            Dim fechaCbteStr As String = cbte.FechaComp.ToString("yyyyMMdd")
            Dim horaCbteStr As String = DateTime.Now.ToString("HHmmss")

            ' 4. Armar el Encabezado por separado
            Dim miEncabezado As New EncabezadoMensajeData() With {
                .TipoMsj = 200,
                .CodAccion = 290420,
                .IdMsj = 1,
                .InicioTrx = New InicioTrxData() With {
                    .Fecha = fechaCbteStr,
                    .Hora = horaCbteStr
                },
                .Software = New SoftwareData() With {
                    .CodigoADESFA = "",
                    .Nombre = "SiCoFa SRL",
                    .Version = "1"
                },
                .Prestador = New PrestadorData() With {
                    .GLN = If(cbte.Empresa.GLN IsNot Nothing, cbte.Empresa.GLN.ToString().Trim(), ""),
                    .CodigoImed = 98734563271L,
                    .Cuit = If(IsNumeric(cbte.Empresa.CUIT), Convert.ToInt64(cbte.Empresa.CUIT.Replace("-", "")), 0L),
                    .CUFE = If(cbte.Empresa.CUFE IsNot Nothing, cbte.Empresa.CUFE.ToString().Trim(), "")
                }
            }

            ' 5. Armar el Comprobante usando directamente las propiedades
            Dim miComprobante As New ComprobanteData() With {
                .Id = objOpera.IdOperacion,
                .PtoVenta = If(Not String.IsNullOrEmpty(cbte.PuntoVta), Convert.ToInt32(cbte.PuntoVta), 0),
                .TipoComprobanteARCA = If(cbte.TipoComprobante IsNot Nothing, Convert.ToInt32(cbte.TipoComprobante.TipoComprobante), 0),
                .Fecha = fechaCbteStr,
                .Hora = horaCbteStr,
                .Tipo = "V",
                .Fiscal = "E",
                .CAE = If(cbte.CAE IsNot Nothing AndAlso Not String.IsNullOrEmpty(cbte.CAE.NumCAE), Convert.ToInt64(cbte.CAE.NumCAE), 0L),
                .ImporteTotal = cbte.ImpBto,
                .TipoDocumento = cbte.Cliente.CodiTD_AFIP,
                .NroDocumento = If(IsNumeric(cbte.Cliente.NumDoc), Convert.ToInt64(cbte.Cliente.NumDoc), 0L),
                .PersonaRetiro = New PersonaRetiroData() With {
                    .TipoDocumentoR = cbte.Cliente.CodiTD_AFIP,
                    .NroDocumentoR = If(IsNumeric(cbte.Cliente.NumDoc), Convert.ToInt64(cbte.Cliente.NumDoc), 0L),
                    .FechaNacimientoR = "19850412",
                    .SexoR = "F"
                }
            }

            ' 6. Recorrer el detalle de la venta (cbte.Detalle) para armar la lista DatosVentas
            Dim listaVentas As New List(Of VentaItemData)()
            If cbte.Detalle IsNot Nothing Then
                Dim nroOrdenItem As Integer = 1
                For Each det In cbte.Detalle
                    Dim itemVentaData As New VentaItemData() With {
                        .Item = New ItemContenidoData() With {
                            .NroOrden = nroOrdenItem,
                            .ItemVenta = New ItemVentaData() With {
                                .OrdenItemVenta = nroOrdenItem,
                                .NroReferenciaSGF = objOpera.IdOperacion.ToString(),
                                .CodBarras = 0L,
                                .CapturaCodBarras = "E",
                                .CodTroquel = 0L,
                                .Alfabeta = 0,
                                .Descripcion = "Artículo de Venta",
                                .CantidadVenta = 1,
                                .ImporteUnitario = 0D
                            },
                            .Validador = New ValidadorData() With {
                                .CodigoADESFA = "A0030",
                                .Nombre = "FARMALINK",
                                .NroReferencia = "",
                                .CantidadAprobada = 1
                            }
                        }
                    }
                    listaVentas.Add(itemVentaData)
                    nroOrdenItem += 1
                Next
            End If

            ' 7. Unir todo en el payload principal
            Dim payload As New ImedPayload()
            payload.MensajeImedSigedad = New MensajeImedSigedadData() With {
                .MensajeImedSigedadVersion = "3.2.0",
                .EncabezadoMensaje = miEncabezado,
                .Comprobante = miComprobante,
                .DatosVentas = listaVentas
            }

            ' 8. Serialización a JSON
            Dim js As New JavaScriptSerializer()
            Dim jsonPayload As String = js.Serialize(payload)

            ' 9. Enviar por HTTP POST a IMED utilizando el Token JWT obtenido
            Using client As New HttpClient()
                client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", tokenImed)
                client.Timeout = TimeSpan.FromSeconds(20)

                Dim content As New StringContent(jsonPayload, Encoding.UTF8, "application/json")
                Dim response = client.PostAsync(urlImed, content).Result
                Dim responseString = response.Content.ReadAsStringAsync().Result

                If response.IsSuccessStatusCode Then
                    Return True
                Else
                    RegistrarPendienteLocal(objOpera.IdOperacion, jsonPayload, responseString)
                    Return False
                End If
            End Using

        Catch ex As Exception
            RegistrarErrorLocal(objOpera.IdOperacion, ex.Message)
            Return False
        End Try
    End Function

    Private Sub RegistrarPendienteLocal(ByVal idOperacion As Long, ByVal jsonTrama As String, ByVal mensajeError As String)
    End Sub

    Private Sub RegistrarErrorLocal(ByVal idOperacion As Long, ByVal mensajeError As String)
    End Sub

End Class

' Excepción personalizada para evitar filtrar claves en logs de error
Public Class ErrorImedException
    Inherits Exception
    Public Sub New(message As String)
        MyBase.New(message)
    End Sub
End Class
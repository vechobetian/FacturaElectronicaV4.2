Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Web.Script.Serialization
Imports SiCoFa.Datos
Imports SiCoFa.Entidades

Public Class N_AdminIMED

    Private Const ARCHIVO_TOKEN As String = "imed_token.txt"
    Private m_AdminDB As New D_AdminDB
    Private m_NumPrestador As String

    Public Sub ProcesarDeclaracionDispensa(ByVal argOpera As Operacion, ByVal argCbte As Comprobante)
        Try
            Dim filtros As New Dictionary(Of String, Object) From {
                                                                    {"CodiAut", "SIGEDAD"}
                                                                   }

            Dim parametros = m_AdminDB.ObtenerRegistro("TblAutorizadores",, filtros)
            Dim claveImed As String = parametros("Licencia").ToString
            m_NumPrestador = parametros("NumPrestador").ToString

            ' 1. Definir entorno (Homologación o Producción según tu lógica de negocio)
            Dim esHomologacion As Boolean = True ' Cambiar a False cuando pasen a producción

            ' 2. Obtener credenciales desde la empresa u operación
            Dim glnFarmacia As String = argCbte.Empresa.GLN

            ' 3. Obtener el Token JWT vigente (reutiliza el txt o pide uno nuevo)
            Dim token As String = ObtenerTokenVigente(glnFarmacia, claveImed, esHomologacion)

            ' 4. Armar el payload con los datos de la receta/dispensa
            Dim jsonPayload As String = ConstruirJsonDispensa(argOpera, argCbte)

            ' 5. Enviar la declaración a la API de IMED
            EnviarDeclaracionAlServidor(jsonPayload, token, esHomologacion)

        Catch ex As Exception
            Throw New Exception("Error en dispensa IMED: " & ex.Message)
        End Try
    End Sub

    Private Function ObtenerTokenVigente(ByVal gln As String, ByVal clave As String, ByVal esHomologacion As Boolean) As String
        Dim rutaToken As String = ObtenerRutaToken()

        ' 1. Verificar si el archivo del token existe en el servidor
        If File.Exists(rutaToken) Then
            ' Margen de seguridad de 7 horas (el token dura 8)
            Dim fechaCreacion As DateTime = File.GetLastWriteTime(rutaToken)
            If DateTime.Now.Subtract(fechaCreacion).TotalHours < 7.0 Then
                Dim tokenGuardado As String = File.ReadAllText(rutaToken).Trim()
                If Not String.IsNullOrEmpty(tokenGuardado) Then
                    Return tokenGuardado
                End If
            End If
        End If

        ' 2. Si no existe o expiró, lo solicitamos a la API de IMED
        Dim nuevoToken As String = SolicitarTokenNuevoAlServidor(gln, clave, esHomologacion)

        ' 3. Guardamos el nuevo token en el archivo txt del servidor
        Dim carpetaDestino As String = Path.GetDirectoryName(rutaToken)
        If Not Directory.Exists(carpetaDestino) Then
            Directory.CreateDirectory(carpetaDestino)
        End If
        File.WriteAllText(rutaToken, nuevoToken)

        Return nuevoToken
    End Function

    Private Function ObtenerRutaToken() As String
        ' Utiliza la función centralizada de Vecho para leer la ruta desde la RAM
        Dim pathServidor As String = Vecho.ObtenerPathServidor()
        Return Path.Combine(pathServidor, "SiCoFa_Server\CompE", ARCHIVO_TOKEN)
    End Function

    Private Function SolicitarTokenNuevoAlServidor(ByVal gln As String, ByVal clave As String, ByVal esHomologacion As Boolean) As String
        Dim urlToken As String = If(esHomologacion, "https://homo.imed-sigedad.com/token", "https://ingest.imed-sigedad.com/token")

        Using client As New HttpClient()
            client.DefaultRequestHeaders.Clear()
            client.DefaultRequestHeaders.Add("x-farma-id", gln)
            client.DefaultRequestHeaders.Add("x-farma-key", clave)
            client.Timeout = TimeSpan.FromSeconds(15)

            Dim response = client.GetAsync(urlToken).Result

            If response.IsSuccessStatusCode Then
                Dim tokenJwt As String = response.Content.ReadAsStringAsync().Result
                Return tokenJwt.Trim().Replace("""", "")
            Else
                Throw New Exception("Error de autenticación IMED. Código: " & CInt(response.StatusCode))
            End If
        End Using
    End Function

    Private Function ObtenerIdMensaje() As Long

        Dim filtros As New Dictionary(Of String, Object) From {
                                                               {"CodiAut", "SIGEDAD"}
                                                              }

        ' Hace la lectura y la actualización en una sola operación segura
        Return m_AdminDB.IncrementarYObtenerContador("TblAutorizadores", "UltiNumFil", filtros)

    End Function

    Private Function ConstruirJsonDispensa(ByVal argOpera As Operacion, ByVal argCbte As Comprobante) As String
        Dim adminRtas As New D_AdminRtas
        Dim receta As Receta = adminRtas.ObtenerRecetaPorIdOperacion(argOpera.IdOperacion)
        Dim listaDatosVentas As New List(Of Object)
        Dim nroOrden As Integer = 1

        For Each item In argCbte.Detalle

            Dim nodoItem = New With {
                .NroOrden = nroOrden,
                .ItemVenta = New With {
                    .OrdenItemVenta = nroOrden,
                    .NroReferenciaSGF = argCbte.NumComp.ToString(),
                    .NroReferenciaSGFOrigen = Nothing,
                    .CodBarras = Nothing,
                    .CapturaCodBarras = "E",
                    .CodTroquel = Nothing,
                    .Alfabeta = Nothing,
                    .Kairos = Nothing,
                    .Codigo = Nothing,
                    .IdProducto = Nothing,
                    .Descripcion = item.Descripcion,
                    .CantidadVenta = item.Cantidad,
                    .ImporteUnitario = CDec(item.PUnit)
                },
                .RecetaDigital = New With {
                    .Fecha = argCbte.FechaComp.ToString("yyyyMMdd"),
                    .NroRecetaDigital = receta.NumRta,
                    .NroReferenciaConsulta = receta.NumConsulta,
                    .TipoPrescriptor = receta.TipoPrescriptor,
                    .TipoMatricula = receta.TipoMatricula,
                    .ProvinciaMatricula = receta.Provincia,
                    .NroMatricula = receta.NumMatricula,
                    .IdProductoSolicitado = Nothing,
                    .AlfabetaSolicitado = Nothing,
                    .DrogaCodigoSolicitado = Nothing,
                    .CodDrogaSnomedSolicitado = Nothing,
                    .DrogaDescripcionSolicitado = "DESCRIPCION",
                    .CantidadSolicitado = item.Cantidad,
                    .EstadoSugerido = 0,
                    .Financiador = New With {
                        .CodigoADESFA = Nothing,
                        .Codigo = "41"
                    }
                },
                .Validador = New With {
                    .CodigoADESFA = "A0030",
                    .Nombre = "FARMALINK",
                    .NroReferencia = receta.NumAutorizacion, ' <-- ¡Acá faltaba la coma!
                    .CantidadAprobada = item.Cantidad
                }
            }

            listaDatosVentas.Add(New With {.Item = nodoItem})
            nroOrden += 1
        Next

        ' 2. Armamos la estructura principal que envuelve todo
        Dim estructuraPrincipal = New With {
            .MensajeImedSigedad = New With {
                .MensajeImedSigedadVersion = "3.2.0",
                .EncabezadoMensaje = New With {
                    .TipoMsj = 200,
                    .CodAccion = 290420,
                    .IdMsj = ObtenerIdMensaje(),
                    .InicioTrx = New With {
                        .Fecha = DateTime.Now.ToString("yyyyMMdd"),
                        .Hora = DateTime.Now.ToString("HHmmss")
                    },
                    .Software = New With {
                        .CodigoADESFA = "",
                        .Nombre = "SiCoFa",
                        .Version = "4.2.0"
                    },
                    .Prestador = New With {
                        .CodigoADESFA = Nothing,
                        .GLN = argCbte.Empresa.GLN,
                        .CodigoImed = "argCbte.Empresa.ClaveIMED",
                        .Cuit = CLng(argCbte.Empresa.CUIT.Replace("-", "")), ' Asegurar formato numérico
                        .CUFE = argCbte.Empresa.CUFE
                    }
                },
                .Comprobante = New With {
                    .Id = argCbte.NumComp,
                    .PtoVenta = CInt(argCbte.PuntoVta),
                    .TipoComprobanteARCA = CInt(argCbte.TipoComprobante.CodiTC_AFIP),
                    .Fecha = argOpera.Fin.ToString("yyyyMMdd"),
                    .Hora = argOpera.Fin.ToString("HHmmss"),
                    .Tipo = "V",
                    .Fiscal = "E",
                    .CAE = argCbte.CAE.NumCAE,
                    .ImporteTotal = CDec(argCbte.ImpBto),
                    .TipoDocumento = argCbte.Cliente.CodiTD_AFIP,
                    .NroDocumento = CLng(argCbte.Cliente.NumDoc.Replace("-", "")),
                    .PersonaRetiro = New With {
                        .TipoDocumentoR = 96,
                        .NroDocumentoR = 30444555,
                        .FechaNacimientoR = "19850412",
                        .SexoR = "F",
                        .LecturaDNI = "00512345690@PEREZ@MARIA LAURA@F@30444555@B@12/04/1991@15/03/2021@274"
                    } ' <-- Si no hay retiro, podrías igualarlo a Nothing según lo que exija IMED
                },
                .DatosVentas = listaDatosVentas
            }
        }

        ' 3. Convertimos todo a formato JSON
        Dim serializer As New JavaScriptSerializer()
        Return serializer.Serialize(estructuraPrincipal)
    End Function

    Private Sub EnviarDeclaracionAlServidor(ByVal jsonPayload As String, ByVal token As String, ByVal esHomologacion As Boolean)
        Dim urlDispensa As String = If(esHomologacion, "https://homo.imed-sigedad.com/dispensa", "https://ingest.imed-sigedad.com/dispensa")

        Using client As New HttpClient()
            client.DefaultRequestHeaders.Clear()
            client.DefaultRequestHeaders.Authorization = New Headers.AuthenticationHeaderValue("Bearer", token)
            client.Timeout = TimeSpan.FromSeconds(20)

            Dim content As New StringContent(jsonPayload, System.Text.Encoding.UTF8, "application/json")
            Dim response = client.PostAsync(urlDispensa, content).Result

            If Not response.IsSuccessStatusCode Then
                Dim errorDetalle As String = response.Content.ReadAsStringAsync().Result
                Throw New Exception("Rechazo de IMED: " & errorDetalle)
            End If
        End Using
    End Sub

End Class
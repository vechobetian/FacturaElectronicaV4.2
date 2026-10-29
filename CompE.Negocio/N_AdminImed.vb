Imports System.Collections.Generic
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Web.Script.Serialization
Imports SiCoFa.Datos
Imports SiCoFa.Entidades

Public Class N_AdminIMED

    Private Const ARCHIVO_TOKEN As String = "imed_token.txt"
    Private m_AdminDB As New D_AdminDB
    Private m_NumPrestadorSIGEDAD As String

    Public Sub ProcesarDeclaracionDispensa(ByVal argOpera As Operacion, ByVal argCbte As Comprobante)
        Try
            Dim filtros As New Dictionary(Of String, Object) From {
                {"CodiAut", "SIGEDAD"}
            }

            Dim parametros = m_AdminDB.ObtenerRegistro("TblAutorizadores",, filtros)
            Dim claveImed As String = parametros("Licencia").ToString
            m_NumPrestadorSIGEDAD = parametros("NumPrestador").ToString

            ' 1. Definir entorno (Homologación o Producción)
            Dim esHomologacion As Boolean = True ' Cambiar a False en producción

            ' 2. Obtener credenciales desde la empresa u operación
            Dim glnFarmacia As String = argCbte.Empresa.GLN

            ' 3. Obtener el Token JWT vigente (reutiliza el txt o pide uno nuevo)
            Dim token As String = ObtenerTokenVigente(glnFarmacia, claveImed, esHomologacion)

            ' 4. Armar el payload JSON con trazabilidad
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

        ' 3. Guardamos el nuevo token
        Dim carpetaDestino As String = Path.GetDirectoryName(rutaToken)
        If Not Directory.Exists(carpetaDestino) Then
            Directory.CreateDirectory(carpetaDestino)
        End If
        File.WriteAllText(rutaToken, nuevoToken)

        Return nuevoToken
    End Function

    Private Function ObtenerRutaToken() As String
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
        Return m_AdminDB.IncrementarYObtenerContador("TblAutorizadores", "UltiNumFil", filtros)
    End Function

    Private Function ConstruirJsonDispensa(ByVal argOpera As Operacion, ByVal argCbte As Comprobante) As String
        Dim adminRtas As New D_AdminRtas
        Dim receta As Receta = adminRtas.ObtenerRecetaPorIdOperacion(argOpera.IdOperacion)
        Dim listaDatosVentas As New List(Of Object)
        Dim nroOrden As Integer = 1

        For Each item As ItemReceta In receta.ItemsReceta
            Dim alfabetataSolicitado As String = ""

            If item.Estado = 1 Then
                alfabetataSolicitado = item.Alfabeta.ToString()
            End If

            ' 1. Armamos ItemVenta en un Dictionary para poder adjuntar UnidadesTrazables dinámicamente
            Dim itemVentaDict As New Dictionary(Of String, Object) From {
                {"OrdenItemVenta", nroOrden},
                {"NroReferenciaSGF", argOpera.IdOperacion},
                {"NroReferenciaSGFOrigen", Nothing},
                {"CodBarras", If(IsNumeric(item.CodBarras), CLng(item.CodBarras), item.CodBarras)},
                {"CapturaCodBarras", "E"},
                {"CodTroquel", Nothing},
                {"Alfabeta", item.Alfabeta},
                {"Kairos", Nothing},
                {"Codigo", Nothing},
                {"IdProducto", Nothing},
                {"Descripcion", item.Descripcion},
                {"CantidadVenta", item.Cant},
                {"ImporteUnitario", CDec(item.PrecioUnit)}
            }

            ' 2. Mapeamos directamente la lista de objetos UnidadTrazable de ItemReceta
            If item.UnidadesTrazables IsNot Nothing AndAlso item.UnidadesTrazables.Count > 0 Then
                Dim listaJsonTrazables As New List(Of Object)

                For Each ut As UnidadTrazable In item.UnidadesTrazables
                    listaJsonTrazables.Add(New With {
                        .GTIN = ut.Gtin.Trim().PadLeft(14, "0"c),
                        .NroSerie = ut.Serial.Trim(),
                        .NroLote = If(Not String.IsNullOrEmpty(ut.Lote), ut.Lote.Trim(), "S/L"),
                        .FechaVencimiento = If(ut.Vencimiento <> Date.MinValue, ut.Vencimiento.ToString("yyyyMMdd"), DateTime.Now.ToString("yyyyMMdd"))
                    })
                Next

                itemVentaDict.Add("UnidadesTrazables", listaJsonTrazables)
            End If

            ' 3. Estructura interna del nodo Item
            Dim nodoItem = New With {
                .NroOrden = nroOrden,
                .ItemVenta = itemVentaDict,
                .RecetaDigital = New With {
                    .Fecha = receta.FechaPresc.ToString("yyyyMMdd"),
                    .NroRecetaDigital = receta.NumRta,
                    .NroReferenciaConsulta = receta.NumConsulta,
                    .TipoPrescriptor = receta.TipoPrescriptor,
                    .TipoMatricula = receta.TipoMatricula,
                    .ProvinciaMatricula = receta.Provincia,
                    .NroMatricula = receta.NumMatricula,
                    .IdProductoSolicitado = Nothing,
                    .AlfabetaSolicitado = alfabetataSolicitado,
                    .DrogaCodigoSolicitado = Nothing,
                    .CodDrogaSnomedSolicitado = Nothing,
                    .DrogaDescripcionSolicitado = Nothing,
                    .CantidadSolicitado = item.CantPrescripta,
                    .EstadoSugerido = item.Estado,
                    .Financiador = New With {
                        .CodigoADESFA = Nothing,
                        .Codigo = receta.Financiador
                    }
                },
                .Validador = New With {
                    .CodigoADESFA = "A0030",
                    .Nombre = "FARMALINK",
                    .NroReferencia = If(Not String.IsNullOrEmpty(item.NumAutorItem), item.NumAutorItem, receta.NumAutorizacion),
                    .CantidadAprobada = item.Cant
                }
            }

            listaDatosVentas.Add(New With {.Item = nodoItem})
            nroOrden += 1
        Next

        ' 4. Estructura principal completa
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
                        .CodigoImed = m_NumPrestadorSIGEDAD,
                        .Cuit = CLng(argCbte.Empresa.CUIT.Replace("-", "")),
                        .CUFE = argCbte.Empresa.CUFE
                    }
                },
                .Comprobante = New With {
                    .Id = CLng(argCbte.NumComp),
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
                        .NroDocumentoR = CLng(argCbte.Cliente.NumDoc.Replace("-", "")),
                        .FechaNacimientoR = "19850412",
                        .SexoR = "F",
                        .LecturaDNI = ""
                    }
                },
                .DatosVentas = listaDatosVentas
            }
        }

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
Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports SiCoFa.Entidades

Public Class D_AdminRtas

    Public Function ObtenerRecetaPorIdOperacion(argIdOperacion As Long) As Receta
        Try
            Dim objRta As Receta = Nothing

            Dim sql As String = "SELECT * FROM TblRtas WHERE IdOperación=@IdOperacion"

            Using cn As New SqlConnection(Mod_D_Admin.strConexionDB)
                cn.Open()

                Using cmd As SqlCommand = cn.CreateCommand()
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = sql
                    cmd.Parameters.Add("@IdOperacion", SqlDbType.BigInt).Value = argIdOperacion

                    Using datos As SqlDataReader = cmd.ExecuteReader()
                        If Not datos.Read() Then
                            Throw New Exception("Receta no Encontrada para la Operación indicada.")
                        End If

                        Dim idRta As Long = If(IsDBNull(datos("IdRta")), 0L, Convert.ToInt64(datos("IdRta")))
                        Dim idOpera As Long = If(IsDBNull(datos("IdOperacion")), 0L, Convert.ToInt64(datos("IdOperacion")))
                        Dim financiador As String = If(IsDBNull(datos("Financiador")), String.Empty, Convert.ToString(datos("Financiador")))
                        Dim fechaPresc As Date = If(IsDBNull(datos("FechaPresc")), Date.MinValue, Convert.ToDateTime(datos("FechaPresc")))
                        Dim fechaDisp As Date = If(IsDBNull(datos("FechaDisp")), Date.MinValue, Convert.ToDateTime(datos("FechaDisp")))
                        Dim numRta As String = If(IsDBNull(datos("NumRta")), String.Empty, Convert.ToString(datos("NumRta")))
                        Dim tipoDoc As String = If(IsDBNull(datos("TipoDoc")), String.Empty, Convert.ToString(datos("TipoDoc")))
                        Dim numDoc As String = If(IsDBNull(datos("NumDoc")), String.Empty, Convert.ToString(datos("NumDoc")))
                        Dim numAf As String = If(IsDBNull(datos("NumAf")), String.Empty, Convert.ToString(datos("NumAf")))
                        Dim nombreAf As String = If(IsDBNull(datos("nombreAf")), String.Empty, Convert.ToString(datos("nombreAf")))
                        Dim tipoMatricula As String = If(IsDBNull(datos("TipoMatricula")), String.Empty, Convert.ToString(datos("TipoMatricula")))
                        Dim provincia As String = If(IsDBNull(datos("Provincia")), String.Empty, Convert.ToString(datos("Provincia")))
                        Dim numMatricula As String = If(IsDBNull(datos("NumMatricula")), String.Empty, Convert.ToString(datos("NumMatricula")))
                        Dim tipoPrescriptor As String = If(IsDBNull(datos("TipoPrescriptor")), String.Empty, Convert.ToString(datos("TipoPrescriptor")))
                        Dim impTotalRta As Decimal = If(IsDBNull(datos("ImpTotalRta")), 0D, Convert.ToDecimal(datos("ImpTotalRta")))
                        Dim impOS As Decimal = If(IsDBNull(datos("ImpOS")), 0D, Convert.ToDecimal(datos("ImpOS")))
                        Dim impAf As Decimal = If(IsDBNull(datos("ImpAf")), 0D, Convert.ToDecimal(datos("ImpAf")))
                        Dim numAutorizacion As String = If(IsDBNull(datos("NumAutorizacion")), String.Empty, Convert.ToString(datos("NumAutorizacion")))
                        Dim estadoRta As String = If(IsDBNull(datos("EstadoRta")), String.Empty, Convert.ToString(datos("EstadoRta")))
                        Dim numConsulta As String = If(IsDBNull(datos("NumConsulta")), String.Empty, Convert.ToString(datos("NumConsulta")))

                        Dim itemsRta As List(Of ItemReceta) = ObtenerItemsRta(idRta)

                        objRta = New Receta(idRta, idOpera, financiador, fechaPresc, fechaDisp, numRta, tipoDoc, numDoc, numAf, nombreAf, tipoMatricula, provincia, numMatricula, tipoPrescriptor, impTotalRta, impOS, impAf, numAutorizacion, estadoRta, numConsulta, itemsRta)

                        Return objRta
                    End Using
                End Using
            End Using

        Catch ex As Exception
            Throw New Exception(Vecho.MensajeError(Me.ToString, "ObtenerRecetaPorIdOperacion", ex.Message), ex)
        End Try
    End Function

    Private Function ObtenerItemsRta(argIdRta As Long) As List(Of ItemReceta)
        Dim objItemsR As New List(Of ItemReceta)

        Try
            Dim SqlItems As String = "SELECT IdDP, IdRta, Descripcion, Cant, PrecioUnit, POS, DesOS, " &
                                    "NumAutorItem, CantPrescripta, Estado, Alfabeta, CodBarras " &
                                    "FROM TblDPRtas WHERE IdRta=@IdRta"

            Using cn As New SqlConnection(Mod_D_Admin.strConexionDB)
                cn.Open()

                ' 1. Carga inicial de ítems (cierra el DataReader al terminar el bloque Using)
                Using cmd As SqlCommand = cn.CreateCommand()
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = SqlItems
                    cmd.Parameters.Add("@IdRta", SqlDbType.BigInt).Value = argIdRta

                    Using datos As SqlDataReader = cmd.ExecuteReader()
                        While datos.Read()
                            Dim idDP As Long = If(IsDBNull(datos("IdDP")), 0L, Convert.ToInt64(datos("IdDP")))
                            Dim idRta As Long = If(IsDBNull(datos("IdRta")), 0L, Convert.ToInt64(datos("IdRta")))
                            Dim descripcion As String = If(IsDBNull(datos("Descripcion")), String.Empty, Convert.ToString(datos("Descripcion")))
                            Dim cant As Integer = If(IsDBNull(datos("Cant")), 0, Convert.ToInt32(datos("Cant")))
                            Dim pUnit As Decimal = If(IsDBNull(datos("PrecioUnit")), 0D, Convert.ToDecimal(datos("PrecioUnit")))
                            Dim pos As Decimal = If(IsDBNull(datos("POS")), 0D, Convert.ToDecimal(datos("POS")))
                            Dim desOS As Decimal = If(IsDBNull(datos("DesOS")), 0D, Convert.ToDecimal(datos("DesOS")))
                            Dim numAutorItem As String = If(IsDBNull(datos("NumAutorItem")), String.Empty, Convert.ToString(datos("NumAutorItem")))
                            Dim cantPrescripta As Integer = If(IsDBNull(datos("CantPrescripta")), 0, Convert.ToInt32(datos("CantPrescripta")))
                            Dim estado As Integer = If(IsDBNull(datos("Estado")), 0, Convert.ToInt32(datos("Estado")))
                            Dim alfabeta As Integer = If(IsDBNull(datos("Alfabeta")), 0, Convert.ToInt32(datos("Alfabeta")))
                            Dim codBarras As String = If(IsDBNull(datos("CodBarras")), String.Empty, Convert.ToString(datos("CodBarras")))

                            Dim objItemR As New ItemReceta(idDP, idRta, descripcion, cant, pUnit, pos, desOS,
                                                           numAutorItem, cantPrescripta, estado, alfabeta, codBarras, Nothing)

                            objItemsR.Add(objItemR)
                        End While
                    End Using
                End Using

                ' 2. Con el DataReader de ítems cerrado, leemos las trazabilidades asociadas sobre la misma conexión
                For Each item As ItemReceta In objItemsR
                    item.UnidadesTrazables = ObtenerTrazablesPorIdDP(item.IdDP, cn)
                Next
            End Using

            Return objItemsR

        Catch ex As Exception
            Throw New Exception(Vecho.MensajeError(Me.ToString, "ObtenerItemsRta", ex.Message), ex)
        End Try
    End Function

    ''' <summary>
    ''' Recupera las N unidades trazables asociadas a un renglón de receta específico (IdDP)
    ''' </summary>
    Private Function ObtenerTrazablesPorIdDP(argIdDP As Long, cn As SqlConnection) As List(Of UnidadTrazable)
        Dim lista As New List(Of UnidadTrazable)

        Dim sql As String = "SELECT gtin, numero_serial, lote, vencimiento FROM TblTransacciones WHERE IdDP = @IdDP"

        Using cmd As SqlCommand = cn.CreateCommand()
            cmd.CommandType = CommandType.Text
            cmd.CommandText = sql
            cmd.Parameters.Add("@IdDP", SqlDbType.BigInt).Value = argIdDP

            Using datos As SqlDataReader = cmd.ExecuteReader()
                While datos.Read()
                    Dim gtin As String = If(IsDBNull(datos("gtin")), String.Empty, Convert.ToString(datos("gtin")).Trim())
                    Dim serial As String = If(IsDBNull(datos("numero_serial")), String.Empty, Convert.ToString(datos("numero_serial")).Trim())
                    Dim lote As String = If(IsDBNull(datos("lote")), String.Empty, Convert.ToString(datos("lote")).Trim())
                    Dim vencimiento As Date = If(IsDBNull(datos("vencimiento")), Date.MinValue, Convert.ToDateTime(datos("vencimiento")))

                    If Not String.IsNullOrEmpty(serial) Then
                        lista.Add(New UnidadTrazable(gtin, serial, lote, vencimiento))
                    End If
                End While
            End Using
        End Using

        Return lista
    End Function

End Class
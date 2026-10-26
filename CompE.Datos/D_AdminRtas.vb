Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports SiCoFa.Entidades

Public Class D_AdminRtas
    Public Function ObtenerRecetaPorIdOperacion(argIdOperacion As Long) As Receta
        Try
            Dim objRta As Receta
            'Dim objDetalleR As List(Of ItemComprobanteRecetas) = Nothing
            Dim idRta As Long
            Dim idOpera As Long
            Dim fechaPresc As Date
            Dim fechaDisp As Date
            Dim numRta As String
            Dim tipoDoc As String
            Dim numDoc As String
            Dim numAf As String
            Dim nombreAf As String
            Dim tipoMatricula As String
            Dim provincia As String
            Dim numMatricula As String
            Dim tipoPrescriptor As String
            Dim impTotalRta As Decimal
            Dim impOS As Decimal
            Dim impAf As Decimal
            Dim numAutorizacion As String
            Dim estadoRta As String
            Dim numConsulta As String

            Dim sql As String = "SELECT * FROM TblRtas WHERE IdOperación=@IdOperacion"

            Using cn As New SqlConnection(Mod_D_Admin.strConexionDB)
                cn.Open()

                Using cmd As SqlCommand = cn.CreateCommand()
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = sql
                    cmd.Parameters.AddWithValue("@IdOperacion", argIdOperacion)

                    Using datos As SqlDataReader = cmd.ExecuteReader()
                        datos.Read()

                        If datos.HasRows = False Then
                            Throw New Exception("Receta no Encontrada")
                        End If

                        idRta = datos("IdRta")
                        idOpera = datos("IdOperacion")
                        fechaPresc = datos("FechaPresc")
                        fechaDisp = datos("FechaDisp")
                        numRta = datos("NumRta")
                        tipoDoc = datos("TipoDoc")
                        numDoc = datos("NumDoc")
                        numAf = datos("NumAf")
                        nombreAf = datos("nombreAf")
                        tipoMatricula = datos("TipoMatricula")
                        provincia = datos("Provincia")
                        numMatricula = datos("NumMatricula")
                        tipoPrescriptor = datos("TipoPrescriptor")
                        impTotalRta = datos("ImpTotalRta")
                        impOS = datos("ImpOS")
                        impAf = datos("ImpAf")
                        numAutorizacion = datos("NumAutorizacion")
                        estadoRta = datos("EstadoRta")
                        numConsulta = datos("NumConsulta")

                    End Using
                End Using
            End Using


            Return objRta

        Catch ex As Exception
            Throw New Exception(Vecho.MensajeError(Me.ToString, "ObtenerRecetaPorIdOperacion", ex.Message))
            Return Nothing

        End Try

    End Function
End Class

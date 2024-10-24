Imports System.Data.SqlClient
Imports SiCoFa.Entidades
Public Class D_AdminOperaciones
    Public Sub RegistrarError(ByVal argIdOpera As String, argDesError As String)
        Try
            Dim sql As String = "UPDATE TblOpera SET EstadoOpera='Error',DesError='" & Replace(argDesError, "'", "") & "' WHERE IdOperación=" & argIdOpera

            Using cn As New SqlConnection(Mod_D_Admin.strConexionDB)
                cn.Open()

                Using cmd As SqlCommand = cn.CreateCommand()
                    cmd.ExecuteNonQuery()
                End Using

            End Using

        Catch Ex As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "RegistrarError", Ex.Message))
        End Try

    End Sub
    Public Sub RegistrarFinalizado(ByVal argOperacion As Operacion)
        Try
            Using cn As New SqlConnection(Mod_D_Admin.strConexionDB)
                cn.Open()

                Using cmd As SqlCommand = New SqlCommand("FinalizarOperacion", cn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.AddWithValue("@IdOpera", argOperacion.IdOperacion)
                    cmd.Parameters.AddWithValue("@IdEmp", argOperacion.IdEmpleado)
                    cmd.Parameters.AddWithValue("@Estado", "Finalizado")
                    cmd.Parameters.AddWithValue("@Obser", argOperacion.Observaciones)
                    cmd.ExecuteNonQuery()
                End Using
            End Using

        Catch Ex As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "RegistrarFinalizado", Ex.Message))

        End Try

    End Sub
    Public Function ObtenerOperacion(ByVal argIdOperacion As Long) As Operacion
        Dim objOpera As Operacion = Nothing
        Try

            Dim sql As String = "SELECT IdOperación,Inicio,Fin,IdEmpr,IdPC,IdCaja,IdEmpleado,CodiTO,EstadoOpera,Observaciones FROM TblOpera WHERE IdOperación=@IdOperacion"

            Using cn As New SqlConnection(Mod_D_Admin.strConexionDB)
                cn.Open()

                Using cmd As SqlCommand = cn.CreateCommand()
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = sql
                    cmd.Parameters.AddWithValue("@IdOperacion", argIdOperacion)

                    Using datos As SqlDataReader = cmd.ExecuteReader()
                        datos.Read()

                        If datos.HasRows Then
                            objOpera = New Operacion(
                                                datos("IdOperación"),
                                                datos("Inicio"),
                                                datos("Fin"),
                                                datos("IdEmpr"),
                                                datos("IdPC"),
                                                datos("IdCaja"),
                                                datos("IdEmpleado"),
                                                datos("CodiTO"),
                                                datos("EstadoOpera"),
                                                datos("Observaciones").ToString
                                                )

                        End If

                    End Using
                End Using
            End Using

            Return objOpera

        Catch ex As Exception
            Throw New Exception(Vecho.MensajeError(Me.ToString, "ObtenerOperacion", ex.Message))
            Return Nothing
        End Try

    End Function
    Public Sub AcutalizarStock(ByVal argOperacion As Operacion, ByVal argEfInf As Integer)

        Try
            Using cn As New SqlConnection(Mod_D_Admin.strConexionDB)
                cn.Open()

                Using cmd As SqlCommand = New SqlCommand("Actualizar_Stock_EnvCer", cn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.AddWithValue("@IdOpera", argOperacion.IdOperacion)
                    cmd.Parameters.AddWithValue("@EfInv", argEfInf)
                    cmd.ExecuteNonQuery()
                End Using

                Using cmd As SqlCommand = New SqlCommand("Actualizar_Stock_EnvFrac", cn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.AddWithValue("@IdOpera", argOperacion.IdOperacion)
                    cmd.Parameters.AddWithValue("@EfInv", argEfInf)
                    cmd.ExecuteNonQuery()
                End Using
            End Using

        Catch Ex As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "ActualizarStock", Ex.Message))

        End Try

    End Sub

End Class
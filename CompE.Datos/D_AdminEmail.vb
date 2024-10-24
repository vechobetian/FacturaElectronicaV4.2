Imports System.Data.SqlClient
Imports SiCoFa.Entidades
Public Class D_AdminEmail
    Public Function ObtenerEmailEmpresa() As Email
        Dim objEmail As Email = Nothing
        Try

            Dim sql As String = "SELECT IdMail,Port,Host,Usuario,Contraseña,Mail FROM TblEmail"

            Using cn As New SqlConnection(Mod_D_Admin.strConexionDB)
                cn.Open()

                Using cmd As SqlCommand = cn.CreateCommand()
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = sql

                    Using datos As SqlDataReader = cmd.ExecuteReader()
                        datos.Read()
                        If datos.HasRows Then
                            objEmail = New Email(
                            datos("IdMail"),
                            datos("Port"),
                            datos("Host"),
                            datos("Usuario"),
                            datos("Contraseña"),
                            datos("Mail")
                             )
                        End If

                    End Using
                End Using
            End Using

        Catch ex As Exception
            Throw ex
        End Try
        Return objEmail

    End Function
End Class
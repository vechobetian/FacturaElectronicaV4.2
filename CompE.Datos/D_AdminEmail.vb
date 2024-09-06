Imports System.Data.SqlClient
Imports SiCoFa.Entidades
Public Class D_AdminEmail
    Public Function ObtenerEmailEmpresa() As Email
        Dim objEmail As Email = Nothing
        Try

            Dim sql As String

            sql = "SELECT IdMail,Port,Host,Usuario,Contraseña,Mail FROM TblEmail"

            Dim cmd As SqlCommand = D_Admin.ConexionDB.conn.CreateCommand()
            cmd.CommandType = CommandType.Text
            cmd.CommandText = sql
            Dim datos As SqlDataReader = cmd.ExecuteReader()
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

        Catch ex As Exception
            Throw ex
        End Try
        Return objEmail

    End Function
End Class
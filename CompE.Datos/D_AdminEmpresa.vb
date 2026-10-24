Imports System.Data.SqlClient
Imports SiCoFa.Entidades
Public Class D_AdminEmpresa
    Public Function ObtenerEmpresa() As Empresa
        Dim objEmpr As Empresa
        Try
            Dim sql As String = "SELECT Nombre, Domicilio, Localidad, Provincia, TE, CUIT, IB, IVA, InicActiv, GLN, CUFE FROM TblEmpresa"

            Using cn As New SqlConnection(Mod_D_Admin.strConexionDB)
                cn.Open()

                Using cmd As SqlCommand = cn.CreateCommand()
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = sql

                    Using datos As SqlDataReader = cmd.ExecuteReader()
                        If datos.Read() Then
                            objEmpr = New Empresa(
                            If(IsDBNull(datos("Nombre")), "", datos("Nombre").ToString()),
                            If(IsDBNull(datos("Domicilio")), "", datos("Domicilio").ToString()),
                            If(IsDBNull(datos("Localidad")), "", datos("Localidad").ToString()),
                            If(IsDBNull(datos("Provincia")), "", datos("Provincia").ToString()),
                            If(IsDBNull(datos("TE")), "", datos("TE").ToString()),
                            If(IsDBNull(datos("CUIT")), "", datos("CUIT").ToString()),
                            If(IsDBNull(datos("IB")), "", datos("IB").ToString()),
                            If(IsDBNull(datos("IVA")), "", datos("IVA").ToString()),
                            If(IsDBNull(datos("InicActiv")), "", datos("InicActiv").ToString()),
                            If(IsDBNull(datos("GLN")), "", datos("GLN").ToString().Trim()),
                            If(IsDBNull(datos("CUFE")), "", datos("CUFE").ToString().Trim())
                        )
                        Else
                            Throw New Exception("No se encontró el registro de la empresa en TblEmpresa.")
                        End If
                    End Using
                End Using
            End Using

            Return objEmpr

        Catch ex As Exception
            Throw New Exception(Vecho.MensajeError(Me.ToString, "ObtenerEmpresa", ex.Message))
            Return Nothing
        End Try
    End Function

End Class
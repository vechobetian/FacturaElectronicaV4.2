Imports System.Data.SqlClient
Imports SiCoFa.Entidades
Public Class D_AdminEmpresa
    Public Function ObtenerEmpresa() As Empresa
        Dim objEmpr As Empresa
        Try

            Dim sql As String = "SELECT Nombre,Domicilio,Localidad,Provincia,TE,CUIT,IB,IVA,InicActiv,GLN FROM TblEmpresa"
            Dim cn As New Conexion

            Using cmd As SqlCommand = cn.conn.CreateCommand()
                cmd.CommandType = CommandType.Text
                cmd.CommandText = sql

                Using datos As SqlDataReader = cmd.ExecuteReader()
                    datos.Read()
                    objEmpr = New Empresa(
                                    datos("Nombre"),
                                    datos("Domicilio"),
                                    datos("Localidad"),
                                    datos("Provincia"),
                                    datos("TE").ToString,
                                    datos("CUIT"),
                                    datos("IB"),
                                    datos("IVA"),
                                    datos("InicActiv"),
                                    datos("GLN")
                                    )

                End Using

            End Using

            cn.CerrarConexion()
            cn = Nothing

            Return objEmpr

        Catch ex As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "ObtenerEmpresa", ex.Message))
            Return Nothing

        End Try

    End Function

End Class
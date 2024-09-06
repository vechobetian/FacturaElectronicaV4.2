Imports System.Data.SqlClient
Imports SiCoFa.Entidades
Public Class D_AdminEmpresa
    Public Function ObtenerEmpresa() As Empresa
        Dim objEmpr As Empresa
        Try

            Dim sql As String
            sql = "SELECT Nombre,Domicilio,Localidad,Provincia,TE,CUIT,IB,IVA,InicActiv,GLN FROM TblEmpresa"

            Dim cmd As SqlCommand = D_Admin.ConexionDB.conn.CreateCommand()
            cmd.CommandType = CommandType.Text
            cmd.CommandText = sql
            Dim datos As SqlDataReader = cmd.ExecuteReader()
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
            datos.Close()
            cmd.Dispose()
            Return objEmpr

        Catch ex As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "ObtenerEmpresa", ex.Message))
            Return Nothing

        End Try

    End Function

End Class
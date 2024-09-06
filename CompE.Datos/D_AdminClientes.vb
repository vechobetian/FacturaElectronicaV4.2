Imports System.Data.SqlClient
Imports SiCoFa.Entidades

Public Class D_AdminClientes
    Public Function ObtenerCliente(ByVal argIdCliente As Long) As Cliente
        Dim objCli As Cliente
        Try
            Dim sql As String

            sql = "SELECT IdCliente,IdGpoCli,RazónSocial,DNI,Calle,Número,Localidad,Provincia,Telefono,EstadoCuenta,Límite,IVA,CUIT,Email,GLN FROM TblClientes WHERE IdCliente=" & argIdCliente

            Dim cmd As SqlCommand = D_Admin.ConexionDB.conn.CreateCommand()
            cmd.CommandType = CommandType.Text
            cmd.CommandText = sql
            Dim datosC As SqlDataReader = cmd.ExecuteReader()
            datosC.Read()

            If datosC.HasRows Then
                Dim IdCliente As Long = datosC("IdCliente")
                Dim IdGpoCli As Long = datosC("IdGpoCli")
                Dim RazónSocial As String = datosC("RazónSocial")
                Dim DNI As String = datosC("DNI")
                Dim Calle As String = datosC("Calle").ToString
                Dim Número As String = datosC("Número").ToString
                Dim Localidad As String = datosC("Localidad").ToString
                Dim Provincia As String = datosC("Provincia").ToString
                Dim Telefono As String = datosC("Telefono").ToString
                Dim EstadoCuenta As String = datosC("EstadoCuenta")
                Dim Límite As String = datosC("Límite")
                Dim IVA As String = datosC("IVA")
                Dim CUIT As String = datosC("CUIT").ToString
                Dim GLN As String = datosC("GLN")
                Dim Email As String = datosC("Email")

                objCli = New Cliente(IdCliente, IdGpoCli, RazónSocial, DNI, Calle, Número, Localidad, Provincia, Telefono, EstadoCuenta, Límite, IVA, CUIT, GLN, Email)

            Else
                objCli = New Cliente(0, 0, "CONSUMIDOR FINAL", "0", "NO APLICA", "", "", "", "", "", 0, "SI", "", "", "")
            End If

            datosC.Close()
            cmd.Dispose()
            Return objCli

        Catch ex As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "ObtenerCliente", ex.Message))
            Return Nothing

        End Try

    End Function

End Class
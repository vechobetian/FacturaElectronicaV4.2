Imports SiCoFa.Datos
Imports SiCoFa.Entidades

Public Class N_AdminClientes

    Private mobjD_AdminClientes As D_AdminClientes
    Public Sub New()
        mobjD_AdminClientes = New D_AdminClientes
    End Sub
    Public Function ObtenerCliente(ByVal argIdCliente As Long) As Cliente
        Try
            Dim objCli As Cliente
            objCli = mobjD_AdminClientes.ObtenerCliente(argIdCliente)
            Return objCli

        Catch ex As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "ObtenerCliente", ex.Message))
            Return Nothing

        End Try

    End Function

End Class
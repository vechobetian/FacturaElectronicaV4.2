Imports SiCoFa.Datos
Imports SiCoFa.Entidades

Public Class N_AdminComprobantes

    Private mobjD_AdminComprobantes As D_AdminComprobantes
    Public Sub New()
        mobjD_AdminComprobantes = New D_AdminComprobantes
    End Sub
    Public Sub RegistrarError(ByVal argIdOperacion As Long)
        Try
            mobjD_AdminComprobantes.RegistrarError(argIdOperacion)

        Catch ex As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "RegistrarError", ex.Message))

        End Try
    End Sub
    Public Sub ActualizarComprobante(ByVal argCbte As Comprobante)

        Try
            mobjD_AdminComprobantes.ActualizarComprobante(argCbte)

        Catch ex As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "ActualizarComprobante", ex.Message))

        End Try

    End Sub

    Public Function ObtenerComprobante(ByVal argOperacion As Operacion) As Comprobante
        Try
            Dim objCbte As Comprobante
            objCbte = mobjD_AdminComprobantes.ObtenerComprobante(argOperacion)
            Return objCbte

        Catch ex As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "ObtenerComprobante", ex.Message))
            Return Nothing

        End Try

    End Function

End Class
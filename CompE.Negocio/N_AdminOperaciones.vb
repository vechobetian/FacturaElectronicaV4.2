Imports SiCoFa.Datos
Imports SiCoFa.Entidades
Public Class N_AdminOperaciones
    Private mobjD_AdminOperaciones As D_AdminOperaciones
    Public Sub New()
        mobjD_AdminOperaciones = New D_AdminOperaciones
    End Sub
    Public Function ObtenerOperacion(ByVal argIdOpera As Long) As Operacion
        Try
            Dim objOpera As Operacion
            objOpera = mobjD_AdminOperaciones.ObtenerOperacion(argIdOpera)
            Return objOpera

        Catch ex As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "ObtenerOperacion", ex.Message))
            Return Nothing

        End Try

    End Function
    Public Sub RegistrarError(ByVal argIdOperacion As Long, argObservaciones As String)
        Try
            mobjD_AdminOperaciones.RegistrarError(argIdOperacion, argObservaciones)

        Catch ex As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "RegistrarError", ex.Message))
        End Try

    End Sub
    Public Sub RegistrarFinalizado(ByVal argOperacion As Operacion)

        Try
            mobjD_AdminOperaciones.RegistrarFinalizado(argOperacion)

        Catch ex As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "RegistrarFinalizado", ex.Message))
        End Try

    End Sub
    Public Sub ActualizarStock(ByVal argOperacion As Operacion, ByVal argEfInv As Integer)
        Try
            mobjD_AdminOperaciones.AcutalizarStock(argOperacion, argEfInv)

        Catch ex As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "ActualizarStock", ex.Message))
        End Try

    End Sub

End Class
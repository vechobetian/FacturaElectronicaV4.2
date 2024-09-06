Imports SiCoFa.Datos
Imports SiCoFa.Entidades
Public Class N_AdminEmpresa

    Private mobjD_AdminEmpresa As D_AdminEmpresa
    Public Sub New()
        mobjD_AdminEmpresa = New D_AdminEmpresa
    End Sub
    Public Function ObtenerEmpresa() As Empresa
        Dim empr As Empresa
        Try
            empr = mobjD_AdminEmpresa.ObtenerEmpresa
            Return empr

        Catch ex As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "ObtenerEmpresa", ex.Message))
            Return Nothing

        End Try

    End Function

End Class
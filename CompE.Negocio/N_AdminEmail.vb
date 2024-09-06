Imports SiCoFa.Datos
Imports SiCoFa.Entidades
Public Class N_AdminEmail
    Private mobjD_AdminEmail As D_AdminEmail
    Public Sub New()
        mobjD_AdminEmail = New D_AdminEmail
    End Sub
    Public Function ObtenerEmailEmpresa() As Email
        Try
            Dim email As Email
            email = mobjD_AdminEmail.ObtenerEmailEmpresa
            Return email

        Catch ex As Exception
            Throw ex
            Return Nothing

        End Try

    End Function

End Class
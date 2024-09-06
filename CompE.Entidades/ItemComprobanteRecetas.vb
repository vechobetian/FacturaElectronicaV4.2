Public Class ItemComprobanteRecetas
    Property NombreOS As String
    Property CantRtas As Integer
    Property ImpOS As Decimal
    Public Sub New(
                  ByVal argNombreOS As String,
                  ByVal argCantRtas As Integer,
                  ByVal argImpOS As Decimal
                  )
        Try
            Me.NombreOS = argNombreOS
            Me.CantRtas = argCantRtas
            Me.ImpOS = argImpOS

        Catch ex As Exception
            Throw ex
        End Try

    End Sub
End Class
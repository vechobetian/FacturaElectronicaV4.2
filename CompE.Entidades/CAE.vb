
Public Class CAE
    Property NumComp As Long
    Property NumCAE As String
    Property VtoCAE As Date

    Public Sub New(
                    ByVal argNumComp As Long,
                    ByVal argNumCAE As String,
                    ByVal argVtoCAE As Date
                    )

        Me.NumComp = argNumComp
        Me.NumCAE = argNumCAE
        Me.VtoCAE = argVtoCAE.ToString("dd-MM-yyyy")
    End Sub

End Class
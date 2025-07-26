Public Class FarmaPuntos
    Property IdOperacion As Long
    Property IdCliente As Long
    Property Puntos As Decimal
    Property Importe As Decimal
    Property Vecimiento As DateTime
    Property PuntosAcumulados As Decimal
    Property ImporteAcumulado As Decimal

    Public Sub New(ByVal argIdOperacion As Long, ByVal argIdCliente As Long, ByVal argPuntos As Decimal, ByVal argImporte As Decimal, ByVal argVencimiento As Date, ByVal argPuntosAcumulados As Decimal, ByVal argImporteAcumulado As Decimal)
        Me.IdOperacion = argIdOperacion
        Me.IdCliente = argIdCliente
        Me.Puntos = argPuntos
        Me.Importe = argImporte
        Me.Vecimiento = argVencimiento
        Me.PuntosAcumulados = argPuntosAcumulados
        Me.ImporteAcumulado = argImporteAcumulado
    End Sub

End Class

Public Class UnidadTrazable
    Public Property Gtin As String
    Public Property Serial As String
    Public Property Lote As String
    Public Property Vencimiento As Date

    Public Sub New()
    End Sub

    Public Sub New(argGtin As String, argSerial As String, argLote As String, argVencimiento As Date)
        Me.Gtin = argGtin
        Me.Serial = argSerial
        Me.Lote = argLote
        Me.Vencimiento = argVencimiento
    End Sub
End Class

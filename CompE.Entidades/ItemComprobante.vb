Public Class ItemComprobante
    Property Descripcion As String
    Property Cantidad As Integer
    Property AlicIVA As Decimal
    Property PUnit As Decimal
    Property DUnit As Decimal
    Property PDes As Decimal
    Property OtraDescripcion As String
    Property Importe As Decimal
    Property ImporteDescuento As Decimal
    Property ImporteConDescuento As Decimal
    Public Sub New(
                  ByVal argDescripcion As String,
                  ByVal argCantidad As Integer,
                  ByVal argPUnit As Decimal,
                  ByVal argGravado As Boolean,
                  ByVal argDisIva As Boolean,
                  ByVal argDUnit As Decimal,
                  ByVal argPDes As Decimal,
                  ByVal argOtraDescripcion As String
                  )
        Try
            Me.Descripcion = argDescripcion
            Me.Cantidad = argCantidad

            If argDisIva And argGravado Then
                Me.PUnit = Math.Round(argPUnit / 1.21, 2, MidpointRounding.ToEven)
                Me.DUnit = Math.Round(argDUnit / 1.21, 2, MidpointRounding.ToEven)
            Else
                Me.PUnit = Math.Round(argPUnit, 2)
                Me.DUnit = Math.Round(argDUnit, 2)
            End If

            If argGravado Then
                Me.AlicIVA = 21
            Else
                Me.AlicIVA = 0
            End If

            Me.Importe = Math.Round(Me.Cantidad * Me.PUnit, 2, MidpointRounding.ToEven)
            Me.PDes = argPDes
            Me.OtraDescripcion = argOtraDescripcion
            Me.ImporteDescuento = Math.Round(Me.Cantidad * Me.DUnit, 2, MidpointRounding.ToEven)
            Me.ImporteConDescuento = Me.Importe - Me.ImporteDescuento

        Catch ex As Exception
            Throw ex
        End Try

    End Sub
End Class
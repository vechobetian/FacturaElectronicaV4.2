Public Class Comprobante
    Property Operacion As Operacion
    Property TipoComprobante As TipoComprobante
    Property PuntoVta As String
    Property NumComp As String
    Property FechaComp As Date
    Property ImpBto As Decimal
    Property ImpEx As Decimal
    Property ImpGrav As Decimal
    Property ImpNeto As Decimal
    Property IVA As Decimal
    Property ImpEf As Decimal
    Property ImpCC As Decimal
    Property ImpTar As Decimal
    Property ImpOS As Decimal
    Property ImpDes As Decimal
    Property CAE As CAE
    Property Cliente As Cliente
    Property CompAsoc As Comprobante
    Property Empresa As Empresa
    Property Detalle As List(Of ItemComprobante)
    Property DetalleRtas As List(Of ItemComprobanteRecetas)
    Property QR As QRCompE
    Public Sub New(
                  ByVal argOpera As Operacion,
                  ByVal argCodiTC_SiCoFa As String,
                  ByVal argPuntoVta As String,
                  ByVal argNumComp As String,
                  ByVal argFechaComp As Date,
                  ByVal argImpBto As Decimal,
                  ByVal argImpEx As Decimal,
                  ByVal argImpGrav As Decimal,
                  ByVal argImpNeto As Decimal,
                  ByVal argIVA As Decimal,
                  ByVal argImpOS As Decimal,
                  ByVal argImpEf As Decimal,
                  ByVal argImpCC As Decimal,
                  ByVal argImpTar As Decimal,
                  ByVal argImpDes As Decimal,
                  ByVal argCAE As CAE,
                  ByVal argCliente As Cliente,
                  ByVal argCompAsoc As Comprobante,
                  ByVal argEmpresa As Empresa,
                  ByVal argDetalle As List(Of ItemComprobante),
                  ByVal argDetalleRtas As List(Of ItemComprobanteRecetas)
                  )
        Me.Operacion = argOpera
        Me.TipoComprobante = New TipoComprobante(argCodiTC_SiCoFa)
        Me.PuntoVta = argPuntoVta
        Me.NumComp = argNumComp
        Me.FechaComp = argFechaComp
        Me.ImpBto = Math.Round(argImpBto, 2)
        Me.ImpEx = Math.Round(argImpEx, 2)
        Me.ImpGrav = Math.Round(argImpGrav, 2)
        Me.ImpNeto = Math.Round(argImpNeto, 2)
        Me.IVA = Math.Round(argIVA, 2)
        Me.ImpOS = Math.Round(argImpOS, 2)
        Me.ImpEf = Math.Round(argImpEf, 2)
        Me.ImpCC = Math.Round(argImpCC, 2)
        Me.ImpTar = Math.Round(argImpTar, 2)
        Me.ImpDes = Math.Round(argImpDes, 2)
        Me.CAE = argCAE
        Me.Cliente = argCliente
        Me.CompAsoc = argCompAsoc
        Me.Empresa = argEmpresa
        Me.Detalle = argDetalle
        Me.DetalleRtas = argDetalleRtas

    End Sub

End Class
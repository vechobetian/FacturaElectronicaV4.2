Public Class Receta
    Public Property IdRta As Long
    Public Property IdOperacion As Long
    Public Property Financiador As String
    Public Property FechaPresc As Date
    Public Property FechaDisp As Date
    Public Property NumRta As String
    Public Property TipoDoc As String
    Public Property NumDoc As String
    Public Property NumAf As String
    Public Property NombreAf As String
    Public Property TipoMatricula As String
    Public Property Provincia As String
    Public Property NumMatricula As String
    Public Property TipoPrescriptor As String
    Public Property ImpTotalRta As Decimal
    Public Property ImpOS As Decimal
    Public Property ImpAf As Decimal
    Public Property NumAutorizacion As String
    Public Property EstadoRta As String
    Public Property NumConsulta As String
    Public Property ItemsReceta As List(Of ItemReceta)

    Public Sub New(
                   argIdRta As Long,
                   argIdOperacion As Long,
                   argFinanciador As String,
                   argFechaPresc As Date,
                   argFechaDip As Date,
                   argNumRta As String,
                   argTipoDoc As String,
                   argNumDoc As String,
                   argNumAf As String,
                   argNombreAf As String,
                   argTipoMatricula As String,
                   argProvincia As String,
                   argNumMatricula As String,
                   argTipoPrescriptor As String,
                   argImpTotalRta As Decimal,
                   argImpOS As Decimal,
                   argImpAf As Decimal,
                   argNumAutorizacion As String,
                   argEstadoRta As String,
                   argNumConsulta As String,
                   argItemsReceta As List(Of ItemReceta)
                   )
        Me.IdRta = argIdRta
        Me.IdOperacion = argIdOperacion
        Me.Financiador = argFinanciador
        Me.FechaPresc = argFechaPresc
        Me.FechaDisp = argFechaDip
        Me.NumRta = argNumRta
        Me.TipoDoc = argTipoDoc
        Me.NumDoc = argNumDoc
        Me.NumAf = argNumAf
        Me.NombreAf = argNombreAf
        Me.TipoMatricula = argTipoMatricula
        Me.Provincia = argProvincia
        Me.NumMatricula = argNumMatricula
        Me.TipoPrescriptor = argTipoPrescriptor
        Me.ImpTotalRta = argImpTotalRta
        Me.ImpOS = argImpOS
        Me.ImpAf = argImpAf
        Me.NumAutorizacion = argNumAutorizacion
        Me.EstadoRta = argEstadoRta
        Me.NumConsulta = argNumConsulta
        Me.ItemsReceta = argItemsReceta
    End Sub
End Class

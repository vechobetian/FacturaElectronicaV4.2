Public Class ItemReceta
    Public Property IdDP As Long
    Public Property IdRta As Long
    Public Property Descripcion As String
    Public Property Cant As Integer
    Public Property PrecioUnit As Decimal
    Public Property POS As Decimal
    Public Property DesOS As Decimal
    Public Property NumAutorItem As String
    Public Property CantPrescripta As Integer
    Public Property Estado As Integer
    Public Property Alfabeta As Integer
    Public Property CodBarras As String

    Public Sub New(argIdDP As Long,
                   argIdRta As Long,
                   argDescripcion As String,
                   argCant As Integer,
                   argPrecioUnit As Decimal,
                   argPOS As Decimal,
                   argDesOS As Decimal,
                   argNumAutorItem As String,
                   argCantPrescripta As Integer,
                   argEstado As Integer,
                   argAlfabeta As Integer,
                   argCodBarras As String)

        Me.IdDP = argIdDP
        Me.IdRta = argIdRta
        Me.Descripcion = argDescripcion
        Me.Cant = argCant
        Me.PrecioUnit = argPrecioUnit
        Me.POS = argPOS
        Me.DesOS = argDesOS
        Me.NumAutorItem = argNumAutorItem
        Me.CantPrescripta = argCantPrescripta
        Me.Estado = argEstado
        Me.Alfabeta = argAlfabeta
        Me.CodBarras = argCodBarras
    End Sub

End Class

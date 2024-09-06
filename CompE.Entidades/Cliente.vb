Public Class Cliente
    Property IdCliente As Long
    Property IdGCli As Long
    Property Nombre As String
    Property DNI As String
    Property Calle As String
    Property Numero As String
    Property Localidad As String
    Property Provincia As String
    Property Domicilio As String
    Property Telefono As String
    Property Estado As String
    Property Limite As Decimal
    Property IVA As String
    Property IVADescripcion As String
    Property CUIT As String
    Property TipoDoc As String
    Property CodiTD_AFIP As Integer
    Property NumDoc As String
    Property GLN As String
    Property Email As String

    Public Sub New(
                ByVal argIdCliente As Long,
                ByVal argGCli As Long,
                ByVal argNombre As String,
                ByVal argDNI As String,
                ByVal argCalle As String,
                ByVal argNumero As String,
                ByVal argLocalidad As String,
                ByVal argProvincia As String,
                ByVal argTelefono As String,
                ByVal argEstado As String,
                ByRef argLimite As Decimal,
                ByVal argIVA As String,
                ByVal argCUIT As String,
                ByVal argGLN As String,
                ByVal argEmail As String
                )

        Me.IdCliente = argIdCliente
        Me.IdGCli = argGCli
        Me.Nombre = argNombre
        Me.DNI = argDNI
        Me.Calle = argCalle
        Me.Numero = argNumero
        Me.Localidad = argLocalidad
        Me.Provincia = argProvincia
        Me.Domicilio = Me.Calle & " " & Me.Numero & " " & Me.Localidad & "-" & Me.Provincia
        Me.Telefono = argTelefono
        Me.Estado = argEstado
        Me.Limite = argLimite
        Me.IVA = argIVA
        Me.CUIT = argCUIT
        Me.TipoDoc = Me.ObetenerTipoDoc
        Me.CodiTD_AFIP = Me.ObtenerCodiTD_AFIP
        Me.NumDoc = Me.ObtenerNumDoc
        Me.IVADescripcion = Me.ObtenerIVADescripcion
        Me.GLN = argGLN
        Me.Email = argEmail

    End Sub

    Private Function ObtenerIVADescripcion() As String
        Select Case Me.IVA
            Case "RI"
                Return "IVA Responsable Inscripto"
            Case "MT"
                Return "Responsable Monotributo"
            Case "EX"
                Return "IVA Sujeto Exento"
            Case "CF"
                Return "Consumidor Final"
            Case Else
                Return "No Identificado"
        End Select
    End Function

    Private Function ObetenerTipoDoc() As String
        Select Case Me.IVA
            Case "RI", "EX", "MT"
                Return "CUIT"
            Case "CF"
                Return "DNI"
            Case Else
                Return "T.DOC"
        End Select

    End Function

    Private Function ObtenerNumDoc() As String
        Select Case Me.IVA
            Case "RI", "EX", "MT"
                Return Me.CUIT
            Case "CF"
                Return Me.DNI
            Case Else
                Return "0"
        End Select
    End Function

    Private Function ObtenerCodiTD_AFIP() As Integer
        Select Case Me.TipoDoc
            Case "CUIT"
                Return 80
            Case "DNI"
                Return 96
            Case Else
                Return 99

        End Select
    End Function
End Class
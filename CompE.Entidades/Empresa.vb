Public Class Empresa
    Property Nombre As String
    Property Domicilio As String
    Property Localidad As String
    Property Provincia As String
    Property Telefono As String
    Property CUIT As String
    Property IB As String
    Property IVA As String
    Property IVADescripcion As String
    Property InicioActividad As String
    Property GLN As String
    Public Sub New(
                  ByVal argNombre As String,
                  ByVal argDomicilio As String,
                  ByVal argLocalidad As String,
                  ByVal argProvincia As String,
                  ByVal argTelefono As String,
                  ByVal argCUIT As String,
                  ByVal argIB As String,
                  ByVal argIVA As String,
                  ByVal argInicioActividad As String,
                  ByVal argGLN As String
                  )


        Me.Nombre = argNombre
        Me.Domicilio = argDomicilio
        Me.Localidad = argLocalidad
        Me.Provincia = argProvincia
        Me.Telefono = argTelefono
        Me.CUIT = argCUIT
        Me.IB = argIB
        Me.IVA = argIVA
        Me.IVADescripcion = ObtenerTipoIVA()
        Me.InicioActividad = argInicioActividad
        Me.GLN = argGLN

    End Sub
    Private Function ObtenerTipoIVA() As String
        Select Case Me.IVA
            Case "RI", "RIM"
                Return "Responsable Inscripto"
            Case "MT"
                Return "Responsable Monotributo"
            Case "Ex"
                Return "Sujeto Exento"
            Case "CF"
                Return "Consumidor Final"
            Case Else
                Return "No Identificado"
        End Select
    End Function

End Class
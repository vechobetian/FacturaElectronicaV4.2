''' <summary>
''' Representa un error genérico producido en las reglas de negocio.
''' De esta clase deben heredar las clases que representen errores 
''' específicos de reglas de negocio.
''' </summary>
Public Class NegocioException
    Inherits ApplicationException

    ''' <summary>
    ''' Construye una instancia con un mensaje de error, 
    ''' invocando al constructor de la clase base.
    ''' </summary>
    ''' <param name="mensaje">El mensaje de error.</param>
    ''' <param name="original">El error original.</param>
    Public Sub New(ByVal mensaje As String, ByVal original As Exception)
        MyBase.New(mensaje, original)
    End Sub

    ''' <summary>
    ''' Construye una instancia con un mensaje de error, 
    ''' invocando al constructor de la clase base.
    ''' </summary>
    ''' <param name="mensaje">El mensaje de error.</param>
    Public Sub New(ByVal mensaje As String)
        MyBase.new(mensaje)
    End Sub

End Class

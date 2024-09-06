Public Class LoginTicket
    Property UniqueId As UInt32 ' Entero de 32 bits sin signo que identifica el requerimiento
    Property GenerationTime As DateTime ' Momento en que fue generado el requerimiento
    Property ExpirationTime As DateTime ' Momento en el que expira la solicitud
    Property Sign As String ' Firma de seguridad recibida en la respuesta
    Property Token As String ' Token de seguridad recibido en la respuesta

End Class
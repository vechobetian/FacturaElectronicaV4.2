Public Class Email
    Property IdMail As Integer
    Property Port As Integer
    Property Host As String
    Property Usuario As String
    Property Contraseña As String
    Property Mail As String

    Public Sub New(
                  ByVal argIdMail As Integer,
                  ByVal argPort As Integer,
                  ByVal argHost As String,
                  ByVal argUsuario As String,
                  ByVal argContraseña As String,
                  ByVal argMail As String
                  )
        Me.IdMail = argIdMail
        Me.Port = argPort
        Me.Host = argHost
        Me.Usuario = argUsuario
        Me.Contraseña = argContraseña
        Me.Mail = argMail
    End Sub

End Class
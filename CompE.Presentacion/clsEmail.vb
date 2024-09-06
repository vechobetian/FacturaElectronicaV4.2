Imports System.Net.Mail
Imports SiCoFa.Negocio
Imports SiCoFa.Entidades
Public Class clsEmail

    Property Email As Email
    Public Function ObtenerEmailEmpresa() As Boolean

        Dim objN_AdminEmail As New N_AdminEmail

        Try
            Email = objN_AdminEmail.ObtenerEmailEmpresa
            Return True

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message,
                            "Error al ObtenerEmail",
                             MessageBoxButtons.OK)
            Return False
        End Try

    End Function
    Public Function EnviarMail(ByVal argNombre As String, ByVal argMail As String, ByVal argAsunto As String, ByVal argMensaje As String, ByVal argPathAdjunto As String) As Boolean

        Dim smtp As New SmtpClient
        Dim correo As New MailMessage
        Dim adjunto As Attachment

        With smtp
            .Port = Email.Port 'para gmail y hotmail es 587
            .Host = Email.Host 'para gmail=smtp.gmail.com, para hotmail=smtp.office365.com
            .Credentials = New System.Net.NetworkCredential(Email.Usuario, Email.Contraseña) 'en gmail y hotmail el usuario es el correo
            .EnableSsl = True
        End With

        adjunto = New Attachment(argPathAdjunto)
        With correo
            .From = New MailAddress(Email.Mail, argNombre)
            .To.Add(argMail)
            .Subject = argAsunto
            .Body = argMensaje
            .IsBodyHtml = False
            .Priority = MailPriority.Normal
            .Attachments.Add(adjunto)

        End With

        Try
            smtp.Send(correo)
            Return True

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message,
                            "Error al enviar correo",
                             MessageBoxButtons.OK)
            Return False
        End Try

        smtp.Dispose()
        correo.Dispose()
        adjunto.Dispose()

    End Function


End Class

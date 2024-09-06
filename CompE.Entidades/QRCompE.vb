Imports System.Drawing
Imports System.Text
Imports MessagingToolkit.QRCode.Codec
Imports Newtonsoft.Json
Public Class QRCompE
    Property ver As Integer = 1
    Property fecha As String
    Property cuit As Long
    Property ptoVta As Integer
    Property tipoCmp As Integer
    Property nroCmp As Long
    Property importe As Decimal
    Property moneda As String = "PES"
    Property ctz As Decimal = 1
    Property tipoDocRec As Integer
    Property nroDocRec As Long
    Property tipoCodAut As String = "E"
    Property codAut As Long
    Property QR As Byte()
    Public Sub New(
                  ByVal argFecha As Date,
                  ByVal argCuit As Long,
                  ByVal argPtoVta As Integer,
                  ByVal argCodiTC_AFIP As Integer,
                  ByVal argNroComp As Long,
                  ByVal argImporte As Decimal,
                  ByVal argCliTipoDoc As Integer,
                  ByVal argCliNroDoc As Long,
                  ByVal argCAE As Long
                  )
        Me.fecha = argFecha.ToString("yyyy-MM-dd")
        Me.cuit = argCuit
        Me.ptovta = argPtoVta
        Me.tipoCmp = argCodiTC_AFIP
        Me.nroCmp = argNroComp
        Me.importe = argImporte
        Me.tipoDocRec = argCliTipoDoc
        Me.nroDocRec = argCliNroDoc
        Me.codAut = argCAE
        Me.QR = Me.GenerarQR(Me.QRBase64(Me))

    End Sub
    Private Function QRBase64(objQR As QRCompE) As String
        Dim WS_QR_URL As String = "https://www.afip.gob.ar/fe/qr/"
        Dim url As String = WS_QR_URL + "?p="
        Dim jsQR As String = JsonConvert.SerializeObject(objQR)
        Dim Base64String As String = Convert.ToBase64String(Encoding.UTF8.GetBytes(jsQR))
        Return url + Base64String

    End Function
    Private Function GenerarQR(argStr As String) As Byte()
        Try
            Dim qrCode As New QRCodeEncoder
            qrCode.QRCodeEncodeMode = QRCodeEncoder.ENCODE_MODE.BYTE
            qrCode.QRCodeErrorCorrect = QRCodeEncoder.ERROR_CORRECTION.L

            Dim QRB As Byte() = Me.Imagen_Bytes(qrCode.Encode(argStr, System.Text.Encoding.UTF8))
            Return QRB

        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function Imagen_Bytes(ByVal Foto As Image) As Byte()
        If Not Foto Is Nothing Then
            Dim Codi As New IO.MemoryStream
            Foto.Save(Codi, Imaging.ImageFormat.Jpeg)
            Return Codi.GetBuffer
        Else
            Return Nothing
        End If
    End Function
    Private Function Bytes_Imagen(ByVal Foto As Byte()) As Image
        If Not Foto Is Nothing Then
            Dim Codi As New IO.MemoryStream(Foto)
            Dim resultado As Image = Image.FromStream(Codi)
            Return resultado
        Else
            Return Nothing
        End If
    End Function

End Class
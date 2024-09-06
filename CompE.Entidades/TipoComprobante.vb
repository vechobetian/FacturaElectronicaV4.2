Public Class TipoComprobante
    Property CodiTC_SiCoFa As String
    Property CodiTC_AFIP As String
    Property TipoComprobante As String
    Property Letra As String
    Public Sub New(argCodiTC_SiCoFa As String)
        Me.CodiTC_SiCoFa = argCodiTC_SiCoFa
        Me.CodiTC_AFIP = Me.ObtenerCodiTC_AFIP
        Me.TipoComprobante = Me.ObtenerTipoComprobante

        If CodiTC_SiCoFa = "REC" Or CodiTC_SiCoFa = "RECR" Or CodiTC_SiCoFa = "PRESU" Or CodiTC_SiCoFa = "RTO" Then
            Me.Letra = "X"

        Else
            Me.Letra = Right(argCodiTC_SiCoFa, 1)

        End If

    End Sub
    Private Function ObtenerCodiTC_AFIP() As String
        Select Case Me.CodiTC_SiCoFa
            Case "FAA"
                Return "01"
            Case "FAB"
                Return "06"
            Case "FAC"
                Return "11"
            Case "FAM"
                Return "51"
            Case "NCA"
                Return "03"
            Case "NCB"
                Return "08"
            Case "NCC"
                Return "13"
            Case "NCM"
                Return "53"
            Case Else
                Return 0
        End Select

    End Function
    Private Function ObtenerTipoComprobante() As String
        Select Case Me.CodiTC_SiCoFa
            Case "FAA", "FAB", "FAC", "FAM"
                Return "FACTURA"

            Case "NCA", "NCB", "NCC", "NCM", "NCR", "NCX"
                Return "NOTA DE CREDITO"

            Case "REC"
                Return "RECIBO"

            Case "RECR"
                Return "RECIBO RECETAS"

            Case "PRESU"
                Return "PRESUPUESTO"

            Case "RTO"
                Return "REMITO"

            Case Else
                Return "DESCONOCIDO"
        End Select

    End Function

End Class
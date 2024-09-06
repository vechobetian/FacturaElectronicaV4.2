Imports System.IO
Imports Microsoft.Reporting.WinForms
Imports SiCoFa.Entidades
Public Class clsGenerarPDF
    Property Operacion As New List(Of Operacion)
    Property Empresa As New List(Of Empresa)
    Property Cliente As New List(Of Cliente)
    Property TipoComp As New List(Of TipoComprobante)
    Property Encabezado As New List(Of Comprobante)
    Property Detalle As New List(Of ItemComprobante)
    Property CAE As New List(Of CAE)
    Property QR As New List(Of QRCompE)
    Property Copia As String
    Property CompAsoc As String
    Property Path As String
    Property CantiLetras As String

    Public Sub Export(ByVal report As LocalReport)

        Dim pdfContent As Byte() = report.Render("PDF")
        Dim pdfPath As String = Path '"C:\SiCoFa_Cliente\Temp\" & NameFil & ".pdf"
        Dim pdfFile As New System.IO.FileStream(pdfPath, System.IO.FileMode.Create)
        pdfFile.Write(pdfContent, 0, pdfContent.Length)
        pdfFile.Close()

    End Sub


    'Crear un informe local para Report.rdlc, cargar los datos, exportar el informe a un archivo .emf e imprimirlo.
    Public Sub Run()
        Dim strPS As String = ParamTerminal.ReadINI("C:\SiCoFa_Cliente\config.ini", "SiCoFa", "strPathS")
        Dim report As New LocalReport()
        Dim param As New List(Of ReportParameter)
        param.Add(New ReportParameter("Copia", Me.Copia))
        param.Add(New ReportParameter("CompAsoc", Me.CompAsoc))
        If Me.TipoComp(0).CodiTC_SiCoFa = "REC" Then
            param.Add(New ReportParameter("CantiLetras", Me.CantiLetras))
        End If

        Select Case Me.TipoComp(0).Letra
            Case "A", "M"
                report.ReportPath = strPS & "\SiCoFa_Server\CompE\rptCompAA4.rdlc"
            Case "B"
                report.ReportPath = strPS & "\SiCoFa_Server\CompE\rptCompBA4.rdlc"
            Case "C"
                report.ReportPath = strPS & "\SiCoFa_Server\CompE\rptCompCA4.rdlc"
            Case "X"
                Select Case Me.TipoComp(0).CodiTC_SiCoFa
                    Case "RTO"
                        report.ReportPath = strPS & "\SiCoFa_Server\CompE\rptCompRA4.rdlc"
                    Case "PRESU"
                        report.ReportPath = "C:\SiCoFa_Server\CompE\rptCompBA4.rdlc"
                    Case "REC"
                        report.ReportPath = strPS & "\SiCoFa_Server\CompE\rptReciboA4.rdlc"
                End Select
        End Select

        report.DataSources.Add(New ReportDataSource("Operacion", Operacion))
        report.DataSources.Add(New ReportDataSource("Empresa", Empresa))
        report.DataSources.Add(New ReportDataSource("Cliente", Cliente))
        report.DataSources.Add(New ReportDataSource("TipoComprobante", TipoComp))
        report.DataSources.Add(New ReportDataSource("Encabezado", Encabezado))
        report.DataSources.Add(New ReportDataSource("Detalle", Detalle))
        report.DataSources.Add(New ReportDataSource("CAE", CAE))
        report.DataSources.Add(New ReportDataSource("QR", QR))
        report.SetParameters(param)

        Export(report)

    End Sub


End Class

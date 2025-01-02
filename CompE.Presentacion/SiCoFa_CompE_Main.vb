Imports System.Net
Imports System.IO
Imports SiCoFa.Entidades
Imports SiCoFa.Negocio
Module SiCoFa_CompE_Main

    Private mobjN_AdminOpera As N_AdminOperaciones
    Private mobjN_AdminCbtes As N_AdminComprobantes
    Private mobjCbte As Comprobante
    Public Sub Main(ByVal cmdArgs() As String)

        Dim IdOpera As Long = 36459
        Dim Hoja As String = "TK"
        Dim NumCopias As Integer = 1
        Dim Email As String = "NO"

        Try
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            For argNum As Integer = 0 To UBound(cmdArgs, 1)
                Select Case argNum
                    Case 0
                        IdOpera = CLng(cmdArgs(argNum))
                    Case 1
                        Hoja = cmdArgs(argNum)
                    Case 2
                        NumCopias = CInt(cmdArgs(argNum))
                    Case 3
                        Email = cmdArgs(argNum)
                End Select
            Next

            If IdOpera = 0 Then
                MsgBox("IdOperacion no establecido", vbInformation, "SiCoFa_CompE")
                Exit Sub
            End If

            mobjN_AdminOpera = New N_AdminOperaciones
            Dim objOpera As Operacion = mobjN_AdminOpera.ObtenerOperacion(IdOpera)

            If objOpera Is Nothing Then
                MsgBox("Operación no encontrada", vbInformation, "SiCoFa_Comp")
                Exit Sub
            End If

            If objOpera.Estado <> "Finalizado" Then
                Exit Sub
            End If

            mobjN_AdminCbtes = New N_AdminComprobantes
            mobjCbte = mobjN_AdminCbtes.ObtenerComprobante(objOpera)

            If mobjCbte Is Nothing Then
                Throw New Exception(vecho.MensajeError("SiCoFa_CompE_Main", "Main", "Comprobante no encontrado"))
            End If

            If mobjCbte.NumComp = "E" Then
                If GenerarFacturaElectronica() = False Then
                    Exit Sub
                End If

                mobjN_AdminCbtes.ActualizarComprobante(mobjCbte)

                'mobjN_AdminOpera.RegistrarFinalizado(mobjCbte.Operacion)

                'Select Case mobjCbte.Operacion.CodiTO
                'Case "VTAM", "INTF"
                'mobjN_AdminOpera.ActualizarStock(mobjCbte.Operacion, -1)
                'Case "NC"
                'mobjN_AdminOpera.ActualizarStock(mobjCbte.Operacion, 1)
                'End Select
            Else
                GenerarQR()
            End If

            If Email = "SI" Then
                EnviarMail()
            End If

            If NumCopias = 0 Then
                Exit Sub
            End If

            Select Case Hoja
                Case "A4"
                    Call ImprimirA4(NumCopias)

                Case "TK"
                    Call ImprimirTK(NumCopias)

                Case "TK58"
                    Call ImprimirTK58(NumCopias)

                Case "PDF"
                    Call GuardarComo()

            End Select

        Catch ex As Exception
            Dim msjMostrar As String = ex.Message
            Dim msjGuardar As String = Replace(msjMostrar, vbCrLf, "")

            If mobjCbte.NumComp = "E" Then
                mobjN_AdminCbtes.RegistrarError(IdOpera)
                mobjN_AdminOpera.RegistrarError(IdOpera, msjGuardar)
            End If

            MsgBox(msjMostrar, vbCritical, "SiCoFa_CompE")

        End Try

    End Sub
    Private Function GenerarFacturaElectronica() As Boolean

        Try
            If SolicitarCAE() = True Then
                GenerarQR()
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Throw New Exception(vecho.MensajeError("SiCoFa_CompE_Main", "GenerarFacturaElectronica", ex.Message))
            Return False
        End Try

    End Function
    Private Function SolicitarCAE() As Boolean

        Try

            Dim objN_AdminCAE As New N_AdminCAE
            mobjCbte.CAE = objN_AdminCAE.ObtenerCAE(mobjCbte)

            If mobjCbte.CAE Is Nothing Then
                Return False
                Exit Function
            End If

            mobjCbte.NumComp = Format(mobjCbte.CAE.NumComp, "00000000")

            Return True

        Catch ex As Exception
            Throw New Exception(vecho.MensajeError("SiCoFa_CompE_Main", "SolicitarCAE", ex.Message))
            Return False

        End Try

    End Function
    Private Sub GenerarQR()

        Try
            If mobjCbte.CAE IsNot Nothing Then
                Dim CUIT As Long = CLng(Replace(mobjCbte.Empresa.CUIT, "-", ""))
                Dim PVta As Integer = CInt(mobjCbte.PuntoVta)
                Dim NumComp As Long = CLng(mobjCbte.NumComp)

                mobjCbte.QR = New QRCompE(mobjCbte.FechaComp, CUIT, PVta, mobjCbte.TipoComprobante.CodiTC_AFIP, NumComp, mobjCbte.ImpBto, mobjCbte.Cliente.CodiTD_AFIP, mobjCbte.Cliente.NumDoc, mobjCbte.CAE.NumCAE)

            End If

        Catch ex As Exception
            Throw New Exception(vecho.MensajeError("SiCoFa_CompE_Main", "GenerarQR", ex.Message))
        End Try
    End Sub
    Private Sub pdfA4(Path As String)

        Try
            Dim objPdf As New clsGenerarPDF
            With objPdf
                .Operacion.Add(mobjCbte.Operacion)
                .Empresa.Add(mobjCbte.Empresa)
                .Cliente.Add(mobjCbte.Cliente)
                .TipoComp.Add(mobjCbte.TipoComprobante)
                .Encabezado.Add(mobjCbte)
                .Detalle = mobjCbte.Detalle
                .CAE.Add(mobjCbte.CAE)
                .QR.Add(mobjCbte.QR)
                .Copia = "ORIGINAL"

                If mobjCbte.CompAsoc IsNot Nothing Then
                    .CompAsoc = "Comprobante Asociado: " & mobjCbte.CompAsoc.TipoComprobante.TipoComprobante & " " & mobjCbte.CompAsoc.TipoComprobante.Letra & " " & mobjCbte.CompAsoc.PuntoVta & "-" & mobjCbte.CompAsoc.NumComp
                Else
                    .CompAsoc = ""
                End If

                If mobjCbte.TipoComprobante.CodiTC_SiCoFa = "REC" Then
                    .CantiLetras = UCase(vecho.NumEnLetras(mobjCbte.ImpBto))
                End If

                .Path = Path
            End With

            objPdf.Run()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub ImprimirA4(ByVal argNumCopias As Integer)
        Dim objImp As New clsImprimir
        Dim Copia As String = ""

        Try
            If argNumCopias < 0 Then
                Select Case Math.Abs(argNumCopias)
                    Case 1
                        Copia = "ORIGINAL"
                    Case 2
                        Copia = "DUPLICADO"
                    Case 3
                        Copia = "TRIPLICADO"
                End Select

                With objImp
                    .Operacion.Add(mobjCbte.Operacion)
                    .Empresa.Add(mobjCbte.Empresa)
                    .Cliente.Add(mobjCbte.Cliente)
                    .TipoComp.Add(mobjCbte.TipoComprobante)
                    .Encabezado.Add(mobjCbte)
                    .Detalle = mobjCbte.Detalle
                    .CAE.Add(mobjCbte.CAE)
                    .QR.Add(mobjCbte.QR)
                    .Copia = Copia
                    If mobjCbte.CompAsoc IsNot Nothing Then
                        .CompAsoc = "Comprobante Asociado: " & mobjCbte.CompAsoc.TipoComprobante.TipoComprobante & " " & mobjCbte.CompAsoc.TipoComprobante.Letra & " " & mobjCbte.CompAsoc.PuntoVta & "-" & mobjCbte.CompAsoc.NumComp
                    Else
                        .CompAsoc = ""
                    End If
                End With

                objImp.Run()
                argNumCopias = 0
            End If

            For x = 1 To argNumCopias
                Select Case x
                    Case 1
                        Copia = "ORIGINAL"
                    Case 2
                        Copia = "DUPLICADO"
                    Case 3
                        Copia = "TRIPLICADO"
                End Select

                With objImp
                    .Operacion.Add(mobjCbte.Operacion)
                    .Empresa.Add(mobjCbte.Empresa)
                    .Cliente.Add(mobjCbte.Cliente)
                    .TipoComp.Add(mobjCbte.TipoComprobante)
                    .Encabezado.Add(mobjCbte)
                    .Detalle = mobjCbte.Detalle
                    .CAE.Add(mobjCbte.CAE)
                    .QR.Add(mobjCbte.QR)
                    .Copia = Copia
                    If mobjCbte.CompAsoc IsNot Nothing Then
                        .CompAsoc = "Comprobante Asociado: " & mobjCbte.CompAsoc.TipoComprobante.TipoComprobante & " " & mobjCbte.CompAsoc.TipoComprobante.Letra & " " & mobjCbte.CompAsoc.PuntoVta & "-" & mobjCbte.CompAsoc.NumComp
                    Else
                        .CompAsoc = ""
                    End If
                End With

                objImp.Run()
            Next
            objImp.Dispose()

        Catch ex As Exception
            Throw ex
        End Try

        objImp.Dispose()

    End Sub
    Private Sub ImprimirTK(ByVal argNumCopias As Integer)

        Try
            Dim objTkt As New clsTicket
            Dim Copia As String = ""
            objTkt.Comprobante = mobjCbte

            If argNumCopias < 0 Then
                Select Case Math.Abs(argNumCopias)
                    Case 1
                        Copia = "                ORIGINAL                 "
                    Case 2
                        Copia = "               DUPLICADO                 "
                    Case 3
                        Copia = "               TRIPLICADO                "
                End Select

                objTkt.Imprimir(Copia)
            End If

            For x = 1 To argNumCopias
                Select Case x
                    Case 1
                        Copia = "                ORIGINAL                 "
                    Case 2
                        Copia = "               DUPLICADO                 "
                    Case 3
                        Copia = "               TRIPLICADO                "
                End Select

                objTkt.Imprimir(Copia)
            Next

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub ImprimirTK58(ByVal argNumCopias As Integer)

        Try
            Dim objTkt As New clsTicket58
            Dim Copia As String = ""
            objTkt.Comprobante = mobjCbte

            If argNumCopias < 0 Then
                Select Case Math.Abs(argNumCopias)
                    Case 1
                        Copia = "        ORIGINAL"
                    Case 2
                        Copia = "        DUPLICADO"
                    Case 3
                        Copia = "        TRIPLICADO"
                End Select

                objTkt.Imprimir(Copia)
            End If

            For x = 1 To argNumCopias
                Select Case x
                    Case 1
                        Copia = "        ORIGINAL"
                    Case 2
                        Copia = "        DUPLICADO"
                    Case 3
                        Copia = "        TRIPLICADO"
                End Select

                objTkt.Imprimir(Copia)
            Next

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub EnviarMail()

        Dim PathAdjunto As String = "C:\SiCoFa_Cliente\Temp\" & mobjCbte.TipoComprobante.CodiTC_SiCoFa & "-" & mobjCbte.PuntoVta & "-" & mobjCbte.NumComp & ".pdf"

        Try

            Dim objEmail As New clsEmail

            If objEmail.ObtenerEmailEmpresa() = False Then
                MsgBox("No se estableció el Correo de la Farmacia", "CompE")
                Exit Sub
            End If

            If objEmail.Email Is Nothing Then
                MessageBox.Show("No se pudo obtener los parametros para enviar el Email",
                            "SiCoFa",
                             MessageBoxButtons.OK)

                Exit Sub
            End If

            If mobjCbte.Cliente.Email = "" Then
                Exit Sub
            End If

            Call pdfA4("C:\SiCoFa_Cliente\Temp\" & mobjCbte.TipoComprobante.CodiTC_SiCoFa & "-" & mobjCbte.PuntoVta & "-" & mobjCbte.NumComp & ".pdf")

            Dim Mensaje As String = "Estimado Cliente, adjuntamos comprobante " & vbCrLf & "-Tipo de Comprobante: " & mobjCbte.TipoComprobante.TipoComprobante & vbCrLf & "-Número de Comprobante: " & mobjCbte.PuntoVta & "-" & mobjCbte.NumComp & vbCrLf & vbCrLf & mobjCbte.Empresa.Nombre

            If File.Exists("C:\SiCoFa_Cliente\Temp\" & mobjCbte.TipoComprobante.CodiTC_SiCoFa & "-" & mobjCbte.PuntoVta & "-" & mobjCbte.NumComp & ".pdf") = False Then
                MessageBox.Show("No se encontro el archivo " & mobjCbte.TipoComprobante.CodiTC_SiCoFa & "-" & mobjCbte.PuntoVta & "-" & mobjCbte.NumComp & ".pdf")
                Exit Sub
            End If

            objEmail.EnviarMail(mobjCbte.Empresa.Nombre, mobjCbte.Cliente.Email, "Comprobante Electrónico", Mensaje, PathAdjunto)

        Catch ex As Exception
            MsgBox(ex.Message) 'no hago Throw ex porque me registra error en la base de datos

        End Try

    End Sub
    Private Sub GuardarComo()
        Dim saveFileDialog1 As New SaveFileDialog()

        Try
            With saveFileDialog1
                .Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
                .FilterIndex = 2
                .FileName = mobjCbte.TipoComprobante.CodiTC_SiCoFa & "-" & mobjCbte.PuntoVta & "-" & mobjCbte.NumComp
                .DefaultExt = ".pdf"
                .RestoreDirectory = True
            End With

            If saveFileDialog1.ShowDialog() = DialogResult.OK Then
                Call pdfA4(saveFileDialog1.FileName)
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

End Module
Imports System.Drawing.Printing
Imports SiCoFa.Entidades
Public Class clsTicket
    Property Comprobante As Comprobante
    Property Copia As String

    Private Sub FACTURA(ByVal sender As Object, ByVal e As PrintPageEventArgs)
        Dim fuenteGrande As Font = New Font("consolas", 15)
        Dim printFont As Font = New Font("consolas", 8)
        Dim topMargin As Double = e.MarginBounds.Top
        Dim yPos As Double
        Dim intItems As Integer
        Dim intUnidades As Integer
        Dim strDescripcion As String
        Dim strLinea2Item As String
        Dim strCantPUnit As String
        Dim strImpItem As String
        Dim strImpDesItem As String
        Dim strSubTotal As String
        Dim strImpEx As String
        Dim strImpNeto As String
        Dim strIVA As String
        Dim strTotal As String
        Dim strImpDes As String
        Dim strOS As String
        Dim strTar As String
        Dim strCC As String
        Dim strEf As String
        Dim Tab As String

        Const IncrementoYPreTexto As Integer = 15
        Const IncrementoYPreLinea As Integer = 5
        Const IncrementoYPreItem As Integer = 20
        Const MargenIzquierdo As Integer = 10
        Const Linea As String = "__________________________________________"

        e.Graphics.DrawString(Copia, printFont, Brushes.Black, MargenIzquierdo, 5)

        If Len(Comprobante.Empresa.Nombre) > 20 Then
            e.Graphics.DrawString(Comprobante.Empresa.Nombre, printFont, Brushes.Black, MargenIzquierdo, 30)
        Else
            e.Graphics.DrawString(Comprobante.Empresa.Nombre, fuenteGrande, Brushes.Black, MargenIzquierdo, 30)
        End If

        yPos = 60

        e.Graphics.DrawString(Comprobante.Empresa.Domicilio, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString(Comprobante.Empresa.Localidad & "-" & Comprobante.Empresa.Provincia, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("Telefono: " & Comprobante.Empresa.Telefono, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("Tipo Iva: " & Comprobante.Empresa.IVADescripcion, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("CUIT: " & Comprobante.Empresa.CUIT, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("Ing.Btos: " & Comprobante.Empresa.IB, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("Inicio Actividades: " & Comprobante.Empresa.InicioActividad, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreLinea
        e.Graphics.DrawString(Linea, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreTexto
        Dim fuenteGigante As Font = New Font("consolas", 30)
        Dim rectF1 As New RectangleF(MargenIzquierdo + 3, yPos + 4, 50, 50)
        Dim stringFormat As New StringFormat()

        stringFormat.Alignment = StringAlignment.Center
        stringFormat.LineAlignment = StringAlignment.Center

        e.Graphics.DrawString(Comprobante.TipoComprobante.Letra, fuenteGigante, Brushes.Black, rectF1, stringFormat)
        e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1))
        e.Graphics.DrawString("Cod." & Comprobante.TipoComprobante.CodiTC_AFIP, printFont, Brushes.Black, MargenIzquierdo + 9, yPos + 42)
        e.Graphics.DrawString(Comprobante.TipoComprobante.TipoComprobante, printFont, Brushes.Black, MargenIzquierdo + 58, yPos + 3)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("P.Vta:" & Comprobante.PuntoVta & "           Nro:" & Comprobante.NumComp, printFont, Brushes.Black, MargenIzquierdo + 58, yPos + 10)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("Fecha:" & Comprobante.FechaComp, printFont, Brushes.Black, MargenIzquierdo + 58, yPos + 10)

        Dim hora As String = TimeString
        e.Graphics.DrawString("Hora:" & hora, printFont, Brushes.Black, MargenIzquierdo + 183, yPos + 10)
        yPos += IncrementoYPreLinea + 15
        e.Graphics.DrawString(Linea, printFont, Brushes.Black, MargenIzquierdo, yPos)

        Dim strCliente As String = Left(Comprobante.Cliente.Nombre, 34)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("Cliente:" & strCliente, printFont, Brushes.Black, MargenIzquierdo, yPos)

        If Len(Comprobante.Cliente.Nombre) > 34 Then
            yPos += IncrementoYPreTexto
            strCliente = Mid(Comprobante.Cliente.Nombre, 35, 42)
            e.Graphics.DrawString(strCliente, printFont, Brushes.Black, MargenIzquierdo, yPos)
        End If

        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("Domicilio:" & Comprobante.Cliente.Localidad & "-" & Comprobante.Cliente.Provincia, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("IVA:" & Comprobante.Cliente.IVADescripcion, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("Tipo Doc:" & Comprobante.Cliente.TipoDoc, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("Num.Doc:" & Comprobante.Cliente.NumDoc, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreLinea
        e.Graphics.DrawString(Linea, printFont, Brushes.Black, MargenIzquierdo, yPos)
        'Fin del Encabezado

        'Aca comienza el detalle del comprobante
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("Descripción", printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("Cant/P.Unit.   %IVA      Desc.     Importe", printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreLinea
        e.Graphics.DrawString(Linea, printFont, Brushes.Black, MargenIzquierdo, yPos)

        intItems = 0
        intUnidades = 0

        For Each Item As ItemComprobante In Comprobante.Detalle
            intItems += 1
            intUnidades += Item.Cantidad

            strCantPUnit = Item.Cantidad & "/" & Format(Item.PUnit, "Fixed")
            strIVA = "(" & Format(Item.AlicIVA, "Fixed") & ")"
            Tab = StrDup(13 - Len(strCantPUnit), " ")
            strDescripcion = Left(Item.Descripcion, 42)
            yPos += IncrementoYPreItem
            e.Graphics.DrawString(LTrim(strDescripcion), printFont, Brushes.Black, MargenIzquierdo, yPos)

            If Item.ImporteDescuento > 0 Then

                If Left(Item.OtraDescripcion, 2) <> "OS" And Left(Item.OtraDescripcion, 2) <> "CS" Then
                    strImpItem = Format(Item.ImporteConDescuento, "Fixed")
                    strImpDesItem = Format(Item.ImporteDescuento, "Fixed")
                    strLinea2Item = strCantPUnit & Tab & StrDup(7 - Len(strIVA), " ") & strIVA & StrDup(10 - Len(strImpDesItem), " ") & strImpDesItem & StrDup(12 - Len(strImpItem), " ") & strImpItem

                    yPos += IncrementoYPreTexto
                    e.Graphics.DrawString(strLinea2Item, printFont, Brushes.Black, MargenIzquierdo, yPos)

                    yPos += IncrementoYPreTexto
                    e.Graphics.DrawString(Item.OtraDescripcion, printFont, Brushes.Black, MargenIzquierdo, yPos)

                Else
                    strImpItem = Format(Item.Importe, "Fixed")
                    strImpDesItem = Format(0, "Fixed")
                    strLinea2Item = strCantPUnit & Tab & StrDup(7 - Len(strIVA), " ") & strIVA & StrDup(10 - Len(strImpDesItem), " ") & strImpDesItem & StrDup(12 - Len(strImpItem), " ") & strImpItem

                    yPos += IncrementoYPreTexto
                    e.Graphics.DrawString(strLinea2Item, printFont, Brushes.Black, MargenIzquierdo, yPos)

                    yPos += IncrementoYPreTexto
                    e.Graphics.DrawString(Item.OtraDescripcion, printFont, Brushes.Black, MargenIzquierdo, yPos)
                End If
            Else
                strImpItem = Format(Item.Importe, "Fixed")
                strImpDesItem = Format(0, "Fixed")
                strLinea2Item = strCantPUnit & Tab & StrDup(7 - Len(strIVA), " ") & strIVA & StrDup(10 - Len(strImpDesItem), " ") & strImpDesItem & StrDup(12 - Len(strImpItem), " ") & strImpItem

                yPos += IncrementoYPreTexto
                e.Graphics.DrawString(strLinea2Item, printFont, Brushes.Black, MargenIzquierdo, yPos)
            End If

        Next

        yPos += IncrementoYPreLinea + 5
        e.Graphics.DrawString(Linea, printFont, Brushes.Black, MargenIzquierdo, yPos)

        If Me.Comprobante.TipoComprobante.Letra = "A" Then
            strSubTotal = Format(Comprobante.ImpNeto + Comprobante.ImpEx, "Standard")
            yPos += IncrementoYPreTexto
            e.Graphics.DrawString("Subtotal: " & StrDup(32 - Len(strSubTotal), " ") & strSubTotal, printFont, Brushes.Black, MargenIzquierdo, yPos)

            strImpEx = Format(Comprobante.ImpEx, "Standard")
            yPos += IncrementoYPreTexto
            e.Graphics.DrawString("Imp.Exento: " & StrDup(30 - Len(strImpEx), " ") & strImpEx, printFont, Brushes.Black, MargenIzquierdo, yPos)

            strImpNeto = Format(Comprobante.ImpNeto, "Standard")
            yPos += IncrementoYPreTexto
            e.Graphics.DrawString("Imp.Neto: " & StrDup(32 - Len(strImpNeto), " ") & strImpNeto, printFont, Brushes.Black, MargenIzquierdo, yPos)

            strIVA = Format(Comprobante.IVA, "Standard")
            yPos += IncrementoYPreTexto
            e.Graphics.DrawString("I.V.A: " & StrDup(35 - Len(strIVA), " ") & strIVA, printFont, Brushes.Black, MargenIzquierdo, yPos)

            strTotal = Format(Comprobante.ImpBto, "Standard")
            yPos += IncrementoYPreTexto
            e.Graphics.DrawString("TOTAL: " & StrDup(15 - Len(strTotal), " ") & strTotal, fuenteGrande, Brushes.Black, MargenIzquierdo, yPos)

        Else
            strTotal = Format(Comprobante.ImpBto, "Standard")
            yPos += IncrementoYPreTexto
            e.Graphics.DrawString("TOTAL: " & StrDup(15 - Len(strTotal), " ") & strTotal, fuenteGrande, Brushes.Black, MargenIzquierdo, yPos)

        End If

        yPos += 2 * IncrementoYPreTexto
        e.Graphics.DrawString("RECIBI(MOS)", printFont, Brushes.Black, MargenIzquierdo, yPos)

        If Comprobante.ImpOS > 0 Then
            If Comprobante.DetalleRtas.Count > 0 Then
                'strOS = Format(Comprobante.ImpOS, "Fixed")
                'yPos += IncrementoYPreTexto
                'e.Graphics.DrawString("Obra Social: " & StrDup(29 - Len(strOS), " ") & strOS, printFont, Brushes.Black, MargenIzquierdo, yPos)
                For Each ItemR As ItemComprobanteRecetas In Comprobante.DetalleRtas
                    strOS = Format(ItemR.ImpOS, "Standard")
                    strDescripcion = "AC." & ItemR.NombreOS & ":"
                    yPos += IncrementoYPreTexto
                    e.Graphics.DrawString(strDescripcion & StrDup(42 - Len(strDescripcion) - Len(strOS), " ") & strOS, printFont, Brushes.Black, MargenIzquierdo, yPos)
                Next

            Else
                strOS = Format(Comprobante.ImpOS, "Standard")
                yPos += IncrementoYPreTexto
                e.Graphics.DrawString("Obra Social: " & StrDup(29 - Len(strOS), " ") & strOS, printFont, Brushes.Black, MargenIzquierdo, yPos)

            End If
        End If

        If Comprobante.ImpTar > 0 Then
            strTar = Format(Comprobante.ImpTar, "Standard")
            yPos += IncrementoYPreTexto
            e.Graphics.DrawString("Tarjeta/s: " & StrDup(31 - Len(strTar), " ") & strTar, printFont, Brushes.Black, MargenIzquierdo, yPos)
        End If

        If Comprobante.ImpCC > 0 Then
            strCC = Format(Comprobante.ImpCC, "Standard")
            yPos += IncrementoYPreTexto
            e.Graphics.DrawString("Cuenta Corriente: " & StrDup(24 - Len(strCC), " ") & strCC, printFont, Brushes.Black, MargenIzquierdo, yPos)
        End If

        If Comprobante.ImpEf > 0 Then
            strEf = Format(Comprobante.ImpEf, "Standard")
            yPos += IncrementoYPreTexto
            e.Graphics.DrawString("Efectivo: " & StrDup(32 - Len(strEf), " ") & strEf, printFont, Brushes.Black, MargenIzquierdo, yPos)
        End If

        If Comprobante.ImpDes > 0 Then
            strImpDes = Format(Comprobante.ImpDes, "Standard")
            yPos += 2 * IncrementoYPreTexto
            e.Graphics.DrawString("DESC.APLICADOS: " & StrDup(26 - Len(strImpDes), " ") & strImpDes, printFont, Brushes.Black, MargenIzquierdo, yPos)
        End If

        yPos += IncrementoYPreLinea
        e.Graphics.DrawString(Linea, printFont, Brushes.Black, MargenIzquierdo, yPos)

        If Comprobante.CAE IsNot Nothing Then
            yPos += IncrementoYPreTexto
            e.Graphics.DrawString("      Dirección de comercio interior", printFont, Brushes.Black, MargenIzquierdo, yPos)
            yPos += IncrementoYPreTexto
            e.Graphics.DrawString("        Teléfono 0800-444-03346", printFont, Brushes.Black, MargenIzquierdo, yPos)

            Dim imgQR As Image = Me.Bytes_Imagen(Me.Comprobante.QR.QR)
            Dim x As Single = MargenIzquierdo
            yPos += IncrementoYPreTexto
            Dim y As Single = yPos
            e.Graphics.DrawImage(imgQR, x, y, 100, 100)
            imgQR.Dispose()

            yPos += 60
            e.Graphics.DrawString("CAE:" & Comprobante.CAE.NumCAE, printFont, Brushes.Black, 115, yPos)
            yPos += IncrementoYPreTexto
            e.Graphics.DrawString("Vto:" & Comprobante.CAE.VtoCAE, printFont, Brushes.Black, 115, yPos)
        End If

        If Copia.Trim = "ORIGINAL" Then
            yPos += 30
            e.Graphics.DrawString("           GRACIAS POR SU COMPRA      ", printFont, Brushes.Black, MargenIzquierdo, yPos)
        Else
            yPos += 50
            e.Graphics.DrawString(Linea, printFont, Brushes.Black, MargenIzquierdo, yPos)
            yPos += 15
            e.Graphics.DrawString("                  FIRMA               ", printFont, Brushes.Black, MargenIzquierdo, yPos)
            yPos += 30
            e.Graphics.DrawString(Linea, printFont, Brushes.Black, MargenIzquierdo, yPos)
            yPos += 15
            e.Graphics.DrawString("                ACLARACIÓN            ", printFont, Brushes.Black, MargenIzquierdo, yPos)
        End If

        If Comprobante.CAE Is Nothing Then
            yPos += 30
            e.Graphics.DrawString("DOCUMENTO NO VALIDO COMO FACTURA", printFont, Brushes.Black, MargenIzquierdo, yPos)
        End If

        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("Items:" & intItems, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("Unidades:" & intUnidades, printFont, Brushes.Black, MargenIzquierdo, yPos)

        If Comprobante.Operacion.Observaciones <> "" Then
            Dim intNumCarRestantes As Integer
            Dim strLineaObservaciones As String
            Dim x As Integer
            Dim y As Integer = 1
            Dim lineas() As String = Split(Comprobante.Operacion.Observaciones, vbCrLf)

            yPos += IncrementoYPreLinea
            e.Graphics.DrawString(Linea, printFont, Brushes.Black, MargenIzquierdo, yPos)
            yPos += IncrementoYPreTexto
            e.Graphics.DrawString("Observaciones:", printFont, Brushes.Black, MargenIzquierdo, yPos)

            For Each strLinea As String In lineas
                intNumCarRestantes = Len(strLinea)
                If intNumCarRestantes > 42 Then
                    Do While intNumCarRestantes > 0
                        x += 1
                        strLineaObservaciones = Mid(strLinea, y, 42)
                        y += 42
                        intNumCarRestantes = Len(strLinea) - 42 * x
                        yPos += IncrementoYPreTexto
                        e.Graphics.DrawString(strLineaObservaciones.TrimStart(" "), printFont, Brushes.Black, MargenIzquierdo, yPos)
                    Loop
                Else
                    yPos += IncrementoYPreTexto
                    e.Graphics.DrawString(strLinea.TrimStart(" "), printFont, Brushes.Black, MargenIzquierdo, yPos)
                End If
            Next

        End If
        e.Graphics.Dispose()

    End Sub

    Private Sub RECIBO(ByVal sender As Object, ByVal e As PrintPageEventArgs)
        Dim fuenteGrande As Font = New Font("consolas", 15)
        Dim printFont As Font = New Font("consolas", 8)
        Dim topMargin As Double = e.MarginBounds.Top
        Dim yPos As Double
        Dim strTotal As String
        Dim strTar As String
        Dim strEf As String

        Const IncrementoYPreTexto As Integer = 15
        Const IncrementoYPreLinea As Integer = 5
        Const IncrementoYPreItem As Integer = 20
        Const MargenIzquierdo As Integer = 10
        Const Linea As String = "__________________________________________"

        e.Graphics.DrawString(Copia, printFont, Brushes.Black, MargenIzquierdo, 5)

        If Len(Comprobante.Empresa.Nombre) > 20 Then
            e.Graphics.DrawString(Comprobante.Empresa.Nombre, printFont, Brushes.Black, MargenIzquierdo, 30)
        Else
            e.Graphics.DrawString(Comprobante.Empresa.Nombre, fuenteGrande, Brushes.Black, MargenIzquierdo, 30)
        End If

        yPos = 60

        e.Graphics.DrawString(Comprobante.Empresa.Domicilio, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString(Comprobante.Empresa.Localidad & "-" & Comprobante.Empresa.Provincia, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("Telefono: " & Comprobante.Empresa.Telefono, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("Tipo Iva: " & Comprobante.Empresa.IVADescripcion, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("CUIT: " & Comprobante.Empresa.CUIT, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("Ing.Btos: " & Comprobante.Empresa.IB, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("Inicio Actividades: " & Comprobante.Empresa.InicioActividad, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreLinea
        e.Graphics.DrawString(Linea, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreTexto
        Dim fuenteGigante As Font = New Font("consolas", 30)
        Dim rectF1 As New RectangleF(MargenIzquierdo + 3, yPos + 4, 50, 50)
        Dim stringFormat As New StringFormat()

        stringFormat.Alignment = StringAlignment.Center
        stringFormat.LineAlignment = StringAlignment.Center

        e.Graphics.DrawString(Comprobante.TipoComprobante.Letra, fuenteGigante, Brushes.Black, rectF1, stringFormat)
        e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectF1))
        e.Graphics.DrawString("Cod." & Comprobante.TipoComprobante.CodiTC_AFIP, printFont, Brushes.Black, MargenIzquierdo + 9, yPos + 42)
        e.Graphics.DrawString(Comprobante.TipoComprobante.TipoComprobante, printFont, Brushes.Black, MargenIzquierdo + 58, yPos + 3)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("P.Vta:" & Comprobante.PuntoVta & "           Nro:" & Comprobante.NumComp, printFont, Brushes.Black, MargenIzquierdo + 58, yPos + 10)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("Fecha:" & Comprobante.FechaComp, printFont, Brushes.Black, MargenIzquierdo + 58, yPos + 10)

        Dim hora As String = TimeString
        e.Graphics.DrawString("Hora:" & hora, printFont, Brushes.Black, MargenIzquierdo + 183, yPos + 10)
        yPos += IncrementoYPreLinea + 15
        e.Graphics.DrawString(Linea, printFont, Brushes.Black, MargenIzquierdo, yPos)

        Dim strCliente As String = Left(Comprobante.Cliente.Nombre, 34)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("Cliente:" & strCliente, printFont, Brushes.Black, MargenIzquierdo, yPos)

        If Len(Comprobante.Cliente.Nombre) > 34 Then
            yPos += IncrementoYPreTexto
            strCliente = Mid(Comprobante.Cliente.Nombre, 35, 42)
            e.Graphics.DrawString(strCliente, printFont, Brushes.Black, MargenIzquierdo, yPos)
        End If

        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("Domicilio:" & Comprobante.Cliente.Localidad & "-" & Comprobante.Cliente.Provincia, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("IVA:" & Comprobante.Cliente.IVADescripcion, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("Tipo Doc:" & Comprobante.Cliente.TipoDoc, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("Num.Doc:" & Comprobante.Cliente.NumDoc, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreLinea
        e.Graphics.DrawString(Linea, printFont, Brushes.Black, MargenIzquierdo, yPos)
        'Fin del Encabezado

        'Aca comienza el detalle del comprobante
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("RECIBI(MOS) LA SUMA DE PESOS: ", printFont, Brushes.Black, MargenIzquierdo, yPos)

        Dim strTextoImporte As String = UCase(vecho.NumEnLetras(Format(Comprobante.ImpBto, "Fixed")))
        Dim CaracteresLeidos As Integer
        Dim TotalCaracteres As Integer = Len(strTextoImporte)
        Dim LeerCantidad As Integer
        Dim CaracteresRestantes As Integer
        Dim strTextoParcial As String

        If TotalCaracteres > 41 Then
            LeerCantidad = 41
        Else
            LeerCantidad = TotalCaracteres
        End If

        Do While LeerCantidad > 0
            strTextoParcial = Mid(strTextoImporte, CaracteresLeidos + 1, LeerCantidad)
            CaracteresLeidos += LeerCantidad

            yPos += IncrementoYPreTexto
            e.Graphics.DrawString(strTextoParcial, printFont, Brushes.Black, MargenIzquierdo, yPos)
            CaracteresRestantes = TotalCaracteres - CaracteresLeidos
            If CaracteresRestantes > 41 Then
                LeerCantidad = 41
            Else
                LeerCantidad = CaracteresRestantes
            End If
        Loop

        yPos += 2 * IncrementoYPreTexto
        e.Graphics.DrawString("EN CONCEPTO DE: ", printFont, Brushes.Black, MargenIzquierdo, yPos)

        For Each Item As ItemComprobante In Comprobante.Detalle

            yPos += IncrementoYPreItem
            e.Graphics.DrawString(Left(Item.Descripcion, 33), printFont, Brushes.Black, MargenIzquierdo, yPos)

        Next

        yPos += IncrementoYPreTexto
        strTotal = "$" & Format(Comprobante.ImpBto, "Fixed")
        e.Graphics.DrawString("SON PESOS: " & StrDup(11 - Len(strTotal), " ") & strTotal, fuenteGrande, Brushes.Black, MargenIzquierdo, yPos)
        yPos += 2 * IncrementoYPreTexto
        e.Graphics.DrawString("RECIBI(MOS)", printFont, Brushes.Black, MargenIzquierdo, yPos)

        If Comprobante.ImpTar > 0 Then
            yPos += IncrementoYPreTexto
            strTar = "$" & Format(Comprobante.ImpTar, "Fixed")
            e.Graphics.DrawString("Tarjeta/s: " & StrDup(31 - Len(strTar), " ") & strTar, printFont, Brushes.Black, MargenIzquierdo, yPos)
        End If

        If Comprobante.ImpEf > 0 Then
            yPos += IncrementoYPreTexto
            strEf = "$" & Format(Comprobante.ImpEf, "Fixed")
            e.Graphics.DrawString("Efectivo: " & StrDup(32 - Len(strEf), " ") & strEf, printFont, Brushes.Black, MargenIzquierdo, yPos)
        End If

        yPos += IncrementoYPreLinea
        e.Graphics.DrawString(Linea, printFont, Brushes.Black, MargenIzquierdo, yPos)
        yPos += IncrementoYPreTexto
        e.Graphics.DrawString("Documento no válido como Factura", printFont, Brushes.Black, MargenIzquierdo, yPos)

        If Comprobante.Operacion.Observaciones <> "" Then
            Dim intNumCarRestantes As Integer
            Dim strLineaObservaciones As String
            Dim x As Integer
            Dim y As Integer = 1
            Dim lineas() As String = Split(Comprobante.Operacion.Observaciones, vbCrLf)

            yPos += IncrementoYPreLinea
            e.Graphics.DrawString(Linea, printFont, Brushes.Black, MargenIzquierdo, yPos)
            yPos += IncrementoYPreTexto
            e.Graphics.DrawString("Observaciones:", printFont, Brushes.Black, MargenIzquierdo, yPos)

            For Each strLinea As String In lineas
                intNumCarRestantes = Len(strLinea)
                If intNumCarRestantes > 42 Then
                    Do While intNumCarRestantes > 0
                        x += 1
                        strLineaObservaciones = Mid(strLinea, y, 42)
                        y += 42
                        intNumCarRestantes = Len(strLinea) - 42 * x
                        yPos += IncrementoYPreTexto
                        e.Graphics.DrawString(strLineaObservaciones.TrimStart(" "), printFont, Brushes.Black, MargenIzquierdo, yPos)
                    Loop
                Else
                    yPos += IncrementoYPreTexto
                    e.Graphics.DrawString(strLinea.TrimStart(" "), printFont, Brushes.Black, MargenIzquierdo, yPos)
                End If
            Next

        End If
        e.Graphics.Dispose()

    End Sub

    Public Sub Imprimir(ByVal argCopia As String)

        Dim strImpresora As String = ParametrosTerminal.ReadINI("C:\SiCoFa_Cliente\config.ini", "SiCoFa", "strImpresoraFE")
        Dim printDoc As New PrintDocument()

        If strImpresora <> "" Then
            printDoc.PrinterSettings.PrinterName = strImpresora
        End If

        If Not printDoc.PrinterSettings.IsValid Then
            Throw New Exception("Error: No se puede encontrar la impresora")
            Exit Sub
        End If

        Copia = argCopia
        Select Case Comprobante.TipoComprobante.CodiTC_SiCoFa
            Case "FAA", "FAB", "FAC", "NCA", "NCB", "NCC", "NCX", "RECR", "PRESU", "RTO"
                AddHandler printDoc.PrintPage, AddressOf FACTURA
            Case "REC"
                AddHandler printDoc.PrintPage, AddressOf RECIBO
        End Select

        printDoc.Print()

        printDoc.Dispose()

    End Sub

    Private Function Bytes_Imagen(ByVal Foto As Byte()) As Image
        If Not Foto Is Nothing Then
            Dim Codi As New IO.MemoryStream(Foto)
            Dim resultado As Image = Image.FromStream(Codi)
            Return resultado
            Codi.Dispose()
            resultado.Dispose()

        Else
            Return Nothing
        End If
    End Function



End Class

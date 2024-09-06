Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports SiCoFa.Entidades

Public Class D_AdminDetalleComprobante
    Public Function ObtenerDetalle(ByVal argCodiTC As String, ByVal argIdOperacion As Long, ByVal argDisIva As Boolean) As List(Of ItemComprobante)
        Select Case argCodiTC
            Case "REC"
                ObtenerDetalle = ObtenerDetalleR(argIdOperacion, argDisIva)
            Case Else
                ObtenerDetalle = ObtenerDetalleC(argIdOperacion, argDisIva)
        End Select
    End Function
    Private Function ObtenerDetalleC(ByVal argIdOperacion As Long, ByVal argDisIva As Boolean) As List(Of ItemComprobante)
        Dim objDetC As New List(Of ItemComprobante)

        Try
            Dim Sql As String = "SELECT IdOperación,Descripcion,Cantidad,PUnit,Gravado,Descuento,MotivoDes,PDes FROM TblDetComprobantes WHERE IdOperación=" & argIdOperacion

            Dim cmd As SqlCommand = D_Admin.ConexionDB.conn.CreateCommand()
            cmd.CommandType = CommandType.Text
            cmd.CommandText = Sql
            Dim datos As SqlDataReader = cmd.ExecuteReader()
            Dim objItemC As ItemComprobante = Nothing

            While datos.Read()
                Dim Descripcion As String = datos("Descripcion")
                Dim Cantidad As Integer = datos("Cantidad")
                Dim PUnit As Decimal = datos("PUnit")
                Dim Gravado As Boolean = datos("Gravado")
                Dim Descuento As Decimal = datos("Descuento")
                Dim PDes As Decimal = datos("PDes")
                Dim MotivoDes As String = datos("MotivoDes")

                objItemC = New ItemComprobante(Descripcion, Cantidad, PUnit, Gravado, argDisIva, Descuento, PDes, MotivoDes)

                objDetC.Add(objItemC)
            End While

            datos.Close()
            cmd.Dispose()
            Return objDetC

        Catch ex As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "ObtenerDetalleC", ex.Message))
            Return Nothing
        End Try
    End Function
    Private Function ObtenerDetalleR(ByVal argIdOperacion As Long, ByVal argDisIVA As Boolean) As List(Of ItemComprobante)
        Dim objDetR As New List(Of ItemComprobante)

        Try
            Dim sql As String = "SELECT TipoOperación,Resu,ImpClientes FROM ConRecibo WHERE IdOperación=" & argIdOperacion

            Dim cmd As SqlCommand = D_Admin.ConexionDB.conn.CreateCommand()
            cmd.CommandType = CommandType.Text
            cmd.CommandText = sql
            Dim datos As SqlDataReader = cmd.ExecuteReader()
            Dim objItemR As ItemComprobante = Nothing
            Dim TipoOperacion As String
            Dim Importe As Decimal
            Dim Resumen As String

            datos.Read()
            TipoOperacion = Replace(datos("TipoOperación"), "Cancelación", "Canc.")
            Importe = datos("ImpClientes")
            If IsDBNull(datos("Resu")) Then
                Resumen = "NO APLICA"
            Else
                Resumen = Left(datos("Resu"), 2) & "/" & Right(datos("Resu"), 2)
            End If

            datos.Close()

            objItemR = New ItemComprobante(TipoOperacion, 1, Importe, 0, argDisIVA, 0, 0, "")
            objDetR.Add(objItemR)
            objItemR = New ItemComprobante("Resúmen: " & Resumen, 1, Importe, 0, argDisIVA, 0, 0, "")
            objDetR.Add(objItemR)

            cmd.Dispose()
            Return objDetR

        Catch ex As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "ObtenerDetalleR", ex.Message))
            Return Nothing
        End Try
    End Function

End Class
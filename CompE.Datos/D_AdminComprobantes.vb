Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports SiCoFa.Entidades

Public Class D_AdminComprobantes

    Private mobjD_AdminCliente As D_AdminClientes
    Private mobjD_AdminEmpresa As D_AdminEmpresa
    Private mobjD_AdminDetalleC As D_AdminDetalleComprobante
    Public Sub New()

        mobjD_AdminCliente = New D_AdminClientes
        mobjD_AdminEmpresa = New D_AdminEmpresa
        mobjD_AdminDetalleC = New D_AdminDetalleComprobante

    End Sub
    Public Sub RegistrarError(ByVal argIdOperacion As Long)

        Try
            Dim sql As String = "UPDATE TblComprobantes SET NumComp='Error' WHERE IdOperación=" & argIdOperacion
            Dim cmd As SqlCommand = New SqlCommand(sql, D_Admin.ConexionDB.conn)
            cmd.ExecuteNonQuery()
            cmd.Dispose()

        Catch ex As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "RegistrarError", ex.Message))

        End Try

    End Sub
    Public Sub ActualizarComprobante(ByVal argCbte As Comprobante)

        Try
            Dim cmd As SqlCommand = New SqlCommand("Actualizar_FE", D_Admin.ConexionDB.conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@IdOpera", argCbte.Operacion.IdOperacion)
            cmd.Parameters.AddWithValue("@NumComp", argCbte.NumComp)
            cmd.Parameters.AddWithValue("@CAE", argCbte.CAE.NumCAE)
            cmd.Parameters.AddWithValue("@VtoCAE", argCbte.CAE.VtoCAE)
            cmd.ExecuteNonQuery()
            cmd.Dispose()

        Catch Ex As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "ActualizarComprobante", Ex.Message))

        End Try

    End Sub
    Public Function ObtenerComprobante(ByVal argOperacion As Operacion) As Comprobante

        Try
            Dim objCbte As Comprobante
            Dim objCli As Cliente
            Dim objEmpr As Empresa
            Dim objCAE As CAE = Nothing
            Dim objDetalleC As List(Of ItemComprobante)
            Dim objDetalleR As List(Of ItemComprobanteRecetas) = Nothing
            Dim CodiTC As String
            Dim PrefComp As String
            Dim NumComp As String
            Dim IdCliente As Long
            Dim FechaComp As Date
            Dim ImpBto As Decimal
            Dim ImpEx As Decimal
            Dim ImpGrav As Decimal
            Dim ImpNeto As Decimal
            Dim ImpIVA As Decimal
            Dim ImpOS As Decimal
            Dim ImpEf As Decimal
            Dim ImpCC As Decimal
            Dim ImpTar As Decimal
            Dim ImpDes As Decimal
            Dim IdOperAsoc As Long
            Dim CAE As String
            Dim VtoCAE As Date

            objEmpr = mobjD_AdminEmpresa.ObtenerEmpresa
            Dim sql As String = "SELECT IdOperación,CodiTC,PrefComp,NumComp,FechaComp,IdCliente,ImpBto,ImpEx,ImpGrav,ImpNeto,ImpIVA,ImpOS,ImpEf,ImpCC,ImpTar,IdOperAsoc,CAE,VtoCAE,ImpDes FROM TblComprobantes WHERE IdOperación=" & argOperacion.IdOperacion

            Dim cmd As SqlCommand = D_Admin.ConexionDB.conn.CreateCommand()
            cmd.CommandType = CommandType.Text
            cmd.CommandText = sql
            Dim datosC As SqlDataReader = cmd.ExecuteReader()
            datosC.Read()

            If datosC.HasRows = False Then
                datosC.Close()
                cmd.Dispose()
                Throw New Exception("Comprobante no Encontrado")
            End If

            CodiTC = datosC("CodiTC")
            PrefComp = datosC("PrefComp")
            NumComp = datosC("NumComp").ToString
            IdCliente = datosC("IdCliente")
            FechaComp = datosC("FechaComp")
            ImpBto = datosC("ImpBto")
            ImpEx = datosC("ImpEx")
            ImpGrav = datosC("ImpGrav")
            ImpNeto = datosC("ImpNeto")
            ImpIVA = datosC("ImpIVA")
            ImpOS = datosC("ImpOS")
            ImpEf = datosC("ImpEf")
            ImpCC = datosC("ImpCC")
            ImpTar = datosC("ImpTar")
            ImpDes = datosC("ImpDes")
            IdOperAsoc = datosC("IdOperAsoc")
            CAE = datosC("CAE").ToString

            If datosC("VtoCAE") IsNot DBNull.Value Then
                VtoCAE = datosC("VtoCAE")
            End If

            datosC.Close()
            cmd.Dispose()

            objCli = mobjD_AdminCliente.ObtenerCliente(IdCliente)
            objDetalleC = mobjD_AdminDetalleC.ObtenerDetalle(CodiTC, argOperacion.IdOperacion, Me.DisIva(CodiTC))

            If ImpOS > 0 Then
                objDetalleR = Me.ObtenerDetalleRtas(argOperacion.IdOperacion)
            End If

            If CAE <> "" Then
                objCAE = New CAE(NumComp, CAE, VtoCAE)
            End If

            If IdOperAsoc > 0 Then
                Dim objCompAsoc As Comprobante = ObtenerComprobanteAsoc(IdOperAsoc)
                objCbte = New Comprobante(argOperacion, CodiTC, PrefComp, NumComp, FechaComp, ImpBto, ImpEx, ImpGrav, ImpNeto, ImpIVA, ImpOS, ImpEf, ImpCC, ImpTar, ImpDes, objCAE, objCli, objCompAsoc, objEmpr, objDetalleC, objDetalleR)
            Else
                objCbte = New Comprobante(argOperacion, CodiTC, PrefComp, NumComp, FechaComp, ImpBto, ImpEx, ImpGrav, ImpNeto, ImpIVA, ImpOS, ImpEf, ImpCC, ImpTar, ImpDes, objCAE, objCli, Nothing, objEmpr, objDetalleC, objDetalleR)
            End If

            Return objCbte

        Catch ex As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "ObtenerComprobante", ex.Message))
            Return Nothing

        End Try

    End Function
    Private Function ObtenerComprobanteAsoc(ByVal argIdOperAsoc As Long) As Comprobante

        Try
            Dim objD_AdminOpera As New D_AdminOperaciones
            Dim objOperaAsoc As Operacion = objD_AdminOpera.ObtenerOperacion(argIdOperAsoc)
            Dim objCbte As Comprobante
            Dim objCli As Cliente
            Dim objEmpr As Empresa
            Dim objCAE As CAE = Nothing
            Dim objDetalleC As List(Of ItemComprobante)
            Dim objDetalleR As List(Of ItemComprobanteRecetas) = Nothing
            Dim CodiTC As String
            Dim PrefComp As String
            Dim NumComp As String
            Dim IdCliente As Long
            Dim FechaComp As Date
            Dim ImpBto As Decimal
            Dim ImpEx As Decimal
            Dim ImpGrav As Decimal
            Dim ImpNeto As Decimal
            Dim ImpIVA As Decimal
            Dim ImpOS As Decimal
            Dim ImpEf As Decimal
            Dim ImpCC As Decimal
            Dim ImpTar As Decimal
            Dim ImpDes As Decimal
            Dim IdOperAsoc As Long
            Dim CAE As String
            Dim VtoCAE As Date

            objEmpr = mobjD_AdminEmpresa.ObtenerEmpresa
            Dim sql As String = "SELECT IdOperación,CodiTC,PrefComp,NumComp,FechaComp,IdCliente,ImpBto,ImpEx,ImpGrav,ImpNeto,ImpIVA,ImpOS,ImpEf,ImpCC,ImpTar,IdOperAsoc,CAE,VtoCAE,ImpDes FROM TblComprobantes WHERE IdOperación=" & argIdOperAsoc

            Dim cmd As SqlCommand = D_Admin.ConexionDB.conn.CreateCommand()
            cmd.CommandType = CommandType.Text
            cmd.CommandText = sql
            Dim datosC As SqlDataReader = cmd.ExecuteReader()
            datosC.Read()

            If datosC.HasRows = False Then
                datosC.Close()
                cmd.Dispose()
                Throw New Exception("Comprobante no Encontrado")
            End If

            CodiTC = datosC("CodiTC")
            PrefComp = datosC("PrefComp")
            NumComp = datosC("NumComp").ToString
            IdCliente = datosC("IdCliente")
            FechaComp = datosC("FechaComp")
            ImpBto = datosC("ImpBto")
            ImpEx = datosC("ImpEx")
            ImpGrav = datosC("ImpGrav")
            ImpNeto = datosC("ImpNeto")
            ImpIVA = datosC("ImpIVA")
            ImpOS = datosC("ImpOS")
            ImpEf = datosC("ImpEf")
            ImpCC = datosC("ImpCC")
            ImpTar = datosC("ImpTar")
            ImpDes = datosC("ImpDes")
            IdOperAsoc = datosC("IdOperAsoc")
            CAE = datosC("CAE").ToString

            If datosC("VtoCAE") IsNot DBNull.Value Then
                VtoCAE = datosC("VtoCAE")
            End If

            datosC.Close()
            cmd.Dispose()

            objCli = mobjD_AdminCliente.ObtenerCliente(IdCliente)
            objDetalleC = mobjD_AdminDetalleC.ObtenerDetalle(CodiTC, objOperaAsoc.IdOperacion, Me.DisIva(CodiTC))

            If ImpOS > 0 Then
                objDetalleR = Me.ObtenerDetalleRtas(argIdOperAsoc)
            End If

            If CAE <> "" Then
                objCAE = New CAE(NumComp, CAE, VtoCAE)
            End If
            objCbte = New Comprobante(objOperaAsoc, CodiTC, PrefComp, NumComp, FechaComp, ImpBto, ImpEx, ImpGrav, ImpNeto, ImpIVA, ImpOS, ImpEf, ImpCC, ImpTar, ImpDes, objCAE, objCli, Nothing, objEmpr, objDetalleC, objDetalleR)
            Return objCbte

        Catch ex As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "ObtenerComprobanteAsoc", ex.Message))
            Return Nothing

        End Try

    End Function
    Private Function ObtenerDetalleRtas(ByVal argIdOperacion As Long) As List(Of ItemComprobanteRecetas)
        Dim objDetR As New List(Of ItemComprobanteRecetas)

        Try
            Dim Sql As String = "SELECT NombreOS,CantRtas,ImpOS FROM ConRtasPorIdOpera WHERE IdOperación=" & argIdOperacion

            Dim cmd As SqlCommand = D_Admin.ConexionDB.conn.CreateCommand()
            cmd.CommandType = CommandType.Text
            cmd.CommandText = Sql
            Dim datos As SqlDataReader = cmd.ExecuteReader()
            Dim objItemR As ItemComprobanteRecetas = Nothing

            While datos.Read()
                Dim NombreOS As String = datos("NombreOS")
                Dim CantRtas As Integer = datos("CantRtas")
                Dim ImpOS As Decimal = datos("ImpOS")

                objItemR = New ItemComprobanteRecetas(NombreOS, CantRtas, ImpOS)

                objDetR.Add(objItemR)
            End While

            datos.Close()
            cmd.Dispose()
            Return objDetR

        Catch ex As Exception
            Throw New Exception(vecho.MensajeError(Me.ToString, "ObtenerDetalleRtas", ex.Message))
            Return Nothing
        End Try
    End Function
    Private Function DisIva(argCodiTC_SiCoFa As String) As Boolean
        Select Case argCodiTC_SiCoFa
            Case "FAA", "NCA", "FAM", "NCM"
                Return True
            Case Else
                Return False
        End Select
    End Function

End Class
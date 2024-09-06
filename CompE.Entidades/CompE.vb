Imports System.Data
Imports System.Data.SqlClient
Imports CompE.DatoSiCoFa

Public Class CompE

    Property ComprobantesEnCola As List(Of Comprobante)

    Public Sub New()
        Call ObtenerComprobantesEnCola()
    End Sub

    Public Sub ObtenerComprobantesEnCola()
        Dim CompEnC As List(Of Comprobante) = New List(Of Comprobante)
        Dim Sql As String = ""

        Try

            Sql = "SELECT IdOperación,CodiTC,PrefComp,FechaComp,ImpBto,ImpEx,ImpGrav,IdCliente FROM TblComprobantes WHERE IdOperación=231"

            Dim Db As Datos = New Datos()
            Db.Conectar()

            Db.CrearComando(Sql)

            Dim Datos As SqlDataReader = Db.EjecutarConsulta()
            Dim c As Comprobante = Nothing

            While Datos.Read()
                Try
                    c = New Comprobante(Datos("IdOperación"), Datos("CodiTC"), Datos("PrefComp"), Datos("FechaComp"), Datos("ImpBto"), Datos("ImpEx"), Datos("ImpGrav"), Datos("IdCliente"))

                    CompEnC.Add(c)
                Catch ex As InvalidCastException
                    Throw New NegocioException("Los tipos no coinciden.", ex)
                Catch ex As DataException
                    Throw New NegocioException("Error de ADO.NET.", ex)
                End Try
            End While

            Datos.Close()
            Db.Desconectar()
        Catch ex As DatosException
            Throw New NegocioException("Error al acceder a la base de datos para obtener los Comprobantes.")
        Catch ex As NegocioException
            Throw New NegocioException("Error al obtener los Comprobantes.")
        Finally

        End Try

        ComprobantesEnCola = CompEnC
    End Sub

End Class

Public Class Comprobante
    Property IdOpera As Long
    Property CodiTC As String
    Property PuntoVta As String
    Property NumComp As String
    Property FechaComp As Date
    Property ImpBto As Decimal
    Property ImpEx As Decimal
    Property ImpGrav As Decimal
    Property CAE As CAE
    Property IdCliente As Long
    Property Cliente As Cliente
    Public Sub New(
                  ByVal argIdOpera As Long,
                  ByVal argCodiTC As String,
                  ByVal argPuntoVta As String,
                  ByVal argFechaComp As Date,
                  ByVal argImpBto As Decimal,
                  ByVal argImpEx As Decimal,
                  ByVal argImpGrav As Decimal,
                  ByVal argIdCliente As Long
                  )
        Me.IdOpera = argIdOpera
        Me.CodiTC = argCodiTC
        Me.PuntoVta = argPuntoVta
        Me.FechaComp = argFechaComp
        Me.ImpBto = argImpBto
        Me.ImpEx = argImpEx
        Me.ImpGrav = argImpGrav
        Me.IdCliente = argIdCliente
        Me.Cliente = New Cliente(argIdCliente)
    End Sub

    Public Function ObtenerCAE() As String
        Me.CAE = New CAE

        If CAE.SolicitarCAE("20210362712", 1, 4, CbteTipo, DocTipo, DocNro, Me.ImpEx, Me.ImpGrav) = "A" Then
            Me.NumComp = CAE.NroCompAutorizado
            Return "A"
        Else
            Return "R"
        End If

    End Function

    Private Function CbteTipo() As Integer
        Select Case Me.CodiTC
            Case "FAA"
                Return 1
            Case "FAB"
                Return 6
            Case "FAC"
                Return 11
            Case "NCA"
                Return 112
            Case "NCB"
                Return 113
            Case "NCC"
                Return 114
        End Select
    End Function
    Private Function DocTipo() As Integer
        Select Case Me.Cliente.IVA
            Case "CF"
                Return 96
            Case "RI", "EX"
                Return 80
        End Select
    End Function

    Private Function DocNro() As String
        Select Case Me.Cliente.IVA
            Case "CF"
                Return Me.Cliente.DNI
            Case "RI", "EX"
                Return Me.Cliente.CUIT
        End Select
    End Function
End Class

Public Class DetalleComprobante

End Class

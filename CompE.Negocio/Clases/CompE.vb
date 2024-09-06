Imports System.Data
Imports System.Data.SqlClient
Imports DatoSiCoFa

Public Class CompE

    Property ComprobantesEnCola As List(Of Comprobante)

    Public Function ObtenerComprobantesEnCola() As List(Of Comprobante)
        Dim CompEnC As List(Of Comprobante) = New List(Of Comprobante)
        Dim Sql As String = ""

        Try

            Sql = "SELECT CodiTC,PrefComp,Fecha,ImpBto,ImpEx,ImpGrav FROM ConCompEnCola"

            Dim Db As BaseDatos = New BaseDatos()
            Db.Conectar()

            Db.CrearComando(Sql)

            Dim Datos As SqlDataReader = Db.EjecutarConsulta()
            Dim c As Comprobante = Nothing

            While Datos.Read()
                Try
                    c = New Comprobante(Datos.GetValue(0), Datos.GetValue(1), Datos.GetValue(2), Datos.GetValue(3), Datos.GetValue(4), Datos.GetValue(5))
                    CompEnC.Add(c)
                Catch ex As InvalidCastException
                    Throw New NegocioException("Los tipos no coinciden.", ex)
                Catch ex As DataException
                    Throw New NegocioException("Error de ADO.NET.", ex)
                End Try
            End While

            Datos.Close()
            Db.Desconectar()
        Catch ex As BaseDatosException
            Throw New NegocioException("Error al acceder a la base de datos para obtener los productos.")
        Catch ex As NegocioException
            Throw New NegocioException("Error al obtener los productos.")
        Finally

        End Try

        Return CompEnC
    End Function

End Class

Public Class Comprobante
    Property CodiTC As String
    Property PrefComp As String
    Property NumComp As String
    Property FechaComp As Date
    Property ImpBto As Decimal
    Property ImpEx As Decimal
    Property ImpGrav As Decimal
    Property CAE As CAE
    Property Cliente As Cliente

    Public Sub New(
                  ByVal argCodiTC As String,
                  ByVal argPrefComp As String,
                  ByVal argFechaComp As Date,
                  ByVal argImpBto As Decimal,
                  ByVal argImpEx As Decimal,
                  ByVal argImpGrav As Decimal
                  )
        CodiTC = argCodiTC
        PrefComp = argPrefComp
        FechaComp = argFechaComp
        ImpBto = argImpBto
        ImpEx = argImpEx
        ImpGrav = argImpGrav
    End Sub
End Class


Public Class Operacion
    Property IdOperacion() As Long
    Property Inicio() As Date
    Property Fin() As Date
    Property IdEmpr() As Integer
    Property IdPC() As String
    Property IdCaja() As Integer
    Property IdEmpleado() As Integer
    Property CodiTO() As String
    Property Estado() As String
    Property Observaciones() As String
    Public Sub New(
                    ByVal argIdOperacion As Long,
                    ByVal argInicio As Date,
                    ByVal argFin As Date,
                    ByVal argIdEmpr As Integer,
                    ByVal argIdPc As String,
                    ByVal argIdCaja As Integer,
                    ByVal argIdEmpleado As Integer,
                    ByVal argCodiTO As String,
                    ByVal argEstado As String,
                    ByVal argObservaciones As String
                    )

        Me.IdOperacion = argIdOperacion
        Me.Inicio = argInicio
        Me.Fin = argFin
        Me.IdEmpr = argIdEmpr
        Me.IdPC = argIdPc
        Me.IdCaja = argIdCaja
        Me.IdEmpleado = argIdEmpleado
        Me.CodiTO = argCodiTO
        Me.Estado = argEstado
        Me.Observaciones = argObservaciones
    End Sub

End Class
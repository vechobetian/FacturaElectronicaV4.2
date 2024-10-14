Imports System.Data.SqlClient
Imports SiCoFa.Entidades

Public Class Conexion
    Property conn As SqlConnection
    Property Servidor As String
    Property Usuario As String
    Property Clave As String
    Property Seguridad As Boolean = True
    Public Sub New()

        Dim strPS As String = ParamTerminal.ReadINI("C:\SiCoFa_Cliente\config.ini", "SiCoFa", "strPathS")

        If strPS = "C:\" Then
            strPS = "(local)\"
        Else
            strPS = Replace(strPS, "\\", "")
        End If

        Me.Servidor = strPS & "SQLEXPRESS"
        Me.Usuario = "usuario_sicofa"
        Me.Clave = "vecho"

        If Me.conn Is Nothing Then
            Me.conn = New SqlConnection(CrearCadena)
            If conn.State = ConnectionState.Closed Then
                Me.conn.Open()
            End If
        End If

    End Sub
    Public Function CrearCadena() As String
        Dim cadena As String
        cadena = "Data Source=" & Me.Servidor & ";Initial Catalog=SiCoFa_Org;User Id=" & Me.Usuario & ";Password=" & Me.Clave
        Return cadena
    End Function
    Public Sub CerrarConexion()
        If conn IsNot Nothing AndAlso conn.State = ConnectionState.Open Then
            conn.Close()
        End If
    End Sub

End Class
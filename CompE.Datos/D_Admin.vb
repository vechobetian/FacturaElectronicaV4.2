Module D_Admin
    Private mobjConexionDB As Conexion

    Public Function ConexionDB() As Conexion
        If mobjConexionDB Is Nothing Then
            mobjConexionDB = New Conexion
        End If

        Return mobjConexionDB

    End Function

End Module
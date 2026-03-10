Imports SiCoFa.Entidades
Module Mod_D_Admin
    Public Function strConexionDB() As String

        Dim PaTer As New ParametrosTerminal

        Dim Servidor As String = PaTer.ServerSql
        Dim Usuario As String = "usuario_sicofa"
        Dim Clave As String = "vecho"

        Dim cadena As String = "Data Source=" & Servidor & ";Initial Catalog=SiCoFa_Org;User Id=" & Usuario & ";Password=" & Clave & ";Encrypt=True;TrustServerCertificate=True;Pooling=True;Min Pool Size=5;Max Pool Size=100;"
        Return cadena

    End Function

End Module
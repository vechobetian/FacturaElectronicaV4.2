Imports SiCoFa.Entidades
Module Mod_D_Admin
    Public Function strConexionDB() As String

        Dim PaTer As New ParametrosTerminal
        Dim strPS As String

        If PaTer.PathServer = "C:\" Then
            strPS = "(local)\"
        Else
            strPS = Replace(PaTer.PathServer, "\\", "")
        End If

        Dim Servidor As String = strPS & "SQLEXPRESS"
        Dim Usuario As String = "usuario_sicofa"
        Dim Clave As String = "vecho"

        Dim cadena As String = "Data Source=" & Servidor & ";Initial Catalog=SiCoFa_Org;User Id=" & Usuario & ";Password=" & Clave & ";Pooling=True;Min Pool Size=5;Max Pool Size=100;"
        Return cadena

    End Function

End Module
Imports System.Collections.Generic
Imports System.Linq
Imports System.Data.SqlClient
Public Class D_AdminDB

    Private Sub AplicarFiltrosAComando(cmd As SqlCommand, sqlBase As String, filtros As Dictionary(Of String, Object))
        If filtros IsNot Nothing AndAlso filtros.Count > 0 Then
            Dim listaWhere As New List(Of String)
            For Each filtro In filtros
                ' Se usan corchetes [] para compatibilidad con SQL Server
                listaWhere.Add($"[{filtro.Key}] = @f_{filtro.Key}")
                cmd.Parameters.AddWithValue($"@f_{filtro.Key}", filtro.Value)
            Next
            cmd.CommandText = $"{sqlBase} WHERE {String.Join(" AND ", listaWhere)}"
        Else
            cmd.CommandText = sqlBase
        End If
    End Sub

    Public Function ObtenerValor(tabla As String, campo As String, Optional filtros As Dictionary(Of String, Object) = Nothing) As Object
        Try
            ' Usamos TOP 1 por seguridad, ya que ExecuteScalar solo lee la primera celda de la primera fila
            Dim sqlBase As String = $"SELECT TOP 1 [{campo}] FROM [{tabla}]"
            Dim valor As Object

            Using cn As New SqlConnection(Mod_D_Admin.strConexionDB)
                cn.Open()

                Using cmd As New SqlCommand()
                    cmd.Connection = cn

                    ' Reutilizamos el método auxiliar que te pasé antes para armar el WHERE
                    AplicarFiltrosAComando(cmd, sqlBase, filtros)

                    valor = cmd.ExecuteScalar()
                End Using
            End Using

            ' Verificamos tanto Nothing como DBNull (el nulo real de SQL Server)
            If valor IsNot Nothing AndAlso Not IsDBNull(valor) Then
                Return valor
            Else
                Return Nothing
            End If

        Catch ex As Exception
            Throw New Exception(Vecho.MensajeError(Me.ToString, "ObtenerValor", ex.Message))
        End Try
    End Function

    Public Function IncrementarYObtenerContador(tabla As String, campo As String, filtros As Dictionary(Of String, Object)) As Long
        Try
            If filtros Is Nothing OrElse filtros.Count = 0 Then
                Throw New ArgumentException("Debe proporcionar un filtro para identificar el contador.")
            End If

            ' UPDATE con OUTPUT 'deleted' actualiza el valor y devuelve el valor ANTERIOR al +1,
            ' exactamente como lo hacía tu lógica original.
            Dim sqlBase As String = $"UPDATE [{tabla}] SET [{campo}] = [{campo}] + 1 OUTPUT deleted.[{campo}]"

            Using cn As New SqlConnection(Mod_D_Admin.strConexionDB)
                cn.Open()

                Using cmd As New SqlCommand()
                    cmd.Connection = cn

                    ' Reutilizamos nuestra función para inyectar el WHERE y los parámetros
                    AplicarFiltrosAComando(cmd, sqlBase, filtros)

                    Dim resultado As Object = cmd.ExecuteScalar()

                    If resultado IsNot Nothing AndAlso Not IsDBNull(resultado) Then
                        Return Convert.ToInt64(resultado)
                    Else
                        Return 0
                    End If
                End Using
            End Using

        Catch ex As Exception
            Throw New Exception(Vecho.MensajeError(Me.ToString, "IncrementarYObtenerContador", ex.Message))
        End Try
    End Function

    Public Function ObtenerRegistro(tabla As String, Optional columnas As String = "*", Optional filtros As Dictionary(Of String, Object) = Nothing) As Dictionary(Of String, Object)
        Try
            ' Agregamos TOP 1 para que el motor SQL deje de buscar apenas encuentre la primera coincidencia
            Dim sqlBase As String = $"SELECT TOP 1 {columnas} FROM [{tabla}]"

            Using cn As New SqlConnection(Mod_D_Admin.strConexionDB)
                cn.Open()

                Using cmd As New SqlCommand()
                    cmd.Connection = cn

                    ' Aplicamos los filtros de forma parametrizada y segura
                    AplicarFiltrosAComando(cmd, sqlBase, filtros)

                    Using datos As SqlDataReader = cmd.ExecuteReader()

                        If datos.Read() Then
                            Dim resultado As New Dictionary(Of String, Object)()

                            For i As Integer = 0 To datos.FieldCount - 1
                                Dim nombreCampo As String = datos.GetName(i)
                                Dim valorCampo As Object = datos.GetValue(i)

                                ' MEJORA: Transformar DBNull de SQL a Nothing de .NET
                                ' Esto evita que tu aplicación falle si intentas leer un campo vacío
                                If IsDBNull(valorCampo) Then
                                    resultado(nombreCampo) = Nothing
                                Else
                                    resultado(nombreCampo) = valorCampo
                                End If
                            Next

                            Return resultado
                        Else
                            Return Nothing
                        End If

                    End Using
                End Using
            End Using

        Catch ex As Exception
            Throw New Exception(Vecho.MensajeError(Me.ToString, "ObtenerRegistro", ex.Message))
        End Try
    End Function

    Public Function InsertarRegistro(tabla As String, valores As Dictionary(Of String, Object)) As Integer
        Try
            ' Validamos que el diccionario tenga datos para evitar errores
            If valores Is Nothing OrElse valores.Count = 0 Then Return 0

            ' 1. CORRECCIÓN: Usar corchetes [...] en lugar de acentos graves `...` para SQL Server
            Dim columnas As String = String.Join(", ", valores.Keys.Select(Function(k) $"[{k}]"))
            Dim parametrosNombres As String = String.Join(", ", valores.Keys.Select(Function(k) $"@{k}"))

            ' 2. Renombramos argSql a 'tabla' porque conceptualmente es lo que recibe
            Dim sql As String = $"INSERT INTO [{tabla}] ({columnas}) VALUES ({parametrosNombres})"

            Using cn As New SqlConnection(Mod_D_Admin.strConexionDB)
                cn.Open()

                Using cmd As New SqlCommand(sql, cn)

                    For Each item In valores
                        ' 3. CORRECCIÓN: Si tu código .NET envía un "Nothing", SQL Server necesita recibir "DBNull.Value"
                        Dim valorAInsertar As Object = If(item.Value Is Nothing, DBNull.Value, item.Value)
                        cmd.Parameters.AddWithValue($"@{item.Key}", valorAInsertar)
                    Next

                    ' Devolvemos la cantidad de filas insertadas (debería ser 1)
                    Return cmd.ExecuteNonQuery()

                End Using
            End Using

        Catch ex As Exception
            ' Mantenemos tu formato de excepciones
            Throw New Exception(Vecho.MensajeError(Me.ToString, "InsertarRegistro", $"Error: {ex.Message}{Environment.NewLine}StackTrace: {ex.StackTrace}"))
        End Try
    End Function

    Public Function CuentaRegistros(tabla As String, Optional filtros As Dictionary(Of String, Object) = Nothing) As Integer
        Try
            Dim cantidad As Integer = 0

            ' Armamos la consulta base utilizando COUNT(*)
            Dim sqlBase As String = $"SELECT COUNT(*) FROM [{tabla}]"

            Using cn As New SqlConnection(Mod_D_Admin.strConexionDB)
                cn.Open()

                Using cmd As New SqlCommand()
                    cmd.Connection = cn

                    ' Aplicamos los filtros dinámicos (WHERE) que creamos en el primer paso
                    AplicarFiltrosAComando(cmd, sqlBase, filtros)

                    ' ExecuteScalar devuelve la primera celda (el resultado del COUNT)
                    Dim resultado As Object = cmd.ExecuteScalar()

                    ' Validación de seguridad por si la tabla no existe o hay un error extraño que devuelva nulo
                    If resultado IsNot Nothing AndAlso Not IsDBNull(resultado) Then
                        cantidad = Convert.ToInt32(resultado)
                    End If

                End Using
            End Using

            Return cantidad

        Catch ex As Exception
            Throw New Exception(Vecho.MensajeError(Me.ToString, "CuentaRegistros", ex.Message))
        End Try
    End Function

    Public Function EliminarRegistros(tabla As String, filtros As Dictionary(Of String, Object)) As Integer
        Try
            ' Medida de seguridad: Evitar un DELETE sin filtros que borre toda la tabla por accidente
            If filtros Is Nothing OrElse filtros.Count = 0 Then
                Throw New ArgumentException("Operación cancelada: Debe proporcionar al menos un filtro para evitar el borrado completo de la tabla.")
            End If

            Dim filasAfectadas As Integer = 0
            Dim sqlBase As String = $"DELETE FROM [{tabla}]"

            Using cn As New SqlConnection(Mod_D_Admin.strConexionDB)
                cn.Open()

                Using cmd As New SqlCommand()
                    cmd.Connection = cn

                    ' Reutilizamos nuestra función para inyectar el WHERE y los parámetros de forma segura
                    AplicarFiltrosAComando(cmd, sqlBase, filtros)

                    ' ExecuteNonQuery ejecuta el DELETE y devuelve cuántas filas se borraron
                    filasAfectadas = cmd.ExecuteNonQuery()
                End Using

            End Using

            Return filasAfectadas

        Catch ex As Exception
            Throw New Exception(Vecho.MensajeError(Me.ToString, "EliminarRegistros", ex.Message))
        End Try
    End Function

    Public Function ActualizarCampo(tabla As String, campo As String, valor As Object, filtros As Dictionary(Of String, Object)) As Integer
        Try
            ' Medida de seguridad: Evitar un UPDATE masivo por enviar un filtro vacío
            If filtros Is Nothing OrElse filtros.Count = 0 Then
                Throw New ArgumentException("Operación cancelada: Debe proporcionar al menos un filtro para evitar actualizar todos los registros de la tabla.")
            End If

            Dim filasAfectadas As Integer = 0

            ' Usamos corchetes para proteger nombres de tablas o campos con espacios
            Dim sqlBase As String = $"UPDATE [{tabla}] SET [{campo}] = @v_nuevoValor"

            Using cn As New SqlConnection(Mod_D_Admin.strConexionDB)
                cn.Open()

                Using cmd As New SqlCommand()
                    cmd.Connection = cn

                    ' 1. Parametrizamos el nuevo valor (y verificamos si es Nothing para enviar DBNull)
                    Dim valorAActualizar As Object = If(valor Is Nothing, DBNull.Value, valor)
                    cmd.Parameters.AddWithValue("@v_nuevoValor", valorAActualizar)

                    ' 2. Reutilizamos la función auxiliar para inyectar el WHERE dinámicamente
                    AplicarFiltrosAComando(cmd, sqlBase, filtros)

                    ' Ejecutamos la consulta
                    filasAfectadas = cmd.ExecuteNonQuery()
                End Using
            End Using

            Return filasAfectadas

        Catch ex As Exception
            Throw New Exception(Vecho.MensajeError(Me.ToString, "ActualizarCampo", ex.Message))
        End Try
    End Function

End Class
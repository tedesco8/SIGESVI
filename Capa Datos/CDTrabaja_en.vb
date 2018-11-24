Imports System.Data.Odbc
Imports System.Data.SqlClient
Public Class CDTrabaja_en
    Private ObjConexion As New CDConexion
    Dim cn As New OdbcConnection
    Dim cm As New OdbcCommand
    Dim da As New OdbcDataAdapter
    Dim dr As OdbcDataReader
    Dim Tabla As New Data.DataTable

    Public Function listartrabajaen() As DataTable
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandText = "SELECT ci, id_sucursal, nombre, apellido FROM trabaja_en WHERE estado = 't'"
        dr = cm.ExecuteReader
        Try
            Tabla.Load(dr)
            ObjConexion.Cerrar()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return Tabla
        cn.Dispose()
        cm.Dispose()
        dr.Dispose()
    End Function
    Public Function registrarTrabajanen(ByVal ci As Integer, ByVal id_sucursal As Integer, ByVal nombre As String, ByVal apellido As String)
        cm.Connection = ObjConexion.Conectar
        cm.CommandText = "insert into trabaja_en (ci, id_sucursal, nombre, apellido, estado) values ('" & ci & "','" & id_sucursal & "', '" & nombre & "', '" & apellido & "', 't' )"
        Try
            cm.ExecuteNonQuery()
            cm.Parameters.Clear()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return cm
        cm.Dispose()
    End Function
    Public Function modificarTrabajaen(ByVal ci, ByVal id_sucursal, ByVal nombre, ByVal apellido)
        Dim sql As String = ""
        cm.CommandType = CommandType.Text
        cm.Connection = ObjConexion.Conectar
        sql = "UPDATE trabaja_en SET "
        sql += "id_sucursal = " & id_sucursal & ", "
        sql += "nombre = '" & nombre & "', "
        sql += "apellido = '" & apellido & "' "
        sql += "WHERE ci = " & ci & ""
        cm.CommandText = sql
        Try
            cm.ExecuteNonQuery()
            cm.Parameters.Clear()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return cm
        cm.Dispose()
    End Function
    Public Function borrarTrabajaen(ByVal ci)
        Dim sql As String = ""
        cm.CommandType = CommandType.Text
        cm.Connection = ObjConexion.Conectar
        sql = "UPDATE trabaja_en SET "
        sql += "estado = 'f' "
        sql += "WHERE ci = " & ci & ""
        cm.CommandText = sql
        Try
            cm.ExecuteNonQuery()
            cm.Parameters.Clear()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return cm
        cm.Dispose()
    End Function
    Public Function restaurarTrabajaen(ByVal ci)
        Dim sql As String = ""
        cm.CommandType = CommandType.Text
        cm.Connection = ObjConexion.Conectar
        sql = "UPDATE trabaja_en SET "
        sql += "estado = 't' "
        sql += "WHERE ci = " & ci & ""
        cm.CommandText = sql
        Try
            cm.ExecuteNonQuery()
            cm.Parameters.Clear()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return cm
        cm.Dispose()
    End Function
    Public Function eliminarTrabajaen(ByVal ci)
        cm.Connection = ObjConexion.Conectar
        cm.CommandText = "delete from trabaja_en where ci = " & ci & ""
        Try
            cm.ExecuteNonQuery()
            cm.Parameters.Clear()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return cm
        cm.Dispose()
    End Function
End Class

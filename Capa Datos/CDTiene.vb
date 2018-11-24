Imports System.Data.Odbc
Imports System.Data.SqlClient
Public Class CDTiene
    Private ObjConexion As New CDConexion
    Dim cn As New OdbcConnection
    Dim cm As New OdbcCommand
    Dim da As New OdbcDataAdapter
    Dim dr As OdbcDataReader
    Dim Tabla As New Data.DataTable
    Public Function listarTiene() As DataTable
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandText = "SELECT id_lote, nombre_proceso, nombre_estado, fecha_inicio, fecha_final FROM tiene WHERE estado = 't' "
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
    
    Public Function registrarTiene(ByVal id_lote As Integer, ByVal nombre_proceso As String, ByVal nombre_estado As String, ByVal fecha_inicio As Integer, ByVal fecha_final As Integer)
        cm.Connection = ObjConexion.Conectar
        Try
            cm.CommandText = "insert into tiene (id_lote, nombre_proceso, nombre_estado, fecha_inicio, fecha_final, estado) values ('" & id_lote & "','" & nombre_proceso & "', '" & nombre_estado & "', '" & fecha_inicio & "', '" & fecha_final & "','t')"
            cm.ExecuteNonQuery()
            'cm.Parameters.Clear()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return cm
        cm.Dispose()
    End Function
    Public Function modificarTiene(ByVal id_lote, ByVal nombre_proceso, ByVal nombre_estado, ByVal fecha_inicio, ByVal fecha_final)
        Dim sql As String = ""
        cm.CommandType = CommandType.Text
        cm.Connection = ObjConexion.Conectar
        sql = "UPDATE tiene SET "
        sql += "nombre_proceso = '" & nombre_proceso & "', "
        sql += "nombre_estado = '" & nombre_estado & "', "
        sql += "fecha_inicio = " & fecha_inicio & ", "
        sql += "fecha_final = " & fecha_final & " "
        sql += "WHERE id_lote = " & id_lote & ""
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
    Public Function borrarTiene(ByVal id_lote)
        Dim sql As String = ""
        cm.CommandType = CommandType.Text
        cm.Connection = ObjConexion.Conectar
        sql = "UPDATE tiene SET "
        sql += "estado = 'f' "
        sql += "WHERE id_lote = " & id_lote & ""
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
    Public Function restaurarTiene(ByVal id_lote)
        Dim sql As String = ""
        cm.CommandType = CommandType.Text
        cm.Connection = ObjConexion.Conectar
        sql = "UPDATE tiene SET "
        sql += "estado = 't' "
        sql += "WHERE id_lote = " & id_lote & ""
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
End Class

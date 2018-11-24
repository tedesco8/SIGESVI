Imports System.Data.Odbc
Imports System.Data.SqlClient
Public Class CDEstan
    Private ObjConexion As New CDConexion
    Dim cn As New OdbcConnection
    Dim cm As New OdbcCommand
    Dim da As New OdbcDataAdapter
    Dim dr As OdbcDataReader
    Dim Tabla As New Data.DataTable

    Public Function listarlotes_estan() As DataTable
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandText = "SELECT id_lote, id_sucursal, cantidad FROM estan WHERE estado = 't'"
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
    Public Function registrarlotes_estan(ByVal id_lote As Integer, ByVal id_sucursal As Integer, ByVal cantidad As Integer)
        cm.Connection = ObjConexion.Conectar
        cm.CommandText = "insert into estan (id_lote, id_sucursal, cantidad, estado) values ('" & id_lote & "','" & id_sucursal & "', '" & cantidad & "', 't' )"
        Try
            cm.ExecuteNonQuery()
            cm.Parameters.Clear()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return cm
        cm.Dispose()
    End Function
    Public Function modificarlotes_estan(ByVal id_lote, ByVal id_sucursal, ByVal cantidad)
        Try
            Dim sql As String = ""
            cm.CommandType = CommandType.Text
            cm.Connection = ObjConexion.Conectar
            sql = "UPDATE estan SET "
            sql += "id_sucursal = " & id_sucursal & ", "
            sql += "cantidad = " & cantidad & " "
            sql += "WHERE id_lote = " & id_lote & ""
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return cm
        cm.Dispose()
    End Function
    Public Function borrarEstan(ByVal id_lote)
        Dim sql As String = ""
        cm.CommandType = CommandType.Text
        cm.Connection = ObjConexion.Conectar
        sql = "UPDATE estan SET "
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
    Public Function restaurarEstan(ByVal id_lote)
        Dim sql As String = ""
        cm.CommandType = CommandType.Text
        cm.Connection = ObjConexion.Conectar
        sql = "UPDATE estan SET "
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

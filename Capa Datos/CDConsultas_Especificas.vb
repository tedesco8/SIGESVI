Imports System.Data.Odbc
Imports System.Data.SqlClient
Public Class CDConsultas_Especificas
    Private ObjConexion As New CDConexion
    Dim cn As New OdbcConnection
    Dim cm As New OdbcCommand
    Dim da As New OdbcDataAdapter
    Dim dr As OdbcDataReader
    Dim Tabla As New Data.DataTable
    Function Contar_Lotes_Tinto() As OdbcDataReader
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandType = CommandType.Text
        cm.CommandText = ("SELECT count(id_lote)FROM lote WHERE tipo = 'Tinto' ")
        Try
            dr = cm.ExecuteReader
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return dr
        ObjConexion.Cerrar()
        dr.Dispose()
        cn.Dispose()
        cm.Dispose()
    End Function
    Function Contar_Lotes_Rosado() As OdbcDataReader
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandType = CommandType.Text
        cm.CommandText = ("SELECT count(id_lote)FROM lote WHERE tipo = 'Rosado' ")
        Try
            dr = cm.ExecuteReader
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return dr
        ObjConexion.Cerrar()
        dr.Dispose()
        cn.Dispose()
        cm.Dispose()
    End Function
    Function Contar_Lotes_Blanco() As OdbcDataReader
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandType = CommandType.Text
        cm.CommandText = ("SELECT count(id_lote)FROM lote WHERE tipo = 'Blanco' ")
        Try
            dr = cm.ExecuteReader
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return dr
        ObjConexion.Cerrar()
        dr.Dispose()
        cn.Dispose()
        cm.Dispose()
    End Function
    Function Cantidad_MateriaPrima() As OdbcDataReader
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandType = CommandType.Text
        cm.CommandText = ("SELECT count(id_lote)FROM tiene WHERE nombre_proceso = 'Vendimia' ")
        Try
            dr = cm.ExecuteReader
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return dr
        ObjConexion.Cerrar()
        dr.Dispose()
        cn.Dispose()
        cm.Dispose()
    End Function
    Function Cantidad_ProductoIntermedio() As OdbcDataReader
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandType = CommandType.Text
        cm.CommandText = ("SELECT count(id_lote)FROM tiene WHERE nombre_proceso = 'Prensado' ")
        Try
            dr = cm.ExecuteReader
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return dr
        ObjConexion.Cerrar()
        dr.Dispose()
        cn.Dispose()
        cm.Dispose()
    End Function

    Function Cantidad_ProductoFinal() As OdbcDataReader
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandType = CommandType.Text
        cm.CommandText = ("SELECT count(id_lote)FROM tiene WHERE nombre_proceso = 'Embotellado' ")
        Try
            dr = cm.ExecuteReader
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return dr
        ObjConexion.Cerrar()
        dr.Dispose()
        cn.Dispose()
        cm.Dispose()
    End Function
    Function Cantidad_Materia_porTipo() As OdbcDataReader
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandType = CommandType.Text
        cm.CommandText = ("SELECT count(id_lote)FROM tiene WHERE nombre_estado = 'Descarga' or nombre_estado = 'Despalillado' or nombre_proceso = 'Vendimia'")
        Try
            dr = cm.ExecuteReader
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return dr
        ObjConexion.Cerrar()
        dr.Dispose()
        cn.Dispose()
        cm.Dispose()
    End Function

End Class

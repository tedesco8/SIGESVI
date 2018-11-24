Imports System.Data.Odbc
Imports System.Data.SqlClient
Public Class CDLote
    Private ObjConexion As New CDConexion

    Dim cn As New OdbcConnection
    Dim cm As New OdbcCommand
    Dim da As New OdbcDataAdapter
    Dim dr As OdbcDataReader
    Dim Tabla As New Data.DataTable

    Public Function listadoLotes() As DataTable
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandText = "SELECT id_lote, origen, cepa, cantidad, fecha_produccion, capacidad, estado_sanitario, comentarios, tipo FROM lote WHERE estado = 't'"
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
    Public Function listadoidsucursal() As DataTable
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandText = "SELECT id_sucursal FROM sucursal"
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
    Public Function listarloteeliminado() As DataTable
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandText = "SELECT id_lote, origen, cepa, cantidad, fecha_produccion, capacidad, estado_sanitario, comentarios, tipo FROM lote WHERE estado = 'f'"
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
    Public Function listarBusqueda(ByVal Opcion As Integer) As DataTable
        cn = ObjConexion.Conectar
        cm.Connection = cn
        Select Case Opcion
            Case 1
                cm.CommandText = "SELECT id_lote, nombre_proceso, nombre_estado, fecha_inicio, fecha_final FROM tiene WHERE estado = 't' "
                dr = cm.ExecuteReader
                Try
                    Tabla.Load(dr)
                    ObjConexion.Cerrar()
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            Case 2
                cm.CommandText = "SELECT id_lote, nombre_proceso, nombre_estado, fecha_inicio, fecha_final FROM tiene WHERE estado = 't' "
                dr = cm.ExecuteReader
                Try
                    Tabla.Load(dr)
                    ObjConexion.Cerrar()
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            Case 3
                cm.CommandText = "SELECT id_lote, nombre_proceso, nombre_estado, fecha_inicio, fecha_final FROM tiene WHERE estado = 't' "
                dr = cm.ExecuteReader
                Try
                    Tabla.Load(dr)
                    ObjConexion.Cerrar()
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            Case 4
                cm.CommandText = "SELECT id_lote, nombre_proceso, nombre_estado, fecha_inicio, fecha_final FROM tiene WHERE estado = 't' "
                dr = cm.ExecuteReader
                Try
                    Tabla.Load(dr)
                    ObjConexion.Cerrar()
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            Case 5
                cm.CommandText = "SELECT id_lote, nombre_proceso, nombre_estado, fecha_inicio, fecha_final FROM tiene WHERE estado = 't' "
                dr = cm.ExecuteReader
                Try
                    Tabla.Load(dr)
                    ObjConexion.Cerrar()
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            Case 6
                cm.CommandText = "SELECT id_lote, nombre_proceso, nombre_estado, fecha_inicio, fecha_final FROM tiene WHERE estado = 't' "
                dr = cm.ExecuteReader
                Try
                    Tabla.Load(dr)
                    ObjConexion.Cerrar()
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            Case Else
        End Select
        Return Tabla
        cn.Dispose()
        cm.Dispose()
        dr.Dispose()
    End Function
    Public Function registrarLote(ByVal id_lote As Integer, ByVal origen As String, ByVal cepa As String, ByVal cantidad As Integer, ByVal fecha_produccion As Integer, ByVal capacidad As Integer, ByVal estado_sanitario As String, ByVal comentarios As String, ByVal tipo As String)
        cm.Connection = ObjConexion.Conectar
        cm.CommandText = "insert into lote (id_lote, cepa, origen, cantidad, fecha_produccion, capacidad, estado_sanitario, comentarios, tipo, estado) values ( '" & id_lote & "','" & cepa & "', '" & origen & "', '" & cantidad & "','" & fecha_produccion & "','" & capacidad & "', '" & estado_sanitario & "', '" & comentarios & "', '" & tipo & "', 't'  )"
        Try
            cm.ExecuteNonQuery()
            cm.Parameters.Clear()
            'llamar metodo insertar lote en tabla trazabilidad
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return cm
        cm.Dispose()
    End Function
    Public Function modificarLote(ByVal id_lote, ByVal origen, ByVal cepa, ByVal cantidad, ByVal fecha_produccion, ByVal capacidad, ByVal estado_sanitario, ByVal comentarios, ByVal tipo)
        Dim sql As String = ""
        cm.CommandType = CommandType.Text
        cm.Connection = ObjConexion.Conectar
        sql = "UPDATE lote SET "
        sql += "origen = '" & origen & "', "
        sql += "cepa = '" & cepa & "', "
        sql += "cantidad = " & cantidad & ", "
        sql += "fecha_produccion = " & fecha_produccion & ", "
        sql += "capacidad = " & capacidad & ", "
        sql += "estado_sanitario = '" & estado_sanitario & "', "
        sql += "comentarios = '" & comentarios & "', "
        sql += "tipo = '" & tipo & "' "
        sql += "WHERE id_lote = " & id_lote & ""
        cm.CommandText = sql
        Try
            cm.ExecuteNonQuery()
            cm.Parameters.Clear()
            'llamar funcion update tabla trazabilidad
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return cm
        cm.Dispose()
    End Function
    Public Function borrarLote(ByVal id_lote)
        Dim sql As String = ""
        cm.CommandType = CommandType.Text
        cm.Connection = ObjConexion.Conectar
        sql = "UPDATE lote SET "
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
    Public Function restaurarLote(ByVal id_lote)
        Dim sql As String = ""
        cm.CommandType = CommandType.Text
        cm.Connection = ObjConexion.Conectar
        sql = "UPDATE lote SET "
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
    'Sub filtroLotes(ByVal Filtro As String, ByVal dvg As DataGridView)
    '    cm.Connection = ObjConexion.Conectar
    '    Try
    '        cm.CommandText = ("SELECT * FROM lote WHERE id_lote LIKE" & "%" & Filtro + "%" & "", cm)
    '        da.Fill(Tabla)
    '        dgv.DataSource = cm
    '    Catch ex As Exception
    '        MsgBox(ex.ToString)
    '    End Try
    '    cm.Dispose()
End Class

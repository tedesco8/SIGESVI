Imports System.Data.Odbc
Imports System.Data.SqlClient
Public Class CDSucursales
    Dim ObjConexion As New CDConexion
    Dim cn As New OdbcConnection
    Dim da As New OdbcDataAdapter
    Dim cm As New OdbcCommand
    Dim dr As OdbcDataReader
    Dim Tabla As New Data.DataTable

    'Creamos una función para listar las Sucursales
    Public Function listadoSucursal() As DataTable
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandText = "SELECT id_sucursal, calle, numero, capacidad, ciudad FROM sucursal WHERE estado = 't'"
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
    ''Creamos un procedimiento para registrar una Sucursal
    Public Function registrarSucursal(ByVal id_sucursal As Integer, ByVal ciudad As String, ByVal capacidad As String, ByVal calle As String, ByVal numero As Integer)
        cm.Connection = ObjConexion.Conectar
        cm.CommandText = "insert into sucursal (id_sucursal, ciudad, capacidad, calle, numero, estado) values ( '" & id_sucursal & "','" & ciudad & "','" & capacidad & "','" & calle & "', '" & numero & "', 't')"
        Try
            cm.ExecuteNonQuery()
            cm.Parameters.Clear()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return cm
    End Function
    'Creamos el procedimiento para modificar los registros de las sucursales
    Public Function modificarSucursal(ByVal id_sucursal, ByVal ciudad, ByVal capacidad, ByVal calle, ByVal numero)
        Dim sql As String = ""
        cm.CommandType = CommandType.Text
        cm.Connection = ObjConexion.Conectar
        sql = "UPDATE sucursal SET "
        sql += "ciudad = '" & ciudad & "', "
        sql += "capacidad = " & capacidad & ","
        sql += "calle = '" & calle & "', "
        sql += "numero = " & numero & " "
        sql += "WHERE id_sucursal = " & id_sucursal & ""
        cm.CommandText = sql
        Try
            cm.ExecuteNonQuery()
            cm.Parameters.Clear()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return cm
    End Function
    Public Function borrarSucursal(ByVal id_sucursal)
        Dim sql As String = ""
        cm.CommandType = CommandType.Text
        cm.Connection = ObjConexion.Conectar
        sql = "UPDATE sucursal SET "
        sql += "estado = 'f' "
        sql += "WHERE id_sucursal = " & id_sucursal & ""
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

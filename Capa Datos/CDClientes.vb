Imports System.Data.Odbc
Imports System.Data.SqlClient
Public Class CDClientes
    Private ObjConexion As New CDConexion
    Dim cn As New OdbcConnection
    Dim cm As New OdbcCommand
    Dim da As New OdbcDataAdapter
    Dim dr As OdbcDataReader
    Dim Tabla As New Data.DataTable
    'Creamos una función que nos permite listar los clientes
    Public Function listadoClientes() As DataTable
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandText = "SELECT id_cliente, nombre,estado FROM clientes WHERE estado = 't'"
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
    Public Function listadoPersonas() As DataTable
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandText = "SELECT ci, apellido, calle, numero, estado FROM personas WHERE estado = 't'"
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
    Public Function listadoEmpresas() As DataTable
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandText = "SELECT rut, calle, numero, estado FROM empresas WHERE estado = 't'"
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
    Public Function registrarClientes(ByVal id_cliente As Integer, ByVal nombre As String)
        cm.Connection = ObjConexion.Conectar
        cm.CommandText = "insert into clientes (id_cliente, nombre, estado) values ( '" & id_cliente & "','" & nombre & "', 't'  )"
        Try
            cm.ExecuteNonQuery()
            cm.Parameters.Clear()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return cm
        cm.Dispose()
    End Function
    Public Function registrarPersonas(ByVal ci As Integer, ByVal apellido As String, ByVal calle As String, ByVal numero As Integer)
        cm.Connection = ObjConexion.Conectar
        cm.CommandText = "insert into personas (ci, apellido, calle, numero, estado) values ( '" & ci & "','" & apellido & "','" & calle & "','" & numero & "', 't'  )"
        Try
            cm.ExecuteNonQuery()
            cm.Parameters.Clear()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return cm
        cm.Dispose()
    End Function
    Public Function registrarEmpresas(ByVal RUT As Integer, ByVal calle As String, ByVal numero As Integer)
        cm.Connection = ObjConexion.Conectar
        cm.CommandText = "insert into empresas (rut, calle, numero, estado) values ( '" & RUT & "','" & calle & "','" & numero & "', 't'  )"
        Try
            cm.ExecuteNonQuery()
            cm.Parameters.Clear()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return cm
        cm.Dispose()
    End Function
   
    ''Creamos una función que nos permite autogenerar el ID Cliente a traves del Store Procedure
    'Function generaCodigo() As Integer
    '    Try
    '        cn = ObjConexion.Conectar
    '        da = New OdbcDataAdapter("UltCodigoCliente", cn)
    '        Dim r As Integer
    '        cn.Open()
    '        r = da.SelectCommand.ExecuteScalar
    '        Return r + 1
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        cn.Dispose()
    '        da.Dispose()
    '    End Try
    'End Function
    ''Creamos un procedimiento para registrar el cliente
End Class

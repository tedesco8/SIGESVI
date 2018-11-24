Imports System.Data.Odbc
Imports System.Data.SqlClient

Public Class CDPersonal
    Private ObjConexion As New CDConexion
    Dim cn As New OdbcConnection
    Dim cm As New OdbcCommand
    Dim da As New OdbcDataAdapter
    Dim dr As OdbcDataReader
    Dim Tabla As New Data.DataTable

    'Login del personal
    Function IniciarSesion(ByVal Usuario As String, ByVal Contrasena As String) As OdbcDataReader
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandType = CommandType.Text
        cm.CommandText = ("SELECT email,contrasena, nombre, apellido, cargo FROM PERSONAL WHERE email='" & Usuario & "' AND contrasena='" & Contrasena & "' ")
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
    'Creamos una función que nos permite listar el personal de la tabla Personal.
    Public Function listadoPersonal() As DataTable
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandText = "SELECT ci, nombre, apellido, cargo, calle, numero, email, contrasena FROM PERSONAL WHERE estado = 't'"
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
    Public Function listarpersonaleliminado() As DataTable
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandText = "SELECT ci, nombre, apellido, cargo, calle, numero, email, contrasena FROM PERSONAL WHERE estado = 'f'"
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
    Public Function listadoTelefono() As DataTable
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandText = "SELECT ci, telefono_1, telefono_2 FROM personal_tel WHERE estado = 't'"
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
    
    ''Creamos una función para registrar el personal en la bd
    Public Function registrarPersonal(ByVal ci As Integer, ByVal nombre As String, ByVal apellido As String, ByVal cargo As String, ByVal calle As String, ByVal numero As Integer, ByVal email As String, ByVal contrasena As String)
        cm.Connection = ObjConexion.Conectar
        cm.CommandText = "insert into personal (ci, nombre, apellido, cargo, calle, numero, email, contrasena, estado) values ( '" & ci & "','" & nombre & "','" & apellido & "','" & cargo & "','" & calle & "', '" & numero & "', '" & email & "', '" & contrasena & "', 't'  )"
        Try
            cm.ExecuteNonQuery()
            cm.Parameters.Clear()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return cm
        cm.Dispose()
    End Function
    Public Function registrarTelefono(ByVal ci As Integer, ByVal nombre As String, ByVal apellido As String, ByVal telefono_1 As String, ByVal telefono_2 As String)
        cm.Connection = ObjConexion.Conectar
        cm.CommandText = "insert into personal_tel (ci, nombre, apellido, estado, telefono_1, telefono_2) values ( '" & ci & "', '" & nombre & "', '" & apellido & "', 't', '" & telefono_1 & "', '" & telefono_2 & "')"
        Try
            cm.ExecuteNonQuery()
            cm.Parameters.Clear()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return cm
        cm.Dispose()
    End Function
    'Creamos la función para modificar los registros del personal
    Public Function modificarPersonal(ByVal ci, ByVal nombre, ByVal apellido, ByVal cargo, ByVal calle, ByVal numero, ByVal email, ByVal contrasena)
        Dim sql As String = ""
        cm.CommandType = CommandType.Text
        cm.Connection = ObjConexion.Conectar
        sql = "UPDATE personal SET "
        sql += "nombre = '" & nombre & "', "
        sql += "apellido = '" & apellido & "', "
        sql += "cargo = '" & cargo & "', "
        sql += "calle = '" & calle & "', "
        sql += "numero = " & numero & ", "
        sql += "email = '" & email & "', "
        sql += "contrasena = '" & contrasena & "' "
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
    Public Function modificarTelefono(ByVal ci, ByVal telefono_1, ByVal telefono_2)
        Dim sql As String = ""
        cm.CommandType = CommandType.Text
        cm.Connection = ObjConexion.Conectar
        sql = "UPDATE personal_tel SET "
        sql += "telefono_1 = '" & telefono_1 & "', "
        sql += "telefono_2 = '" & telefono_2 & "' "
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
    Public Function borrarPersonal(ByVal ci)
        Dim sql As String = ""
        cm.CommandType = CommandType.Text
        cm.Connection = ObjConexion.Conectar
        sql = "UPDATE personal SET "
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
    Public Function borrarTelefono(ByVal ci)
        Dim sql As String = ""
        cm.CommandType = CommandType.Text
        cm.Connection = ObjConexion.Conectar
        sql = "UPDATE personal_tel SET "
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
    Public Function restaurarPersonal(ByVal ci)
        Dim sql As String = ""
        cm.CommandType = CommandType.Text
        cm.Connection = ObjConexion.Conectar
        sql = "UPDATE personal SET "
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
    Public Function restaurarTelefono(ByVal ci)
        Dim sql As String = ""
        cm.CommandType = CommandType.Text
        cm.Connection = ObjConexion.Conectar
        sql = "UPDATE personal_tel SET "
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
    Public Function eliminarTelefono(ByVal ci) 'Telefono depende de personal, este se tiene que eliminar primero
        cm.Connection = ObjConexion.Conectar
        cm.CommandText = "delete from personal_tel where ci = " & ci & ""
        Try
            cm.ExecuteNonQuery()
            cm.Parameters.Clear()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return cm
        cm.Dispose()
    End Function
    Public Function eliminarPersonal(ByVal ci) ' este despues de telefono, pero el de trabaja en depende de sucursal me parece, si
        cm.Connection = ObjConexion.Conectar
        cm.CommandText = "delete from personal where ci = " & ci & ""
        Try
            cm.ExecuteNonQuery()
            cm.Parameters.Clear()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return cm
        cm.Dispose()
    End Function

    ''Creamos una función que nos permite autogenerar el ID

    'Function generaCodigo() As Integer
    'Try
    'cn = ObjConexion.Conectar
    'da = New OdbcDataAdapter("UltCodigoPersona", cn)
    'Dim r As Integer
    'cn.Open()
    'r = da.SelectCommand.ExecuteScalar
    'Return r + 1
    'Catch ex As Exception
    'Throw ex
    'Finally
    'cn.Dispose()
    'da.Dispose()
    'End Try
    'End Function

    ''Creamos procedimiento que permite mostrar el cargo
    'Function listarCargo() As DataSet
    '    cn = ObjConexion.Conectar
    '    da = New OdbcDataAdapter("ListarCargo", cn)
    '    Dim ds As New DataSet
    '    da.Fill(ds, "Cargo")
    '    Return ds
    '    ds.Dispose()
    '    da.Dispose()
    '    cn.Dispose()
    'End Function
End Class

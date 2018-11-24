Imports System.Data.Odbc
Imports System.Data.SqlClient
Public Class CDTrazabilidad
    Private ObjConexion As New CDConexion
    Dim cn As New OdbcConnection
    Dim cm As New OdbcCommand
    Dim da As New OdbcDataAdapter
    Dim dr As OdbcDataReader
    Dim Tabla As New Data.DataTable

    Public Function listarTrazabilidad() As DataTable
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandText = "SELECT * FROM trazabilidad"
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
    Public Function registrarTrazabilidad(ByVal id_lote As Integer, ByVal nombre_proceso As String, ByVal nombre_estado As String, ByVal origen As String, ByVal tipo As String, ByVal cepa As String, ByVal id_sucursal As Integer, ByVal comentarios As String, ByVal estado_sanitario As String, ByVal fecha_inicio As Integer, ByVal fecha_final As Integer)
        Try
            cm.CommandText = "insert into trazabilidad (id_lote, nombre_proceso, nombre_estado, origen, tipo, cepa, id_sucursal, comentarios, estado_sanitario, fecha_inicio, fecha_final) values ( '" & id_lote & "','" & nombre_proceso & "', '" & nombre_estado & "', '" & origen & "', '" & tipo & "', '" & cepa & "','" & id_sucursal & "', '" & comentarios & "', '" & estado_sanitario & "', '" & fecha_inicio & "', '" & fecha_final & "')"
            cm.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return cm
        cm.Dispose()
    End Function
End Class

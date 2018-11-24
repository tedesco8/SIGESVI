﻿Imports System.Data.Odbc
Imports System.Data.SqlClient
Public Class CDEstado
    Private ObjConexion As New CDConexion
    Dim cn As New OdbcConnection
    Dim cm As New OdbcCommand
    Dim da As New OdbcDataAdapter
    Dim dr As OdbcDataReader
    Dim Tabla As New Data.DataTable
    Public Function listadoEstados() As DataTable
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandText = "SELECT nombre_estado FROM estado"
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
End Class

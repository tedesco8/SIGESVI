Imports System.Data.Odbc
Imports System.Data.SqlClient
Public Class CDEmbotellado
    Private ObjConexion As New CDConexion
    Dim cn As New OdbcConnection
    Dim cm As New OdbcCommand
    Dim da As New OdbcDataAdapter
    Dim dr As OdbcDataReader
    Dim Tabla As New Data.DataTable

    Function cantidad_embotellado(ByVal id_lote As Integer)
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandType = CommandType.Text
        cm.CommandText = ("SELECT cantidad FROM lote WHERE id_lote= '" & id_lote & "' ")
        Try
            dr = cm.ExecuteReader()
            Tabla.Load(dr)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Dim cant As String = Tabla.Rows(0)(0).ToString()
        ObjConexion.Cerrar()
        dr.Dispose()
        dr.Close()
        cn.Dispose()
        cm.Dispose()
        Return cant
    End Function
    Function tipo_botella(ByVal id_lote As Integer)
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandType = CommandType.Text
        cm.CommandText = ("SELECT tipo_botella FROM lote WHERE id_lote= '" & id_lote & "' ")
        Try
            dr = cm.ExecuteReader()
            Tabla.Load(dr)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Dim cant As String = Tabla.Rows(0)(0).ToString()
        ObjConexion.Cerrar()
        dr.Dispose()
        dr.Close()
        cn.Dispose()
        cm.Dispose()
        Return cant
    End Function
End Class

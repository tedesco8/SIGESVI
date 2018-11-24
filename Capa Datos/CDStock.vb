Imports System.Data.Odbc
Imports System.Data.SqlClient
Public Class CDStock
    Private ObjConexion As New CDConexion
    Private Obj_Embotellado As New CDEmbotellado
    Dim cn As New OdbcConnection
    Dim cm As New OdbcCommand
    Dim da As New OdbcDataAdapter
    Dim dr As OdbcDataReader
    Dim Tabla As New Data.DataTable
    Public Function listadoStock() As DataTable
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandText = "SELECT id_stock, tipo_vino, tipo_botella, cantidad, precio FROM stock WHERE estado = 't'"
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
    Function tipo_vino(ByVal id_lote As Integer)
        cn = ObjConexion.Conectar
        cm.Connection = cn
        cm.CommandType = CommandType.Text
        cm.CommandText = ("SELECT tipo FROM lote WHERE id_lote= '" & id_lote & "' ")
        Try
            dr = cm.ExecuteReader
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Tabla.Load(dr)
        ObjConexion.Cerrar()
        dr.Dispose()
        dr.Close()
        cn.Dispose()
        cm.Dispose()
        Return Tabla.Rows(0)(0).ToString()
        
    End Function
    Sub llenar_stock(ByVal id_lote As Integer)

        Dim tipo As String
        tipo = tipo_vino(id_lote).ToString()
        cm.Connection = ObjConexion.Conectar
        cm.CommandText = "insert into stock (cantidad, tipo_vino, estado) values ( '" & Obj_Embotellado.cantidad_embotellado(id_lote).ToString() & "','" & tipo & "', 't'  )"
        Try
            cm.ExecuteNonQuery()
            cm.Parameters.Clear()
            'llamar metodo insertar lote en tabla trazabilidad
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        cm.Dispose()
    End Sub
End Class

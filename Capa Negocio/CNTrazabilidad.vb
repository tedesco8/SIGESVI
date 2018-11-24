Imports Capa_Datos
Imports System.Data.Odbc
Imports System.Data.SqlClient
Public Class CNTrazabilidad
    Dim ObjCapaDatos As New CDTrazabilidad
    ''Creamos las variables generales con las cuales trabajaremos en toda la clase.
    Dim dr As OdbcDataReader
    Dim Tabla As New Data.DataTable
    Function listado_trazabilidad() As DataTable
        Tabla = ObjCapaDatos.listarTrazabilidad
        Return Tabla
    End Function
    Sub registrar_Trazabilidad(ByVal id_lote As Integer, ByVal nombre_proceso As String, ByVal nombre_estado As String, ByVal origen As String, ByVal tipo As String, ByVal cepa As String, ByVal id_sucursal As Integer, ByVal comentarios As String, ByVal estado_sanitario As String, ByVal fecha_inicio As Integer, ByVal fecha_final As Integer)
        ObjCapaDatos.registrarTrazabilidad(id_lote, nombre_proceso, nombre_estado, origen, tipo, cepa, id_sucursal, comentarios, estado_sanitario, fecha_inicio, fecha_final)
    End Sub
End Class

Imports Capa_Datos
Imports System.Data.Odbc
Imports System.Data.SqlClient
Public Class CNStock
    Dim ObjCapaDatos As New CDStock
    ''Creamos las variables generales con las cuales trabajaremos en toda la clase.
    Dim dr As OdbcDataReader
    Dim Tabla As New Data.DataTable
    ''Creamos los atributos y propiedades de la clase.
    Function listado_Stock() As DataTable
        Tabla = ObjCapaDatos.listadoStock
        Return Tabla
    End Function
    Sub llenar_Stock(ByVal id_lote As Integer)
        ObjCapaDatos.llenar_stock(id_lote)
    End Sub
End Class

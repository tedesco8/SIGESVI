Imports Capa_Datos
Imports System.Data.Odbc
Imports System.Data.SqlClient
Public Class CNConsultas_especificas
    ''Creamos un objeto de la clase personal de la capa de datos.
    Dim ObjCapaDatos As New CDConsultas_Especificas
    ''Creamos las variables generales con las cuales trabajaremos en toda la clase.
    Dim dr As OdbcDataReader
    Dim Tabla As New Data.DataTable
    Function Contar_Lotes_Tinto() As OdbcDataReader
        dr = ObjCapaDatos.Contar_Lotes_Tinto()
        Return dr
    End Function
    Function Contar_Lotes_Rosado() As OdbcDataReader
        dr = ObjCapaDatos.Contar_Lotes_Rosado()
        Return dr
    End Function
    Function Contar_Lotes_Blanco() As OdbcDataReader
        dr = ObjCapaDatos.Contar_Lotes_Blanco()
        Return dr
    End Function
    Function Cantidad_Materia_Prima() As OdbcDataReader
        dr = ObjCapaDatos.Cantidad_MateriaPrima
        Return dr
    End Function
    Function Cantidad_Producto_Intermedio() As OdbcDataReader
        dr = ObjCapaDatos.Cantidad_ProductoIntermedio
        Return dr
    End Function
    Function Cantidad_Producto_Final() As OdbcDataReader
        dr = ObjCapaDatos.Cantidad_ProductoFinal
        Return dr
    End Function
End Class

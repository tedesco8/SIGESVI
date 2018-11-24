Imports Capa_Datos
Imports System.Data.Odbc
Imports System.Data.SqlClient
Public Class CNEstado
    Dim ObjCapaDatos As New CDEstado
    ''Creamos las variables generales con las cuales trabajaremos en toda la clase.
    Dim dr As OdbcDataReader
    Dim Tabla As New Data.DataTable
    ''Creamos los atributos y propiedades de la clase.
    Private _Nombre_Estado As String
    Public Property Nombre_Estado() As String
        Get
            Return _Nombre_Estado
        End Get
        Set(ByVal value As String)
            _Nombre_Estado = value
        End Set
    End Property
    Public Sub New()

    End Sub
    'Referencias
    Public Sub New(ByVal Nombre_Estado As String)
        Nombre_Estado = _Nombre_Estado
    End Sub
    Function listado_Estados() As DataTable
        Tabla = ObjCapaDatos.listadoEstados
        Return Tabla
    End Function
End Class

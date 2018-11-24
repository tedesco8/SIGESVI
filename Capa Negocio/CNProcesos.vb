Imports Capa_Datos
Imports System.Data.Odbc
Imports System.Data.SqlClient
Public Class CNProcesos
    Dim ObjCapaDatos As New CDProcesos
    Dim dr As OdbcDataReader
    Dim Tabla As New Data.DataTable
    Private _Nombre_Proceso As String
    Private _Fecha_Inicio, _Fecha_Final As Integer
    Public Property Nombre_Proceso() As String
        Get
            Return _Nombre_Proceso
        End Get
        Set(ByVal value As String)
            _Nombre_Proceso = value
        End Set
    End Property
    Public Sub New()

    End Sub
    'Referencias
    Public Sub New(ByVal Nombre_Proceso As String)
        Nombre_Proceso = _Nombre_Proceso
    End Sub
    Function listado_Procesos() As DataTable
        Tabla = ObjCapaDatos.listadoProcesos
        Return Tabla
    End Function
End Class

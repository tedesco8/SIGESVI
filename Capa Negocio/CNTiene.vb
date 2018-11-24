Imports Capa_Datos
Imports System.Data.Odbc
Imports System.Data.SqlClient
Public Class CNTiene
    Dim ObjCapaDatos As New CDTiene
    ''Creamos las variables generales con las cuales trabajaremos en toda la clase.
    Dim dr As OdbcDataReader
    Dim Tabla As New Data.DataTable

    Private _ID_Lote, _Fecha_Inicio, _Fecha_Final As Integer
    Private _Nombre_Estado, _Nombre_Proceso As String

    Public Property ID_Lote() As Integer
        Get
            Return _ID_Lote
        End Get
        Set(ByVal value As Integer)
            _ID_Lote = ID_Lote
        End Set
    End Property
    Public Property Nombre_Estado() As String
        Get
            Return _Nombre_Estado
        End Get
        Set(ByVal value As String)
            _Nombre_Estado = Nombre_Estado
        End Set
    End Property
    Public Property Nombre_Proceso() As String
        Get
            Return _Nombre_Proceso
        End Get
        Set(ByVal value As String)
            _Nombre_Proceso = Nombre_Proceso
        End Set
    End Property
    Public Property Fecha_Inicio() As Integer
        Get
            Return _Fecha_Inicio
        End Get
        Set(ByVal value As Integer)
            _Fecha_Inicio = Fecha_Inicio
        End Set
    End Property
    Public Property Fecha_Final() As Integer
        Get
            Return _Fecha_Final
        End Get
        Set(ByVal value As Integer)
            _Fecha_Final = Fecha_Final
        End Set
    End Property
    Public Sub New()

    End Sub
    Public Sub New(ByVal Id_Lote As Integer, ByVal Nombre_Estado As String, ByVal Nombre_Proceso As String, ByVal Fecha_Inicio As Integer, ByVal Fecha_Final As Integer)
        Id_Lote = _ID_Lote
        Nombre_Estado = _Nombre_Estado
        Nombre_Proceso = _Nombre_Proceso
        Fecha_Inicio = _Fecha_Inicio
        Fecha_Final = _Fecha_Final
    End Sub
    Function listado_Tiene() As DataTable
        Tabla = ObjCapaDatos.listarTiene
        Return Tabla
    End Function
    
    Sub registrar_Tiene(ByVal id_lote As Integer, ByVal nombre_proceso As String, ByVal nombre_estado As String, ByVal fecha_inicio As Integer, ByVal fecha_final As Integer)
        ObjCapaDatos.registrarTiene(id_lote, nombre_proceso, nombre_estado, fecha_inicio, fecha_final)
    End Sub
    Sub modificar_Tiene(ByVal id_lote As Integer, ByVal nombre_proceso As String, ByVal nombre_estado As String, ByVal fecha_inicio As Integer, ByVal fecha_final As Integer)
        ObjCapaDatos.modificarTiene(id_lote, nombre_proceso, nombre_estado, fecha_inicio, fecha_final)
    End Sub
    Sub borrar_Tiene(ByVal id_lote)
        ObjCapaDatos.borrarTiene(id_lote)
    End Sub
    Sub restaurar_Tiene(ByVal id_lote)
        ObjCapaDatos.restaurarTiene(id_lote)
    End Sub
End Class

Imports Capa_Datos
Imports System.Data.Odbc
Imports System.Data.SqlClient
Public Class CNLote
    Dim ObjCapaDatos As New CDLote
    ''Creamos las variables generales con las cuales trabajaremos en toda la clase.
    Dim dr As OdbcDataReader
    Dim Tabla As New Data.DataTable

    Private _Origen, _Cepa, _Comentarios, _Tipo, _estado_sanitario As String
    Private _Fecha_produccion, _Capacidad, _Cantidad, _ID_Lote As Integer
    Public Property ID_Lote() As Integer
        Get
            Return _ID_Lote
        End Get
        Set(ByVal value As Integer)
            _ID_Lote = value
        End Set
    End Property
    Public Property Origen() As String
        Get
            Return _Origen
        End Get
        Set(ByVal value As String)
            _Origen = value
        End Set
    End Property
    Public Property Cepa() As String
        Get
            Return _Cepa
        End Get
        Set(ByVal value As String)
            _Cepa = value
        End Set
    End Property
    Public Property Tipo() As String
        Get
            Return _Tipo
        End Get
        Set(ByVal value As String)
            _Tipo = value
        End Set
    End Property
    Public Property Comentarios() As String
        Get
            Return _Comentarios
        End Get
        Set(ByVal value As String)
            _Comentarios = value
        End Set
    End Property
    Public Property Fecha_Produccion() As Integer
        Get
            Return _Fecha_produccion
        End Get
        Set(ByVal value As Integer)
            _Fecha_produccion = value
        End Set
    End Property
    Public Property Capacidad() As Integer
        Get
            Return _Cantidad
        End Get
        Set(ByVal value As Integer)
            _Capacidad = value
        End Set
    End Property
    Public Property Estado_Sanitario() As String
        Get
            Return _estado_sanitario
        End Get
        Set(ByVal value As String)
            _estado_sanitario = value
        End Set
    End Property
    Public Property Cantidad() As Integer
        Get
            Return _Cantidad
        End Get
        Set(ByVal value As Integer)
            _Cantidad = value
        End Set
    End Property
    Public Sub New()

    End Sub
    'Referencias
    Public Sub New(ByVal Id_Lote As Integer, ByVal Origen As String, ByVal Cepa As String, ByVal Tipo As String, ByVal Comentarios As String, ByVal Fecha_Produccion As Integer, ByVal Capacidad As Integer, ByVal Cantidad As Integer)
        Id_Lote = _ID_Lote
        Origen = _Origen
        Cepa = _Cepa
        Tipo = _Tipo
        Comentarios = _Comentarios
        Fecha_Produccion = _Fecha_produccion
        Capacidad = _Capacidad
        Estado_Sanitario = _estado_sanitario
        Cantidad = _Cantidad
    End Sub

    Function listado_Lotes() As DataTable
        Tabla = ObjCapaDatos.listadoLotes
        Return Tabla
    End Function
    Function listado_idsucursal() As DataTable
        Tabla = ObjCapaDatos.listadoidsucursal
        Return Tabla
    End Function
    Function listado_LoteEliminado() As DataTable
        Tabla = ObjCapaDatos.listarloteeliminado
        Return Tabla
    End Function
    Sub registrar_Lote(ByVal id_lote As Integer, ByVal origen As String, ByVal cepa As String, ByVal cantidad As Integer, ByVal fecha_produccion As Integer, ByVal capacidad As Integer, ByVal estado_sanitario As String, ByVal comentarios As String, ByVal tipo As String)
        ObjCapaDatos.registrarLote(id_lote, origen, cepa, cantidad, fecha_produccion, capacidad, estado_sanitario, comentarios, tipo)
    End Sub
    Sub modificar_Lote(ByVal id_lote As Integer, ByVal origen As String, ByVal cepa As String, ByVal cantidad As Integer, ByVal fecha_produccion As Integer, ByVal capacidad As Integer, ByVal estado_sanitario As String, ByVal comentarios As String, ByVal tipo As String)
        ObjCapaDatos.modificarLote(id_lote, origen, cepa, cantidad, fecha_produccion, capacidad, estado_sanitario, comentarios, tipo)
    End Sub
    Sub borrar_Lote(ByVal id_lote)
        ObjCapaDatos.borrarLote(id_lote)
    End Sub
    Sub restaurar_Lote(ByVal id_lote)
        ObjCapaDatos.restaurarLote(id_lote)
    End Sub
End Class

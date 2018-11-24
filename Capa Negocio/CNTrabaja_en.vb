Imports Capa_Datos
Imports System.Data.Odbc
Imports System.Data.SqlClient
Public Class CNTrabaja_en
    ''Creamos un objeto de la clase personal de la capa de datos.
    Dim ObjCapaDatos As New CDTrabaja_en
    ''Creamos las variables generales con las cuales trabajaremos en toda la clase.
    Dim dr As OdbcDataReader
    Dim Tabla As New Data.DataTable
    ''Creamos los atributos y propiedades de la clase.
    Private _Nombre, _Apellido As String
    Private _CI, _ID_Sucursal As Integer

    Public Property CI() As Integer
        Get
            Return _CI
        End Get
        Set(ByVal value As Integer)
            _CI = value
        End Set
    End Property
    Public Property ID_Sucursal() As Integer
        Get
            Return _ID_Sucursal
        End Get
        Set(ByVal value As Integer)
            _ID_Sucursal = value
        End Set
    End Property
    Public Property Nombre() As String
        Get
            Return _Nombre
        End Get
        Set(ByVal value As String)
            _Nombre = value
        End Set
    End Property
    Public Property Apellido() As String
        Get
            Return _Apellido
        End Get
        Set(ByVal value As String)
            _Apellido = value
        End Set
    End Property
    Public Sub New()

    End Sub
    Public Sub New(ByVal CI As Integer, ByVal ID_Sucursal As Integer, ByVal Nombre As String, ByVal Apellido As String)
        CI = _CI
        ID_Sucursal = _ID_Sucursal
        Nombre = _Nombre
        Apellido = _Apellido
    End Sub
    Function listar_trabajanen() As DataTable
        Tabla = ObjCapaDatos.listartrabajaen
        Return Tabla
    End Function
    ''Creamos un procedimiento para insertar en la tabla relacional entre personal y sucursal
    Sub registrar_Trabajanen(ByVal ci As Integer, ByVal id_sucursal As Integer, ByVal nombre As String, ByVal apellido As String)
        ObjCapaDatos.registrarTrabajanen(ci, id_sucursal, nombre, apellido)
    End Sub
    Sub modificar_Trabajaen(ByVal ci, ByVal id_sucursal, ByVal nombre, ByVal apellido)
        ObjCapaDatos.modificarTrabajaen(ci, id_sucursal, nombre, apellido)
    End Sub
    Sub borrar_trabajaen(ByVal ci)
        ObjCapaDatos.borrarTrabajaen(ci)
    End Sub
    Sub restaurar_Trabajaen(ByVal ci)
        ObjCapaDatos.restaurarTrabajaen(ci)
    End Sub
    Sub eliminar_Personal(ByVal ci)
        ObjCapaDatos.eliminarTrabajaen(ci)
    End Sub
End Class

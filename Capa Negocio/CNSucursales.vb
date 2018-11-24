Imports Capa_Datos
Imports System.Data.Odbc
Imports System.Data.SqlClient
Public Class CNSucursales
    Dim dr As OdbcDataReader
    Dim Tabla As New Data.DataTable
    ''Creamos el objeto Capa Datos
    Dim ObjCapaDatos As New CDSucursales
    Private _Ciudad, _Capacidad, _Calle As String
    Private _idsucursal, _Numero As Integer
    Public Property ID_Sucursal() As Integer
        Get
            Return _idsucursal
        End Get
        Set(ByVal value As Integer)
            _idsucursal = value
        End Set
    End Property
    Public Property Ciudad() As String
        Get
            Return _Ciudad
        End Get
        Set(ByVal value As String)
            _Ciudad = value
        End Set
    End Property
    Public Property Capacidad() As Integer
        Get
            Return _Capacidad
        End Get
        Set(ByVal value As Integer)
            _Capacidad = value
        End Set
    End Property
    Public Property Calle() As String
        Get
            Return _Calle
        End Get
        Set(ByVal value As String)
            _Calle = value
        End Set
    End Property
    Public Property Numero() As Integer
        Get
            Return _Numero
        End Get
        Set(ByVal value As Integer)
            _Numero = value
        End Set
    End Property
    Public Sub New()

    End Sub
    'Referencias
    Public Sub New(ByVal id_sucursal As Integer, ByVal ciudad As String, ByVal capacidad As Integer, ByVal calle As String, ByVal numero As Integer)
        id_sucursal = _idsucursal
        ciudad = _Ciudad
        capacidad = _Capacidad
        calle = _Calle
        numero = _Numero
    End Sub
    'Implementamos la función que retorna el Data Set para listar la sucursal
    Function listado_Sucursal() As DataTable
        Tabla = ObjCapaDatos.listadoSucursal
        Return Tabla
    End Function
    ''Creamos un procedimiento para capturar la funcion que carga sucursales
    Sub registrar_Sucursal(ByVal id_sucursal As Integer, ByVal ciudad As String, ByVal capacidad As String, ByVal calle As String, ByVal numero As Integer)
        ObjCapaDatos.registrarSucursal(id_sucursal, ciudad, capacidad, calle, numero)
    End Sub
    ''Creamos un procedimiento para capturar la funcion que modifica sucursales
    Sub modificar_Sucursal(ByVal id_sucursal, ByVal ciudad, ByVal capacidad, ByVal calle, ByVal numero)
        ObjCapaDatos.modificarSucursal(id_sucursal, ciudad, capacidad, calle, numero)
    End Sub
    Sub borrar_sucursal(ByVal id_sucursal)
        ObjCapaDatos.borrarSucursal(id_sucursal)
    End Sub
    ' ''Llamamos a la funcion que autogenera el ID
    'Function generaCodigo() As Integer
    '    Return ObjCapaDatos.generaCodigo
    'End Function
    ''Llamamos a la funcion que lista el cargo
    'Function listarCiudad() As DataSet
    '    Return ObjCapaDatos.listarCiudad
    'End Function
    ''Llamamos al procedimiento para cargar el personal
    'Sub registrarSucursal(ByVal obj As CESucursales)
    '    ObjCapaDatos.registrarSucursales(obj)
    'End Sub
End Class

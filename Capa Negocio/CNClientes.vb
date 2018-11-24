Imports Capa_Datos
Imports System.Data.Odbc
Imports System.Data.SqlClient
Public Class CNClientes
    ''Creamos un objeto de la clase Clientes de la capa de datos.
    Dim ObjCapaDatos As New CDClientes
    ''Creamos las variables generales con las cuales trabajaremos en toda la clase.
    Dim dr As OdbcDataReader
    Dim Tabla As New Data.DataTable
    ''Implementamos la función que retorna el Data Set para listar los clientes
    Private _Nombre, _Apellido, _Calle As String
    Private _ID_Cliente, _CI, _RUT, _Numero, _Telefono As Integer
    Public Property ID_Cliente() As Integer
        Get
            Return _ID_Cliente
        End Get
        Set(ByVal value As Integer)
            _ID_Cliente = value
        End Set
    End Property
    Public Property CI() As Integer
        Get
            Return _CI
        End Get
        Set(ByVal value As Integer)
            _CI = value
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
    Public Property Telefono() As Integer
        Get
            Return _Telefono
        End Get
        Set(ByVal value As Integer)
            _Telefono = value
        End Set
    End Property
    Public Sub New()

    End Sub
    Public Sub New(ByVal ID_Cliente As Integer, ByVal CI As Integer, ByVal RUT As Integer, ByVal Nombre As String, ByVal Apellido As String, ByVal Calle As String, ByVal Numero As Integer, ByVal Telefono As Integer)
        ID_Cliente = _ID_Cliente
        RUT = _RUT
        CI = _CI
        Nombre = _Nombre
        Apellido = _Apellido
        Calle = _Calle
        Numero = _Numero
        Telefono = _Telefono
    End Sub
    Function listado_Clientes() As DataTable
        Tabla = ObjCapaDatos.listadoClientes
        Return Tabla
    End Function
    Function listado_Personas() As DataTable
        Tabla = ObjCapaDatos.listadoPersonas
        Return Tabla
    End Function
    Function listado_Empresas() As DataTable
        Tabla = ObjCapaDatos.listadoEmpresas
        Return Tabla
    End Function
    Sub registrar_Clientes(ByVal id_cliente As Integer, ByVal nombre As String)
        ObjCapaDatos.registrarClientes(id_cliente, nombre)
    End Sub
    Sub registrar_Personas(ByVal ci As Integer, ByVal apellido As String, ByVal calle As String, ByVal numero As Integer)
        ObjCapaDatos.registrarPersonas(ci, apellido, Calle, Numero)
    End Sub
    Sub registrar_Empresas(ByVal rut As Integer, ByVal calle As String, ByVal numero As Integer)
        ObjCapaDatos.registrarEmpresas(rut, calle, numero)
    End Sub
    ''Llamamos a la funcion que autogenera el ID
    'Function generaCodigo() As Integer
    '    Return ObjCapaDatos.generaCodigo
    'End Function
End Class

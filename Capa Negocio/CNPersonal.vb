Imports Capa_Datos
Imports System.Data.Odbc
Imports System.Data.SqlClient
Public Class CNPersonal
    ''Creamos un objeto de la clase personal de la capa de datos.
    Dim ObjCapaDatos As New CDPersonal
    ''Creamos las variables generales con las cuales trabajaremos en toda la clase.
    Dim dr As OdbcDataReader
    Dim Tabla As New Data.DataTable
    ''Creamos los atributos y propiedades de la clase.
    Private _Nombre, _Apellido, _Calle, _Cargo, _Usuario, _Contrasena As String
    Private _Numero, _Telefono, _CI As Integer
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
    Public Property CI() As Integer
        Get
            Return _CI
        End Get
        Set(ByVal value As Integer)
            _CI = value
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
    Public Property Cargo() As String
        Get
            Return _Cargo
        End Get
        Set(ByVal value As String)
            _Cargo = value
        End Set
    End Property
    Public Property Usuario() As String
        Get
            Return _Usuario
        End Get
        Set(ByVal value As String)
            _Usuario = value
        End Set
    End Property
    Public Property Contraseña() As String
        Get
            Return _Contrasena
        End Get
        Set(ByVal value As String)
            _Contrasena = value
        End Set
    End Property
    Public Sub New()

    End Sub
    'Referencias
    Public Sub New(ByVal Nombre As String, ByVal Apellido As String, ByVal Calle As String, ByVal Numero As Integer, ByVal CI As Integer, ByVal Telefono As Integer, ByVal Cargo As String, ByVal Usuario As String, ByVal Contraseña As String)
        Nombre = _Nombre
        Apellido = _Apellido
        Calle = _Calle
        Numero = _Numero
        CI = _CI
        Telefono = _Telefono
        Cargo = _Cargo
        Usuario = _Usuario
        Contraseña = _Contrasena
    End Sub

    'Implementamos la función del Login
    Function Iniciar_Sesion() As OdbcDataReader
        dr = ObjCapaDatos.IniciarSesion(Usuario, Contraseña)
        Return dr
    End Function
    ''Implementamos la función que retorna el Data Set para listar el personal
    Function listado_Personal() As DataTable
        Tabla = ObjCapaDatos.listadoPersonal
        Return Tabla
    End Function
    Function listado_PersonalEliminado() As DataTable
        Tabla = ObjCapaDatos.listarpersonaleliminado
        Return Tabla
    End Function
    Function listado_Telefono() As DataTable
        Tabla = ObjCapaDatos.listadoTelefono
        Return Tabla
    End Function
    Function listado_idsucursal() As DataTable
        Tabla = ObjCapaDatos.listadoidsucursal
        Return Tabla
    End Function
    ''Creamos un procedimiento para capturar la funcion que carga personal
    Sub registrar_Personal(ByVal ci As Integer, ByVal nombre As String, ByVal apellido As String, ByVal cargo As String, ByVal calle As String, ByVal numero As Integer, ByVal usuario As String, ByVal contrasena As String)
        ObjCapaDatos.registrarPersonal(ci, nombre, apellido, cargo, calle, numero, usuario, contrasena)
    End Sub
    Sub registrar_Telefono(ByVal ci As Integer, ByVal nombre As String, ByVal apellido As String, ByVal telefono_1 As String, ByVal telefono_2 As String)
        ObjCapaDatos.registrarTelefono(ci, nombre, apellido, telefono_1, telefono_2)
    End Sub
    ''Creamos un procedimiento para capturar la funcion que modifica personal
    Sub modificar_Personal(ByVal ci, ByVal nombre, ByVal apellido, ByVal cargo, ByVal calle, ByVal numero, ByVal email, ByVal contrasena)
        ObjCapaDatos.modificarPersonal(ci, nombre, apellido, cargo, calle, numero, email, contrasena)
    End Sub
    Sub modificar_Telefono(ByVal ci, ByVal telefono_1, ByVal telefono_2)
        ObjCapaDatos.modificarTelefono(ci, telefono_1, telefono_2)
    End Sub
    Sub borrar_personal(ByVal ci)
        ObjCapaDatos.borrarPersonal(ci)
    End Sub
    Sub borrar_telefono(ByVal ci)
        ObjCapaDatos.borrarTelefono(ci)
    End Sub
    Sub restaurar_Personal(ByVal ci)
        ObjCapaDatos.restaurarPersonal(ci)
    End Sub
    Sub restaurar_Telefono(ByVal ci)
        ObjCapaDatos.restaurarTelefono(ci)
    End Sub
    Sub eliminar_Personal(ByVal ci)
        ObjCapaDatos.eliminarPersonal(ci)
    End Sub
    Sub eliminar_Telefono(ByVal ci)
        ObjCapaDatos.eliminarTelefono(ci)
    End Sub


    ''Llamamos a la funcion que lista el cargo
    'Function listarCargo() As DataSet
    '    Return ObjCapaDatos.listarCargo
    'End Function
    ''Llamamos a la funcion que autogenera el ID

    'Function generaCodigo() As Integer
    'Return ObjCapaDatos.generaCodigo
    'End Function
End Class


Public Class CESucursales
    ''Declaración de atributos
    Private _IDSucursal As Integer
    Private _Nombre As String
    Private _Telefono As Integer
    Private _Direccion As String
    Private _Ciudad As String
    Private _Capacidad As Integer

    ''Implementación de las propiedades
    Public Property IDSucursal() As Integer
        Get
            Return _IDSucursal
        End Get
        Set(ByVal value As Integer)
            _IDSucursal = value
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
    Public Property Telefono() As Integer
        Get
            Return _Telefono
        End Get
        Set(ByVal value As Integer)
            _Telefono = value
        End Set
    End Property
    Public Property Direccion() As String
        Get
            Return _Direccion
        End Get
        Set(ByVal value As String)
            _Direccion = value
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
End Class

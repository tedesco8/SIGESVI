Imports Capa_Datos
Imports System.Data.Odbc
Imports System.Data.SqlClient
Public Class CNEstan
    ''Creamos un objeto de la clase personal de la capa de datos.
    Dim ObjCapaDatos As New CDEstan
    ''Creamos las variables generales con las cuales trabajaremos en toda la clase.
    Dim dr As OdbcDataReader
    Dim Tabla As New Data.DataTable
    ''Creamos los atributos y propiedades de la clase.
    Private _ID_Lote, _ID_Sucursal, _Cantidad As Integer

    Public Property ID_Lote() As Integer
        Get
            Return _ID_Lote
        End Get
        Set(ByVal value As Integer)
            _ID_Lote = ID_Lote
        End Set
    End Property
    Public Property ID_Sucursal() As Integer
        Get
            Return _ID_Sucursal
        End Get
        Set(ByVal value As Integer)
            _ID_Sucursal = ID_Sucursal
        End Set
    End Property
    Public Property Cantidad() As Integer
        Get
            Return _Cantidad
        End Get
        Set(ByVal value As Integer)
            _Cantidad = Cantidad
        End Set
    End Property

    Public Sub New()

    End Sub
    Public Sub New(ByVal ID_Lote As Integer, ByVal ID_Sucursal As Integer, ByVal Cantidad As Integer)
        ID_Lote = _ID_Lote
        ID_Sucursal = _ID_Sucursal
        Cantidad = _Cantidad
    End Sub
    Function listar_lotes_estan() As DataTable
        Tabla = ObjCapaDatos.listarlotes_estan
        Return Tabla
    End Function
    Sub registrar_lotes_estan(ByVal id_lote, ByVal id_sucursal, ByVal cantidad)
        ObjCapaDatos.registrarlotes_estan(id_lote, id_sucursal, cantidad)
    End Sub
    Sub modificar_lotes_estan(ByVal id_lote, ByVal id_sucursal, ByVal cantidad)
        ObjCapaDatos.modificarlotes_estan(id_lote, id_sucursal, cantidad)
    End Sub
    Sub borrar_Estan(ByVal id_lote)
        ObjCapaDatos.borrarEstan(id_lote)
    End Sub
    Sub restaurar_Estan(ByVal id_lote)
        ObjCapaDatos.restaurarEstan(id_lote)
    End Sub
End Class

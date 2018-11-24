Imports System.Data
Imports System.Data.Odbc
Public Class CDConexion
    ''Con este objeto trabajaremos desde toda la aplicación
    Dim Conexion As OdbcConnection
    ''Creamos una funcion publica que nos permite acceder a el para conectarnos a nuestra base de datos y devuelve la cadena conexion.

    Public Function Conectar()
        Dim cn As String
        cn = "FileDsn=C:\Users\Usuario\Documents\con_maquina.dsn;UID=root;PWD=root"
        Try
            Conexion = New OdbcConnection(cn)
            Conexion.Open()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return Conexion
    End Function

    ''Creamos una sub rutina para cerrar la conexion 
    Public Sub Cerrar()
        Conexion.Close()
    End Sub
End Class

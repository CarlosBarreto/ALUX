Imports System
Imports System.Data
Imports System.Data.SqlClient

Public Class DataAccess

    ''' <summary>
    ''' Almacena el ultimo error ocurrido en la clase
    ''' </summary>
    ''' <remarks></remarks>
    Public _LastErrorMessage As String

    '--- Variables de la base de datos
    ''' <summary>
    ''' Almacena el nombre del servidor de Base de datos
    ''' </summary>
    ''' <remarks></remarks>
    Protected _dbServer As String
    ''' <summary>
    ''' Almacena el nombre de la base de datos
    ''' </summary>
    ''' <remarks></remarks>
    Protected _dbDataBase As String
    ''' <summary>
    ''' Almacena el Id de usuario de la base de datos
    ''' </summary>
    ''' <remarks></remarks>
    Protected _dbUser As String
    ''' <summary>
    ''' Almacena el password de acceso a la base de datos
    ''' </summary>
    ''' <remarks></remarks>
    Protected _dbPassword As String

    Protected sCon As String
    Protected RD As SqlDataReader
    Protected Connection As SqlConnection

    '--- Propiedades de la clase
    ''' <summary>
    ''' Contiene el ultimo error ocurrido en el sistema
    ''' </summary>
    ''' <value>El valor inicial es una cadena en vacia</value>
    ''' <returns>Retorna el ultimo error ocurrido en la clase</returns>
    ''' <remarks>El sistema no muestra las exepciones que ocurren, debido a esto es preferible
    '''          que se guarde el ultimo mensaje de error para que pueda ser mostrado en la 
    '''          página
    ''' </remarks>
    Public ReadOnly Property LastErrorMessage() As String
        Get
            If _LastErrorMessage Is Nothing Then _LastErrorMessage = String.Empty
            Return _LastErrorMessage
        End Get
    End Property

    ''' <summary>
    ''' Contiene el nombre del servidor de Base de Datos
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected ReadOnly Property db_Server() As String
        Get
            If _dbServer Is Nothing Then _dbServer = String.Empty
            Return _dbServer
        End Get
    End Property

    ''' <summary>
    ''' Contiene el nombre de la Base de Datos Accesada
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected ReadOnly Property db_DataBase() As String
        Get
            If _dbDataBase Is Nothing Then _dbDataBase = String.Empty
            Return _dbDataBase
        End Get
    End Property

    ''' <summary>
    ''' Contiene el ID de usuario de acceso a la Base de Datos
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected ReadOnly Property db_User() As String
        Get
            If _dbUser Is Nothing Then _dbUser = String.Empty
            Return _dbUser
        End Get
    End Property

    ''' <summary>
    ''' Contiene el ID password de acceso a la Base de Datos
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected ReadOnly Property db_Password() As String
        Get
            If _dbPassword Is Nothing Then _dbPassword = String.Empty
            Return _dbPassword
        End Get
    End Property

    ''' <summary>
    ''' Propiedad que contiene el READER de la Base de datos
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected ReadOnly Property Reader As SqlDataReader
        Get
            Return RD
        End Get
    End Property

    ''' <summary>
    ''' Constructor SobreCargado
    ''' </summary>
    ''' <param name="Server"></param>
    ''' <param name="DataBase"></param>
    ''' <param name="UserID"></param>
    ''' <param name="Password"></param>
    ''' <remarks></remarks>
    Protected Friend Sub New(Server As String, DataBase As String, UserID As String, Password As String)
        Try
            'Inicializar las variables
            _dbServer = Server
            _dbDataBase = DataBase
            _dbUser = UserID
            _dbPassword = Password

            'Generar la cadena de conexion
            sCon = CadenaConexion()
            'Establecer la conexion
            Connection = New SqlConnection(sCon)
            Connection.Open()
            _LastErrorMessage = ""
        Catch ex As Exception
            _LastErrorMessage = ex.Message
        End Try
    End Sub

    ''' <summary>
    ''' Destructor
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Dispose()
        Connection.Close()
        Connection = Nothing
        RD = Nothing
        sCon = Nothing
        _LastErrorMessage = ""
    End Sub

    ''' <summary>
    ''' Genera una cadena de conexión / Make a new Connection String
    ''' </summary>
    ''' <returns>Connection String</returns>
    ''' <remarks></remarks>
    Protected Function CadenaConexion() As String
        'LoadXMLDBConfig()

        Dim csb As New SqlConnectionStringBuilder
        csb.DataSource = _dbServer
        csb.InitialCatalog = _dbDataBase
        csb.UserID = _dbUser
        csb.Password = _dbPassword

        Return csb.ConnectionString
    End Function

    ''' <summary>
    ''' Ejecuta un Query SQL / Execute a SQL Query
    ''' </summary>
    ''' <param name="strSQL">SQL Query</param>
    ''' <returns>Retorna un DataReader / Returns a DataReader</returns>
    ''' <remarks></remarks>
    Public Function GetRecordsetQuery(ByVal strSQL As String) As SqlDataReader
        Dim cmd As New SqlCommand(strSQL, Connection)
        cmd.CommandType = CommandType.Text
        RD = cmd.ExecuteReader
        Return RD
    End Function

    ''' <summary>
    ''' Executa un Query SQL sin retornar valores / Execute a SQL Query without return values
    ''' </summary>
    ''' <param name="strSQL">SQL Query</param>
    ''' <remarks></remarks>
    Public Sub ExecuteQuery(ByVal strSQL As String)
        Dim cmd As New SqlCommand(strSQL, Connection)
        cmd.CommandType = CommandType.Text
        cmd.ExecuteNonQuery()
    End Sub

    ''' <summary>
    ''' Ejecuta un Stored Procedure con parametros / Execute a Stored Procedure with parameters 
    ''' </summary>
    ''' <param name="SPName">Nombre del Stored Procedure / Stored Procedure Name</param>
    ''' <param name="Parameters">Parametros del StoredProcedure / Parameters Array </param>
    ''' <returns>Retorna un DataReader / Returns a SQLDataReader</returns>
    ''' <remarks></remarks>
    Public Overloads Function ExecuteSP(ByVal SPName As String, ByVal ParamArray Parameters() As Object) As SqlDataReader
        Dim cmd As New SqlCommand(SPName, Connection)
        Dim Obj As Object
        Dim Cont As Integer

        cmd.CommandType = CommandType.StoredProcedure
        Cont = 0

        Try
            SqlCommandBuilder.DeriveParameters(cmd)
            For Each Obj In Parameters
                Cont = Cont + 1
                cmd.Parameters(Cont).Value = Obj
            Next
            RD = cmd.ExecuteReader()
            _LastErrorMessage = ""
        Catch ex As Exception
            'Error cachado desde metodo exterior -- Me.Dispose()
            _LastErrorMessage = ex.Message
        End Try

        Return RD
    End Function

    ''' <summary>
    ''' Ejecuta un Stored Procedured sin parametros / Exec a Stored Procedure without parameters
    ''' </summary>
    ''' <param name="SPName">Nombre del Stored Procedure / Stored Procedure Name</param>
    ''' <returns>Retorna un dataReader / Returns SQLDataReader</returns>
    ''' <remarks></remarks>
    Public Overloads Function ExecuteSP(ByVal SPName As String) As SqlDataReader
        Dim cmd As New SqlCommand(SPName, Connection)
        cmd.CommandType = CommandType.StoredProcedure
        RD = cmd.ExecuteReader()
        Return RD
    End Function

    Protected Function toDataTable() As DataTable
        Dim DT As New DataTable
        DT.Load(RD)
        Return DT
    End Function
End Class

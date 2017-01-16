Imports System.Data.SqlClient

''' <summary>
''' SQL Connection Class for Connecting to and executing T-SQL and stored procedures
''' </summary>
''' <remarks></remarks>
Public Class SQL
    Implements IDisposable

    ' Evelutio Ltd
    ' Created By: Andrew Cunningham
    ' Created Date: 01/03/2017
    ' Description: SQL Connection Class for Connecting to and executing T-SQL and stored procedures
    ' Ammended By: 
    ' Ammended Date: 
    ' Description:


#Region "Connection"

    Private mConnectionString As String = ""
    Private mCommandTimeOut As Int16
    Private mConnectTimeOut As Int16
    Private mTracing As Boolean
    Private disposedValue As Boolean = False    ' To detect redundant calls

    Public Property ConnectionString() As String
        Get

            Return mConnectionString
        End Get
        Set(ByVal Value As String)
            mConnectionString = Value
        End Set
    End Property

    Public Property ConnectTimeOut() As Int16
        Get
            Return mConnectTimeOut
        End Get
        Set(ByVal Value As Int16)
            mConnectTimeOut = Value
        End Set
    End Property

    Public Property CommandTimeOut() As Int16
        Get
            Return mCommandTimeOut
        End Get
        Set(ByVal Value As Int16)
            mCommandTimeOut = Value
        End Set
    End Property

    '----------------------------------------------------------------------------
    ' TracingOn
    '----------------------------------------------------------------------------
    ''' <summary>
    '''  Set this TRUE to trace the SQL of all commands to debug or trace 
    '''  (depending on context)  
    ''' </summary>
    Public Property TracingOn() As Boolean
        Get
            Return mTracing
        End Get
        Set(ByVal value As Boolean)
            mTracing = value
        End Set
    End Property

    Public Function ConnectDB() As SqlClient.SqlConnection
        ' Evelutio Ltd
        ' Created By: Andrew Cunningham
        ' Created Date: 01/03/2017
        ' Description: Create a connection to SQL this connection can then be used to execute queries etc
        ' Ammended By: 
        ' Ammended Date: 
        ' Description:

        Dim con As New SqlConnection

        Try
            mConnectionString = ParseConnectionString(
                                  mConnectionString, mCommandTimeOut)
            con = New SqlConnection(mConnectionString.ToString)
            con.Open()


            Return con

        Catch ex As System.Exception
            Throw New System.Exception(ex.ToString)

        End Try

    End Function

    Public Function ConnectDB(ByVal sqlConnectionString As String,
                              Optional ByVal OverrideCommandTimeout As Int32 = -1,
                              Optional ByVal OverrideConnectTimeout As Int32 = -1,
                              Optional ByVal bTracingOn As Boolean = False) As SqlConnection
        ' Evelutio Ltd
        ' Created By: Andrew Cunningham
        ' Created Date: 01/03/2017
        ' Description: Allows the user to pass the parameters, connection string and timeout and sets the private vars then called the ConnectDB
        ' Ammended By: 
        ' Ammended Date: 
        ' Description:

        mConnectionString = sqlConnectionString
        If (OverrideCommandTimeout > -1) Then
            mCommandTimeOut = OverrideCommandTimeout
        End If
        If (OverrideConnectTimeout > -1) Then
            mConnectTimeOut = OverrideConnectTimeout
        End If

        '--------------------------------------------------------------------------
        ' Explicitly force tracing on if specified otherwise use
        '  the value set when the class is initialized based on
        '  whether we are running with tracing enabled in the calling application
        '--------------------------------------------------------------------------
        If (bTracingOn) Then
            mTracing = bTracingOn
        End If

        Return ConnectDB()
    End Function

    '---------------------------------------------------------------------------
    ' ParseConnectionString (Shared)
    '---------------------------------------------------------------------------
    ''' <summary>
    '''  Parse passed in connection string and validate it as well as appending
    '''  any supplied time out  
    ''' </summary>
    ''' <remarks>Code culled from the ConnectDB method in 
    '''  as we need to be able to set up the
    '''  connection without opening it</remarks>
    Public Shared Function ParseConnectionString(ByVal strConnectionString As String,
                                                 ByVal iConnectTimeOut As Int32) _
                                                   As String

        ' See if connecton string is empty or not valid e.g. less than 2 characters
        If IsNothing(strConnectionString) Then

            Throw New System.Exception("Connection string is empty")
            Exit Function

        ElseIf strConnectionString.Length < 2 Then

            Throw New System.Exception("Connection string is empty")
            Exit Function

        End If

        ' Set connection string timeout

        Dim str() As String = strConnectionString.Split(";")
        Dim blnFound As Boolean = False

        For Each s As String In str
            s = s.ToLower
            If s.IndexOf("timeout") > -1 Then
                ' found
                s = "Timeout = " & iConnectTimeOut
                blnFound = True
                Exit For
            End If
        Next

        strConnectionString = ""
        Dim blnFirs As Boolean = True
        For Each s As String In str
            If Not blnFirs Then strConnectionString &= ";"
            strConnectionString &= s
            blnFirs = False
        Next

        If Not blnFound Then
            strConnectionString &= ";Timeout = " & iConnectTimeOut
        End If

        Return strConnectionString

    End Function

    Public Sub CloseConnection(ByRef Con As SqlConnection)
        ' Evelutio Ltd
        ' Created By: Andrew Cunningham
        ' Created Date: 01/03/2017
        ' Description: Closes the SQL Connection Variable that has been passed to this sub
        ' Ammended By: 
        ' Ammended Date: 
        ' Description:

        If Not IsNothing(Con) Then
            If Not Con.State = ConnectionState.Closed Then
                Con.Close()
            End If
            Con = Nothing

        End If
    End Sub
#End Region

#Region "  Parameters"

    ''' <summary>
    ''' Creates and returns an instance of a parameter
    ''' </summary>
    ''' <param name="Name">Name of the parameter</param>
    ''' <param name="sqlType">Type of the parameter</param>
    ''' <param name="Value">Value of the parameter</param>
    ''' <param name="Size">Optional Size of the parameter </param>
    ''' <returns>SqlParameter</returns>
    ''' <remarks></remarks>
    Public Function Createparameter(ByVal Name As String, ByVal sqlType As SqlDbType,
                    ByVal Value As Object, Optional ByVal Size As Int32 = 50) As SqlParameter

        ' Evelutio Ltd
        ' Created By: Andrew Cunningham
        ' Created Date: 01/03/2017
        ' Description: Function creates and returns an instance of a parameter
        ' Ammended By: 
        ' Ammended Date: 
        ' Description:

        Dim p As New SqlParameter

        Try
            'If name does not start with a @ then add one
            If Not Name.StartsWith("@") Then Name = "@" & Name


            With p
                'Create the parameter settings
                p.ParameterName = Name
                p.SqlDbType = sqlType


                If Value Is Nothing Then
                    .Value = DBNull.Value
                ElseIf Not IsDBNull(Value) Then

                    Select Case sqlType
                        Case SqlDbType.BigInt, SqlDbType.Int, SqlDbType.TinyInt, SqlDbType.SmallInt
                            .Value = CInt(Value)
                        Case SqlDbType.Char, SqlDbType.VarChar, SqlDbType.NVarChar, SqlDbType.Text
                            .Value = CStr(Value)
                            .Size = CInt(Size)
                        Case SqlDbType.DateTime, SqlDbType.SmallDateTime
                            .Value = CDate(Value)
                        Case SqlDbType.Decimal, SqlDbType.SmallMoney, SqlDbType.Money, SqlDbType.Real
                            .Value = CDbl(Value)
                        Case SqlDbType.UniqueIdentifier
                            .Value = DirectCast(Value, Guid)
                        Case Else
                            .Value = Value
                    End Select

                Else
                    .Value = Value
                End If

            End With


            Return p

        Catch ex As System.Exception
            Throw New System.Exception(ex.ToString)

        End Try

    End Function

    ''' <summary>
    ''' Creates and returns an instance of an output parameter
    ''' </summary>
    ''' <param name="Name">Name of the parameter</param>
    ''' <param name="sqlType">Type of the parameter</param>
    ''' <param name="Size">Optional Size of the parameter </param>
    ''' <returns>SqlParameter</returns>
    ''' <remarks></remarks>

    Public Function CreateOutputParameter(ByVal Name As String, ByVal sqlType As SqlDbType, Optional ByVal size As Int32 = 0) As SqlParameter
        ' Evelutio Ltd
        ' Created By: Andrew Cunningham
        ' Created Date: 01/03/2017
        ' Description: Function creates and returns an instance of an output parameter
        ' Ammended By: 
        ' Ammended Date: 
        ' Description:

        Dim p As SqlParameter = New SqlParameter
        Dim ParamName As String = ""

        Try

            With p
                'If name does not start with a @ then add one
                If Not Name.StartsWith("@") Then
                    ParamName = "@"
                End If
                ParamName &= Name
                .ParameterName = ParamName

                ' direction
                .Direction = ParameterDirection.Output
                ' type
                .SqlDbType = sqlType
                ' size
                Select Case sqlType
                    Case SqlDbType.Char, SqlDbType.VarChar, SqlDbType.NChar, SqlDbType.NVarChar, SqlDbType.NText, SqlDbType.Text
                        .Size = CInt(size)
                End Select

            End With

            Return p

        Catch ex As System.Exception

            Throw New System.Exception(ex.ToString)

        End Try

    End Function


#End Region

#Region "   Execute Query / Command"

    ''' <summary>
    ''' Create an SQLCommand Object with n parameters
    ''' </summary>
    ''' <param name="cmd">SqlCommand</param>
    ''' <param name="params">An array of Sql Parameters</param>
    ''' <returns>SqlCommand</returns>
    ''' <remarks></remarks>
    Private Function AddParams(ByVal cmd As SqlCommand, ByVal params() As SqlParameter) As SqlCommand
        ' Evelutio Ltd
        ' Created By: Andrew Cunningham
        ' Created Date: 01/03/2017
        ' Description: Create an SQLCommand Object with n parameters
        ' Ammended By: 
        ' Ammended Date: 
        ' Description:

        Try
            With cmd
                .CommandTimeout = mCommandTimeOut
                .CommandType = CommandType.StoredProcedure
                For Each p As SqlParameter In params
                    If (Not IsNothing(p)) Then
                        .Parameters.Add(CheckPrefix(p))
                    End If
                Next

            End With
            Return cmd

        Catch ex As System.Exception

            Throw

        End Try

    End Function

    ''' <summary>
    ''' Create an SQLCommand Object with upto 10 parameters
    ''' </summary>
    ''' <param name="cmd">SqlCommand</param>
    ''' <param name="param1"></param>
    ''' <param name="param2"></param>
    ''' <param name="param3"></param>
    ''' <param name="param4"></param>
    ''' <param name="param5"></param>
    ''' <param name="param6"></param>
    ''' <param name="param7"></param>
    ''' <param name="param8"></param>
    ''' <param name="param9"></param>
    ''' <param name="param10"></param>
    ''' <returns>SqlCommand</returns>
    ''' <remarks></remarks>
    Private Function AddParams(ByRef cmd As SqlCommand,
            Optional ByVal param1 As SqlParameter = Nothing,
            Optional ByVal param2 As SqlParameter = Nothing,
            Optional ByVal param3 As SqlParameter = Nothing,
            Optional ByVal param4 As SqlParameter = Nothing,
            Optional ByVal param5 As SqlParameter = Nothing,
            Optional ByVal param6 As SqlParameter = Nothing,
            Optional ByVal param7 As SqlParameter = Nothing,
            Optional ByVal param8 As SqlParameter = Nothing,
            Optional ByVal param9 As SqlParameter = Nothing,
            Optional ByVal param10 As SqlParameter = Nothing) As SqlCommand

        ' Evelutio Ltd
        ' Created By: Andrew Cunningham
        ' Created Date: 01/03/2017
        ' Description: Create an SQLCommand Object with upto 10 parameters
        ' Ammended By: 
        ' Ammended Date: 
        ' Description:

        Try

            ' checks for prefix "@"
            With cmd
                cmd.CommandType = CommandType.StoredProcedure
                If Not IsNothing(param1) Then : .Parameters.Add(CheckPrefix(param1)) : End If
                If Not IsNothing(param2) Then : .Parameters.Add(CheckPrefix(param2)) : End If
                If Not IsNothing(param3) Then : .Parameters.Add(CheckPrefix(param3)) : End If
                If Not IsNothing(param4) Then : .Parameters.Add(CheckPrefix(param4)) : End If
                If Not IsNothing(param5) Then : .Parameters.Add(CheckPrefix(param5)) : End If
                If Not IsNothing(param6) Then : .Parameters.Add(CheckPrefix(param6)) : End If
                If Not IsNothing(param7) Then : .Parameters.Add(CheckPrefix(param7)) : End If
                If Not IsNothing(param8) Then : .Parameters.Add(CheckPrefix(param8)) : End If
                If Not IsNothing(param9) Then : .Parameters.Add(CheckPrefix(param9)) : End If
                If Not IsNothing(param10) Then : .Parameters.Add(CheckPrefix(param10)) : End If
            End With
            Return cmd
        Catch ex As System.Exception
            Throw
        End Try


    End Function

    ''' <summary>
    ''' Checks for prefix @ and includes it if missing
    ''' </summary>
    ''' <param name="prm">SqlParameter</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckPrefix(ByVal prm As SqlParameter) As SqlParameter
        ' Evelutio Ltd
        ' Created By: Andrew Cunningham
        ' Created Date: 01/03/2017
        ' Description: Checks for prefix @ and includes it if missing
        ' Ammended By: 
        ' Ammended Date: 
        ' Description:

        Try
            Dim strName As String = ""
            strName = prm.ParameterName

            If Not strName.StartsWith("@") Then
                prm.ParameterName = "@" & strName
            End If
            Return prm
        Catch ex As System.Exception
            Throw
        End Try

    End Function

    ''' <summary>
    ''' Executes an SQL Text Statement using the Execute Reader Command data returned as a data reader object
    ''' </summary>
    ''' <param name="con">SqlConnection</param>
    ''' <param name="SQL_Text">SQL Text</param>
    ''' <returns>SqlDataReader</returns>
    ''' <remarks></remarks>




    Public Function ExecuteSQLQueryReader(ByVal con As SqlConnection, ByVal SQL_Text As String) As SqlDataReader
        ' Evelutio Ltd
        ' Created By: Andrew Cunningham
        ' Created Date: 01/03/2017
        ' Description: Executes an SQL Text Statement using the Execute Reader Command data returned as a data reader object
        ' Ammended By: 
        ' Ammended Date: 
        ' Description:

        Dim sdr As SqlDataReader
        Dim cmd As New SqlCommand

        Try

            cmd.Connection = con
            cmd.CommandText = SQL_Text
            cmd.CommandTimeout = mCommandTimeOut
            cmd.CommandType = CommandType.Text

            DebugWrite(SQL_Text)
            ' Retrieve all rows
            sdr = cmd.ExecuteReader()

            cmd = Nothing


            Return sdr

        Catch ex As System.Exception
            cmd = Nothing

            Throw
        End Try

    End Function

    ''' <summary>
    ''' Executes an SQL Stored Proc using the Execute Reader Command data returned as a data reader object
    ''' </summary>
    ''' <param name="Con">SqlConnection</param>
    ''' <param name="sp_Name">Stored Procedure</param>
    ''' <param name="DataTableName">Data Table Name</param>
    ''' <param name="param1"></param>
    ''' <param name="param2"></param>
    ''' <param name="param3"></param>
    ''' <param name="param4"></param>
    ''' <param name="param5"></param>
    ''' <param name="param6"></param>
    ''' <param name="param7"></param>
    ''' <param name="param8"></param>
    ''' <param name="param9"></param>
    ''' <param name="param10"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecuteQueryReaderSP(ByRef Con As SqlConnection,
               ByVal sp_Name As String,
               ByVal DataTableName As String,
               Optional ByVal param1 As SqlParameter = Nothing,
               Optional ByVal param2 As SqlParameter = Nothing,
               Optional ByVal param3 As SqlParameter = Nothing,
               Optional ByVal param4 As SqlParameter = Nothing,
               Optional ByVal param5 As SqlParameter = Nothing,
               Optional ByVal param6 As SqlParameter = Nothing,
               Optional ByVal param7 As SqlParameter = Nothing,
               Optional ByVal param8 As SqlParameter = Nothing,
               Optional ByVal param9 As SqlParameter = Nothing,
               Optional ByVal param10 As SqlParameter = Nothing) As SqlDataReader

        ' Evelutio Ltd
        ' Created By: Andrew Cunningham
        ' Created Date: 01/03/2017
        ' Description: Executes an SQL Stored Proc using the Execute Reader Command data returned as a data reader object
        ' Ammended By: 
        ' Ammended Date: 
        ' Description:

        Dim cmd As New SqlCommand(sp_Name)

        cmd = AddParams(cmd, param1, param2, param3, param4, param5, param6, param7, param8, param9, param10)

        Try
            cmd.Connection = Con
            cmd.CommandTimeout = mCommandTimeOut

            Dim sdr As SqlDataReader

            mDebugSQLCommand(cmd)
            ' Retrieve all rows
            sdr = cmd.ExecuteReader()
            mDebugSQLCommand(cmd)

            cmd = Nothing

            Return sdr

        Catch ex As System.Exception
            cmd = Nothing

            Throw
        End Try

    End Function

    ''' <summary>
    ''' Executes an SQL Stored Proc using the Execute Reader Command data returned as a data reader object
    ''' </summary>
    ''' <param name="Con">SqlConnection</param>
    ''' <param name="sp_Name">Stored Procedure</param>
    ''' <param name="params">Parameter array of SqlParameters to add to the command</param>  
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecuteQueryReaderSP(ByRef Con As SqlConnection,
                                         ByVal sp_Name As String,
                                         ByVal ParamArray params() As SqlParameter) As SqlDataReader
        '--------------------------------------------------------------------------
        ' Local Variables
        '--------------------------------------------------------------------------
        Dim cmd As New SqlCommand(sp_Name)
        Dim sdr As SqlDataReader

        cmd = AddParams(cmd, params)

        Try
            cmd.Connection = Con
            cmd.CommandTimeout = mCommandTimeOut

            mDebugSQLCommand(cmd)
            sdr = cmd.ExecuteReader()
            mDebugSQLCommand(cmd)

            cmd = Nothing

            Return sdr

        Catch ex As System.Exception
            cmd = Nothing
            Throw
        End Try

    End Function

    ''' <summary>
    ''' Executes an SQL Stored Proc using the Execute Data Adapeter Command data returned as a DataSet
    ''' </summary>
    ''' <param name="Con">SqlConnection</param>
    ''' <param name="sp_Name">Stored Procedure Name</param>
    ''' <param name="params"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecuteDatasetQuery(ByRef Con As SqlConnection, ByVal sp_Name As String,
            ByVal params() As SqlParameter) As DataSet

        Dim cmd As New SqlCommand(sp_Name)
        Dim ds As New System.Data.DataSet
        cmd = AddParams(cmd, params)

        Try

            cmd.Connection = Con
            cmd.CommandTimeout = mCommandTimeOut
            Dim adp As New SqlDataAdapter(cmd)
            mDebugSQLCommand(cmd)
            adp.Fill(ds)
            mDebugSQLCommand(cmd)
            cmd = Nothing

        Catch ex As System.Exception
            cmd = Nothing

            Throw
        End Try

        Return ds

    End Function

    ''' <summary>
    ''' Executes an SQL Stored Proc using the Execute Data Adapeter Command data returned as a DataSet
    ''' </summary>
    ''' <param name="Con">SqlConnection</param>
    ''' <param name="sp_Name">Stored Procedure Name</param>
    ''' <param name="param1"></param>
    ''' <param name="param2"></param>
    ''' <param name="param3"></param>
    ''' <param name="param4"></param>
    ''' <param name="param5"></param>
    ''' <param name="param6"></param>
    ''' <param name="param7"></param>
    ''' <param name="param8"></param>
    ''' <param name="param9"></param>
    ''' <param name="param10"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecuteDatasetQuery(ByRef Con As SqlConnection, ByVal sp_Name As String,
             Optional ByVal param1 As SqlParameter = Nothing,
            Optional ByVal param2 As SqlParameter = Nothing,
            Optional ByVal param3 As SqlParameter = Nothing,
            Optional ByVal param4 As SqlParameter = Nothing,
            Optional ByVal param5 As SqlParameter = Nothing,
            Optional ByVal param6 As SqlParameter = Nothing,
            Optional ByVal param7 As SqlParameter = Nothing,
            Optional ByVal param8 As SqlParameter = Nothing,
            Optional ByVal param9 As SqlParameter = Nothing,
            Optional ByVal param10 As SqlParameter = Nothing) As DataSet

        Dim cmd As New SqlCommand(sp_Name)
        Dim ds As New System.Data.DataSet
        cmd = AddParams(cmd, param1, param2, param3, param4, param5, param6, param7, param8, param9, param10)

        Try

            cmd.Connection = Con
            cmd.CommandTimeout = mCommandTimeOut
            Dim adp As New SqlDataAdapter(cmd)
            mDebugSQLCommand(cmd)
            adp.Fill(ds)
            mDebugSQLCommand(cmd)
            cmd = Nothing

        Catch ex As System.Exception
            cmd = Nothing

            Throw
        End Try

        Return ds

    End Function


    ''' <summary>
    ''' Executes an SQL Stored Proc using the Execute Data Adapeter Command data returned as a datatable
    ''' </summary>
    ''' <param name="Con">SqlConnection</param>
    ''' <param name="sp_Name">Stored Procedure Name</param>
    ''' <param name="DataTableName">Data Table Name</param>
    ''' <param name="params"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecuteQuery(ByRef Con As SqlConnection, ByVal sp_Name As String,
            ByVal DataTableName As String, ByVal params() As SqlParameter) As DataTable

        ' Evelutio Ltd
        ' Created By: Andrew Cunningham
        ' Created Date: 01/03/2017
        ' Description: Executes an SQL Stored Proc using the Execute Data Adapeter Command data returned as a datatable
        ' Ammended By: 
        ' Ammended Date: 
        ' Description:

        Dim cmd As New SqlCommand(sp_Name)
        Dim tbl As New DataTable(DataTableName)

        cmd = AddParams(cmd, params)

        Try

            cmd.Connection = Con
            cmd.CommandTimeout = mCommandTimeOut
            Dim adp As New SqlDataAdapter(cmd)
            mDebugSQLCommand(cmd)
            adp.Fill(tbl)
            mDebugSQLCommand(cmd)
            cmd = Nothing

        Catch ex As System.Exception
            cmd = Nothing

            Throw
        End Try


        Return tbl
    End Function

    ''' <summary>
    ''' Executes an SQL Stored Proc using the Execute Data Adapeter Command data returned as a datatable
    ''' </summary>
    ''' <param name="Con">SqlConnection</param>
    ''' <param name="sp_Name">Stored Procedure Name</param>
    ''' <param name="DataTableName">Data Table Name</param>
    ''' <param name="param1"></param>
    ''' <param name="param2"></param>
    ''' <param name="param3"></param>
    ''' <param name="param4"></param>
    ''' <param name="param5"></param>
    ''' <param name="param6"></param>
    ''' <param name="param7"></param>
    ''' <param name="param8"></param>
    ''' <param name="param9"></param>
    ''' <param name="param10"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecuteQuery(ByRef Con As SqlConnection,
            ByVal sp_Name As String,
            ByVal DataTableName As String,
            Optional ByVal param1 As SqlParameter = Nothing,
            Optional ByVal param2 As SqlParameter = Nothing,
            Optional ByVal param3 As SqlParameter = Nothing,
            Optional ByVal param4 As SqlParameter = Nothing,
            Optional ByVal param5 As SqlParameter = Nothing,
            Optional ByVal param6 As SqlParameter = Nothing,
            Optional ByVal param7 As SqlParameter = Nothing,
            Optional ByVal param8 As SqlParameter = Nothing,
            Optional ByVal param9 As SqlParameter = Nothing,
            Optional ByVal param10 As SqlParameter = Nothing) As DataTable

        ' Evelutio Ltd
        ' Created By: Andrew Cunningham
        ' Created Date: 01/03/2017
        ' Description: Executes an SQL Stored Proc using the Execute Data Adapeter Command data returned as a datatable
        ' Ammended By: 
        ' Ammended Date: 
        ' Description:

        Dim cmd As New SqlCommand(sp_Name)
        Dim tbl As New DataTable(DataTableName)

        cmd = AddParams(cmd, param1, param2, param3, param4, param5, param6, param7, param8, param9, param10)

        Try

            cmd.Connection = Con
            cmd.CommandTimeout = mCommandTimeOut
            Dim adp As New SqlDataAdapter(cmd)
            mDebugSQLCommand(cmd)
            adp.Fill(tbl)
            mDebugSQLCommand(cmd)
            cmd = Nothing

        Catch ex As System.Exception
            cmd = Nothing

            Throw
        End Try


        Return tbl
    End Function

    ''' <summary>
    ''' Executes an SQL Text Statement using the Execute Data Adapeter Command data returned as a datatable 
    ''' </summary>
    ''' <param name="con">SqlConnection</param>
    ''' <param name="SQL_Text">Sql Text</param>
    ''' <param name="DataTableName">Data Table Name</param>
    ''' <returns></returns>
    ''' <remarks>this functino takes in sql text and returns a data table object</remarks>

    Public Function ExecuteSQLQuery(ByVal con As SqlConnection, ByVal SQL_Text As String, ByVal DataTableName As String) As DataTable
        ' Evelutio Ltd
        ' Created By: Andrew Cunningham
        ' Created Date: 01/03/2017
        ' Description: Executes an SQL Text Statement using the Execute Data Adapeter Command data returned as a datatable
        ' Ammended By: 
        ' Ammended Date: 
        ' Description:

        Dim tbl As New DataTable(DataTableName)
        Dim cmd As New SqlCommand

        Try

            cmd.Connection = con
            cmd.CommandText = SQL_Text
            cmd.CommandTimeout = mCommandTimeOut
            cmd.CommandType = CommandType.Text
            Dim adp As New SqlDataAdapter(cmd)
            DebugWrite(SQL_Text)
            adp.Fill(tbl)
            cmd = Nothing

        Catch ex As System.Exception
            cmd = Nothing

            Throw
        End Try
        Return tbl
    End Function

    ''' <summary>
    ''' Executes an SQL Text Statement using the ExecuteNonQuery the returned value int indicates if the process was sucessful e.g. if = 0
    ''' </summary>
    ''' <param name="con">SqlConnection</param>
    ''' <param name="SQL_Text">Sql Text</param>
    ''' <returns></returns>
    ''' <remarks>this function takes in sql text and executes against database returning integer</remarks>
    Public Function ExecuteSQLCommand(ByVal con As SqlConnection, ByVal SQL_Text As String) As Integer
        ' Evelutio Ltd
        ' Created By: Andrew Cunningham
        ' Created Date: 01/03/2017
        ' Description: Executes an SQL Text Statement using the ExecuteNonQuery the returned value int indicates if the process was sucessful e.g. if = 0
        ' Ammended By: 
        ' Ammended Date: 
        ' Description:

        Dim intRet As Integer


        Try
            If con.State <> ConnectionState.Open Then
                con = ConnectDB()
            End If
        Catch ex As System.Exception

        End Try

        Dim cmd As New SqlCommand

        Try
            cmd.Connection = con
            cmd.CommandText = SQL_Text
            cmd.CommandTimeout = mCommandTimeOut
            cmd.CommandType = CommandType.Text
            DebugWrite(SQL_Text)
            intRet = cmd.ExecuteNonQuery
            cmd = Nothing



        Catch ex As System.Exception
            cmd = Nothing

            Throw

        End Try


        Return intRet

    End Function

    ''' <summary>
    ''' Executes an SQL Stored Proc with n parameters using the ExecuteNonQuery the returned value int indicates if the process was sucessful e.g. if = 0
    ''' </summary>
    ''' <param name="con">SqlConnection</param>
    ''' <param name="sp_Name">Stored Procedure Name</param>
    ''' <param name="params">An Array of Sql Parameters</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecuteCommand(ByRef con As SqlConnection, ByVal sp_Name As String,
                                            ByVal params() As SqlParameter) As Integer

        ' Evelutio Ltd
        ' Created By: Andrew Cunningham
        ' Created Date: 01/03/2017
        ' Description: Executes an SQL Stored Proc with n parameters using the ExecuteNonQuery the returned value int indicates if the process was sucessful e.g. if = 0
        ' Ammended By: 
        ' Ammended Date: 
        ' Description:

        Dim cmd As New SqlCommand(sp_Name)
        Dim intRet As Integer

        cmd = AddParams(cmd, params)

        Try
            If con.State <> ConnectionState.Open Then
                con = ConnectDB()
            End If
        Catch ex As System.Exception
            cmd = Nothing

            Throw
            Exit Function
        End Try

        Try
            cmd.Connection = con
            cmd.CommandTimeout = mCommandTimeOut
            mDebugSQLCommand(cmd)
            intRet = cmd.ExecuteNonQuery
            mDebugSQLCommand(cmd)
            cmd = Nothing



        Catch ex As System.Exception
            cmd = Nothing

            Throw

        End Try
        Return intRet

    End Function


    ''' <summary>
    ''' ' Executes an SQL Stored Proc with upto 10 parameters using the ExecuteNonQuery the returned value int indicates if the process was sucessful e.g. if = 0
    ''' </summary>
    ''' <param name="con">SqlConnection</param>
    ''' <param name="sp_Name">Stored Procedure Name</param>
    ''' <param name="param1"></param>
    ''' <param name="param2"></param>
    ''' <param name="param3"></param>
    ''' <param name="param4"></param>
    ''' <param name="param5"></param>
    ''' <param name="param6"></param>
    ''' <param name="param7"></param>
    ''' <param name="param8"></param>
    ''' <param name="param9"></param>
    ''' <param name="param10"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecuteCommand(ByRef con As SqlConnection, ByVal sp_Name As String,
            Optional ByVal param1 As SqlParameter = Nothing,
            Optional ByVal param2 As SqlParameter = Nothing,
            Optional ByVal param3 As SqlParameter = Nothing,
            Optional ByVal param4 As SqlParameter = Nothing,
            Optional ByVal param5 As SqlParameter = Nothing,
            Optional ByVal param6 As SqlParameter = Nothing,
            Optional ByVal param7 As SqlParameter = Nothing,
            Optional ByVal param8 As SqlParameter = Nothing,
            Optional ByVal param9 As SqlParameter = Nothing,
            Optional ByVal param10 As SqlParameter = Nothing) As Integer

        ' Evelutio Ltd
        ' Created By: Andrew Cunningham
        ' Created Date: 01/03/2017
        ' Description: Executes an SQL Stored Proc with upto 10 parameters using the ExecuteNonQuery the returned value int indicates if the process was sucessful e.g. if = 0
        ' Ammended By: 
        ' Ammended Date: 
        ' Description:

        Dim cmd As New SqlCommand(sp_Name)
        Dim intRet As Integer

        cmd = AddParams(cmd, param1, param2, param3, param4, param5, param6, param7, param8, param9, param10)

        Try
            If con.State <> ConnectionState.Open Then
                con = ConnectDB()
            End If

        Catch ex As System.Exception
            cmd = Nothing

            Throw
            Exit Function
        End Try

        Try

            cmd.Connection = con
            cmd.CommandTimeout = mCommandTimeOut
            mDebugSQLCommand(cmd)
            intRet = cmd.ExecuteNonQuery
            mDebugSQLCommand(cmd)
            cmd = Nothing


        Catch ex As System.Exception
            cmd = Nothing


            Throw
        End Try


        Return intRet
    End Function

    ''' <summary>
    ''' Returns first column of a select statement
    ''' Executes an SQL Stored Proc with upto 10 parameters using the ExecuteScalar and it returns first column of select statement
    ''' </summary>
    ''' <param name="con"></param>
    ''' <param name="sp_name">SQL Query</param>
    ''' <param name="param1"></param>
    ''' <param name="param2"></param>
    ''' <param name="param3"></param>
    ''' <param name="param4"></param>
    ''' <param name="param5"></param>
    ''' <param name="param6"></param>
    ''' <param name="param7"></param>
    ''' <param name="param8"></param>
    ''' <param name="param9"></param>
    ''' <param name="param10"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecuteScalar(ByRef con As SqlConnection, ByVal sp_name As String,
          Optional ByVal param1 As SqlParameter = Nothing,
          Optional ByVal param2 As SqlParameter = Nothing,
          Optional ByVal param3 As SqlParameter = Nothing,
          Optional ByVal param4 As SqlParameter = Nothing,
          Optional ByVal param5 As SqlParameter = Nothing,
          Optional ByVal param6 As SqlParameter = Nothing,
          Optional ByVal param7 As SqlParameter = Nothing,
          Optional ByVal param8 As SqlParameter = Nothing,
          Optional ByVal param9 As SqlParameter = Nothing,
          Optional ByVal param10 As SqlParameter = Nothing) As Object

        ' Evelutio Ltd
        ' Created By: JS
        ' Created Date: 04/09/2009
        ' Description: Executes an SQL Stored Proc with upto 10 parameters using the ExecuteScalar and it returns first column of select statement
        ' Ammended By: 
        ' Ammended Date: 
        ' Description:

        Dim cmd As New SqlCommand(sp_name)
        Dim ObjRet As Object

        cmd = AddParams(cmd, param1, param2, param3, param4, param5, param6, param7, param8, param9, param10)

        Try

            If con.State <> ConnectionState.Open Then
                con = ConnectDB()
            End If

        Catch ex As System.Exception
            cmd = Nothing

            Throw
            Exit Function
        End Try

        Try

            cmd.Connection = con
            cmd.CommandTimeout = mCommandTimeOut
            mDebugSQLCommand(cmd)
            ObjRet = cmd.ExecuteScalar
            mDebugSQLCommand(cmd)
            cmd = Nothing


        Catch ex As System.Exception
            cmd = Nothing


            Throw
        End Try


        Return ObjRet
    End Function

    ''' <summary>
    ''' Returns first column of a select statement
    ''' Executes an SQL Stored Proc with upto 10 parameters using the ExecuteScalar and it returns first column of select statement
    ''' </summary>
    ''' <param name="con"></param>
    ''' <param name="sp_name">SQL Query</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecuteScalar(ByRef con As SqlConnection, ByVal sp_name As String,
        ByVal params() As SqlParameter) As Object

        ' Evelutio Ltd
        ' Created By: JS
        ' Created Date: 04/09/2009
        ' Description: Executes an SQL Stored Proc with upto 10 parameters using the ExecuteScalar and it returns first column of select statement
        ' Ammended By: 
        ' Ammended Date: 
        ' Description:

        Dim cmd As New SqlCommand(sp_name)
        Dim ObjRet As Object

        cmd = AddParams(cmd, params)

        Try

            If con.State <> ConnectionState.Open Then
                con = ConnectDB()
            End If

        Catch ex As System.Exception
            cmd = Nothing

            Throw
            Exit Function
        End Try

        Try

            cmd.Connection = con
            cmd.CommandTimeout = mCommandTimeOut
            mDebugSQLCommand(cmd)
            ObjRet = cmd.ExecuteScalar
            mDebugSQLCommand(cmd)
            cmd = Nothing


        Catch ex As System.Exception
            cmd = Nothing


            Throw
        End Try


        Return ObjRet
    End Function

    ''' <summary>
    ''' Returns first column of a select statement
    ''' Executes an SQL string with upto 10 parameters using the ExecuteScalar and it returns first column of a select statement
    ''' </summary>
    ''' <param name="con"></param>
    ''' <param name="SQL_Text">SQL Query</param>
    ''' <param name="param1"></param>
    ''' <param name="param2"></param>
    ''' <param name="param3"></param>
    ''' <param name="param4"></param>
    ''' <param name="param5"></param>
    ''' <param name="param6"></param>
    ''' <param name="param7"></param>
    ''' <param name="param8"></param>
    ''' <param name="param9"></param>
    ''' <param name="param10"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecuteSQLScalar(ByRef con As SqlConnection, ByVal SQL_Text As String,
          Optional ByVal param1 As SqlParameter = Nothing,
          Optional ByVal param2 As SqlParameter = Nothing,
          Optional ByVal param3 As SqlParameter = Nothing,
          Optional ByVal param4 As SqlParameter = Nothing,
          Optional ByVal param5 As SqlParameter = Nothing,
          Optional ByVal param6 As SqlParameter = Nothing,
          Optional ByVal param7 As SqlParameter = Nothing,
          Optional ByVal param8 As SqlParameter = Nothing,
          Optional ByVal param9 As SqlParameter = Nothing,
          Optional ByVal param10 As SqlParameter = Nothing) As Object

        ' Evelutio Ltd
        ' Created By: JS
        ' Created Date: 04/09/2009
        ' Description: Executes an SQL string with upto 10 parameters using the ExecuteScalar and it returns first column of a select statement
        ' Ammended By: 
        ' Ammended Date: 
        ' Description:

        Dim cmd As New SqlCommand
        Dim ObjRet As Object

        cmd = AddParams(cmd, param1, param2, param3, param4, param5, param6, param7, param8, param9, param10)

        Try

            cmd.CommandText = SQL_Text
            cmd.CommandType = CommandType.Text

            If con.State <> ConnectionState.Open Then
                con = ConnectDB()
            End If

        Catch ex As System.Exception
            cmd = Nothing

            Throw
            Exit Function
        End Try

        Try

            cmd.Connection = con
            cmd.CommandTimeout = mCommandTimeOut
            mDebugSQLCommand(cmd)
            ObjRet = cmd.ExecuteScalar
            mDebugSQLCommand(cmd)
            cmd = Nothing


        Catch ex As System.Exception
            cmd = Nothing


            Throw
        End Try

        Return ObjRet
    End Function

#End Region

#Region "Debug and tracing"
    '-----------------------------------------------------------------------------
    ' mDebugSQLCommand
    '-----------------------------------------------------------------------------
    ''' <summary>
    '''  Member version of DebugSQLCommand that checks tracing property on the class
    ''' </summary>
    Private Function mDebugSQLCommand(ByVal SQLcmd1 As SqlClient.SqlCommand) As String
        '---------------------------------------------------------------------------
        ' If we are not tracing then exit
        '---------------------------------------------------------------------------
        If (Not TracingOn) Then
            Return Nothing
        End If

        Return DebugSQLCommand(SQLcmd1)

    End Function

    '-----------------------------------------------------------------------------
    ' DebugSQLCommand
    '-----------------------------------------------------------------------------
    ''' <summary>
    '''  Take the passed in SQL command object and write out the full SQL
    '''  syntax of the command (so it can be executed in SQL Management studio)  
    '''  to debug window if in windows context or trace if web context  
    ''' </summary>
    Public Shared Function DebugSQLCommand(ByVal SqlCmd1 As SqlClient.SqlCommand) As String
        '---------------------------------------------------------------------------
        ' Local variables
        '---------------------------------------------------------------------------
        Dim strLogDetail As String = ""
        Dim strDeclare As String = ""
        Dim strOutput As String = ""
        Dim strParamName As String = ""
        Dim strParamType As String = ""
        Dim paramThis As SqlClient.SqlParameter
        Dim iIndx As Integer

        '---------------------------------------------------------------------------
        ' Build up the command text for both tracing and logging
        '---------------------------------------------------------------------------
        If (SqlCmd1.CommandType = CommandType.StoredProcedure) Then
            strLogDetail &= vbNewLine & "EXEC @ret = " & SqlCmd1.CommandText
            strDeclare &= vbNewLine & "DECLARE @ret int"
            strOutput = vbNewLine & "@ret AS [ReturnValue],"
        End If

        '---------------------------------------------------------------------------
        ' Go through all the parameters and build SQL statement to log to file
        '---------------------------------------------------------------------------
        For iIndx = 0 To SqlCmd1.Parameters.Count - 1
            '-------------------------------------------------------------------------
            ' Process one parameter
            '-------------------------------------------------------------------------
            paramThis = SqlCmd1.Parameters.Item(iIndx)
            strParamName = paramThis.ParameterName

            '-------------------------------------------------------------------------
            ' Ensure parameter name starts with "@"
            '-------------------------------------------------------------------------
            If (strParamName.IndexOf("@"c) = -1) Then
                strParamName = "@" & strParamName
            End If

            If (paramThis.Direction = ParameterDirection.ReturnValue) Then
                strParamName = "--" & strParamName
            End If

            strLogDetail &= vbNewLine
            If (iIndx > 1) Then
                strLogDetail &= ","
            End If

            strLogDetail &= strParamName & "="

            '-------------------------------------------------------------------------
            ' Convert parameter data type to string, and add size if appropriate
            '-------------------------------------------------------------------------
            strParamType = paramThis.SqlDbType.ToString.ToLower
            If (paramThis.Size > 0) Then
                Select Case paramThis.SqlDbType

                    Case SqlDbType.Char, SqlDbType.NChar, SqlDbType.NVarChar,
                         SqlDbType.VarBinary, SqlDbType.VarChar
                        strParamType &= "(" & paramThis.Size.ToString & ")"

                    Case Else

                End Select
            End If

            '-------------------------------------------------------------------------
            ' Generate declare and select statements for OUTPUT parameters
            '-------------------------------------------------------------------------
            If ((paramThis.Direction = ParameterDirection.InputOutput) OrElse
                (paramThis.Direction = ParameterDirection.Output)) Then
                strLogDetail &= strParamName & " OUTPUT -- "
                strOutput &= vbNewLine & strParamName & " AS [" &
                             strParamName.TrimStart("@"c) & "],"
                strDeclare &= vbNewLine & "DECLARE " &
                              strParamName & " " & strParamType
                If ((paramThis.Direction = ParameterDirection.InputOutput) AndAlso
                    (Not IsNothing(paramThis.Value))) Then
                    '---------------------------------------------------------------------
                    ' Input and output, so set local variable value on way in
                    '---------------------------------------------------------------------
                    strDeclare &= vbNewLine & "SET " &
                                  strParamName & " = " & FormatSqlParamValue(paramThis)
                End If
            End If

            strLogDetail &= FormatSqlParamValue(paramThis) &
                            " --(" & strParamType & ", " &
                            paramThis.Direction.ToString & ")"
        Next

        If (strDeclare.Length > 0) Then
            strLogDetail = strDeclare &
                           strLogDetail & vbNewLine &
                           "SELECT " & strOutput.TrimEnd(","c)
        End If

        strLogDetail = "--" & SqlCmd1.CommandType.ToString & ": " & SqlCmd1.CommandText &
                       strLogDetail & vbNewLine & vbNewLine

        '---------------------------------------------------------------------------
        ' Write to debug / trace
        '---------------------------------------------------------------------------
        DebugWrite(strLogDetail)

        '---------------------------------------------------------------------------
        ' return the command SQL for good measure
        '---------------------------------------------------------------------------
        Return strLogDetail

    End Function

    '-----------------------------------------------------------------------------
    ' mDebugWrite
    '-----------------------------------------------------------------------------
    ''' <summary>
    '''  Member version of DebugWrite that checks tracing property on the class
    ''' </summary>
    Private Sub mDebugWrite(ByVal strText As String)
        If (TracingOn) Then
            DebugWrite(strText)
        End If
    End Sub

    '-----------------------------------------------------------------------------
    ' DebugWrite
    '-----------------------------------------------------------------------------
    ''' <summary>
    '''  Write the passed in string to either Trace if we are in Web Context
    '''  or the Debug window if we are in Windows context  
    ''' </summary>
    Public Shared Sub DebugWrite(ByVal strText As String)
        '---------------------------------------------------------------------------
        ' If we are in web context Trace the command otherwise send it to Debug
        ' TODO: FOr some reason HttpContext.Current returns nothing but
        '   the if statement does not respond correctly
        '---------------------------------------------------------------------------
        Dim bHTTP As Boolean = Not IsNothing(System.Web.HttpContext.Current)
        If (bHTTP) Then
            System.Web.HttpContext.Current.Trace.Write(strText)
        Else
            System.Diagnostics.Trace.Write(strText)
        End If
    End Sub

    '-----------------------------------------------------------------------------
    ' FormatSqlParamValue
    '-----------------------------------------------------------------------------
    ''' <summary>
    ''' Format the value of a SqlParameter according to its type, 
    ''' i.e. quote strings
    ''' </summary>
    Shared Function FormatSqlParamValue(ByVal paramThis As SqlClient.SqlParameter) _
                    As String
        '---------------------------------------------------------------------------
        ' Local variables
        '---------------------------------------------------------------------------
        Dim strText As String

        If (IsDBNull(paramThis.Value)) Then
            '-------------------------------------------------------------------------
            ' NULL parameter
            '-------------------------------------------------------------------------
            strText = "NULL"
        ElseIf (IsNothing(paramThis.Value)) Then
            '-------------------------------------------------------------------------
            ' VB Nothing, which we think means the same as setting the parameter 
            ' to keyword "default" which is the same as not specifying it at all
            '-------------------------------------------------------------------------
            strText = "DEFAULT -- (Nothing)"
        Else
            '-------------------------------------------------------------------------
            ' Non-NULL parameter.  Format according to type.
            '-------------------------------------------------------------------------
            Try
                Select Case paramThis.SqlDbType
                    Case SqlDbType.VarChar, SqlDbType.Char, SqlDbType.Text,
                         SqlDbType.NVarChar, SqlDbType.NChar, SqlDbType.NText,
                         SqlDbType.Xml
                        strText = "'" & Replace(paramThis.Value, "'", "''") & "'"

                    Case SqlDbType.SmallDateTime, SqlDbType.DateTime
                        strText = "'" & CDate(paramThis.Value).ToString("dd MMM yyy HH:mm:ss") & "'"

                    Case SqlDbType.Timestamp, SqlDbType.VarBinary, SqlDbType.Binary, SqlDbType.Image
                        strText = FormatBinary(paramThis.Value)

                    Case SqlDbType.Bit
                        strText = IIf(paramThis.Value, "1", "0")

                    Case Else
                        strText = paramThis.Value

                End Select
            Catch ex As Exception
                strText = "[Cannnot convert value to a string for display]"
            End Try
        End If

        Return strText
    End Function

    '-----------------------------------------------------------------------------
    ' FormatBinary
    '-----------------------------------------------------------------------------
    ''' <summary>
    ''' Format binary array as long hex number, to match
    ''' the way SQL management studio displays it.
    ''' </summary>
    ''' <returns>Hex string, e.g. 0x00000000000012ab</returns>
    Public Shared Function FormatBinary(ByVal abytBinary() As Byte) _
                             As String
        '---------------------------------------------------------------------------
        ' Local variables
        '---------------------------------------------------------------------------
        Dim iIndex As Integer
        Dim strText As New Text.StringBuilder("0x")

        For iIndex = 0 To abytBinary.Length - 1
            strText.Append(Convert.ToString(abytBinary(iIndex), 16).PadLeft(2, "0"c))
        Next

        Return strText.ToString

    End Function
#End Region

    '----------------------------------------------------------------------------
    ' New (Constructor)
    '----------------------------------------------------------------------------
    Public Sub New()
        '--------------------------------------------------------------------------
        ' Set tracing on automatically based on whether tracing is enabled
        '  in the currently running top level assembly
        '--------------------------------------------------------------------------
        mCommandTimeOut = 600
        mConnectTimeOut = 600


    End Sub


    '----------------------------------------------------------------------------
    ' New (Constructor)
    '----------------------------------------------------------------------------
    Public Sub New(CommandTimeOut As Int16, ConnectTimeOut As Int16)
        '--------------------------------------------------------------------------
        ' Set tracing on automatically based on whether tracing is enabled
        '  in the currently running top level assembly
        '--------------------------------------------------------------------------
        mCommandTimeOut = CommandTimeOut
        mConnectTimeOut = ConnectTimeOut

    End Sub

#Region " IDisposable Support "

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: free managed resources when explicitly called
                ' Currently nothing to do as member variables only have simple types
            End If
            ' TODO: free shared unmanaged resources
            ' Currently nothing to do no unmanaged resources
        End If
        Me.disposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class


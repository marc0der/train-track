Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration

Namespace Defra.TrainTrack.DataAccess
    Public Class DatabaseHelper
        Implements IDisposable

        Private _connection As SqlConnection

        Public Sub New()
            Dim connectionString As String = ConfigurationManager.ConnectionStrings("TrainTrackDatabase").ConnectionString
            _connection = New SqlConnection(connectionString)
        End Sub

        Public Sub New(connectionStringName As String)
            Dim connectionString As String = ConfigurationManager.ConnectionStrings(connectionStringName).ConnectionString
            _connection = New SqlConnection(connectionString)
        End Sub

        Public ReadOnly Property Connection() As SqlConnection
            Get
                Return _connection
            End Get
        End Property

        Public Sub OpenConnection()
            Try
                If _connection.State = ConnectionState.Closed Then
                    _connection.Open()
                End If
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error opening database connection: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to connect to database", ex)
            End Try
        End Sub

        Public Sub CloseConnection()
            Try
                If _connection.State <> ConnectionState.Closed Then
                    _connection.Close()
                End If
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error closing database connection: {ex.Message}", EventLogEntryType.Warning)
                ' Don't throw exception for close operations
            End Try
        End Sub

        Public Function ExecuteScalar(sql As String, parameters As SqlParameter()) As Object
            Try
                OpenConnection()
                Using cmd As New SqlCommand(sql, _connection)
                    If parameters IsNot Nothing Then
                        cmd.Parameters.AddRange(parameters)
                    End If
                    Return cmd.ExecuteScalar()
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error executing scalar query: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Database query failed", ex)
            End Try
        End Function

        Public Function ExecuteNonQuery(sql As String, parameters As SqlParameter()) As Integer
            Try
                OpenConnection()
                Using cmd As New SqlCommand(sql, _connection)
                    If parameters IsNot Nothing Then
                        cmd.Parameters.AddRange(parameters)
                    End If
                    Return cmd.ExecuteNonQuery()
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error executing non-query: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Database operation failed", ex)
            End Try
        End Function

        Public Function ExecuteReader(sql As String, parameters As SqlParameter()) As SqlDataReader
            Try
                OpenConnection()
                Dim cmd As New SqlCommand(sql, _connection)
                If parameters IsNot Nothing Then
                    cmd.Parameters.AddRange(parameters)
                End If
                Return cmd.ExecuteReader()
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error executing reader query: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Database query failed", ex)
            End Try
        End Function

        Public Function ExecuteDataTable(sql As String, parameters As SqlParameter()) As DataTable
            Try
                OpenConnection()
                Using cmd As New SqlCommand(sql, _connection)
                    If parameters IsNot Nothing Then
                        cmd.Parameters.AddRange(parameters)
                    End If
                    Using adapter As New SqlDataAdapter(cmd)
                        Dim table As New DataTable()
                        adapter.Fill(table)
                        Return table
                    End Using
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error executing data table query: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Database query failed", ex)
            End Try
        End Function

        Public Function ExecuteDataSet(sql As String, parameters As SqlParameter()) As DataSet
            Try
                OpenConnection()
                Using cmd As New SqlCommand(sql, _connection)
                    If parameters IsNot Nothing Then
                        cmd.Parameters.AddRange(parameters)
                    End If
                    Using adapter As New SqlDataAdapter(cmd)
                        Dim dataSet As New DataSet()
                        adapter.Fill(dataSet)
                        Return dataSet
                    End Using
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error executing data set query: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Database query failed", ex)
            End Try
        End Function

        Public Function ExecuteStoredProcedure(procedureName As String, parameters As SqlParameter()) As DataTable
            Try
                OpenConnection()
                Using cmd As New SqlCommand(procedureName, _connection)
                    cmd.CommandType = CommandType.StoredProcedure
                    If parameters IsNot Nothing Then
                        cmd.Parameters.AddRange(parameters)
                    End If
                    Using adapter As New SqlDataAdapter(cmd)
                        Dim table As New DataTable()
                        adapter.Fill(table)
                        Return table
                    End Using
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error executing stored procedure {procedureName}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Stored procedure {procedureName} failed", ex)
            End Try
        End Function

        Public Function BeginTransaction() As SqlTransaction
            Try
                OpenConnection()
                Return _connection.BeginTransaction()
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error beginning transaction: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to begin database transaction", ex)
            End Try
        End Function

        Public Shared Function CreateParameter(name As String, value As Object) As SqlParameter
            If value Is Nothing Then
                Return New SqlParameter(name, DBNull.Value)
            Else
                Return New SqlParameter(name, value)
            End If
        End Function

        Public Shared Function CreateParameter(name As String, sqlDbType As SqlDbType, value As Object) As SqlParameter
            Dim param As New SqlParameter(name, sqlDbType)
            If value Is Nothing Then
                param.Value = DBNull.Value
            Else
                param.Value = value
            End If
            Return param
        End Function

        Public Shared Function CreateParameter(name As String, sqlDbType As SqlDbType, size As Integer, value As Object) As SqlParameter
            Dim param As New SqlParameter(name, sqlDbType, size)
            If value Is Nothing Then
                param.Value = DBNull.Value
            Else
                param.Value = value
            End If
            Return param
        End Function

        Public Shared Function SafeGetString(reader As SqlDataReader, columnName As String) As String
            Try
                Dim ordinal As Integer = reader.GetOrdinal(columnName)
                If reader.IsDBNull(ordinal) Then
                    Return String.Empty
                Else
                    Return reader.GetString(ordinal)
                End If
            Catch ex As Exception
                Return String.Empty
            End Try
        End Function

        Public Shared Function SafeGetInteger(reader As SqlDataReader, columnName As String) As Integer
            Try
                Dim ordinal As Integer = reader.GetOrdinal(columnName)
                If reader.IsDBNull(ordinal) Then
                    Return 0
                Else
                    Return reader.GetInt32(ordinal)
                End If
            Catch ex As Exception
                Return 0
            End Try
        End Function

        Public Shared Function SafeGetNullableInteger(reader As SqlDataReader, columnName As String) As Integer?
            Try
                Dim ordinal As Integer = reader.GetOrdinal(columnName)
                If reader.IsDBNull(ordinal) Then
                    Return Nothing
                Else
                    Return reader.GetInt32(ordinal)
                End If
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        Public Shared Function SafeGetDecimal(reader As SqlDataReader, columnName As String) As Decimal
            Try
                Dim ordinal As Integer = reader.GetOrdinal(columnName)
                If reader.IsDBNull(ordinal) Then
                    Return 0D
                Else
                    Return reader.GetDecimal(ordinal)
                End If
            Catch ex As Exception
                Return 0D
            End Try
        End Function

        Public Shared Function SafeGetBoolean(reader As SqlDataReader, columnName As String) As Boolean
            Try
                Dim ordinal As Integer = reader.GetOrdinal(columnName)
                If reader.IsDBNull(ordinal) Then
                    Return False
                Else
                    Return reader.GetBoolean(ordinal)
                End If
            Catch ex As Exception
                Return False
            End Try
        End Function

        Public Shared Function SafeGetDateTime(reader As SqlDataReader, columnName As String) As DateTime
            Try
                Dim ordinal As Integer = reader.GetOrdinal(columnName)
                If reader.IsDBNull(ordinal) Then
                    Return DateTime.MinValue
                Else
                    Return reader.GetDateTime(ordinal)
                End If
            Catch ex As Exception
                Return DateTime.MinValue
            End Try
        End Function

        Public Shared Function SafeGetNullableDateTime(reader As SqlDataReader, columnName As String) As DateTime?
            Try
                Dim ordinal As Integer = reader.GetOrdinal(columnName)
                If reader.IsDBNull(ordinal) Then
                    Return Nothing
                Else
                    Return reader.GetDateTime(ordinal)
                End If
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        Public Shared Function SafeGetTimeSpan(reader As SqlDataReader, columnName As String) As TimeSpan
            Try
                Dim ordinal As Integer = reader.GetOrdinal(columnName)
                If reader.IsDBNull(ordinal) Then
                    Return TimeSpan.Zero
                Else
                    Return reader.GetTimeSpan(ordinal)
                End If
            Catch ex As Exception
                Return TimeSpan.Zero
            End Try
        End Function

        Public Function TestConnection() As Boolean
            Try
                OpenConnection()
                Using cmd As New SqlCommand("SELECT 1", _connection)
                    Dim result = cmd.ExecuteScalar()
                    Return result IsNot Nothing AndAlso result.ToString() = "1"
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Database connection test failed: {ex.Message}", EventLogEntryType.Warning)
                Return False
            End Try
        End Function

        Public Shared Function EscapeSqlString(input As String) As String
            If String.IsNullOrEmpty(input) Then
                Return String.Empty
            End If
            Return input.Replace("'", "''")
        End Function

        Public Shared Function BuildWhereClause(conditions As List(Of String)) As String
            If conditions Is Nothing OrElse conditions.Count = 0 Then
                Return String.Empty
            End If

            Dim validConditions = conditions.Where(Function(c) Not String.IsNullOrWhiteSpace(c)).ToList()
            If validConditions.Count = 0 Then
                Return String.Empty
            End If

            Return " WHERE " + String.Join(" AND ", validConditions)
        End Function

        Public Shared Function BuildOrderByClause(orderBy As String, orderDirection As String) As String
            If String.IsNullOrWhiteSpace(orderBy) Then
                Return String.Empty
            End If

            Dim direction As String = If(String.IsNullOrWhiteSpace(orderDirection) OrElse orderDirection.ToUpper() <> "DESC", "ASC", "DESC")
            Return $" ORDER BY {orderBy} {direction}"
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            Try
                If _connection IsNot Nothing Then
                    CloseConnection()
                    _connection.Dispose()
                    _connection = Nothing
                End If
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error disposing database helper: {ex.Message}", EventLogEntryType.Warning)
            End Try
        End Sub

    End Class
End Namespace
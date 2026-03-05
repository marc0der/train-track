Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports Defra.TrainTrack.Models

Namespace Defra.TrainTrack.DataAccess
    Public Class WaitingListRepository
        Implements IDisposable

        Private _db As DatabaseHelper

        Public Sub New()
            _db = New DatabaseHelper()
        End Sub

        Public Function GetWaitingListBySession(sessionId As Integer) As List(Of WaitingListEntry)
            Dim entries As New List(Of WaitingListEntry)()
            Try
                Dim sql As String = "SELECT * FROM WaitingList WHERE SessionId = @SessionId ORDER BY Position"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@SessionId", SqlDbType.Int, sessionId)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    While reader.Read()
                        entries.Add(MapEntryFromReader(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetWaitingListBySession ({sessionId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve waiting list for session {sessionId}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return entries
        End Function

        Public Function GetWaitingListByEmployee(employeeId As Integer) As List(Of WaitingListEntry)
            Dim entries As New List(Of WaitingListEntry)()
            Try
                Dim sql As String = "SELECT * FROM WaitingList WHERE EmployeeId = @EmployeeId ORDER BY RequestDate DESC"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@EmployeeId", SqlDbType.Int, employeeId)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    While reader.Read()
                        entries.Add(MapEntryFromReader(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetWaitingListByEmployee ({employeeId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve waiting list entries for employee {employeeId}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return entries
        End Function

        Public Function AddToWaitingList(entry As WaitingListEntry) As Integer
            Try
                Dim sql As String = "INSERT INTO WaitingList " &
                                    "(SessionId, EmployeeId, RequestDate, Position, Status, OfferedDate, " &
                                    "ResponseDeadline, Notes, CreatedBy) " &
                                    "VALUES " &
                                    "(@SessionId, @EmployeeId, @RequestDate, @Position, @Status, @OfferedDate, " &
                                    "@ResponseDeadline, @Notes, @CreatedBy); " &
                                    "SELECT SCOPE_IDENTITY();"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@SessionId", SqlDbType.Int, entry.SessionId),
                    DatabaseHelper.CreateParameter("@EmployeeId", SqlDbType.Int, entry.EmployeeId),
                    DatabaseHelper.CreateParameter("@RequestDate", SqlDbType.DateTime, entry.RequestDate),
                    DatabaseHelper.CreateParameter("@Position", SqlDbType.Int, entry.Position),
                    DatabaseHelper.CreateParameter("@Status", SqlDbType.NVarChar, 50, entry.Status),
                    DatabaseHelper.CreateParameter("@OfferedDate", SqlDbType.DateTime, If(entry.OfferedDate.HasValue, CObj(entry.OfferedDate.Value), DBNull.Value)),
                    DatabaseHelper.CreateParameter("@ResponseDeadline", SqlDbType.DateTime, If(entry.ResponseDeadline.HasValue, CObj(entry.ResponseDeadline.Value), DBNull.Value)),
                    DatabaseHelper.CreateParameter("@Notes", entry.Notes),
                    DatabaseHelper.CreateParameter("@CreatedBy", SqlDbType.NVarChar, 100, entry.CreatedBy)
                }

                Dim result As Object = _db.ExecuteScalar(sql, parameters)
                If result IsNot Nothing AndAlso result IsNot DBNull.Value Then
                    Return Convert.ToInt32(result)
                End If
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in AddToWaitingList: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to add entry to waiting list", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return 0
        End Function

        Public Function UpdateWaitingListEntry(entry As WaitingListEntry) As Boolean
            Try
                Dim sql As String = "UPDATE WaitingList SET " &
                                    "SessionId = @SessionId, " &
                                    "EmployeeId = @EmployeeId, " &
                                    "RequestDate = @RequestDate, " &
                                    "Position = @Position, " &
                                    "Status = @Status, " &
                                    "OfferedDate = @OfferedDate, " &
                                    "ResponseDeadline = @ResponseDeadline, " &
                                    "Notes = @Notes, " &
                                    "CreatedBy = @CreatedBy " &
                                    "WHERE EntryId = @EntryId"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@EntryId", SqlDbType.Int, entry.EntryId),
                    DatabaseHelper.CreateParameter("@SessionId", SqlDbType.Int, entry.SessionId),
                    DatabaseHelper.CreateParameter("@EmployeeId", SqlDbType.Int, entry.EmployeeId),
                    DatabaseHelper.CreateParameter("@RequestDate", SqlDbType.DateTime, entry.RequestDate),
                    DatabaseHelper.CreateParameter("@Position", SqlDbType.Int, entry.Position),
                    DatabaseHelper.CreateParameter("@Status", SqlDbType.NVarChar, 50, entry.Status),
                    DatabaseHelper.CreateParameter("@OfferedDate", SqlDbType.DateTime, If(entry.OfferedDate.HasValue, CObj(entry.OfferedDate.Value), DBNull.Value)),
                    DatabaseHelper.CreateParameter("@ResponseDeadline", SqlDbType.DateTime, If(entry.ResponseDeadline.HasValue, CObj(entry.ResponseDeadline.Value), DBNull.Value)),
                    DatabaseHelper.CreateParameter("@Notes", entry.Notes),
                    DatabaseHelper.CreateParameter("@CreatedBy", SqlDbType.NVarChar, 100, entry.CreatedBy)
                }

                Dim rowsAffected As Integer = _db.ExecuteNonQuery(sql, parameters)
                Return rowsAffected > 0
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in UpdateWaitingListEntry ({entry.EntryId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to update waiting list entry {entry.EntryId}", ex)
            Finally
                _db.CloseConnection()
            End Try
        End Function

        Public Function RemoveFromWaitingList(entryId As Integer) As Boolean
            Try
                Dim sql As String = "DELETE FROM WaitingList WHERE EntryId = @EntryId"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@EntryId", SqlDbType.Int, entryId)
                }

                Dim rowsAffected As Integer = _db.ExecuteNonQuery(sql, parameters)
                Return rowsAffected > 0
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in RemoveFromWaitingList ({entryId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to remove waiting list entry {entryId}", ex)
            Finally
                _db.CloseConnection()
            End Try
        End Function

        Public Function GetNextInLine(sessionId As Integer) As WaitingListEntry
            Try
                Dim sql As String = "SELECT TOP 1 * FROM WaitingList " &
                                    "WHERE SessionId = @SessionId AND Status = 'Waiting' " &
                                    "ORDER BY Position"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@SessionId", SqlDbType.Int, sessionId)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    If reader.Read() Then
                        Return MapEntryFromReader(reader)
                    End If
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetNextInLine ({sessionId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to get next in line for session {sessionId}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return Nothing
        End Function

        Public Function GetAllActiveWaitingListEntries() As List(Of WaitingListEntry)
            Dim entries As New List(Of WaitingListEntry)()
            Try
                Dim sql As String = "SELECT * FROM WaitingList " &
                                    "WHERE Status IN ('Waiting', 'Offered') " &
                                    "ORDER BY SessionId, Position"

                Using reader As SqlDataReader = _db.ExecuteReader(sql, Nothing)
                    While reader.Read()
                        entries.Add(MapEntryFromReader(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetAllActiveWaitingListEntries: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve active waiting list entries", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return entries
        End Function

        Public Function ReorderWaitingList(sessionId As Integer, entryIds As List(Of Integer)) As Boolean
            Try
                Dim transaction As SqlTransaction = _db.BeginTransaction()
                Try
                    For i As Integer = 0 To entryIds.Count - 1
                        Dim sql As String = "UPDATE WaitingList SET Position = @Position " &
                                            "WHERE EntryId = @EntryId AND SessionId = @SessionId"

                        Using cmd As New SqlCommand(sql, _db.Connection, transaction)
                            cmd.Parameters.Add(DatabaseHelper.CreateParameter("@Position", SqlDbType.Int, i + 1))
                            cmd.Parameters.Add(DatabaseHelper.CreateParameter("@EntryId", SqlDbType.Int, entryIds(i)))
                            cmd.Parameters.Add(DatabaseHelper.CreateParameter("@SessionId", SqlDbType.Int, sessionId))
                            cmd.ExecuteNonQuery()
                        End Using
                    Next

                    transaction.Commit()
                    Return True
                Catch ex As Exception
                    transaction.Rollback()
                    Throw
                End Try
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in ReorderWaitingList ({sessionId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to reorder waiting list for session {sessionId}", ex)
            Finally
                _db.CloseConnection()
            End Try
        End Function

        Private Function MapEntryFromReader(reader As SqlDataReader) As WaitingListEntry
            Dim entry As New WaitingListEntry()
            entry.EntryId = DatabaseHelper.SafeGetInteger(reader, "EntryId")
            entry.SessionId = DatabaseHelper.SafeGetInteger(reader, "SessionId")
            entry.EmployeeId = DatabaseHelper.SafeGetInteger(reader, "EmployeeId")
            entry.RequestDate = DatabaseHelper.SafeGetDateTime(reader, "RequestDate")
            entry.Position = DatabaseHelper.SafeGetInteger(reader, "Position")
            entry.Status = DatabaseHelper.SafeGetString(reader, "Status")
            entry.OfferedDate = DatabaseHelper.SafeGetNullableDateTime(reader, "OfferedDate")
            entry.ResponseDeadline = DatabaseHelper.SafeGetNullableDateTime(reader, "ResponseDeadline")
            entry.Notes = DatabaseHelper.SafeGetString(reader, "Notes")
            entry.CreatedBy = DatabaseHelper.SafeGetString(reader, "CreatedBy")
            Return entry
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            Try
                If _db IsNot Nothing Then
                    _db.Dispose()
                    _db = Nothing
                End If
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error disposing WaitingListRepository: {ex.Message}", EventLogEntryType.Warning)
            End Try
        End Sub

    End Class
End Namespace

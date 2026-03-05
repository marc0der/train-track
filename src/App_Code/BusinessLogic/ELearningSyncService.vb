Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports Defra.TrainTrack.Models
Imports Defra.TrainTrack.DataAccess

Namespace Defra.TrainTrack.BusinessLogic

    ''' <summary>
    ''' Handles the overnight batch sync process from the third-party e-learning platform.
    ''' Parses batch completion files, matches employees and courses, and creates/updates
    ''' training records. Logs sync results to the ELearningSyncLog table.
    ''' </summary>
    Public Class ELearningSyncService
        Implements IDisposable

        Private _db As DatabaseHelper
        Private _employeeRepository As EmployeeRepository
        Private _courseRepository As CourseRepository
        Private _trainingRepository As TrainingRepository

        ' Expected CSV column names in the e-learning batch file
        Private Shared ReadOnly RequiredColumns As String() = {
            "EmployeeNumber", "CourseCode", "CompletionDate", "Score", "Status"
        }

        Public Sub New()
            _db = New DatabaseHelper()
            _employeeRepository = New EmployeeRepository()
            _courseRepository = New CourseRepository()
            _trainingRepository = New TrainingRepository()
        End Sub

        ''' <summary>
        ''' Processes a batch file from the e-learning platform.
        ''' For each completion record: finds employee by EmployeeNumber, finds course by CourseCode,
        ''' creates or updates a TrainingRecord, and logs successes and failures.
        ''' Returns a sync record with statistics.
        ''' </summary>
        Public Function ProcessBatchFile(filePath As String, processedBy As String) As ELearningSyncRecord
            Dim syncRecord As New ELearningSyncRecord(filePath)
            syncRecord.SyncDate = DateTime.Now
            syncRecord.StartTime = DateTime.Now

            Dim unmatchedRecords As New List(Of String)()

            Try
                If String.IsNullOrWhiteSpace(filePath) Then
                    syncRecord.SyncStatus = "Failed"
                    syncRecord.ErrorDetails = "File path cannot be empty"
                    syncRecord.EndTime = DateTime.Now
                    SaveSyncRecord(syncRecord)
                    Return syncRecord
                End If

                If Not File.Exists(filePath) Then
                    syncRecord.SyncStatus = "Failed"
                    syncRecord.ErrorDetails = $"File not found: {filePath}"
                    syncRecord.EndTime = DateTime.Now
                    SaveSyncRecord(syncRecord)
                    Return syncRecord
                End If

                ' Parse the batch file
                Dim batchData As DataTable = ParseBatchFile(filePath)

                If batchData Is Nothing OrElse batchData.Rows.Count = 0 Then
                    syncRecord.SyncStatus = "Failed"
                    syncRecord.ErrorDetails = "Batch file is empty or could not be parsed"
                    syncRecord.EndTime = DateTime.Now
                    SaveSyncRecord(syncRecord)
                    Return syncRecord
                End If

                ' Validate required columns
                Dim missingColumns As New List(Of String)()
                For Each requiredCol As String In RequiredColumns
                    Dim found As Boolean = False
                    For Each col As DataColumn In batchData.Columns
                        If String.Equals(col.ColumnName.Trim(), requiredCol, StringComparison.OrdinalIgnoreCase) Then
                            found = True
                            Exit For
                        End If
                    Next
                    If Not found Then
                        missingColumns.Add(requiredCol)
                    End If
                Next

                If missingColumns.Count > 0 Then
                    syncRecord.SyncStatus = "Failed"
                    syncRecord.ErrorDetails = $"Missing required columns: {String.Join(", ", missingColumns)}"
                    syncRecord.EndTime = DateTime.Now
                    SaveSyncRecord(syncRecord)
                    Return syncRecord
                End If

                syncRecord.TotalRecords = batchData.Rows.Count
                Dim errorMessages As New List(Of String)()

                EventLog.WriteEntry("TrainTrack", $"E-learning batch sync started by {processedBy}: {filePath} ({syncRecord.TotalRecords} records)", EventLogEntryType.Information)

                For rowIndex As Integer = 0 To batchData.Rows.Count - 1
                    Dim row As DataRow = batchData.Rows(rowIndex)
                    Dim rowNumber As Integer = rowIndex + 1

                    Try
                        Dim employeeNumber As String = GetColumnValue(batchData, row, "EmployeeNumber")
                        Dim courseCode As String = GetColumnValue(batchData, row, "CourseCode")
                        Dim completionDateStr As String = GetColumnValue(batchData, row, "CompletionDate")
                        Dim scoreStr As String = GetColumnValue(batchData, row, "Score")
                        Dim status As String = GetColumnValue(batchData, row, "Status")

                        ' Validate required fields
                        If String.IsNullOrWhiteSpace(employeeNumber) OrElse String.IsNullOrWhiteSpace(courseCode) Then
                            errorMessages.Add($"Row {rowNumber}: EmployeeNumber and CourseCode are required")
                            unmatchedRecords.Add($"Row {rowNumber}: Missing EmployeeNumber or CourseCode")
                            syncRecord.FailedRecords += 1
                            Continue For
                        End If

                        ' Find employee by EmployeeNumber
                        Dim employee As Employee = _employeeRepository.GetEmployeeByNumber(employeeNumber)
                        If employee Is Nothing Then
                            errorMessages.Add($"Row {rowNumber}: Employee not found: {employeeNumber}")
                            unmatchedRecords.Add($"Row {rowNumber}: Employee not found: {employeeNumber}")
                            syncRecord.FailedRecords += 1
                            Continue For
                        End If

                        ' Find course by CourseCode
                        Dim course As Course = _courseRepository.GetCourseByCode(courseCode)
                        If course Is Nothing Then
                            errorMessages.Add($"Row {rowNumber}: Course not found: {courseCode}")
                            unmatchedRecords.Add($"Row {rowNumber}: Course not found: {courseCode}")
                            syncRecord.FailedRecords += 1
                            Continue For
                        End If

                        ' Parse completion date
                        Dim completionDate As DateTime = DateTime.Now
                        If Not String.IsNullOrWhiteSpace(completionDateStr) Then
                            If Not DateTime.TryParse(completionDateStr, completionDate) Then
                                errorMessages.Add($"Row {rowNumber}: Invalid completion date: {completionDateStr}")
                                syncRecord.FailedRecords += 1
                                Continue For
                            End If
                        End If

                        ' Parse score
                        Dim score As Integer? = Nothing
                        If Not String.IsNullOrWhiteSpace(scoreStr) Then
                            Dim parsedScore As Integer
                            If Integer.TryParse(scoreStr, parsedScore) Then
                                score = parsedScore
                            End If
                        End If

                        ' Determine completion status
                        Dim completionStatus As String = "COMPLETED"
                        If Not String.IsNullOrWhiteSpace(status) Then
                            Select Case status.ToUpper().Trim()
                                Case "PASSED", "COMPLETED", "COMPLETE"
                                    completionStatus = "COMPLETED"
                                Case "FAILED", "FAIL"
                                    completionStatus = "FAILED"
                                Case "IN_PROGRESS", "IN PROGRESS", "STARTED"
                                    completionStatus = "IN_PROGRESS"
                                Case Else
                                    completionStatus = "COMPLETED"
                            End Select
                        End If

                        ' Check if a training record already exists for this employee/course
                        Dim existingRecords = _trainingRepository.GetTrainingRecordsByEmployee(employee.EmployeeId)
                        Dim existingRecord As TrainingRecord = Nothing
                        For Each record In existingRecords
                            If record.CourseId = course.CourseId AndAlso
                               record.CompletionStatus <> "CANCELLED" Then
                                existingRecord = record
                                Exit For
                            End If
                        Next

                        If existingRecord IsNot Nothing Then
                            ' Update existing training record
                            existingRecord.CompletionStatus = completionStatus
                            existingRecord.CompletionDate = completionDate
                            existingRecord.Score = score
                            existingRecord.ModifiedBy = processedBy
                            existingRecord.ModifiedDate = DateTime.Now

                            If completionStatus = "COMPLETED" AndAlso course.ValidityPeriodMonths > 0 Then
                                existingRecord.ExpiryDate = completionDate.AddMonths(course.ValidityPeriodMonths)
                                existingRecord.RenewalRequired = True
                            End If

                            Dim success As Boolean = _trainingRepository.UpdateTrainingRecord(existingRecord)
                            If success Then
                                syncRecord.ProcessedRecords += 1
                                EventLog.WriteEntry("TrainTrack", $"E-learning sync: Updated training record for employee {employeeNumber}, course {courseCode}", EventLogEntryType.Information)
                            Else
                                errorMessages.Add($"Row {rowNumber}: Failed to update training record for employee {employeeNumber}, course {courseCode}")
                                syncRecord.FailedRecords += 1
                            End If
                        Else
                            ' Create new training record
                            Dim newRecord As New TrainingRecord()
                            newRecord.EmployeeId = employee.EmployeeId
                            newRecord.CourseId = course.CourseId
                            newRecord.EnrollmentDate = completionDate
                            newRecord.CompletionDate = completionDate
                            newRecord.CompletionStatus = completionStatus
                            newRecord.AttendanceStatus = "COMPLETED"
                            newRecord.Score = score
                            newRecord.CreatedBy = processedBy
                            newRecord.CreatedDate = DateTime.Now

                            If completionStatus = "COMPLETED" AndAlso course.ValidityPeriodMonths > 0 Then
                                newRecord.ExpiryDate = completionDate.AddMonths(course.ValidityPeriodMonths)
                                newRecord.RenewalRequired = True
                            End If

                            Dim newRecordId As Integer = _trainingRepository.CreateTrainingRecord(newRecord)
                            If newRecordId > 0 Then
                                syncRecord.ProcessedRecords += 1
                                EventLog.WriteEntry("TrainTrack", $"E-learning sync: Created training record (ID: {newRecordId}) for employee {employeeNumber}, course {courseCode}", EventLogEntryType.Information)
                            Else
                                errorMessages.Add($"Row {rowNumber}: Failed to create training record for employee {employeeNumber}, course {courseCode}")
                                syncRecord.FailedRecords += 1
                            End If
                        End If

                    Catch rowEx As Exception
                        errorMessages.Add($"Row {rowNumber}: {rowEx.Message}")
                        syncRecord.FailedRecords += 1
                        EventLog.WriteEntry("TrainTrack", $"E-learning sync error at row {rowNumber}: {rowEx.Message}", EventLogEntryType.Warning)
                    End Try
                Next

                ' Determine overall sync status
                If syncRecord.FailedRecords = 0 Then
                    syncRecord.SyncStatus = "Success"
                ElseIf syncRecord.ProcessedRecords > 0 Then
                    syncRecord.SyncStatus = "Partial"
                Else
                    syncRecord.SyncStatus = "Failed"
                End If

                ' Store error details and unmatched records
                If errorMessages.Count > 0 Then
                    syncRecord.ErrorDetails = String.Join(Environment.NewLine, errorMessages)
                End If

                syncRecord.EndTime = DateTime.Now

                ' Save sync record to database
                SaveSyncRecord(syncRecord)

                ' Store unmatched records against the sync ID
                If unmatchedRecords.Count > 0 Then
                    SaveUnmatchedRecords(syncRecord.SyncId, unmatchedRecords)
                End If

                EventLog.WriteEntry("TrainTrack",
                    $"E-learning batch sync completed by {processedBy}: {syncRecord.ProcessedRecords}/{syncRecord.TotalRecords} processed, {syncRecord.FailedRecords} failed ({syncRecord.SyncStatus})",
                    EventLogEntryType.Information)

            Catch ex As Exception
                syncRecord.SyncStatus = "Failed"
                syncRecord.ErrorDetails = $"Sync failed: {ex.Message}"
                syncRecord.EndTime = DateTime.Now
                SaveSyncRecord(syncRecord)

                EventLog.WriteEntry("TrainTrack", $"E-learning batch sync failed: {ex.Message}", EventLogEntryType.Error)
            End Try

            Return syncRecord
        End Function

        ''' <summary>
        ''' Retrieves sync history between the specified dates.
        ''' </summary>
        Public Function GetSyncHistory(fromDate As DateTime, toDate As DateTime) As List(Of ELearningSyncRecord)
            Dim records As New List(Of ELearningSyncRecord)()
            Try
                Dim sql As String = "SELECT * FROM ELearningSyncLog WHERE SyncDate >= @FromDate AND SyncDate <= @ToDate ORDER BY SyncDate DESC"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@FromDate", SqlDbType.DateTime, fromDate),
                    DatabaseHelper.CreateParameter("@ToDate", SqlDbType.DateTime, toDate)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    While reader.Read()
                        records.Add(MapSyncRecord(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting sync history: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve sync history", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return records
        End Function

        ''' <summary>
        ''' Gets the most recent sync record. Quick check: did last night's sync succeed?
        ''' </summary>
        Public Function GetLastSyncStatus() As ELearningSyncRecord
            Try
                Dim sql As String = "SELECT TOP 1 * FROM ELearningSyncLog ORDER BY SyncDate DESC"

                Using reader As SqlDataReader = _db.ExecuteReader(sql, Nothing)
                    If reader.Read() Then
                        Return MapSyncRecord(reader)
                    End If
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting last sync status: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve last sync status", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return Nothing
        End Function

        ''' <summary>
        ''' Gets records from a batch file that couldn't be matched to employees or courses.
        ''' Addresses the reliability issues described in the transcript.
        ''' </summary>
        Public Function GetUnmatchedRecords(syncId As Integer) As List(Of String)
            Dim unmatchedList As New List(Of String)()
            Try
                If syncId <= 0 Then
                    Throw New ArgumentException("Sync ID must be greater than 0", "syncId")
                End If

                ' Retrieve the sync record and parse unmatched details from ErrorDetails
                Dim sql As String = "SELECT ErrorDetails FROM ELearningSyncLog WHERE SyncId = @SyncId"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@SyncId", SqlDbType.Int, syncId)
                }

                Dim result As Object = _db.ExecuteScalar(sql, parameters)

                If result IsNot Nothing AndAlso result IsNot DBNull.Value Then
                    Dim errorDetails As String = result.ToString()
                    If Not String.IsNullOrWhiteSpace(errorDetails) Then
                        ' Parse error lines that indicate unmatched records
                        Dim lines As String() = errorDetails.Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
                        For Each line As String In lines
                            If line.Contains("Employee not found") OrElse
                               line.Contains("Course not found") OrElse
                               line.Contains("Missing EmployeeNumber or CourseCode") Then
                                unmatchedList.Add(line)
                            End If
                        Next
                    End If
                End If
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting unmatched records for sync {syncId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve unmatched records for sync {syncId}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return unmatchedList
        End Function

        ' ===== Private Helper Methods =====

        ''' <summary>
        ''' Saves a sync record to the ELearningSyncLog table.
        ''' </summary>
        Private Sub SaveSyncRecord(syncRecord As ELearningSyncRecord)
            Try
                Dim sql As String = "INSERT INTO ELearningSyncLog (SyncDate, SyncStatus, TotalRecords, ProcessedRecords, FailedRecords, ErrorDetails, BatchFilePath, StartTime, EndTime) " &
                                    "VALUES (@SyncDate, @SyncStatus, @TotalRecords, @ProcessedRecords, @FailedRecords, @ErrorDetails, @BatchFilePath, @StartTime, @EndTime); " &
                                    "SELECT SCOPE_IDENTITY();"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@SyncDate", SqlDbType.DateTime, syncRecord.SyncDate),
                    DatabaseHelper.CreateParameter("@SyncStatus", SqlDbType.NVarChar, syncRecord.SyncStatus),
                    DatabaseHelper.CreateParameter("@TotalRecords", SqlDbType.Int, syncRecord.TotalRecords),
                    DatabaseHelper.CreateParameter("@ProcessedRecords", SqlDbType.Int, syncRecord.ProcessedRecords),
                    DatabaseHelper.CreateParameter("@FailedRecords", SqlDbType.Int, syncRecord.FailedRecords),
                    DatabaseHelper.CreateParameter("@ErrorDetails", SqlDbType.NVarChar, If(syncRecord.ErrorDetails, CObj(DBNull.Value))),
                    DatabaseHelper.CreateParameter("@BatchFilePath", SqlDbType.NVarChar, If(syncRecord.BatchFilePath, CObj(DBNull.Value))),
                    DatabaseHelper.CreateParameter("@StartTime", SqlDbType.DateTime, syncRecord.StartTime),
                    DatabaseHelper.CreateParameter("@EndTime", SqlDbType.DateTime, If(syncRecord.EndTime.HasValue, CObj(syncRecord.EndTime.Value), CObj(DBNull.Value)))
                }

                Dim result As Object = _db.ExecuteScalar(sql, parameters)
                If result IsNot Nothing AndAlso result IsNot DBNull.Value Then
                    syncRecord.SyncId = Convert.ToInt32(result)
                End If
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error saving sync record: {ex.Message}", EventLogEntryType.Error)
            Finally
                _db.CloseConnection()
            End Try
        End Sub

        ''' <summary>
        ''' Stores unmatched record details against the sync record in the ErrorDetails field.
        ''' Since unmatched records are already captured in ErrorDetails during processing,
        ''' this method is a no-op but reserved for future use if a separate table is added.
        ''' </summary>
        Private Sub SaveUnmatchedRecords(syncId As Integer, unmatchedRecords As List(Of String))
            ' Unmatched records are already stored as part of ErrorDetails in SaveSyncRecord.
            ' This method is reserved for future expansion if a dedicated unmatched records table is added.
            EventLog.WriteEntry("TrainTrack", $"Sync {syncId}: {unmatchedRecords.Count} unmatched records logged", EventLogEntryType.Information)
        End Sub

        ''' <summary>
        ''' Maps a SqlDataReader row to an ELearningSyncRecord model.
        ''' </summary>
        Private Function MapSyncRecord(reader As SqlDataReader) As ELearningSyncRecord
            Dim record As New ELearningSyncRecord()
            record.SyncId = DatabaseHelper.SafeGetInteger(reader, "SyncId")
            record.SyncDate = DatabaseHelper.SafeGetDateTime(reader, "SyncDate")
            record.SyncStatus = DatabaseHelper.SafeGetString(reader, "SyncStatus")
            record.TotalRecords = DatabaseHelper.SafeGetInteger(reader, "TotalRecords")
            record.ProcessedRecords = DatabaseHelper.SafeGetInteger(reader, "ProcessedRecords")
            record.FailedRecords = DatabaseHelper.SafeGetInteger(reader, "FailedRecords")
            record.ErrorDetails = DatabaseHelper.SafeGetString(reader, "ErrorDetails")
            record.BatchFilePath = DatabaseHelper.SafeGetString(reader, "BatchFilePath")
            record.StartTime = DatabaseHelper.SafeGetDateTime(reader, "StartTime")
            record.EndTime = DatabaseHelper.SafeGetNullableDateTime(reader, "EndTime")
            Return record
        End Function

        ''' <summary>
        ''' Parses a batch CSV file from the e-learning platform.
        ''' Expected columns: EmployeeNumber, CourseCode, CompletionDate, Score, Status
        ''' </summary>
        Private Function ParseBatchFile(filePath As String) As DataTable
            Dim table As New DataTable()

            Using reader As New StreamReader(filePath)
                ' Read the header line
                Dim headerLine As String = reader.ReadLine()
                If String.IsNullOrWhiteSpace(headerLine) Then
                    Return table
                End If

                ' Parse header columns
                Dim headers As String() = ParseCsvLine(headerLine)
                For Each header As String In headers
                    Dim cleanHeader As String = header.Trim().Trim(""""c)
                    If Not String.IsNullOrWhiteSpace(cleanHeader) Then
                        table.Columns.Add(cleanHeader)
                    End If
                Next

                ' Read data rows
                Dim line As String = reader.ReadLine()
                While line IsNot Nothing
                    If Not String.IsNullOrWhiteSpace(line) Then
                        Dim values As String() = ParseCsvLine(line)
                        Dim dataRow As DataRow = table.NewRow()

                        For i As Integer = 0 To Math.Min(values.Length, table.Columns.Count) - 1
                            dataRow(i) = values(i).Trim().Trim(""""c)
                        Next

                        table.Rows.Add(dataRow)
                    End If

                    line = reader.ReadLine()
                End While
            End Using

            Return table
        End Function

        ''' <summary>
        ''' Parses a single CSV line, handling quoted fields that may contain commas.
        ''' </summary>
        Private Function ParseCsvLine(line As String) As String()
            Dim fields As New List(Of String)()
            Dim currentField As New System.Text.StringBuilder()
            Dim inQuotes As Boolean = False

            For i As Integer = 0 To line.Length - 1
                Dim c As Char = line(i)

                If c = """"c Then
                    If inQuotes AndAlso i + 1 < line.Length AndAlso line(i + 1) = """"c Then
                        ' Escaped quote (double quote inside quotes)
                        currentField.Append(""""c)
                        i += 1
                    Else
                        ' Toggle quote mode
                        inQuotes = Not inQuotes
                    End If
                ElseIf c = ","c AndAlso Not inQuotes Then
                    ' Field separator
                    fields.Add(currentField.ToString())
                    currentField.Clear()
                Else
                    currentField.Append(c)
                End If
            Next

            ' Add the last field
            fields.Add(currentField.ToString())

            Return fields.ToArray()
        End Function

        ''' <summary>
        ''' Gets a column value from a DataRow, performing case-insensitive column name matching.
        ''' Returns empty string if the column does not exist.
        ''' </summary>
        Private Function GetColumnValue(batchData As DataTable, row As DataRow, columnName As String) As String
            For Each col As DataColumn In batchData.Columns
                If String.Equals(col.ColumnName.Trim(), columnName, StringComparison.OrdinalIgnoreCase) Then
                    Dim value As Object = row(col)
                    If value IsNot Nothing AndAlso value IsNot DBNull.Value Then
                        Return value.ToString().Trim()
                    End If
                    Return String.Empty
                End If
            Next
            Return String.Empty
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            If _db IsNot Nothing Then
                _db.Dispose()
                _db = Nothing
            End If
            If _employeeRepository IsNot Nothing Then
                _employeeRepository.Dispose()
                _employeeRepository = Nothing
            End If
            If _courseRepository IsNot Nothing Then
                _courseRepository.Dispose()
                _courseRepository = Nothing
            End If
            If _trainingRepository IsNot Nothing Then
                _trainingRepository.Dispose()
                _trainingRepository = Nothing
            End If
        End Sub

    End Class
End Namespace

Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports Defra.TrainTrack.Models

Namespace Defra.TrainTrack.DataAccess
    Public Class TrainingRepository
        Implements IDisposable

        Private _db As DatabaseHelper

        Public Sub New()
            _db = New DatabaseHelper()
        End Sub

        ' ===========================
        ' TrainingRecord Methods
        ' ===========================

        Public Function GetAllTrainingRecords() As List(Of TrainingRecord)
            Dim records As New List(Of TrainingRecord)()
            Try
                Dim sql As String = "SELECT * FROM TrainingRecords ORDER BY EnrollmentDate DESC"

                Using reader As SqlDataReader = _db.ExecuteReader(sql, Nothing)
                    While reader.Read()
                        records.Add(MapTrainingRecord(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetAllTrainingRecords: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve training records", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return records
        End Function

        Public Function GetTrainingRecordById(recordId As Integer) As TrainingRecord
            Try
                Dim sql As String = "SELECT * FROM TrainingRecords WHERE RecordId = @RecordId"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@RecordId", SqlDbType.Int, recordId)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    If reader.Read() Then
                        Return MapTrainingRecord(reader)
                    End If
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetTrainingRecordById ({recordId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve training record with ID {recordId}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return Nothing
        End Function

        Public Function GetTrainingRecordsByEmployee(employeeId As Integer) As List(Of TrainingRecord)
            Dim records As New List(Of TrainingRecord)()
            Try
                Dim sql As String = "SELECT * FROM TrainingRecords WHERE EmployeeId = @EmployeeId ORDER BY EnrollmentDate DESC"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@EmployeeId", SqlDbType.Int, employeeId)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    While reader.Read()
                        records.Add(MapTrainingRecord(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetTrainingRecordsByEmployee ({employeeId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve training records for employee {employeeId}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return records
        End Function

        Public Function GetTrainingRecordsByCourse(courseId As Integer) As List(Of TrainingRecord)
            Dim records As New List(Of TrainingRecord)()
            Try
                Dim sql As String = "SELECT * FROM TrainingRecords WHERE CourseId = @CourseId ORDER BY EnrollmentDate DESC"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@CourseId", SqlDbType.Int, courseId)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    While reader.Read()
                        records.Add(MapTrainingRecord(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetTrainingRecordsByCourse ({courseId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve training records for course {courseId}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return records
        End Function

        Public Function GetTrainingRecordsBySession(sessionId As Integer) As List(Of TrainingRecord)
            Dim records As New List(Of TrainingRecord)()
            Try
                Dim sql As String = "SELECT * FROM TrainingRecords WHERE SessionId = @SessionId ORDER BY EnrollmentDate DESC"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@SessionId", SqlDbType.Int, sessionId)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    While reader.Read()
                        records.Add(MapTrainingRecord(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetTrainingRecordsBySession ({sessionId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve training records for session {sessionId}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return records
        End Function

        Public Function HasEmployeeCompletedCourse(employeeId As Integer, courseId As Integer) As Boolean
            Try
                Dim sql As String = "SELECT COUNT(*) FROM TrainingRecords " &
                                    "WHERE EmployeeId = @EmployeeId AND CourseId = @CourseId " &
                                    "AND CompletionStatus = 'COMPLETED' " &
                                    "AND (ExpiryDate IS NULL OR ExpiryDate >= GETDATE())"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@EmployeeId", SqlDbType.Int, employeeId),
                    DatabaseHelper.CreateParameter("@CourseId", SqlDbType.Int, courseId)
                }

                Dim result As Object = _db.ExecuteScalar(sql, parameters)
                If result IsNot Nothing AndAlso result IsNot DBNull.Value Then
                    Return Convert.ToInt32(result) > 0
                End If
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in HasEmployeeCompletedCourse (employee {employeeId}, course {courseId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to check course completion for employee {employeeId}, course {courseId}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return False
        End Function

        Public Function CreateTrainingRecord(record As TrainingRecord) As Integer
            Try
                Dim sql As String = "INSERT INTO TrainingRecords " &
                                    "(EmployeeId, SessionId, CourseId, EnrollmentDate, CompletionDate, " &
                                    "AttendanceStatus, CompletionStatus, Score, MaxScore, PassMark, Grade, " &
                                    "CertificateIssued, CertificateNumber, CertificateIssuedDate, ExpiryDate, " &
                                    "IsExpired, Feedback, InstructorComments, CreatedDate, CreatedBy, " &
                                    "RenewalRequired, RenewalNotificationSent, CostCentre, ApprovalRequired, " &
                                    "ApprovedBy, ApprovedDate, RejectedReason) " &
                                    "VALUES " &
                                    "(@EmployeeId, @SessionId, @CourseId, @EnrollmentDate, @CompletionDate, " &
                                    "@AttendanceStatus, @CompletionStatus, @Score, @MaxScore, @PassMark, @Grade, " &
                                    "@CertificateIssued, @CertificateNumber, @CertificateIssuedDate, @ExpiryDate, " &
                                    "@IsExpired, @Feedback, @InstructorComments, @CreatedDate, @CreatedBy, " &
                                    "@RenewalRequired, @RenewalNotificationSent, @CostCentre, @ApprovalRequired, " &
                                    "@ApprovedBy, @ApprovedDate, @RejectedReason); " &
                                    "SELECT SCOPE_IDENTITY();"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@EmployeeId", SqlDbType.Int, record.EmployeeId),
                    DatabaseHelper.CreateParameter("@SessionId", SqlDbType.Int, record.SessionId),
                    DatabaseHelper.CreateParameter("@CourseId", SqlDbType.Int, record.CourseId),
                    DatabaseHelper.CreateParameter("@EnrollmentDate", SqlDbType.DateTime, record.EnrollmentDate),
                    DatabaseHelper.CreateParameter("@CompletionDate", SqlDbType.DateTime, record.CompletionDate),
                    DatabaseHelper.CreateParameter("@AttendanceStatus", SqlDbType.NVarChar, 20, record.AttendanceStatus),
                    DatabaseHelper.CreateParameter("@CompletionStatus", SqlDbType.NVarChar, 20, record.CompletionStatus),
                    DatabaseHelper.CreateParameter("@Score", SqlDbType.Int, record.Score),
                    DatabaseHelper.CreateParameter("@MaxScore", SqlDbType.Int, record.MaxScore),
                    DatabaseHelper.CreateParameter("@PassMark", SqlDbType.Int, record.PassMark),
                    DatabaseHelper.CreateParameter("@Grade", SqlDbType.NVarChar, 10, record.Grade),
                    DatabaseHelper.CreateParameter("@CertificateIssued", SqlDbType.Bit, record.CertificateIssued),
                    DatabaseHelper.CreateParameter("@CertificateNumber", SqlDbType.NVarChar, 50, record.CertificateNumber),
                    DatabaseHelper.CreateParameter("@CertificateIssuedDate", SqlDbType.DateTime, record.CertificateIssuedDate),
                    DatabaseHelper.CreateParameter("@ExpiryDate", SqlDbType.DateTime, record.ExpiryDate),
                    DatabaseHelper.CreateParameter("@IsExpired", SqlDbType.Bit, record.IsExpired),
                    DatabaseHelper.CreateParameter("@Feedback", record.Feedback),
                    DatabaseHelper.CreateParameter("@InstructorComments", record.InstructorComments),
                    DatabaseHelper.CreateParameter("@CreatedDate", SqlDbType.DateTime, record.CreatedDate),
                    DatabaseHelper.CreateParameter("@CreatedBy", SqlDbType.NVarChar, 50, record.CreatedBy),
                    DatabaseHelper.CreateParameter("@RenewalRequired", SqlDbType.Bit, record.RenewalRequired),
                    DatabaseHelper.CreateParameter("@RenewalNotificationSent", SqlDbType.Bit, record.RenewalNotificationSent),
                    DatabaseHelper.CreateParameter("@CostCentre", SqlDbType.NVarChar, 20, record.CostCentre),
                    DatabaseHelper.CreateParameter("@ApprovalRequired", SqlDbType.Bit, record.ApprovalRequired),
                    DatabaseHelper.CreateParameter("@ApprovedBy", SqlDbType.NVarChar, 50, record.ApprovedBy),
                    DatabaseHelper.CreateParameter("@ApprovedDate", SqlDbType.DateTime, record.ApprovedDate),
                    DatabaseHelper.CreateParameter("@RejectedReason", record.RejectedReason)
                }

                Dim result As Object = _db.ExecuteScalar(sql, parameters)
                If result IsNot Nothing AndAlso result IsNot DBNull.Value Then
                    Return Convert.ToInt32(result)
                End If
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in CreateTrainingRecord: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to create training record", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return 0
        End Function

        Public Function UpdateTrainingRecord(record As TrainingRecord) As Boolean
            Try
                Dim sql As String = "UPDATE TrainingRecords SET " &
                                    "EmployeeId = @EmployeeId, " &
                                    "SessionId = @SessionId, " &
                                    "CourseId = @CourseId, " &
                                    "EnrollmentDate = @EnrollmentDate, " &
                                    "CompletionDate = @CompletionDate, " &
                                    "AttendanceStatus = @AttendanceStatus, " &
                                    "CompletionStatus = @CompletionStatus, " &
                                    "Score = @Score, " &
                                    "MaxScore = @MaxScore, " &
                                    "PassMark = @PassMark, " &
                                    "Grade = @Grade, " &
                                    "CertificateIssued = @CertificateIssued, " &
                                    "CertificateNumber = @CertificateNumber, " &
                                    "CertificateIssuedDate = @CertificateIssuedDate, " &
                                    "ExpiryDate = @ExpiryDate, " &
                                    "IsExpired = @IsExpired, " &
                                    "Feedback = @Feedback, " &
                                    "InstructorComments = @InstructorComments, " &
                                    "ModifiedDate = @ModifiedDate, " &
                                    "ModifiedBy = @ModifiedBy, " &
                                    "RenewalRequired = @RenewalRequired, " &
                                    "RenewalNotificationSent = @RenewalNotificationSent, " &
                                    "CostCentre = @CostCentre, " &
                                    "ApprovalRequired = @ApprovalRequired, " &
                                    "ApprovedBy = @ApprovedBy, " &
                                    "ApprovedDate = @ApprovedDate, " &
                                    "RejectedReason = @RejectedReason " &
                                    "WHERE RecordId = @RecordId"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@RecordId", SqlDbType.Int, record.RecordId),
                    DatabaseHelper.CreateParameter("@EmployeeId", SqlDbType.Int, record.EmployeeId),
                    DatabaseHelper.CreateParameter("@SessionId", SqlDbType.Int, record.SessionId),
                    DatabaseHelper.CreateParameter("@CourseId", SqlDbType.Int, record.CourseId),
                    DatabaseHelper.CreateParameter("@EnrollmentDate", SqlDbType.DateTime, record.EnrollmentDate),
                    DatabaseHelper.CreateParameter("@CompletionDate", SqlDbType.DateTime, record.CompletionDate),
                    DatabaseHelper.CreateParameter("@AttendanceStatus", SqlDbType.NVarChar, 20, record.AttendanceStatus),
                    DatabaseHelper.CreateParameter("@CompletionStatus", SqlDbType.NVarChar, 20, record.CompletionStatus),
                    DatabaseHelper.CreateParameter("@Score", SqlDbType.Int, record.Score),
                    DatabaseHelper.CreateParameter("@MaxScore", SqlDbType.Int, record.MaxScore),
                    DatabaseHelper.CreateParameter("@PassMark", SqlDbType.Int, record.PassMark),
                    DatabaseHelper.CreateParameter("@Grade", SqlDbType.NVarChar, 10, record.Grade),
                    DatabaseHelper.CreateParameter("@CertificateIssued", SqlDbType.Bit, record.CertificateIssued),
                    DatabaseHelper.CreateParameter("@CertificateNumber", SqlDbType.NVarChar, 50, record.CertificateNumber),
                    DatabaseHelper.CreateParameter("@CertificateIssuedDate", SqlDbType.DateTime, record.CertificateIssuedDate),
                    DatabaseHelper.CreateParameter("@ExpiryDate", SqlDbType.DateTime, record.ExpiryDate),
                    DatabaseHelper.CreateParameter("@IsExpired", SqlDbType.Bit, record.IsExpired),
                    DatabaseHelper.CreateParameter("@Feedback", record.Feedback),
                    DatabaseHelper.CreateParameter("@InstructorComments", record.InstructorComments),
                    DatabaseHelper.CreateParameter("@ModifiedDate", SqlDbType.DateTime, record.ModifiedDate),
                    DatabaseHelper.CreateParameter("@ModifiedBy", SqlDbType.NVarChar, 50, record.ModifiedBy),
                    DatabaseHelper.CreateParameter("@RenewalRequired", SqlDbType.Bit, record.RenewalRequired),
                    DatabaseHelper.CreateParameter("@RenewalNotificationSent", SqlDbType.Bit, record.RenewalNotificationSent),
                    DatabaseHelper.CreateParameter("@CostCentre", SqlDbType.NVarChar, 20, record.CostCentre),
                    DatabaseHelper.CreateParameter("@ApprovalRequired", SqlDbType.Bit, record.ApprovalRequired),
                    DatabaseHelper.CreateParameter("@ApprovedBy", SqlDbType.NVarChar, 50, record.ApprovedBy),
                    DatabaseHelper.CreateParameter("@ApprovedDate", SqlDbType.DateTime, record.ApprovedDate),
                    DatabaseHelper.CreateParameter("@RejectedReason", record.RejectedReason)
                }

                Dim rowsAffected As Integer = _db.ExecuteNonQuery(sql, parameters)
                Return rowsAffected > 0
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in UpdateTrainingRecord ({record.RecordId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to update training record {record.RecordId}", ex)
            Finally
                _db.CloseConnection()
            End Try
        End Function

        Public Function GetExpiredTrainingRecords() As List(Of TrainingRecord)
            Dim records As New List(Of TrainingRecord)()
            Try
                Dim sql As String = "SELECT * FROM TrainingRecords " &
                                    "WHERE CompletionStatus = 'COMPLETED' " &
                                    "AND ExpiryDate IS NOT NULL AND ExpiryDate < GETDATE() " &
                                    "ORDER BY ExpiryDate"

                Using reader As SqlDataReader = _db.ExecuteReader(sql, Nothing)
                    While reader.Read()
                        records.Add(MapTrainingRecord(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetExpiredTrainingRecords: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve expired training records", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return records
        End Function

        Public Function GetRecordsNearingExpiry(days As Integer) As List(Of TrainingRecord)
            Dim records As New List(Of TrainingRecord)()
            Try
                Dim sql As String = "SELECT * FROM TrainingRecords " &
                                    "WHERE CompletionStatus = 'COMPLETED' " &
                                    "AND ExpiryDate IS NOT NULL " &
                                    "AND ExpiryDate >= GETDATE() " &
                                    "AND ExpiryDate <= DATEADD(day, @Days, GETDATE()) " &
                                    "ORDER BY ExpiryDate"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@Days", SqlDbType.Int, days)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    While reader.Read()
                        records.Add(MapTrainingRecord(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetRecordsNearingExpiry: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve training records nearing expiry", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return records
        End Function

        ' ===========================
        ' TrainingSession Methods
        ' ===========================

        Public Function GetAllTrainingSessions() As List(Of TrainingSession)
            Dim sessions As New List(Of TrainingSession)()
            Try
                Dim sql As String = "SELECT * FROM TrainingSessions ORDER BY SessionDate DESC"

                Using reader As SqlDataReader = _db.ExecuteReader(sql, Nothing)
                    While reader.Read()
                        sessions.Add(MapTrainingSession(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetAllTrainingSessions: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve training sessions", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return sessions
        End Function

        Public Function GetTrainingSessionById(sessionId As Integer) As TrainingSession
            Try
                Dim sql As String = "SELECT * FROM TrainingSessions WHERE SessionId = @SessionId"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@SessionId", SqlDbType.Int, sessionId)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    If reader.Read() Then
                        Return MapTrainingSession(reader)
                    End If
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetTrainingSessionById ({sessionId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve training session with ID {sessionId}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return Nothing
        End Function

        Public Function GetTrainingSessionsByCourse(courseId As Integer) As List(Of TrainingSession)
            Dim sessions As New List(Of TrainingSession)()
            Try
                Dim sql As String = "SELECT * FROM TrainingSessions WHERE CourseId = @CourseId ORDER BY SessionDate DESC"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@CourseId", SqlDbType.Int, courseId)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    While reader.Read()
                        sessions.Add(MapTrainingSession(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetTrainingSessionsByCourse ({courseId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve training sessions for course {courseId}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return sessions
        End Function

        Public Function GetUpcomingTrainingSessions() As List(Of TrainingSession)
            Dim sessions As New List(Of TrainingSession)()
            Try
                Dim sql As String = "SELECT * FROM TrainingSessions " &
                                    "WHERE SessionDate >= GETDATE() " &
                                    "AND SessionStatus IN ('SCHEDULED', 'CONFIRMED') " &
                                    "ORDER BY SessionDate"

                Using reader As SqlDataReader = _db.ExecuteReader(sql, Nothing)
                    While reader.Read()
                        sessions.Add(MapTrainingSession(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetUpcomingTrainingSessions: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve upcoming training sessions", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return sessions
        End Function

        Public Function GetTrainingSessionsByInstructor(instructorId As Integer) As List(Of TrainingSession)
            Dim sessions As New List(Of TrainingSession)()
            Try
                Dim sql As String = "SELECT * FROM TrainingSessions " &
                                    "WHERE (PrimaryInstructorId = @InstructorId OR SecondaryInstructorId = @InstructorId) " &
                                    "ORDER BY SessionDate DESC"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@InstructorId", SqlDbType.Int, instructorId)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    While reader.Read()
                        sessions.Add(MapTrainingSession(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetTrainingSessionsByInstructor ({instructorId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve training sessions for instructor {instructorId}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return sessions
        End Function

        Public Function CreateTrainingSession(session As TrainingSession) As Integer
            Try
                Dim sql As String = "INSERT INTO TrainingSessions " &
                                    "(CourseId, SessionDate, StartTime, EndTime, Location, MaxParticipants, " &
                                    "CurrentParticipants, PrimaryInstructorId, SecondaryInstructorId, SessionStatus, " &
                                    "RegistrationDeadline, SessionNotes, InstructorNotes, EquipmentRequired, " &
                                    "CateringRequired, MaterialsPrepared, RoomBooked, NotificationsSent, " &
                                    "WaitingListEnabled, PrerequisitesChecked, CreatedDate, CreatedBy, " &
                                    "CostPerParticipant, TotalCost, ApprovedBy, ApprovedDate, " &
                                    "CancelledReason, FeedbackRequested) " &
                                    "VALUES " &
                                    "(@CourseId, @SessionDate, @StartTime, @EndTime, @Location, @MaxParticipants, " &
                                    "@CurrentParticipants, @PrimaryInstructorId, @SecondaryInstructorId, @SessionStatus, " &
                                    "@RegistrationDeadline, @SessionNotes, @InstructorNotes, @EquipmentRequired, " &
                                    "@CateringRequired, @MaterialsPrepared, @RoomBooked, @NotificationsSent, " &
                                    "@WaitingListEnabled, @PrerequisitesChecked, @CreatedDate, @CreatedBy, " &
                                    "@CostPerParticipant, @TotalCost, @ApprovedBy, @ApprovedDate, " &
                                    "@CancelledReason, @FeedbackRequested); " &
                                    "SELECT SCOPE_IDENTITY();"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@CourseId", SqlDbType.Int, session.CourseId),
                    DatabaseHelper.CreateParameter("@SessionDate", SqlDbType.DateTime, session.SessionDate),
                    DatabaseHelper.CreateParameter("@StartTime", session.StartTime),
                    DatabaseHelper.CreateParameter("@EndTime", session.EndTime),
                    DatabaseHelper.CreateParameter("@Location", SqlDbType.NVarChar, 200, session.Location),
                    DatabaseHelper.CreateParameter("@MaxParticipants", SqlDbType.Int, session.MaxParticipants),
                    DatabaseHelper.CreateParameter("@CurrentParticipants", SqlDbType.Int, session.CurrentParticipants),
                    DatabaseHelper.CreateParameter("@PrimaryInstructorId", SqlDbType.Int, session.PrimaryInstructorId),
                    DatabaseHelper.CreateParameter("@SecondaryInstructorId", SqlDbType.Int, session.SecondaryInstructorId),
                    DatabaseHelper.CreateParameter("@SessionStatus", SqlDbType.NVarChar, 20, session.SessionStatus),
                    DatabaseHelper.CreateParameter("@RegistrationDeadline", SqlDbType.DateTime, session.RegistrationDeadline),
                    DatabaseHelper.CreateParameter("@SessionNotes", session.SessionNotes),
                    DatabaseHelper.CreateParameter("@InstructorNotes", session.InstructorNotes),
                    DatabaseHelper.CreateParameter("@EquipmentRequired", session.EquipmentRequired),
                    DatabaseHelper.CreateParameter("@CateringRequired", SqlDbType.Bit, session.CateringRequired),
                    DatabaseHelper.CreateParameter("@MaterialsPrepared", SqlDbType.Bit, session.MaterialsPrepared),
                    DatabaseHelper.CreateParameter("@RoomBooked", SqlDbType.Bit, session.RoomBooked),
                    DatabaseHelper.CreateParameter("@NotificationsSent", SqlDbType.Bit, session.NotificationsSent),
                    DatabaseHelper.CreateParameter("@WaitingListEnabled", SqlDbType.Bit, session.WaitingListEnabled),
                    DatabaseHelper.CreateParameter("@PrerequisitesChecked", SqlDbType.Bit, session.PrerequisitesChecked),
                    DatabaseHelper.CreateParameter("@CreatedDate", SqlDbType.DateTime, session.CreatedDate),
                    DatabaseHelper.CreateParameter("@CreatedBy", SqlDbType.NVarChar, 50, session.CreatedBy),
                    DatabaseHelper.CreateParameter("@CostPerParticipant", SqlDbType.Decimal, session.CostPerParticipant),
                    DatabaseHelper.CreateParameter("@TotalCost", SqlDbType.Decimal, session.TotalCost),
                    DatabaseHelper.CreateParameter("@ApprovedBy", SqlDbType.NVarChar, 50, session.ApprovedBy),
                    DatabaseHelper.CreateParameter("@ApprovedDate", SqlDbType.DateTime, session.ApprovedDate),
                    DatabaseHelper.CreateParameter("@CancelledReason", session.CancelledReason),
                    DatabaseHelper.CreateParameter("@FeedbackRequested", SqlDbType.Bit, session.FeedbackRequested)
                }

                Dim result As Object = _db.ExecuteScalar(sql, parameters)
                If result IsNot Nothing AndAlso result IsNot DBNull.Value Then
                    Return Convert.ToInt32(result)
                End If
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in CreateTrainingSession: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to create training session", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return 0
        End Function

        Public Function UpdateTrainingSession(session As TrainingSession) As Boolean
            Try
                Dim sql As String = "UPDATE TrainingSessions SET " &
                                    "CourseId = @CourseId, " &
                                    "SessionDate = @SessionDate, " &
                                    "StartTime = @StartTime, " &
                                    "EndTime = @EndTime, " &
                                    "Location = @Location, " &
                                    "MaxParticipants = @MaxParticipants, " &
                                    "CurrentParticipants = @CurrentParticipants, " &
                                    "PrimaryInstructorId = @PrimaryInstructorId, " &
                                    "SecondaryInstructorId = @SecondaryInstructorId, " &
                                    "SessionStatus = @SessionStatus, " &
                                    "RegistrationDeadline = @RegistrationDeadline, " &
                                    "SessionNotes = @SessionNotes, " &
                                    "InstructorNotes = @InstructorNotes, " &
                                    "EquipmentRequired = @EquipmentRequired, " &
                                    "CateringRequired = @CateringRequired, " &
                                    "MaterialsPrepared = @MaterialsPrepared, " &
                                    "RoomBooked = @RoomBooked, " &
                                    "NotificationsSent = @NotificationsSent, " &
                                    "WaitingListEnabled = @WaitingListEnabled, " &
                                    "PrerequisitesChecked = @PrerequisitesChecked, " &
                                    "ModifiedDate = @ModifiedDate, " &
                                    "ModifiedBy = @ModifiedBy, " &
                                    "CostPerParticipant = @CostPerParticipant, " &
                                    "TotalCost = @TotalCost, " &
                                    "ApprovedBy = @ApprovedBy, " &
                                    "ApprovedDate = @ApprovedDate, " &
                                    "CancelledReason = @CancelledReason, " &
                                    "FeedbackRequested = @FeedbackRequested " &
                                    "WHERE SessionId = @SessionId"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@SessionId", SqlDbType.Int, session.SessionId),
                    DatabaseHelper.CreateParameter("@CourseId", SqlDbType.Int, session.CourseId),
                    DatabaseHelper.CreateParameter("@SessionDate", SqlDbType.DateTime, session.SessionDate),
                    DatabaseHelper.CreateParameter("@StartTime", session.StartTime),
                    DatabaseHelper.CreateParameter("@EndTime", session.EndTime),
                    DatabaseHelper.CreateParameter("@Location", SqlDbType.NVarChar, 200, session.Location),
                    DatabaseHelper.CreateParameter("@MaxParticipants", SqlDbType.Int, session.MaxParticipants),
                    DatabaseHelper.CreateParameter("@CurrentParticipants", SqlDbType.Int, session.CurrentParticipants),
                    DatabaseHelper.CreateParameter("@PrimaryInstructorId", SqlDbType.Int, session.PrimaryInstructorId),
                    DatabaseHelper.CreateParameter("@SecondaryInstructorId", SqlDbType.Int, session.SecondaryInstructorId),
                    DatabaseHelper.CreateParameter("@SessionStatus", SqlDbType.NVarChar, 20, session.SessionStatus),
                    DatabaseHelper.CreateParameter("@RegistrationDeadline", SqlDbType.DateTime, session.RegistrationDeadline),
                    DatabaseHelper.CreateParameter("@SessionNotes", session.SessionNotes),
                    DatabaseHelper.CreateParameter("@InstructorNotes", session.InstructorNotes),
                    DatabaseHelper.CreateParameter("@EquipmentRequired", session.EquipmentRequired),
                    DatabaseHelper.CreateParameter("@CateringRequired", SqlDbType.Bit, session.CateringRequired),
                    DatabaseHelper.CreateParameter("@MaterialsPrepared", SqlDbType.Bit, session.MaterialsPrepared),
                    DatabaseHelper.CreateParameter("@RoomBooked", SqlDbType.Bit, session.RoomBooked),
                    DatabaseHelper.CreateParameter("@NotificationsSent", SqlDbType.Bit, session.NotificationsSent),
                    DatabaseHelper.CreateParameter("@WaitingListEnabled", SqlDbType.Bit, session.WaitingListEnabled),
                    DatabaseHelper.CreateParameter("@PrerequisitesChecked", SqlDbType.Bit, session.PrerequisitesChecked),
                    DatabaseHelper.CreateParameter("@ModifiedDate", SqlDbType.DateTime, session.ModifiedDate),
                    DatabaseHelper.CreateParameter("@ModifiedBy", SqlDbType.NVarChar, 50, session.ModifiedBy),
                    DatabaseHelper.CreateParameter("@CostPerParticipant", SqlDbType.Decimal, session.CostPerParticipant),
                    DatabaseHelper.CreateParameter("@TotalCost", SqlDbType.Decimal, session.TotalCost),
                    DatabaseHelper.CreateParameter("@ApprovedBy", SqlDbType.NVarChar, 50, session.ApprovedBy),
                    DatabaseHelper.CreateParameter("@ApprovedDate", SqlDbType.DateTime, session.ApprovedDate),
                    DatabaseHelper.CreateParameter("@CancelledReason", session.CancelledReason),
                    DatabaseHelper.CreateParameter("@FeedbackRequested", SqlDbType.Bit, session.FeedbackRequested)
                }

                Dim rowsAffected As Integer = _db.ExecuteNonQuery(sql, parameters)
                Return rowsAffected > 0
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in UpdateTrainingSession ({session.SessionId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to update training session {session.SessionId}", ex)
            Finally
                _db.CloseConnection()
            End Try
        End Function

        Public Function CancelTrainingSession(sessionId As Integer, cancelledReason As String, modifiedBy As String) As Boolean
            Try
                Dim sql As String = "UPDATE TrainingSessions SET " &
                                    "SessionStatus = 'CANCELLED', " &
                                    "CancelledReason = @CancelledReason, " &
                                    "ModifiedDate = @ModifiedDate, " &
                                    "ModifiedBy = @ModifiedBy " &
                                    "WHERE SessionId = @SessionId"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@SessionId", SqlDbType.Int, sessionId),
                    DatabaseHelper.CreateParameter("@CancelledReason", cancelledReason),
                    DatabaseHelper.CreateParameter("@ModifiedDate", SqlDbType.DateTime, DateTime.Now),
                    DatabaseHelper.CreateParameter("@ModifiedBy", SqlDbType.NVarChar, 50, modifiedBy)
                }

                Dim rowsAffected As Integer = _db.ExecuteNonQuery(sql, parameters)
                Return rowsAffected > 0
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in CancelTrainingSession ({sessionId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to cancel training session {sessionId}", ex)
            Finally
                _db.CloseConnection()
            End Try
        End Function

        Public Function GetSessionParticipants(sessionId As Integer) As List(Of Employee)
            Dim employees As New List(Of Employee)()
            Try
                Dim sql As String = "SELECT e.* " &
                                    "FROM Employees e " &
                                    "INNER JOIN SessionParticipants sp ON e.EmployeeId = sp.EmployeeId " &
                                    "WHERE sp.SessionId = @SessionId AND sp.EnrollmentStatus = 'ENROLLED' " &
                                    "ORDER BY e.LastName, e.FirstName"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@SessionId", SqlDbType.Int, sessionId)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    While reader.Read()
                        Dim employee As New Employee()
                        employee.EmployeeId = DatabaseHelper.SafeGetInteger(reader, "EmployeeId")
                        employee.EmployeeNumber = DatabaseHelper.SafeGetString(reader, "EmployeeNumber")
                        employee.FirstName = DatabaseHelper.SafeGetString(reader, "FirstName")
                        employee.LastName = DatabaseHelper.SafeGetString(reader, "LastName")
                        employee.Email = DatabaseHelper.SafeGetString(reader, "Email")
                        employee.UserName = DatabaseHelper.SafeGetString(reader, "UserName")
                        employee.Department = DatabaseHelper.SafeGetString(reader, "Department")
                        employee.Position = DatabaseHelper.SafeGetString(reader, "Position")
                        employee.Location = DatabaseHelper.SafeGetString(reader, "Location")
                        employee.ManagerId = DatabaseHelper.SafeGetNullableInteger(reader, "ManagerId")
                        employee.ManagerName = DatabaseHelper.SafeGetString(reader, "ManagerName")
                        employee.HireDate = DatabaseHelper.SafeGetDateTime(reader, "HireDate")
                        employee.LastLoginDate = DatabaseHelper.SafeGetDateTime(reader, "LastLoginDate")
                        employee.IsActive = DatabaseHelper.SafeGetBoolean(reader, "IsActive")
                        employee.PhoneNumber = DatabaseHelper.SafeGetString(reader, "PhoneNumber")
                        employee.LineManagerEmail = DatabaseHelper.SafeGetString(reader, "LineManagerEmail")
                        employee.CostCentre = DatabaseHelper.SafeGetString(reader, "CostCentre")
                        employee.PayBand = DatabaseHelper.SafeGetString(reader, "PayBand")
                        employee.ContractType = DatabaseHelper.SafeGetString(reader, "ContractType")
                        employee.WorkingPattern = DatabaseHelper.SafeGetString(reader, "WorkingPattern")
                        employee.CreatedDate = DatabaseHelper.SafeGetDateTime(reader, "CreatedDate")
                        employee.CreatedBy = DatabaseHelper.SafeGetString(reader, "CreatedBy")
                        employee.ModifiedDate = DatabaseHelper.SafeGetNullableDateTime(reader, "ModifiedDate")
                        employee.ModifiedBy = DatabaseHelper.SafeGetString(reader, "ModifiedBy")
                        employees.Add(employee)
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetSessionParticipants ({sessionId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve participants for session {sessionId}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return employees
        End Function

        Public Function EnrollParticipant(sessionId As Integer, employeeId As Integer, createdBy As String) As Boolean
            Try
                Dim sql As String = "INSERT INTO SessionParticipants (SessionId, EmployeeId, EnrollmentDate, EnrollmentStatus, CreatedDate, CreatedBy) " &
                                    "VALUES (@SessionId, @EmployeeId, @EnrollmentDate, 'ENROLLED', @CreatedDate, @CreatedBy); " &
                                    "UPDATE TrainingSessions SET CurrentParticipants = CurrentParticipants + 1 WHERE SessionId = @SessionId"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@SessionId", SqlDbType.Int, sessionId),
                    DatabaseHelper.CreateParameter("@EmployeeId", SqlDbType.Int, employeeId),
                    DatabaseHelper.CreateParameter("@EnrollmentDate", SqlDbType.DateTime, DateTime.Now),
                    DatabaseHelper.CreateParameter("@CreatedDate", SqlDbType.DateTime, DateTime.Now),
                    DatabaseHelper.CreateParameter("@CreatedBy", SqlDbType.NVarChar, 50, createdBy)
                }

                Dim rowsAffected As Integer = _db.ExecuteNonQuery(sql, parameters)
                Return rowsAffected > 0
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in EnrollParticipant (session {sessionId}, employee {employeeId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to enroll employee {employeeId} in session {sessionId}", ex)
            Finally
                _db.CloseConnection()
            End Try
        End Function

        Public Function WithdrawParticipant(sessionId As Integer, employeeId As Integer) As Boolean
            Try
                Dim sql As String = "UPDATE SessionParticipants SET EnrollmentStatus = 'WITHDRAWN' " &
                                    "WHERE SessionId = @SessionId AND EmployeeId = @EmployeeId AND EnrollmentStatus = 'ENROLLED'; " &
                                    "UPDATE TrainingSessions SET CurrentParticipants = CASE WHEN CurrentParticipants > 0 THEN CurrentParticipants - 1 ELSE 0 END " &
                                    "WHERE SessionId = @SessionId"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@SessionId", SqlDbType.Int, sessionId),
                    DatabaseHelper.CreateParameter("@EmployeeId", SqlDbType.Int, employeeId)
                }

                Dim rowsAffected As Integer = _db.ExecuteNonQuery(sql, parameters)
                Return rowsAffected > 0
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in WithdrawParticipant (session {sessionId}, employee {employeeId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to withdraw employee {employeeId} from session {sessionId}", ex)
            Finally
                _db.CloseConnection()
            End Try
        End Function

        Public Function GetDashboardMetrics() As Dictionary(Of String, Object)
            Dim metrics As New Dictionary(Of String, Object)()
            Try
                Dim table As DataTable = _db.ExecuteStoredProcedure("sp_GetDashboardMetrics", Nothing)

                If table.Rows.Count > 0 Then
                    Dim row As DataRow = table.Rows(0)
                    metrics("ActiveEmployees") = If(row("ActiveEmployees") Is DBNull.Value, 0, Convert.ToInt32(row("ActiveEmployees")))
                    metrics("TotalCourses") = If(row("TotalCourses") Is DBNull.Value, 0, Convert.ToInt32(row("TotalCourses")))
                    metrics("UpcomingSessions") = If(row("UpcomingSessions") Is DBNull.Value, 0, Convert.ToInt32(row("UpcomingSessions")))
                    metrics("CompletedThisMonth") = If(row("CompletedThisMonth") Is DBNull.Value, 0, Convert.ToInt32(row("CompletedThisMonth")))
                    metrics("ComplianceRate") = If(row("ComplianceRate") Is DBNull.Value, 0D, Convert.ToDecimal(row("ComplianceRate")))
                End If
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetDashboardMetrics: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve dashboard metrics", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return metrics
        End Function

        ' ===========================
        ' Mapping Methods
        ' ===========================

        Private Function MapTrainingRecord(reader As SqlDataReader) As TrainingRecord
            Dim record As New TrainingRecord()
            record.RecordId = DatabaseHelper.SafeGetInteger(reader, "RecordId")
            record.EmployeeId = DatabaseHelper.SafeGetInteger(reader, "EmployeeId")
            record.SessionId = DatabaseHelper.SafeGetInteger(reader, "SessionId")
            record.CourseId = DatabaseHelper.SafeGetInteger(reader, "CourseId")
            record.EnrollmentDate = DatabaseHelper.SafeGetDateTime(reader, "EnrollmentDate")
            record.CompletionDate = DatabaseHelper.SafeGetNullableDateTime(reader, "CompletionDate")
            record.AttendanceStatus = DatabaseHelper.SafeGetString(reader, "AttendanceStatus")
            record.CompletionStatus = DatabaseHelper.SafeGetString(reader, "CompletionStatus")
            record.Score = DatabaseHelper.SafeGetNullableInteger(reader, "Score")
            record.MaxScore = DatabaseHelper.SafeGetNullableInteger(reader, "MaxScore")
            record.PassMark = DatabaseHelper.SafeGetNullableInteger(reader, "PassMark")
            record.Grade = DatabaseHelper.SafeGetString(reader, "Grade")
            record.CertificateIssued = DatabaseHelper.SafeGetBoolean(reader, "CertificateIssued")
            record.CertificateNumber = DatabaseHelper.SafeGetString(reader, "CertificateNumber")
            record.CertificateIssuedDate = DatabaseHelper.SafeGetNullableDateTime(reader, "CertificateIssuedDate")
            record.ExpiryDate = DatabaseHelper.SafeGetNullableDateTime(reader, "ExpiryDate")
            record.IsExpired = DatabaseHelper.SafeGetBoolean(reader, "IsExpired")
            record.Feedback = DatabaseHelper.SafeGetString(reader, "Feedback")
            record.InstructorComments = DatabaseHelper.SafeGetString(reader, "InstructorComments")
            record.CreatedDate = DatabaseHelper.SafeGetDateTime(reader, "CreatedDate")
            record.CreatedBy = DatabaseHelper.SafeGetString(reader, "CreatedBy")
            record.ModifiedDate = DatabaseHelper.SafeGetNullableDateTime(reader, "ModifiedDate")
            record.ModifiedBy = DatabaseHelper.SafeGetString(reader, "ModifiedBy")
            record.RenewalRequired = DatabaseHelper.SafeGetBoolean(reader, "RenewalRequired")
            record.RenewalNotificationSent = DatabaseHelper.SafeGetBoolean(reader, "RenewalNotificationSent")
            record.CostCentre = DatabaseHelper.SafeGetString(reader, "CostCentre")
            record.ApprovalRequired = DatabaseHelper.SafeGetBoolean(reader, "ApprovalRequired")
            record.ApprovedBy = DatabaseHelper.SafeGetString(reader, "ApprovedBy")
            record.ApprovedDate = DatabaseHelper.SafeGetNullableDateTime(reader, "ApprovedDate")
            record.RejectedReason = DatabaseHelper.SafeGetString(reader, "RejectedReason")
            Return record
        End Function

        Private Function MapTrainingSession(reader As SqlDataReader) As TrainingSession
            Dim session As New TrainingSession()
            session.SessionId = DatabaseHelper.SafeGetInteger(reader, "SessionId")
            session.CourseId = DatabaseHelper.SafeGetInteger(reader, "CourseId")
            session.SessionDate = DatabaseHelper.SafeGetDateTime(reader, "SessionDate")
            session.StartTime = DatabaseHelper.SafeGetTimeSpan(reader, "StartTime")
            session.EndTime = DatabaseHelper.SafeGetTimeSpan(reader, "EndTime")
            session.Location = DatabaseHelper.SafeGetString(reader, "Location")
            session.MaxParticipants = DatabaseHelper.SafeGetInteger(reader, "MaxParticipants")
            session.CurrentParticipants = DatabaseHelper.SafeGetInteger(reader, "CurrentParticipants")
            session.PrimaryInstructorId = DatabaseHelper.SafeGetInteger(reader, "PrimaryInstructorId")
            session.SecondaryInstructorId = DatabaseHelper.SafeGetNullableInteger(reader, "SecondaryInstructorId")
            session.SessionStatus = DatabaseHelper.SafeGetString(reader, "SessionStatus")
            session.RegistrationDeadline = DatabaseHelper.SafeGetDateTime(reader, "RegistrationDeadline")
            session.SessionNotes = DatabaseHelper.SafeGetString(reader, "SessionNotes")
            session.InstructorNotes = DatabaseHelper.SafeGetString(reader, "InstructorNotes")
            session.EquipmentRequired = DatabaseHelper.SafeGetString(reader, "EquipmentRequired")
            session.CateringRequired = DatabaseHelper.SafeGetBoolean(reader, "CateringRequired")
            session.MaterialsPrepared = DatabaseHelper.SafeGetBoolean(reader, "MaterialsPrepared")
            session.RoomBooked = DatabaseHelper.SafeGetBoolean(reader, "RoomBooked")
            session.NotificationsSent = DatabaseHelper.SafeGetBoolean(reader, "NotificationsSent")
            session.WaitingListEnabled = DatabaseHelper.SafeGetBoolean(reader, "WaitingListEnabled")
            session.PrerequisitesChecked = DatabaseHelper.SafeGetBoolean(reader, "PrerequisitesChecked")
            session.CreatedDate = DatabaseHelper.SafeGetDateTime(reader, "CreatedDate")
            session.CreatedBy = DatabaseHelper.SafeGetString(reader, "CreatedBy")
            session.ModifiedDate = DatabaseHelper.SafeGetNullableDateTime(reader, "ModifiedDate")
            session.ModifiedBy = DatabaseHelper.SafeGetString(reader, "ModifiedBy")
            session.CostPerParticipant = DatabaseHelper.SafeGetDecimal(reader, "CostPerParticipant")
            session.TotalCost = DatabaseHelper.SafeGetDecimal(reader, "TotalCost")
            session.ApprovedBy = DatabaseHelper.SafeGetString(reader, "ApprovedBy")
            session.ApprovedDate = DatabaseHelper.SafeGetNullableDateTime(reader, "ApprovedDate")
            session.CancelledReason = DatabaseHelper.SafeGetString(reader, "CancelledReason")
            session.FeedbackRequested = DatabaseHelper.SafeGetBoolean(reader, "FeedbackRequested")
            Return session
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            Try
                If _db IsNot Nothing Then
                    _db.Dispose()
                    _db = Nothing
                End If
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error disposing TrainingRepository: {ex.Message}", EventLogEntryType.Warning)
            End Try
        End Sub

    End Class
End Namespace
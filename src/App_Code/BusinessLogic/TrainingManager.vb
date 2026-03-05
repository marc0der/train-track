Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports Defra.TrainTrack.Models
Imports Defra.TrainTrack.DataAccess

Namespace Defra.TrainTrack.BusinessLogic
    Public Class TrainingManager
        Implements IDisposable

        Private _repository As TrainingRepository
        Private _waitingListRepository As WaitingListRepository

        Public Sub New()
            _repository = New TrainingRepository()
            _waitingListRepository = New WaitingListRepository()
        End Sub

        ' ===========================
        ' Training Record Methods
        ' ===========================

        Public Function GetAllTrainingRecords() As List(Of TrainingRecord)
            Try
                Return _repository.GetAllTrainingRecords()
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting all training records: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve training record list", ex)
            End Try
        End Function

        Public Function GetTrainingRecordById(recordId As Integer) As TrainingRecord
            Try
                If recordId <= 0 Then
                    Throw New ArgumentException("Record ID must be greater than 0", "recordId")
                End If

                Return _repository.GetTrainingRecordById(recordId)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting training record by ID {recordId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve training record with ID {recordId}", ex)
            End Try
        End Function

        Public Function GetTrainingRecordsByEmployee(employeeId As Integer) As List(Of TrainingRecord)
            Try
                If employeeId <= 0 Then
                    Throw New ArgumentException("Employee ID must be greater than 0", "employeeId")
                End If

                Return _repository.GetTrainingRecordsByEmployee(employeeId)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting training records for employee {employeeId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve training records for employee {employeeId}", ex)
            End Try
        End Function

        Public Function GetTrainingRecordsByCourse(courseId As Integer) As List(Of TrainingRecord)
            Try
                If courseId <= 0 Then
                    Throw New ArgumentException("Course ID must be greater than 0", "courseId")
                End If

                Return _repository.GetTrainingRecordsByCourse(courseId)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting training records for course {courseId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve training records for course {courseId}", ex)
            End Try
        End Function

        Public Function GetTrainingRecordsBySession(sessionId As Integer) As List(Of TrainingRecord)
            Try
                If sessionId <= 0 Then
                    Throw New ArgumentException("Session ID must be greater than 0", "sessionId")
                End If

                Return _repository.GetTrainingRecordsBySession(sessionId)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting training records for session {sessionId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve training records for session {sessionId}", ex)
            End Try
        End Function

        Public Function CreateTrainingRecord(record As TrainingRecord, createdBy As String) As Integer
            Try
                If record Is Nothing Then
                    Throw New ArgumentNullException("record", "Training record cannot be null")
                End If

                If String.IsNullOrWhiteSpace(createdBy) Then
                    Throw New ArgumentException("Created by cannot be empty", "createdBy")
                End If

                ' Validate training record data
                Dim validationErrors = record.GetValidationErrors()
                If validationErrors.Count > 0 Then
                    Throw New ArgumentException($"Training record validation failed: {String.Join("; ", validationErrors)}")
                End If

                ' Check if employee is already enrolled in this session
                Dim existingRecords = _repository.GetTrainingRecordsBySession(record.SessionId)
                For Each existingRecord In existingRecords
                    If existingRecord.EmployeeId = record.EmployeeId AndAlso
                       existingRecord.CompletionStatus <> "CANCELLED" AndAlso
                       existingRecord.AttendanceStatus <> "WITHDRAWN" Then
                        Throw New InvalidOperationException($"Employee {record.EmployeeId} is already enrolled in session {record.SessionId}")
                    End If
                Next

                ' Check if employee has already completed this course and it hasn't expired
                If HasEmployeeCompletedCourse(record.EmployeeId, record.CourseId) Then
                    ' Allow re-enrollment but log a warning
                    EventLog.WriteEntry("TrainTrack", $"Employee {record.EmployeeId} is being re-enrolled for course {record.CourseId} (already completed)", EventLogEntryType.Warning)
                End If

                ' Set audit fields
                record.CreatedBy = createdBy
                record.CreatedDate = DateTime.Now

                ' Create the training record
                Dim newRecordId As Integer = _repository.CreateTrainingRecord(record)

                ' Log the action
                EventLog.WriteEntry("TrainTrack", $"Training record created: ID {newRecordId} for employee {record.EmployeeId}, session {record.SessionId} by {createdBy}", EventLogEntryType.Information)

                Return newRecordId
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error creating training record: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to create training record", ex)
            End Try
        End Function

        Public Function UpdateTrainingRecord(record As TrainingRecord, modifiedBy As String) As Boolean
            Try
                If record Is Nothing Then
                    Throw New ArgumentNullException("record", "Training record cannot be null")
                End If

                If record.RecordId <= 0 Then
                    Throw New ArgumentException("Record ID must be greater than 0", "record")
                End If

                If String.IsNullOrWhiteSpace(modifiedBy) Then
                    Throw New ArgumentException("Modified by cannot be empty", "modifiedBy")
                End If

                ' Validate training record data
                Dim validationErrors = record.GetValidationErrors()
                If validationErrors.Count > 0 Then
                    Throw New ArgumentException($"Training record validation failed: {String.Join("; ", validationErrors)}")
                End If

                ' Check if record exists
                Dim existingRecord = _repository.GetTrainingRecordById(record.RecordId)
                If existingRecord Is Nothing Then
                    Throw New InvalidOperationException($"Training record with ID {record.RecordId} does not exist")
                End If

                ' Set audit fields
                record.ModifiedBy = modifiedBy
                record.ModifiedDate = DateTime.Now

                ' Update the training record
                Dim success As Boolean = _repository.UpdateTrainingRecord(record)

                If success Then
                    ' Log the action
                    EventLog.WriteEntry("TrainTrack", $"Training record updated: ID {record.RecordId} by {modifiedBy}", EventLogEntryType.Information)
                End If

                Return success
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error updating training record: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to update training record", ex)
            End Try
        End Function

        Public Function MarkRecordAsCompleted(recordId As Integer, score As Integer?, grade As String, instructorComments As String, modifiedBy As String) As Boolean
            Try
                If recordId <= 0 Then
                    Throw New ArgumentException("Record ID must be greater than 0", "recordId")
                End If

                If String.IsNullOrWhiteSpace(modifiedBy) Then
                    Throw New ArgumentException("Modified by cannot be empty", "modifiedBy")
                End If

                ' Get existing record
                Dim record = _repository.GetTrainingRecordById(recordId)
                If record Is Nothing Then
                    Throw New InvalidOperationException($"Training record with ID {recordId} does not exist")
                End If

                ' Mark as completed using model business method
                record.MarkAsCompleted(score, grade, instructorComments)
                record.ModifiedBy = modifiedBy

                ' Calculate expiry date based on course validity period
                Using courseManager As New CourseManager()
                    Dim course = courseManager.GetCourseById(record.CourseId)
                    If course IsNot Nothing AndAlso course.ValidityPeriodMonths > 0 Then
                        record.ExpiryDate = DateTime.Now.AddMonths(course.ValidityPeriodMonths)
                        record.RenewalRequired = True
                    End If
                End Using

                ' Update the record
                Dim success As Boolean = _repository.UpdateTrainingRecord(record)

                If success Then
                    EventLog.WriteEntry("TrainTrack", $"Training record completed: ID {recordId} by {modifiedBy}", EventLogEntryType.Information)
                End If

                Return success
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error marking training record {recordId} as completed: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to mark training record {recordId} as completed", ex)
            End Try
        End Function

        Public Function MarkRecordAsFailed(recordId As Integer, score As Integer?, comments As String, modifiedBy As String) As Boolean
            Try
                If recordId <= 0 Then
                    Throw New ArgumentException("Record ID must be greater than 0", "recordId")
                End If

                If String.IsNullOrWhiteSpace(modifiedBy) Then
                    Throw New ArgumentException("Modified by cannot be empty", "modifiedBy")
                End If

                ' Get existing record
                Dim record = _repository.GetTrainingRecordById(recordId)
                If record Is Nothing Then
                    Throw New InvalidOperationException($"Training record with ID {recordId} does not exist")
                End If

                ' Mark as failed using model business method
                record.MarkAsFailed(score, comments)
                record.ModifiedBy = modifiedBy

                ' Update the record
                Dim success As Boolean = _repository.UpdateTrainingRecord(record)

                If success Then
                    EventLog.WriteEntry("TrainTrack", $"Training record marked as failed: ID {recordId} by {modifiedBy}", EventLogEntryType.Information)
                End If

                Return success
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error marking training record {recordId} as failed: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to mark training record {recordId} as failed", ex)
            End Try
        End Function

        Public Function MarkRecordAsNoShow(recordId As Integer, modifiedBy As String) As Boolean
            Try
                If recordId <= 0 Then
                    Throw New ArgumentException("Record ID must be greater than 0", "recordId")
                End If

                If String.IsNullOrWhiteSpace(modifiedBy) Then
                    Throw New ArgumentException("Modified by cannot be empty", "modifiedBy")
                End If

                ' Get existing record
                Dim record = _repository.GetTrainingRecordById(recordId)
                If record Is Nothing Then
                    Throw New InvalidOperationException($"Training record with ID {recordId} does not exist")
                End If

                ' Mark as no show using model business method
                record.MarkAsNoShow()
                record.ModifiedBy = modifiedBy

                ' Update the record
                Dim success As Boolean = _repository.UpdateTrainingRecord(record)

                If success Then
                    EventLog.WriteEntry("TrainTrack", $"Training record marked as no show: ID {recordId} by {modifiedBy}", EventLogEntryType.Information)
                End If

                Return success
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error marking training record {recordId} as no show: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to mark training record {recordId} as no show", ex)
            End Try
        End Function

        Public Function IssueCertificate(recordId As Integer, modifiedBy As String) As String
            Try
                If recordId <= 0 Then
                    Throw New ArgumentException("Record ID must be greater than 0", "recordId")
                End If

                If String.IsNullOrWhiteSpace(modifiedBy) Then
                    Throw New ArgumentException("Modified by cannot be empty", "modifiedBy")
                End If

                ' Get existing record
                Dim record = _repository.GetTrainingRecordById(recordId)
                If record Is Nothing Then
                    Throw New InvalidOperationException($"Training record with ID {recordId} does not exist")
                End If

                ' Check if record is completed
                If Not record.IsCompleted Then
                    Throw New InvalidOperationException($"Cannot issue certificate for incomplete training record {recordId}")
                End If

                ' Check if certificate already issued
                If record.CertificateIssued Then
                    Throw New InvalidOperationException($"Certificate already issued for training record {recordId}")
                End If

                ' Generate certificate number
                Dim certificateNumber As String = record.GenerateCertificateNumber()

                ' Issue certificate using model business method
                record.IssueCertificate(certificateNumber)
                record.ModifiedBy = modifiedBy

                ' Update the record
                Dim success As Boolean = _repository.UpdateTrainingRecord(record)

                If success Then
                    EventLog.WriteEntry("TrainTrack", $"Certificate issued: {certificateNumber} for record {recordId} by {modifiedBy}", EventLogEntryType.Information)
                    Return certificateNumber
                End If

                Return String.Empty
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error issuing certificate for record {recordId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to issue certificate for training record {recordId}", ex)
            End Try
        End Function

        Public Function HasEmployeeCompletedCourse(employeeId As Integer, courseId As Integer) As Boolean
            Try
                If employeeId <= 0 Then
                    Throw New ArgumentException("Employee ID must be greater than 0", "employeeId")
                End If

                If courseId <= 0 Then
                    Throw New ArgumentException("Course ID must be greater than 0", "courseId")
                End If

                Return _repository.HasEmployeeCompletedCourse(employeeId, courseId)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error checking if employee {employeeId} completed course {courseId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to check course completion for employee {employeeId}, course {courseId}", ex)
            End Try
        End Function

        Public Function ApproveTrainingRecord(recordId As Integer, approvedBy As String) As Boolean
            Try
                If recordId <= 0 Then
                    Throw New ArgumentException("Record ID must be greater than 0", "recordId")
                End If

                If String.IsNullOrWhiteSpace(approvedBy) Then
                    Throw New ArgumentException("Approved by cannot be empty", "approvedBy")
                End If

                ' Get existing record
                Dim record = _repository.GetTrainingRecordById(recordId)
                If record Is Nothing Then
                    Throw New InvalidOperationException($"Training record with ID {recordId} does not exist")
                End If

                ' Check if approval is required
                If Not record.ApprovalRequired Then
                    Throw New InvalidOperationException($"Training record {recordId} does not require approval")
                End If

                ' Check if already approved
                If record.IsApproved Then
                    Throw New InvalidOperationException($"Training record {recordId} has already been approved")
                End If

                ' Set approval fields
                record.ApprovedBy = approvedBy
                record.ApprovedDate = DateTime.Now
                record.ModifiedBy = approvedBy
                record.ModifiedDate = DateTime.Now

                ' Update the record
                Dim success As Boolean = _repository.UpdateTrainingRecord(record)

                If success Then
                    EventLog.WriteEntry("TrainTrack", $"Training record approved: ID {recordId} by {approvedBy}", EventLogEntryType.Information)
                End If

                Return success
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error approving training record {recordId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to approve training record {recordId}", ex)
            End Try
        End Function

        Public Function RejectTrainingRecord(recordId As Integer, rejectedReason As String, rejectedBy As String) As Boolean
            Try
                If recordId <= 0 Then
                    Throw New ArgumentException("Record ID must be greater than 0", "recordId")
                End If

                If String.IsNullOrWhiteSpace(rejectedReason) Then
                    Throw New ArgumentException("Rejection reason cannot be empty", "rejectedReason")
                End If

                If String.IsNullOrWhiteSpace(rejectedBy) Then
                    Throw New ArgumentException("Rejected by cannot be empty", "rejectedBy")
                End If

                ' Get existing record
                Dim record = _repository.GetTrainingRecordById(recordId)
                If record Is Nothing Then
                    Throw New InvalidOperationException($"Training record with ID {recordId} does not exist")
                End If

                ' Set rejection fields
                record.RejectedReason = rejectedReason
                record.ModifiedBy = rejectedBy
                record.ModifiedDate = DateTime.Now

                ' Update the record
                Dim success As Boolean = _repository.UpdateTrainingRecord(record)

                If success Then
                    EventLog.WriteEntry("TrainTrack", $"Training record rejected: ID {recordId} by {rejectedBy}. Reason: {rejectedReason}", EventLogEntryType.Information)
                End If

                Return success
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error rejecting training record {recordId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to reject training record {recordId}", ex)
            End Try
        End Function

        ' ===========================
        ' Training Session Methods
        ' ===========================

        Public Function GetAllTrainingSessions() As List(Of TrainingSession)
            Try
                Return _repository.GetAllTrainingSessions()
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting all training sessions: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve training session list", ex)
            End Try
        End Function

        Public Function GetTrainingSessionById(sessionId As Integer) As TrainingSession
            Try
                If sessionId <= 0 Then
                    Throw New ArgumentException("Session ID must be greater than 0", "sessionId")
                End If

                Return _repository.GetTrainingSessionById(sessionId)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting training session by ID {sessionId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve training session with ID {sessionId}", ex)
            End Try
        End Function

        Public Function GetTrainingSessionsByCourse(courseId As Integer) As List(Of TrainingSession)
            Try
                If courseId <= 0 Then
                    Throw New ArgumentException("Course ID must be greater than 0", "courseId")
                End If

                Return _repository.GetTrainingSessionsByCourse(courseId)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting training sessions for course {courseId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve training sessions for course {courseId}", ex)
            End Try
        End Function

        Public Function GetUpcomingTrainingSessions() As List(Of TrainingSession)
            Try
                Return _repository.GetUpcomingTrainingSessions()
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting upcoming training sessions: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve upcoming training sessions", ex)
            End Try
        End Function

        Public Function GetTrainingSessionsByInstructor(instructorId As Integer) As List(Of TrainingSession)
            Try
                If instructorId <= 0 Then
                    Throw New ArgumentException("Instructor ID must be greater than 0", "instructorId")
                End If

                Return _repository.GetTrainingSessionsByInstructor(instructorId)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting training sessions for instructor {instructorId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve training sessions for instructor {instructorId}", ex)
            End Try
        End Function

        Public Function CreateTrainingSession(session As TrainingSession, createdBy As String) As Integer
            Try
                If session Is Nothing Then
                    Throw New ArgumentNullException("session", "Training session cannot be null")
                End If

                If String.IsNullOrWhiteSpace(createdBy) Then
                    Throw New ArgumentException("Created by cannot be empty", "createdBy")
                End If

                ' Validate session data
                Dim validationErrors = session.GetValidationErrors()
                If validationErrors.Count > 0 Then
                    Throw New ArgumentException($"Training session validation failed: {String.Join("; ", validationErrors)}")
                End If

                ' Check if course exists and is active
                Using courseManager As New CourseManager()
                    Dim course = courseManager.GetCourseById(session.CourseId)
                    If course Is Nothing Then
                        Throw New InvalidOperationException($"Course with ID {session.CourseId} does not exist")
                    End If

                    If Not course.IsActive Then
                        Throw New InvalidOperationException($"Course with ID {session.CourseId} is not active")
                    End If
                End Using

                ' Check for instructor scheduling conflicts
                Dim instructorSessions = _repository.GetTrainingSessionsByInstructor(session.PrimaryInstructorId)
                For Each existingSession In instructorSessions
                    If existingSession.SessionDate.Date = session.SessionDate.Date AndAlso
                       existingSession.SessionStatus <> "CANCELLED" Then
                        ' Check for time overlap
                        If session.StartTime < existingSession.EndTime AndAlso session.EndTime > existingSession.StartTime Then
                            Throw New InvalidOperationException($"Instructor {session.PrimaryInstructorId} has a scheduling conflict on {session.SessionDate:dd/MM/yyyy}")
                        End If
                    End If
                Next

                ' Set audit fields
                session.CreatedBy = createdBy
                session.CreatedDate = DateTime.Now

                ' Create training session
                Dim newSessionId As Integer = _repository.CreateTrainingSession(session)

                ' Log the action
                EventLog.WriteEntry("TrainTrack", $"Training session created: ID {newSessionId} for course {session.CourseId} by {createdBy}", EventLogEntryType.Information)

                Return newSessionId
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error creating training session: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to create training session", ex)
            End Try
        End Function

        Public Function UpdateTrainingSession(session As TrainingSession, modifiedBy As String) As Boolean
            Try
                If session Is Nothing Then
                    Throw New ArgumentNullException("session", "Training session cannot be null")
                End If

                If session.SessionId <= 0 Then
                    Throw New ArgumentException("Session ID must be greater than 0", "session")
                End If

                If String.IsNullOrWhiteSpace(modifiedBy) Then
                    Throw New ArgumentException("Modified by cannot be empty", "modifiedBy")
                End If

                ' Validate session data
                Dim validationErrors = session.GetValidationErrors()
                If validationErrors.Count > 0 Then
                    Throw New ArgumentException($"Training session validation failed: {String.Join("; ", validationErrors)}")
                End If

                ' Check if session exists
                Dim existingSession = _repository.GetTrainingSessionById(session.SessionId)
                If existingSession Is Nothing Then
                    Throw New InvalidOperationException($"Training session with ID {session.SessionId} does not exist")
                End If

                ' Don't allow updates to cancelled sessions
                If existingSession.IsCancelled Then
                    Throw New InvalidOperationException($"Cannot update cancelled training session {session.SessionId}")
                End If

                ' Set audit fields
                session.ModifiedBy = modifiedBy
                session.ModifiedDate = DateTime.Now

                ' Update training session
                Dim success As Boolean = _repository.UpdateTrainingSession(session)

                If success Then
                    ' Log the action
                    EventLog.WriteEntry("TrainTrack", $"Training session updated: ID {session.SessionId} by {modifiedBy}", EventLogEntryType.Information)
                End If

                Return success
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error updating training session: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to update training session", ex)
            End Try
        End Function

        Public Function CancelTrainingSession(sessionId As Integer, cancelledReason As String, cancelledBy As String) As Boolean
            Try
                If sessionId <= 0 Then
                    Throw New ArgumentException("Session ID must be greater than 0", "sessionId")
                End If

                If String.IsNullOrWhiteSpace(cancelledReason) Then
                    Throw New ArgumentException("Cancellation reason cannot be empty", "cancelledReason")
                End If

                If String.IsNullOrWhiteSpace(cancelledBy) Then
                    Throw New ArgumentException("Cancelled by cannot be empty", "cancelledBy")
                End If

                ' Get session to check if exists
                Dim session = _repository.GetTrainingSessionById(sessionId)
                If session Is Nothing Then
                    Throw New InvalidOperationException($"Training session with ID {sessionId} does not exist")
                End If

                ' Don't allow cancellation of already cancelled sessions
                If session.IsCancelled Then
                    Throw New InvalidOperationException($"Training session {sessionId} is already cancelled")
                End If

                ' Don't allow cancellation of completed sessions
                If session.SessionStatus = "COMPLETED" Then
                    Throw New InvalidOperationException($"Cannot cancel completed training session {sessionId}")
                End If

                ' Cancel the session
                Dim success As Boolean = _repository.CancelTrainingSession(sessionId, cancelledReason, cancelledBy)

                If success Then
                    ' Log the action
                    EventLog.WriteEntry("TrainTrack", $"Training session cancelled: ID {sessionId} by {cancelledBy}. Reason: {cancelledReason}", EventLogEntryType.Information)
                End If

                Return success
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error cancelling training session {sessionId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to cancel training session {sessionId}", ex)
            End Try
        End Function

        ' ===========================
        ' Enrollment Methods
        ' ===========================

        Public Function EnrollParticipant(sessionId As Integer, employeeId As Integer, enrolledBy As String) As Boolean
            Try
                If sessionId <= 0 Then
                    Throw New ArgumentException("Session ID must be greater than 0", "sessionId")
                End If

                If employeeId <= 0 Then
                    Throw New ArgumentException("Employee ID must be greater than 0", "employeeId")
                End If

                If String.IsNullOrWhiteSpace(enrolledBy) Then
                    Throw New ArgumentException("Enrolled by cannot be empty", "enrolledBy")
                End If

                ' Get session to check availability
                Dim session = _repository.GetTrainingSessionById(sessionId)
                If session Is Nothing Then
                    Throw New InvalidOperationException($"Training session with ID {sessionId} does not exist")
                End If

                ' Check if session is open for registration
                If Not session.IsRegistrationOpen Then
                    Throw New InvalidOperationException($"Training session {sessionId} is not open for registration")
                End If

                ' Check if session is fully booked
                If session.IsFullyBooked Then
                    Throw New InvalidOperationException($"Training session {sessionId} is fully booked ({session.MaxParticipants} participants)")
                End If

                ' Check if session is cancelled
                If session.IsCancelled Then
                    Throw New InvalidOperationException($"Cannot enroll in cancelled training session {sessionId}")
                End If

                ' Check if employee exists and is active
                Using employeeManager As New EmployeeManager()
                    Dim employee = employeeManager.GetEmployeeById(employeeId)
                    If employee Is Nothing Then
                        Throw New InvalidOperationException($"Employee with ID {employeeId} does not exist")
                    End If

                    If Not employee.IsActive Then
                        Throw New InvalidOperationException($"Employee with ID {employeeId} is not active")
                    End If
                End Using

                ' Check if employee has completed course prerequisites
                Using courseManager As New CourseManager()
                    If Not courseManager.CanEmployeeTakeCourse(employeeId, session.CourseId) Then
                        Throw New InvalidOperationException($"Employee {employeeId} has not met the prerequisites for course {session.CourseId}")
                    End If
                End Using

                ' Check if employee is already enrolled in this session
                Dim existingRecords = _repository.GetTrainingRecordsBySession(sessionId)
                For Each existingRecord In existingRecords
                    If existingRecord.EmployeeId = employeeId AndAlso
                       existingRecord.AttendanceStatus = "ENROLLED" Then
                        Throw New InvalidOperationException($"Employee {employeeId} is already enrolled in session {sessionId}")
                    End If
                Next

                ' Enroll the participant
                Dim success As Boolean = _repository.EnrollParticipant(sessionId, employeeId, enrolledBy)

                If success Then
                    EventLog.WriteEntry("TrainTrack", $"Participant enrolled: employee {employeeId} in session {sessionId} by {enrolledBy}", EventLogEntryType.Information)
                End If

                Return success
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error enrolling employee {employeeId} in session {sessionId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to enroll employee {employeeId} in session {sessionId}", ex)
            End Try
        End Function

        Public Function WithdrawParticipant(sessionId As Integer, employeeId As Integer, withdrawnBy As String) As Boolean
            Try
                If sessionId <= 0 Then
                    Throw New ArgumentException("Session ID must be greater than 0", "sessionId")
                End If

                If employeeId <= 0 Then
                    Throw New ArgumentException("Employee ID must be greater than 0", "employeeId")
                End If

                If String.IsNullOrWhiteSpace(withdrawnBy) Then
                    Throw New ArgumentException("Withdrawn by cannot be empty", "withdrawnBy")
                End If

                ' Get session to check status
                Dim session = _repository.GetTrainingSessionById(sessionId)
                If session Is Nothing Then
                    Throw New InvalidOperationException($"Training session with ID {sessionId} does not exist")
                End If

                ' Don't allow withdrawal from completed sessions
                If session.SessionStatus = "COMPLETED" Then
                    Throw New InvalidOperationException($"Cannot withdraw from completed training session {sessionId}")
                End If

                ' Withdraw the participant
                Dim success As Boolean = _repository.WithdrawParticipant(sessionId, employeeId)

                If success Then
                    EventLog.WriteEntry("TrainTrack", $"Participant withdrawn: employee {employeeId} from session {sessionId} by {withdrawnBy}", EventLogEntryType.Information)
                End If

                Return success
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error withdrawing employee {employeeId} from session {sessionId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to withdraw employee {employeeId} from session {sessionId}", ex)
            End Try
        End Function

        Public Function GetSessionParticipants(sessionId As Integer) As List(Of Employee)
            Try
                If sessionId <= 0 Then
                    Throw New ArgumentException("Session ID must be greater than 0", "sessionId")
                End If

                Return _repository.GetSessionParticipants(sessionId)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting participants for session {sessionId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve participants for session {sessionId}", ex)
            End Try
        End Function

        ' ===========================
        ' Compliance Methods
        ' ===========================

        Public Function GetExpiredTrainingRecords() As List(Of TrainingRecord)
            Try
                Return _repository.GetExpiredTrainingRecords()
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting expired training records: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve expired training records", ex)
            End Try
        End Function

        Public Function GetRecordsNearingExpiry(days As Integer) As List(Of TrainingRecord)
            Try
                If days <= 0 Then
                    days = 90 ' Default to 90 days
                End If

                Return _repository.GetRecordsNearingExpiry(days)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting training records nearing expiry: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve training records nearing expiry", ex)
            End Try
        End Function

        Public Function GetEmployeeComplianceStatus(employeeId As Integer) As Dictionary(Of String, Object)
            Try
                If employeeId <= 0 Then
                    Throw New ArgumentException("Employee ID must be greater than 0", "employeeId")
                End If

                Dim complianceInfo As New Dictionary(Of String, Object)()

                ' Get all training records for the employee
                Dim records = _repository.GetTrainingRecordsByEmployee(employeeId)
                Dim completedCount As Integer = 0
                Dim expiredCount As Integer = 0
                Dim nearExpiryCount As Integer = 0
                Dim pendingCount As Integer = 0
                Dim failedCount As Integer = 0

                For Each record In records
                    Select Case record.CompletionStatus.ToUpper()
                        Case "COMPLETED"
                            completedCount += 1
                            If record.CalculatedIsExpired Then
                                expiredCount += 1
                            ElseIf record.IsNearExpiry Then
                                nearExpiryCount += 1
                            End If
                        Case "NOT_STARTED", "IN_PROGRESS"
                            pendingCount += 1
                        Case "FAILED"
                            failedCount += 1
                    End Select
                Next

                ' Get compulsory courses to check compliance
                Dim compulsoryCourseCount As Integer = 0
                Dim compulsoryCompletedCount As Integer = 0
                Using courseManager As New CourseManager()
                    Dim compulsoryCourses = courseManager.GetCompulsoryCourses()
                    compulsoryCourseCount = compulsoryCourses.Count

                    For Each course In compulsoryCourses
                        If _repository.HasEmployeeCompletedCourse(employeeId, course.CourseId) Then
                            compulsoryCompletedCount += 1
                        End If
                    Next
                End Using

                complianceInfo("TotalRecords") = records.Count
                complianceInfo("CompletedCount") = completedCount
                complianceInfo("ExpiredCount") = expiredCount
                complianceInfo("NearExpiryCount") = nearExpiryCount
                complianceInfo("PendingCount") = pendingCount
                complianceInfo("FailedCount") = failedCount
                complianceInfo("CompulsoryCourseCount") = compulsoryCourseCount
                complianceInfo("CompulsoryCompletedCount") = compulsoryCompletedCount

                ' Calculate compliance rate
                If compulsoryCourseCount > 0 Then
                    complianceInfo("ComplianceRate") = CDec(compulsoryCompletedCount) / CDec(compulsoryCourseCount) * 100D
                Else
                    complianceInfo("ComplianceRate") = 100D ' No compulsory courses means 100% compliant
                End If

                ' Determine overall compliance status
                If compulsoryCourseCount > 0 AndAlso compulsoryCompletedCount < compulsoryCourseCount Then
                    complianceInfo("OverallStatus") = "Non-Compliant"
                ElseIf expiredCount > 0 Then
                    complianceInfo("OverallStatus") = "Expired Training"
                ElseIf nearExpiryCount > 0 Then
                    complianceInfo("OverallStatus") = "Renewal Required"
                Else
                    complianceInfo("OverallStatus") = "Compliant"
                End If

                Return complianceInfo
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting compliance status for employee {employeeId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve compliance status for employee {employeeId}", ex)
            End Try
        End Function

        Public Function GetDepartmentComplianceReport(department As String) As Dictionary(Of String, Object)
            Try
                If String.IsNullOrWhiteSpace(department) Then
                    Throw New ArgumentException("Department cannot be empty", "department")
                End If

                Dim reportData As New Dictionary(Of String, Object)()
                Dim compliantCount As Integer = 0
                Dim nonCompliantCount As Integer = 0
                Dim totalEmployees As Integer = 0

                ' Get all employees in the department
                Using employeeManager As New EmployeeManager()
                    Dim employees = employeeManager.GetEmployeesByDepartment(department)
                    totalEmployees = employees.Count

                    For Each employee In employees
                        Dim complianceStatus = GetEmployeeComplianceStatus(employee.EmployeeId)
                        Dim overallStatus As String = CStr(complianceStatus("OverallStatus"))

                        If overallStatus = "Compliant" Then
                            compliantCount += 1
                        Else
                            nonCompliantCount += 1
                        End If
                    Next
                End Using

                reportData("Department") = department
                reportData("TotalEmployees") = totalEmployees
                reportData("CompliantCount") = compliantCount
                reportData("NonCompliantCount") = nonCompliantCount

                If totalEmployees > 0 Then
                    reportData("ComplianceRate") = CDec(compliantCount) / CDec(totalEmployees) * 100D
                Else
                    reportData("ComplianceRate") = 0D
                End If

                Return reportData
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting department compliance report for {department}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve compliance report for department {department}", ex)
            End Try
        End Function

        ' ===========================
        ' Reporting Methods
        ' ===========================

        Public Function GetDashboardMetrics() As Dictionary(Of String, Object)
            Try
                Return _repository.GetDashboardMetrics()
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting dashboard metrics: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve dashboard metrics", ex)
            End Try
        End Function

        Public Function GetTrainingHistory(employeeId As Integer) As List(Of TrainingRecord)
            Try
                If employeeId <= 0 Then
                    Throw New ArgumentException("Employee ID must be greater than 0", "employeeId")
                End If

                Return _repository.GetTrainingRecordsByEmployee(employeeId)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting training history for employee {employeeId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve training history for employee {employeeId}", ex)
            End Try
        End Function

        Public Function GetTrainingCompletionSummary(fromDate As DateTime?, toDate As DateTime?) As Dictionary(Of String, Object)
            Try
                Dim summary As New Dictionary(Of String, Object)()

                ' Get all training records
                Dim allRecords = _repository.GetAllTrainingRecords()

                ' Apply date filter if specified
                Dim filteredRecords As List(Of TrainingRecord) = allRecords
                If fromDate.HasValue Then
                    filteredRecords = filteredRecords.FindAll(Function(r) r.EnrollmentDate >= fromDate.Value)
                End If
                If toDate.HasValue Then
                    filteredRecords = filteredRecords.FindAll(Function(r) r.EnrollmentDate <= toDate.Value)
                End If

                Dim totalRecords As Integer = filteredRecords.Count
                Dim completedRecords As Integer = filteredRecords.FindAll(Function(r) r.CompletionStatus = "COMPLETED").Count
                Dim failedRecords As Integer = filteredRecords.FindAll(Function(r) r.CompletionStatus = "FAILED").Count
                Dim inProgressRecords As Integer = filteredRecords.FindAll(Function(r) r.CompletionStatus = "IN_PROGRESS").Count
                Dim notStartedRecords As Integer = filteredRecords.FindAll(Function(r) r.CompletionStatus = "NOT_STARTED").Count
                Dim noShowRecords As Integer = filteredRecords.FindAll(Function(r) r.AttendanceStatus = "NO_SHOW").Count

                summary("TotalRecords") = totalRecords
                summary("CompletedRecords") = completedRecords
                summary("FailedRecords") = failedRecords
                summary("InProgressRecords") = inProgressRecords
                summary("NotStartedRecords") = notStartedRecords
                summary("NoShowRecords") = noShowRecords

                If totalRecords > 0 Then
                    summary("CompletionRate") = CDec(completedRecords) / CDec(totalRecords) * 100D
                    summary("PassRate") = If(completedRecords + failedRecords > 0,
                                            CDec(completedRecords) / CDec(completedRecords + failedRecords) * 100D,
                                            0D)
                    summary("NoShowRate") = CDec(noShowRecords) / CDec(totalRecords) * 100D
                Else
                    summary("CompletionRate") = 0D
                    summary("PassRate") = 0D
                    summary("NoShowRate") = 0D
                End If

                Return summary
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting training completion summary: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve training completion summary", ex)
            End Try
        End Function

        Public Function GetSessionReadinessReport() As List(Of Dictionary(Of String, Object))
            Try
                Dim report As New List(Of Dictionary(Of String, Object))()

                ' Get upcoming sessions
                Dim upcomingSessions = _repository.GetUpcomingTrainingSessions()

                For Each session In upcomingSessions
                    Dim sessionReport As New Dictionary(Of String, Object)()
                    sessionReport("SessionId") = session.SessionId
                    sessionReport("CourseId") = session.CourseId
                    sessionReport("SessionDate") = session.SessionDate
                    sessionReport("Location") = session.Location
                    sessionReport("ReadinessScore") = session.ReadinessScore
                    sessionReport("ReadinessStatus") = session.ReadinessStatus
                    sessionReport("CurrentParticipants") = session.CurrentParticipants
                    sessionReport("MaxParticipants") = session.MaxParticipants
                    sessionReport("AvailableSpaces") = session.AvailableSpaces
                    sessionReport("DaysUntilSession") = session.DaysUntilSession
                    sessionReport("PreparationTasks") = session.GetPreparationTasks()
                    report.Add(sessionReport)
                Next

                Return report
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting session readiness report: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve session readiness report", ex)
            End Try
        End Function

        Public Function GetRenewalNotificationList(days As Integer) As List(Of TrainingRecord)
            Try
                If days <= 0 Then
                    days = 90 ' Default to 90 days
                End If

                ' Get records nearing expiry that haven't had notification sent
                Dim nearExpiryRecords = _repository.GetRecordsNearingExpiry(days)
                Dim notificationList As New List(Of TrainingRecord)()

                For Each record In nearExpiryRecords
                    If record.RenewalRequired AndAlso Not record.RenewalNotificationSent Then
                        notificationList.Add(record)
                    End If
                Next

                Return notificationList
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting renewal notification list: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve renewal notification list", ex)
            End Try
        End Function

        Public Function MarkRenewalNotificationSent(recordId As Integer, modifiedBy As String) As Boolean
            Try
                If recordId <= 0 Then
                    Throw New ArgumentException("Record ID must be greater than 0", "recordId")
                End If

                If String.IsNullOrWhiteSpace(modifiedBy) Then
                    Throw New ArgumentException("Modified by cannot be empty", "modifiedBy")
                End If

                ' Get existing record
                Dim record = _repository.GetTrainingRecordById(recordId)
                If record Is Nothing Then
                    Throw New InvalidOperationException($"Training record with ID {recordId} does not exist")
                End If

                ' Mark notification as sent
                record.RenewalNotificationSent = True
                record.ModifiedBy = modifiedBy
                record.ModifiedDate = DateTime.Now

                ' Update the record
                Dim success As Boolean = _repository.UpdateTrainingRecord(record)

                If success Then
                    EventLog.WriteEntry("TrainTrack", $"Renewal notification marked as sent for record {recordId} by {modifiedBy}", EventLogEntryType.Information)
                End If

                Return success
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error marking renewal notification sent for record {recordId}: {ex.Message}", EventLogEntryType.Error)
                ' Don't throw exception for notification tracking - it's not critical
                Return False
            End Try
        End Function

        ' ===========================
        ' Waiting List Methods
        ' ===========================

        Public Function AddToWaitingList(sessionId As Integer, employeeId As Integer, createdBy As String) As Integer
            Try
                If sessionId <= 0 Then
                    Throw New ArgumentException("Session ID must be greater than 0", "sessionId")
                End If

                If employeeId <= 0 Then
                    Throw New ArgumentException("Employee ID must be greater than 0", "employeeId")
                End If

                If String.IsNullOrWhiteSpace(createdBy) Then
                    Throw New ArgumentException("Created by cannot be empty", "createdBy")
                End If

                ' Check session exists and is full
                Dim session = _repository.GetTrainingSessionById(sessionId)
                If session Is Nothing Then
                    Throw New InvalidOperationException($"Training session with ID {sessionId} does not exist")
                End If

                If Not session.IsFullyBooked Then
                    Throw New InvalidOperationException($"Training session {sessionId} is not full. Employee should be enrolled directly instead of added to waiting list.")
                End If

                ' Check employee exists and is active
                Using employeeManager As New EmployeeManager()
                    Dim employee = employeeManager.GetEmployeeById(employeeId)
                    If employee Is Nothing Then
                        Throw New InvalidOperationException($"Employee with ID {employeeId} does not exist")
                    End If

                    If Not employee.IsActive Then
                        Throw New InvalidOperationException($"Employee with ID {employeeId} is not active")
                    End If
                End Using

                ' Check employee is not already enrolled in this session
                Dim existingRecords = _repository.GetTrainingRecordsBySession(sessionId)
                For Each existingRecord In existingRecords
                    If existingRecord.EmployeeId = employeeId AndAlso
                       existingRecord.AttendanceStatus = "ENROLLED" Then
                        Throw New InvalidOperationException($"Employee {employeeId} is already enrolled in session {sessionId}")
                    End If
                Next

                ' Check employee is not already on the waiting list for this session
                Dim waitingList = _waitingListRepository.GetWaitingListBySession(sessionId)
                For Each entry In waitingList
                    If entry.EmployeeId = employeeId AndAlso entry.IsActive Then
                        Throw New InvalidOperationException($"Employee {employeeId} is already on the waiting list for session {sessionId}")
                    End If
                Next

                ' Auto-calculate position as max(position) + 1 for that session
                Dim nextPosition As Integer = 1
                If waitingList.Count > 0 Then
                    For Each entry In waitingList
                        If entry.Position >= nextPosition Then
                            nextPosition = entry.Position + 1
                        End If
                    Next
                End If

                ' Create waiting list entry
                Dim newEntry As New WaitingListEntry(sessionId, employeeId, createdBy)
                newEntry.Position = nextPosition

                Dim entryId As Integer = _waitingListRepository.AddToWaitingList(newEntry)

                If entryId > 0 Then
                    EventLog.WriteEntry("TrainTrack", $"Employee {employeeId} added to waiting list for session {sessionId} at position {nextPosition} by {createdBy}", EventLogEntryType.Information)
                End If

                Return entryId
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error adding employee {employeeId} to waiting list for session {sessionId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to add employee {employeeId} to waiting list for session {sessionId}", ex)
            End Try
        End Function

        Public Function RemoveFromWaitingList(entryId As Integer) As Boolean
            Try
                If entryId <= 0 Then
                    Throw New ArgumentException("Entry ID must be greater than 0", "entryId")
                End If

                Dim success As Boolean = _waitingListRepository.RemoveFromWaitingList(entryId)

                If success Then
                    EventLog.WriteEntry("TrainTrack", $"Waiting list entry {entryId} removed", EventLogEntryType.Information)
                End If

                Return success
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error removing waiting list entry {entryId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to remove waiting list entry {entryId}", ex)
            End Try
        End Function

        Public Function PromoteFromWaitingList(sessionId As Integer, promotedBy As String) As Boolean
            Try
                If sessionId <= 0 Then
                    Throw New ArgumentException("Session ID must be greater than 0", "sessionId")
                End If

                If String.IsNullOrWhiteSpace(promotedBy) Then
                    Throw New ArgumentException("Promoted by cannot be empty", "promotedBy")
                End If

                ' Check session exists
                Dim session = _repository.GetTrainingSessionById(sessionId)
                If session Is Nothing Then
                    Throw New InvalidOperationException($"Training session with ID {sessionId} does not exist")
                End If

                ' Get next in line
                Dim nextEntry As WaitingListEntry = _waitingListRepository.GetNextInLine(sessionId)
                If nextEntry Is Nothing Then
                    Throw New InvalidOperationException($"No one is waiting for session {sessionId}")
                End If

                ' Change status to "Offered" and set response deadline (3 business days)
                nextEntry.Status = "Offered"
                nextEntry.OfferedDate = DateTime.Now
                nextEntry.ResponseDeadline = CalculateBusinessDayDeadline(DateTime.Now, 3)

                Dim success As Boolean = _waitingListRepository.UpdateWaitingListEntry(nextEntry)

                If success Then
                    EventLog.WriteEntry("TrainTrack", $"Waiting list entry {nextEntry.EntryId} promoted (offered) for session {sessionId} by {promotedBy}. Employee {nextEntry.EmployeeId} has until {nextEntry.ResponseDeadline:dd/MM/yyyy} to respond.", EventLogEntryType.Information)
                End If

                Return success
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error promoting from waiting list for session {sessionId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to promote from waiting list for session {sessionId}", ex)
            End Try
        End Function

        Public Function GetSessionWaitingList(sessionId As Integer) As List(Of WaitingListEntry)
            Try
                If sessionId <= 0 Then
                    Throw New ArgumentException("Session ID must be greater than 0", "sessionId")
                End If

                Return _waitingListRepository.GetWaitingListBySession(sessionId)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting waiting list for session {sessionId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve waiting list for session {sessionId}", ex)
            End Try
        End Function

        Public Function GetAllWaitingListOverview() As List(Of WaitingListEntry)
            Try
                Return _waitingListRepository.GetAllActiveWaitingListEntries()
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting waiting list overview: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve waiting list overview", ex)
            End Try
        End Function

        Public Function GetWaitingListCountBySession(sessionId As Integer) As Integer
            Try
                If sessionId <= 0 Then
                    Throw New ArgumentException("Session ID must be greater than 0", "sessionId")
                End If

                Dim waitingList = _waitingListRepository.GetWaitingListBySession(sessionId)
                Dim activeCount As Integer = 0
                For Each entry In waitingList
                    If entry.IsActive Then
                        activeCount += 1
                    End If
                Next
                Return activeCount
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting waiting list count for session {sessionId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve waiting list count for session {sessionId}", ex)
            End Try
        End Function

        Private Function CalculateBusinessDayDeadline(startDate As DateTime, businessDays As Integer) As DateTime
            Dim currentDate As DateTime = startDate
            Dim daysAdded As Integer = 0

            While daysAdded < businessDays
                currentDate = currentDate.AddDays(1)
                If currentDate.DayOfWeek <> DayOfWeek.Saturday AndAlso
                   currentDate.DayOfWeek <> DayOfWeek.Sunday Then
                    daysAdded += 1
                End If
            End While

            Return currentDate
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            If _repository IsNot Nothing Then
                _repository.Dispose()
                _repository = Nothing
            End If
            If _waitingListRepository IsNot Nothing Then
                _waitingListRepository.Dispose()
                _waitingListRepository = Nothing
            End If
        End Sub

    End Class
End Namespace

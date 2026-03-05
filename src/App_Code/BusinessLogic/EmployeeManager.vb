Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports Defra.TrainTrack.Models
Imports Defra.TrainTrack.DataAccess

Namespace Defra.TrainTrack.BusinessLogic
    Public Class EmployeeManager
        Implements IDisposable

        Private _repository As EmployeeRepository

        Public Sub New()
            _repository = New EmployeeRepository()
        End Sub

        Public Function GetAllEmployees() As List(Of Employee)
            Try
                Return _repository.GetAllEmployees()
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting all employees: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve employee list", ex)
            End Try
        End Function

        Public Function GetActiveEmployees() As List(Of Employee)
            Try
                Return _repository.GetActiveEmployees()
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting active employees: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve active employee list", ex)
            End Try
        End Function

        Public Function GetEmployeeById(employeeId As Integer) As Employee
            Try
                If employeeId <= 0 Then
                    Throw New ArgumentException("Employee ID must be greater than 0", "employeeId")
                End If

                Return _repository.GetEmployeeById(employeeId)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting employee by ID {employeeId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve employee with ID {employeeId}", ex)
            End Try
        End Function

        Public Function GetEmployeeByNumber(employeeNumber As String) As Employee
            Try
                If String.IsNullOrWhiteSpace(employeeNumber) Then
                    Throw New ArgumentException("Employee number cannot be empty", "employeeNumber")
                End If

                Return _repository.GetEmployeeByNumber(employeeNumber)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting employee by number {employeeNumber}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve employee with number {employeeNumber}", ex)
            End Try
        End Function

        Public Function GetEmployeeByUserName(userName As String) As Employee
            Try
                If String.IsNullOrWhiteSpace(userName) Then
                    Throw New ArgumentException("User name cannot be empty", "userName")
                End If

                ' Clean up domain\username format
                Dim cleanUserName As String = userName
                If userName.Contains("\") Then
                    cleanUserName = userName.Split("\"c)(1)
                End If

                Return _repository.GetEmployeeByUserName(cleanUserName)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting employee by username {userName}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve employee with username {userName}", ex)
            End Try
        End Function

        Public Function SearchEmployees(searchTerm As String, department As String, location As String, isActive As Boolean?) As List(Of Employee)
            Try
                Return _repository.SearchEmployees(searchTerm, department, location, isActive)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error searching employees: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to search employees", ex)
            End Try
        End Function

        Public Function GetEmployeesByDepartment(department As String) As List(Of Employee)
            Try
                If String.IsNullOrWhiteSpace(department) Then
                    Throw New ArgumentException("Department cannot be empty", "department")
                End If

                Return _repository.GetEmployeesByDepartment(department)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting employees by department {department}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve employees for department {department}", ex)
            End Try
        End Function

        Public Function GetEmployeesByManager(managerId As Integer) As List(Of Employee)
            Try
                If managerId <= 0 Then
                    Throw New ArgumentException("Manager ID must be greater than 0", "managerId")
                End If

                Return _repository.GetEmployeesByManager(managerId)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting employees by manager {managerId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve employees for manager {managerId}", ex)
            End Try
        End Function

        Public Function CreateEmployee(employee As Employee, createdBy As String) As Integer
            Try
                If employee Is Nothing Then
                    Throw New ArgumentNullException("employee", "Employee cannot be null")
                End If

                If String.IsNullOrWhiteSpace(createdBy) Then
                    Throw New ArgumentException("Created by cannot be empty", "createdBy")
                End If

                ' Validate employee data
                Dim validationErrors = employee.GetValidationErrors()
                If validationErrors.Count > 0 Then
                    Throw New ArgumentException($"Employee validation failed: {String.Join("; ", validationErrors)}")
                End If

                ' Check for duplicate employee number
                If Not String.IsNullOrEmpty(employee.EmployeeNumber) Then
                    Dim existingEmployee = _repository.GetEmployeeByNumber(employee.EmployeeNumber)
                    If existingEmployee IsNot Nothing Then
                        Throw New InvalidOperationException($"Employee with number {employee.EmployeeNumber} already exists")
                    End If
                End If

                ' Check for duplicate email
                Dim existingEmailEmployee = _repository.GetEmployeeByEmail(employee.Email)
                If existingEmailEmployee IsNot Nothing Then
                    Throw New InvalidOperationException($"Employee with email {employee.Email} already exists")
                End If

                ' Set audit fields
                employee.CreatedBy = createdBy
                employee.CreatedDate = DateTime.Now

                ' Create employee
                Dim newEmployeeId As Integer = _repository.CreateEmployee(employee)

                ' Log the action
                EventLog.WriteEntry("TrainTrack", $"Employee created: {employee.EmployeeNumber} by {createdBy}", EventLogEntryType.Information)

                Return newEmployeeId
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error creating employee: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to create employee", ex)
            End Try
        End Function

        Public Function UpdateEmployee(employee As Employee, modifiedBy As String) As Boolean
            Try
                If employee Is Nothing Then
                    Throw New ArgumentNullException("employee", "Employee cannot be null")
                End If

                If employee.EmployeeId <= 0 Then
                    Throw New ArgumentException("Employee ID must be greater than 0", "employee")
                End If

                If String.IsNullOrWhiteSpace(modifiedBy) Then
                    Throw New ArgumentException("Modified by cannot be empty", "modifiedBy")
                End If

                ' Validate employee data
                Dim validationErrors = employee.GetValidationErrors()
                If validationErrors.Count > 0 Then
                    Throw New ArgumentException($"Employee validation failed: {String.Join("; ", validationErrors)}")
                End If

                ' Check if employee exists
                Dim existingEmployee = _repository.GetEmployeeById(employee.EmployeeId)
                If existingEmployee Is Nothing Then
                    Throw New InvalidOperationException($"Employee with ID {employee.EmployeeId} does not exist")
                End If

                ' Check for duplicate employee number (excluding current employee)
                If Not String.IsNullOrEmpty(employee.EmployeeNumber) AndAlso employee.EmployeeNumber <> existingEmployee.EmployeeNumber Then
                    Dim duplicateEmployee = _repository.GetEmployeeByNumber(employee.EmployeeNumber)
                    If duplicateEmployee IsNot Nothing AndAlso duplicateEmployee.EmployeeId <> employee.EmployeeId Then
                        Throw New InvalidOperationException($"Employee with number {employee.EmployeeNumber} already exists")
                    End If
                End If

                ' Check for duplicate email (excluding current employee)
                If employee.Email <> existingEmployee.Email Then
                    Dim duplicateEmailEmployee = _repository.GetEmployeeByEmail(employee.Email)
                    If duplicateEmailEmployee IsNot Nothing AndAlso duplicateEmailEmployee.EmployeeId <> employee.EmployeeId Then
                        Throw New InvalidOperationException($"Employee with email {employee.Email} already exists")
                    End If
                End If

                ' Set audit fields
                employee.ModifiedBy = modifiedBy
                employee.ModifiedDate = DateTime.Now

                ' Update employee
                Dim success As Boolean = _repository.UpdateEmployee(employee)

                If success Then
                    ' Log the action
                    EventLog.WriteEntry("TrainTrack", $"Employee updated: {employee.EmployeeNumber} by {modifiedBy}", EventLogEntryType.Information)
                End If

                Return success
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error updating employee: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to update employee", ex)
            End Try
        End Function

        Public Function DeactivateEmployee(employeeId As Integer, modifiedBy As String) As Boolean
            Try
                If employeeId <= 0 Then
                    Throw New ArgumentException("Employee ID must be greater than 0", "employeeId")
                End If

                If String.IsNullOrWhiteSpace(modifiedBy) Then
                    Throw New ArgumentException("Modified by cannot be empty", "modifiedBy")
                End If

                ' Get employee to check if exists
                Dim employee = _repository.GetEmployeeById(employeeId)
                If employee Is Nothing Then
                    Throw New InvalidOperationException($"Employee with ID {employeeId} does not exist")
                End If

                ' Check for any active training sessions
                Dim activeTrainingSessions = GetActiveTrainingSessionsForEmployee(employeeId)
                If activeTrainingSessions.Count > 0 Then
                    Throw New InvalidOperationException($"Cannot deactivate employee with active training sessions. Please cancel or transfer {activeTrainingSessions.Count} active session(s) first.")
                End If

                ' Deactivate employee
                Dim success As Boolean = _repository.DeactivateEmployee(employeeId, modifiedBy)

                If success Then
                    ' Log the action
                    EventLog.WriteEntry("TrainTrack", $"Employee deactivated: {employee.EmployeeNumber} by {modifiedBy}", EventLogEntryType.Information)
                End If

                Return success
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error deactivating employee: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to deactivate employee", ex)
            End Try
        End Function

        Public Function ReactivateEmployee(employeeId As Integer, modifiedBy As String) As Boolean
            Try
                If employeeId <= 0 Then
                    Throw New ArgumentException("Employee ID must be greater than 0", "employeeId")
                End If

                If String.IsNullOrWhiteSpace(modifiedBy) Then
                    Throw New ArgumentException("Modified by cannot be empty", "modifiedBy")
                End If

                ' Get employee to check if exists
                Dim employee = _repository.GetEmployeeById(employeeId)
                If employee Is Nothing Then
                    Throw New InvalidOperationException($"Employee with ID {employeeId} does not exist")
                End If

                ' Reactivate employee
                Dim success As Boolean = _repository.ReactivateEmployee(employeeId, modifiedBy)

                If success Then
                    ' Log the action
                    EventLog.WriteEntry("TrainTrack", $"Employee reactivated: {employee.EmployeeNumber} by {modifiedBy}", EventLogEntryType.Information)
                End If

                Return success
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error reactivating employee: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to reactivate employee", ex)
            End Try
        End Function

        Public Function GetActiveTrainingSessionsForEmployee(employeeId As Integer) As List(Of TrainingSession)
            Try
                If employeeId <= 0 Then
                    Throw New ArgumentException("Employee ID must be greater than 0", "employeeId")
                End If

                Return _repository.GetActiveTrainingSessionsForEmployee(employeeId)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting active training sessions for employee {employeeId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve active training sessions for employee {employeeId}", ex)
            End Try
        End Function

        Public Function GetTrainingComplianceStatus(employeeId As Integer) As Dictionary(Of String, Object)
            Try
                If employeeId <= 0 Then
                    Throw New ArgumentException("Employee ID must be greater than 0", "employeeId")
                End If

                Return _repository.GetTrainingComplianceStatus(employeeId)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting training compliance status for employee {employeeId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve training compliance status for employee {employeeId}", ex)
            End Try
        End Function

        Public Function GetDepartments() As List(Of String)
            Try
                Return _repository.GetDepartments()
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting departments: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve departments", ex)
            End Try
        End Function

        Public Function GetLocations() As List(Of String)
            Try
                Return _repository.GetLocations()
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting locations: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve locations", ex)
            End Try
        End Function

        Public Function GetNewStarters(days As Integer) As List(Of Employee)
            Try
                If days <= 0 Then
                    days = 90 ' Default to 90 days
                End If

                Return _repository.GetNewStarters(days)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting new starters: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve new starters", ex)
            End Try
        End Function

        Public Function GetEmployeesWithUpcomingAnniversaries(days As Integer) As List(Of Employee)
            Try
                If days <= 0 Then
                    days = 30 ' Default to 30 days
                End If

                Return _repository.GetEmployeesWithUpcomingAnniversaries(days)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting employees with upcoming anniversaries: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve employees with upcoming anniversaries", ex)
            End Try
        End Function

        Public Function UpdateLastLoginDate(employeeId As Integer) As Boolean
            Try
                If employeeId <= 0 Then
                    Throw New ArgumentException("Employee ID must be greater than 0", "employeeId")
                End If

                Return _repository.UpdateLastLoginDate(employeeId, DateTime.Now)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error updating last login date for employee {employeeId}: {ex.Message}", EventLogEntryType.Warning)
                ' Don't throw exception for login tracking - it's not critical
                Return False
            End Try
        End Function

        Public Function ValidateEmployeePermissions(employeeId As Integer, requiredPermissions As List(Of String)) As Boolean
            Try
                If employeeId <= 0 Then
                    Throw New ArgumentException("Employee ID must be greater than 0", "employeeId")
                End If

                If requiredPermissions Is Nothing OrElse requiredPermissions.Count = 0 Then
                    Return True ' No permissions required
                End If

                Return _repository.ValidateEmployeePermissions(employeeId, requiredPermissions)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error validating employee permissions for employee {employeeId}: {ex.Message}", EventLogEntryType.Error)
                ' For security, deny access if there's an error
                Return False
            End Try
        End Function

        ' ===== Employee Notes Methods =====

        Private Shared ReadOnly ValidNoteTypes As String() = {"General", "Follow-up", "Compliance", "Training"}

        Public Function GetEmployeeNotes(employeeId As Integer) As List(Of EmployeeNote)
            Try
                If employeeId <= 0 Then
                    Throw New ArgumentException("Employee ID must be greater than 0", "employeeId")
                End If

                Return _repository.GetNotesByEmployeeId(employeeId)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting notes for employee {employeeId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve notes for employee {employeeId}", ex)
            End Try
        End Function

        Public Function AddEmployeeNote(employeeId As Integer, noteText As String, noteType As String, createdBy As String) As Integer
            Try
                If employeeId <= 0 Then
                    Throw New ArgumentException("Employee ID must be greater than 0", "employeeId")
                End If

                If String.IsNullOrWhiteSpace(noteText) Then
                    Throw New ArgumentException("Note text cannot be empty", "noteText")
                End If

                If String.IsNullOrWhiteSpace(noteType) Then
                    Throw New ArgumentException("Note type cannot be empty", "noteType")
                End If

                If Not ValidNoteTypes.Contains(noteType) Then
                    Throw New ArgumentException($"Note type must be one of: {String.Join(", ", ValidNoteTypes)}", "noteType")
                End If

                If String.IsNullOrWhiteSpace(createdBy) Then
                    Throw New ArgumentException("Created by cannot be empty", "createdBy")
                End If

                ' Check if employee exists
                Dim employee = _repository.GetEmployeeById(employeeId)
                If employee Is Nothing Then
                    Throw New InvalidOperationException($"Employee with ID {employeeId} does not exist")
                End If

                ' Create the note
                Dim note As New EmployeeNote(noteText, noteType, createdBy)
                note.EmployeeId = employeeId

                Dim newNoteId As Integer = _repository.CreateNote(note)

                ' Log the action
                EventLog.WriteEntry("TrainTrack", $"Note added to employee {employeeId} by {createdBy} (type: {noteType})", EventLogEntryType.Information)

                Return newNoteId
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error adding note to employee {employeeId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to add note to employee {employeeId}", ex)
            End Try
        End Function

        Public Function ResolveEmployeeNote(noteId As Integer, resolvedBy As String) As Boolean
            Try
                If noteId <= 0 Then
                    Throw New ArgumentException("Note ID must be greater than 0", "noteId")
                End If

                If String.IsNullOrWhiteSpace(resolvedBy) Then
                    Throw New ArgumentException("Resolved by cannot be empty", "resolvedBy")
                End If

                Dim success As Boolean = _repository.ResolveNote(noteId, resolvedBy)

                If success Then
                    EventLog.WriteEntry("TrainTrack", $"Note {noteId} resolved by {resolvedBy}", EventLogEntryType.Information)
                End If

                Return success
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error resolving note {noteId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to resolve note {noteId}", ex)
            End Try
        End Function

        Public Function GetAllUnresolvedNotes() As List(Of EmployeeNote)
            Try
                Return _repository.GetUnresolvedNotes()
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting unresolved notes: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve unresolved notes", ex)
            End Try
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            If _repository IsNot Nothing Then
                _repository.Dispose()
                _repository = Nothing
            End If
        End Sub

    End Class
End Namespace
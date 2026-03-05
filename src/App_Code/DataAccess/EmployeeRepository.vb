Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports Defra.TrainTrack.Models

Namespace Defra.TrainTrack.DataAccess
    Public Class EmployeeRepository
        Implements IDisposable

        Private _db As DatabaseHelper

        Public Sub New()
            _db = New DatabaseHelper()
        End Sub

        Public Function GetAllEmployees() As List(Of Employee)
            Dim employees As New List(Of Employee)()
            Try
                Dim sql As String = "SELECT e.*, m.FirstName + ' ' + m.LastName AS ManagerFullName " &
                                    "FROM Employees e " &
                                    "LEFT JOIN Employees m ON e.ManagerId = m.EmployeeId " &
                                    "ORDER BY e.LastName, e.FirstName"

                Using reader As SqlDataReader = _db.ExecuteReader(sql, Nothing)
                    While reader.Read()
                        employees.Add(MapEmployee(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetAllEmployees: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve employees", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return employees
        End Function

        Public Function GetActiveEmployees() As List(Of Employee)
            Dim employees As New List(Of Employee)()
            Try
                Dim sql As String = "SELECT e.*, m.FirstName + ' ' + m.LastName AS ManagerFullName " &
                                    "FROM Employees e " &
                                    "LEFT JOIN Employees m ON e.ManagerId = m.EmployeeId " &
                                    "WHERE e.IsActive = 1 " &
                                    "ORDER BY e.LastName, e.FirstName"

                Using reader As SqlDataReader = _db.ExecuteReader(sql, Nothing)
                    While reader.Read()
                        employees.Add(MapEmployee(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetActiveEmployees: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve active employees", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return employees
        End Function

        Public Function GetEmployeeById(employeeId As Integer) As Employee
            Try
                Dim sql As String = "SELECT e.*, m.FirstName + ' ' + m.LastName AS ManagerFullName " &
                                    "FROM Employees e " &
                                    "LEFT JOIN Employees m ON e.ManagerId = m.EmployeeId " &
                                    "WHERE e.EmployeeId = @EmployeeId"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@EmployeeId", SqlDbType.Int, employeeId)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    If reader.Read() Then
                        Return MapEmployee(reader)
                    End If
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetEmployeeById ({employeeId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve employee with ID {employeeId}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return Nothing
        End Function

        Public Function GetEmployeeByNumber(employeeNumber As String) As Employee
            Try
                Dim sql As String = "SELECT e.*, m.FirstName + ' ' + m.LastName AS ManagerFullName " &
                                    "FROM Employees e " &
                                    "LEFT JOIN Employees m ON e.ManagerId = m.EmployeeId " &
                                    "WHERE e.EmployeeNumber = @EmployeeNumber"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@EmployeeNumber", SqlDbType.NVarChar, 20, employeeNumber)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    If reader.Read() Then
                        Return MapEmployee(reader)
                    End If
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetEmployeeByNumber ({employeeNumber}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve employee with number {employeeNumber}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return Nothing
        End Function

        Public Function GetEmployeeByUserName(userName As String) As Employee
            Try
                Dim sql As String = "SELECT e.*, m.FirstName + ' ' + m.LastName AS ManagerFullName " &
                                    "FROM Employees e " &
                                    "LEFT JOIN Employees m ON e.ManagerId = m.EmployeeId " &
                                    "WHERE e.UserName = @UserName"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@UserName", SqlDbType.NVarChar, 50, userName)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    If reader.Read() Then
                        Return MapEmployee(reader)
                    End If
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetEmployeeByUserName ({userName}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve employee with username {userName}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return Nothing
        End Function

        Public Function GetEmployeeByEmail(email As String) As Employee
            Try
                Dim sql As String = "SELECT e.*, m.FirstName + ' ' + m.LastName AS ManagerFullName " &
                                    "FROM Employees e " &
                                    "LEFT JOIN Employees m ON e.ManagerId = m.EmployeeId " &
                                    "WHERE e.Email = @Email"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@Email", SqlDbType.NVarChar, 100, email)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    If reader.Read() Then
                        Return MapEmployee(reader)
                    End If
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetEmployeeByEmail ({email}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve employee with email {email}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return Nothing
        End Function

        Public Function SearchEmployees(searchTerm As String, department As String, location As String, isActive As Boolean?) As List(Of Employee)
            Dim employees As New List(Of Employee)()
            Try
                Dim conditions As New List(Of String)()
                Dim paramList As New List(Of SqlParameter)()

                If Not String.IsNullOrWhiteSpace(searchTerm) Then
                    conditions.Add("(e.FirstName LIKE @SearchTerm OR e.LastName LIKE @SearchTerm OR e.EmployeeNumber LIKE @SearchTerm OR e.Email LIKE @SearchTerm)")
                    paramList.Add(DatabaseHelper.CreateParameter("@SearchTerm", SqlDbType.NVarChar, 100, "%" & searchTerm & "%"))
                End If

                If Not String.IsNullOrWhiteSpace(department) Then
                    conditions.Add("e.Department = @Department")
                    paramList.Add(DatabaseHelper.CreateParameter("@Department", SqlDbType.NVarChar, 100, department))
                End If

                If Not String.IsNullOrWhiteSpace(location) Then
                    conditions.Add("e.Location = @Location")
                    paramList.Add(DatabaseHelper.CreateParameter("@Location", SqlDbType.NVarChar, 100, location))
                End If

                If isActive.HasValue Then
                    conditions.Add("e.IsActive = @IsActive")
                    paramList.Add(DatabaseHelper.CreateParameter("@IsActive", SqlDbType.Bit, isActive.Value))
                End If

                Dim sql As String = "SELECT e.*, m.FirstName + ' ' + m.LastName AS ManagerFullName " &
                                    "FROM Employees e " &
                                    "LEFT JOIN Employees m ON e.ManagerId = m.EmployeeId" &
                                    DatabaseHelper.BuildWhereClause(conditions) &
                                    " ORDER BY e.LastName, e.FirstName"

                Dim parameters As SqlParameter() = Nothing
                If paramList.Count > 0 Then
                    parameters = paramList.ToArray()
                End If

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    While reader.Read()
                        employees.Add(MapEmployee(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in SearchEmployees: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to search employees", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return employees
        End Function

        Public Function GetEmployeesByDepartment(department As String) As List(Of Employee)
            Dim employees As New List(Of Employee)()
            Try
                Dim sql As String = "SELECT e.*, m.FirstName + ' ' + m.LastName AS ManagerFullName " &
                                    "FROM Employees e " &
                                    "LEFT JOIN Employees m ON e.ManagerId = m.EmployeeId " &
                                    "WHERE e.Department = @Department AND e.IsActive = 1 " &
                                    "ORDER BY e.LastName, e.FirstName"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@Department", SqlDbType.NVarChar, 100, department)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    While reader.Read()
                        employees.Add(MapEmployee(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetEmployeesByDepartment ({department}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve employees for department {department}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return employees
        End Function

        Public Function GetEmployeesByManager(managerId As Integer) As List(Of Employee)
            Dim employees As New List(Of Employee)()
            Try
                Dim sql As String = "SELECT e.*, m.FirstName + ' ' + m.LastName AS ManagerFullName " &
                                    "FROM Employees e " &
                                    "LEFT JOIN Employees m ON e.ManagerId = m.EmployeeId " &
                                    "WHERE e.ManagerId = @ManagerId AND e.IsActive = 1 " &
                                    "ORDER BY e.LastName, e.FirstName"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@ManagerId", SqlDbType.Int, managerId)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    While reader.Read()
                        employees.Add(MapEmployee(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetEmployeesByManager ({managerId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve employees for manager {managerId}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return employees
        End Function

        Public Function CreateEmployee(employee As Employee) As Integer
            Try
                Dim sql As String = "INSERT INTO Employees " &
                                    "(EmployeeNumber, FirstName, LastName, Email, UserName, Department, " &
                                    "Position, Location, ManagerId, ManagerName, HireDate, IsActive, " &
                                    "PhoneNumber, LineManagerEmail, CostCentre, PayBand, ContractType, " &
                                    "WorkingPattern, CreatedDate, CreatedBy) " &
                                    "VALUES " &
                                    "(@EmployeeNumber, @FirstName, @LastName, @Email, @UserName, @Department, " &
                                    "@Position, @Location, @ManagerId, @ManagerName, @HireDate, @IsActive, " &
                                    "@PhoneNumber, @LineManagerEmail, @CostCentre, @PayBand, @ContractType, " &
                                    "@WorkingPattern, @CreatedDate, @CreatedBy); " &
                                    "SELECT SCOPE_IDENTITY();"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@EmployeeNumber", SqlDbType.NVarChar, 20, employee.EmployeeNumber),
                    DatabaseHelper.CreateParameter("@FirstName", SqlDbType.NVarChar, 50, employee.FirstName),
                    DatabaseHelper.CreateParameter("@LastName", SqlDbType.NVarChar, 50, employee.LastName),
                    DatabaseHelper.CreateParameter("@Email", SqlDbType.NVarChar, 100, employee.Email),
                    DatabaseHelper.CreateParameter("@UserName", SqlDbType.NVarChar, 50, employee.UserName),
                    DatabaseHelper.CreateParameter("@Department", SqlDbType.NVarChar, 100, employee.Department),
                    DatabaseHelper.CreateParameter("@Position", SqlDbType.NVarChar, 100, employee.Position),
                    DatabaseHelper.CreateParameter("@Location", SqlDbType.NVarChar, 100, employee.Location),
                    DatabaseHelper.CreateParameter("@ManagerId", SqlDbType.Int, employee.ManagerId),
                    DatabaseHelper.CreateParameter("@ManagerName", SqlDbType.NVarChar, 100, employee.ManagerName),
                    DatabaseHelper.CreateParameter("@HireDate", SqlDbType.DateTime, employee.HireDate),
                    DatabaseHelper.CreateParameter("@IsActive", SqlDbType.Bit, employee.IsActive),
                    DatabaseHelper.CreateParameter("@PhoneNumber", SqlDbType.NVarChar, 20, employee.PhoneNumber),
                    DatabaseHelper.CreateParameter("@LineManagerEmail", SqlDbType.NVarChar, 100, employee.LineManagerEmail),
                    DatabaseHelper.CreateParameter("@CostCentre", SqlDbType.NVarChar, 20, employee.CostCentre),
                    DatabaseHelper.CreateParameter("@PayBand", SqlDbType.NVarChar, 10, employee.PayBand),
                    DatabaseHelper.CreateParameter("@ContractType", SqlDbType.NVarChar, 20, employee.ContractType),
                    DatabaseHelper.CreateParameter("@WorkingPattern", SqlDbType.NVarChar, 50, employee.WorkingPattern),
                    DatabaseHelper.CreateParameter("@CreatedDate", SqlDbType.DateTime, employee.CreatedDate),
                    DatabaseHelper.CreateParameter("@CreatedBy", SqlDbType.NVarChar, 50, employee.CreatedBy)
                }

                Dim result As Object = _db.ExecuteScalar(sql, parameters)
                If result IsNot Nothing AndAlso result IsNot DBNull.Value Then
                    Return Convert.ToInt32(result)
                End If
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in CreateEmployee: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to create employee", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return 0
        End Function

        Public Function UpdateEmployee(employee As Employee) As Boolean
            Try
                Dim sql As String = "UPDATE Employees SET " &
                                    "EmployeeNumber = @EmployeeNumber, " &
                                    "FirstName = @FirstName, " &
                                    "LastName = @LastName, " &
                                    "Email = @Email, " &
                                    "UserName = @UserName, " &
                                    "Department = @Department, " &
                                    "Position = @Position, " &
                                    "Location = @Location, " &
                                    "ManagerId = @ManagerId, " &
                                    "ManagerName = @ManagerName, " &
                                    "HireDate = @HireDate, " &
                                    "IsActive = @IsActive, " &
                                    "PhoneNumber = @PhoneNumber, " &
                                    "LineManagerEmail = @LineManagerEmail, " &
                                    "CostCentre = @CostCentre, " &
                                    "PayBand = @PayBand, " &
                                    "ContractType = @ContractType, " &
                                    "WorkingPattern = @WorkingPattern, " &
                                    "ModifiedDate = @ModifiedDate, " &
                                    "ModifiedBy = @ModifiedBy " &
                                    "WHERE EmployeeId = @EmployeeId"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@EmployeeId", SqlDbType.Int, employee.EmployeeId),
                    DatabaseHelper.CreateParameter("@EmployeeNumber", SqlDbType.NVarChar, 20, employee.EmployeeNumber),
                    DatabaseHelper.CreateParameter("@FirstName", SqlDbType.NVarChar, 50, employee.FirstName),
                    DatabaseHelper.CreateParameter("@LastName", SqlDbType.NVarChar, 50, employee.LastName),
                    DatabaseHelper.CreateParameter("@Email", SqlDbType.NVarChar, 100, employee.Email),
                    DatabaseHelper.CreateParameter("@UserName", SqlDbType.NVarChar, 50, employee.UserName),
                    DatabaseHelper.CreateParameter("@Department", SqlDbType.NVarChar, 100, employee.Department),
                    DatabaseHelper.CreateParameter("@Position", SqlDbType.NVarChar, 100, employee.Position),
                    DatabaseHelper.CreateParameter("@Location", SqlDbType.NVarChar, 100, employee.Location),
                    DatabaseHelper.CreateParameter("@ManagerId", SqlDbType.Int, employee.ManagerId),
                    DatabaseHelper.CreateParameter("@ManagerName", SqlDbType.NVarChar, 100, employee.ManagerName),
                    DatabaseHelper.CreateParameter("@HireDate", SqlDbType.DateTime, employee.HireDate),
                    DatabaseHelper.CreateParameter("@IsActive", SqlDbType.Bit, employee.IsActive),
                    DatabaseHelper.CreateParameter("@PhoneNumber", SqlDbType.NVarChar, 20, employee.PhoneNumber),
                    DatabaseHelper.CreateParameter("@LineManagerEmail", SqlDbType.NVarChar, 100, employee.LineManagerEmail),
                    DatabaseHelper.CreateParameter("@CostCentre", SqlDbType.NVarChar, 20, employee.CostCentre),
                    DatabaseHelper.CreateParameter("@PayBand", SqlDbType.NVarChar, 10, employee.PayBand),
                    DatabaseHelper.CreateParameter("@ContractType", SqlDbType.NVarChar, 20, employee.ContractType),
                    DatabaseHelper.CreateParameter("@WorkingPattern", SqlDbType.NVarChar, 50, employee.WorkingPattern),
                    DatabaseHelper.CreateParameter("@ModifiedDate", SqlDbType.DateTime, employee.ModifiedDate),
                    DatabaseHelper.CreateParameter("@ModifiedBy", SqlDbType.NVarChar, 50, employee.ModifiedBy)
                }

                Dim rowsAffected As Integer = _db.ExecuteNonQuery(sql, parameters)
                Return rowsAffected > 0
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in UpdateEmployee ({employee.EmployeeId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to update employee {employee.EmployeeId}", ex)
            Finally
                _db.CloseConnection()
            End Try
        End Function

        Public Function DeactivateEmployee(employeeId As Integer, modifiedBy As String) As Boolean
            Try
                Dim sql As String = "UPDATE Employees SET IsActive = 0, ModifiedDate = @ModifiedDate, ModifiedBy = @ModifiedBy " &
                                    "WHERE EmployeeId = @EmployeeId"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@EmployeeId", SqlDbType.Int, employeeId),
                    DatabaseHelper.CreateParameter("@ModifiedDate", SqlDbType.DateTime, DateTime.Now),
                    DatabaseHelper.CreateParameter("@ModifiedBy", SqlDbType.NVarChar, 50, modifiedBy)
                }

                Dim rowsAffected As Integer = _db.ExecuteNonQuery(sql, parameters)
                Return rowsAffected > 0
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in DeactivateEmployee ({employeeId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to deactivate employee {employeeId}", ex)
            Finally
                _db.CloseConnection()
            End Try
        End Function

        Public Function ReactivateEmployee(employeeId As Integer, modifiedBy As String) As Boolean
            Try
                Dim sql As String = "UPDATE Employees SET IsActive = 1, ModifiedDate = @ModifiedDate, ModifiedBy = @ModifiedBy " &
                                    "WHERE EmployeeId = @EmployeeId"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@EmployeeId", SqlDbType.Int, employeeId),
                    DatabaseHelper.CreateParameter("@ModifiedDate", SqlDbType.DateTime, DateTime.Now),
                    DatabaseHelper.CreateParameter("@ModifiedBy", SqlDbType.NVarChar, 50, modifiedBy)
                }

                Dim rowsAffected As Integer = _db.ExecuteNonQuery(sql, parameters)
                Return rowsAffected > 0
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in ReactivateEmployee ({employeeId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to reactivate employee {employeeId}", ex)
            Finally
                _db.CloseConnection()
            End Try
        End Function

        Public Function GetActiveTrainingSessionsForEmployee(employeeId As Integer) As List(Of TrainingSession)
            Dim sessions As New List(Of TrainingSession)()
            Try
                Dim sql As String = "SELECT ts.* " &
                                    "FROM TrainingSessions ts " &
                                    "INNER JOIN SessionParticipants sp ON ts.SessionId = sp.SessionId " &
                                    "WHERE sp.EmployeeId = @EmployeeId " &
                                    "AND ts.SessionStatus IN ('SCHEDULED', 'CONFIRMED') " &
                                    "AND ts.SessionDate >= GETDATE() " &
                                    "ORDER BY ts.SessionDate"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@EmployeeId", SqlDbType.Int, employeeId)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    While reader.Read()
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
                        sessions.Add(session)
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetActiveTrainingSessionsForEmployee ({employeeId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve active training sessions for employee {employeeId}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return sessions
        End Function

        Public Function GetTrainingComplianceStatus(employeeId As Integer) As Dictionary(Of String, Object)
            Dim result As New Dictionary(Of String, Object)()
            Try
                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@EmployeeId", SqlDbType.Int, employeeId)
                }

                Dim table As DataTable = _db.ExecuteStoredProcedure("sp_GetEmployeeComplianceStatus", parameters)

                Dim compulsoryTotal As Integer = 0
                Dim compulsoryCompliant As Integer = 0
                Dim compulsoryExpired As Integer = 0
                Dim compulsoryRequired As Integer = 0
                Dim courseStatuses As New List(Of Dictionary(Of String, Object))()

                For Each row As DataRow In table.Rows
                    Dim courseStatus As New Dictionary(Of String, Object)()
                    courseStatus("CourseCode") = row("CourseCode").ToString()
                    courseStatus("Title") = row("Title").ToString()
                    courseStatus("IsCompulsory") = Convert.ToBoolean(row("IsCompulsory"))
                    courseStatus("CompletionStatus") = If(row("CompletionStatus") Is DBNull.Value, "", row("CompletionStatus").ToString())
                    courseStatus("CompletionDate") = If(row("CompletionDate") Is DBNull.Value, Nothing, Convert.ToDateTime(row("CompletionDate")))
                    courseStatus("ExpiryDate") = If(row("ExpiryDate") Is DBNull.Value, Nothing, Convert.ToDateTime(row("ExpiryDate")))
                    courseStatus("ComplianceStatus") = row("ComplianceStatus").ToString()
                    courseStatuses.Add(courseStatus)

                    If Convert.ToBoolean(row("IsCompulsory")) Then
                        compulsoryTotal += 1
                        Select Case row("ComplianceStatus").ToString()
                            Case "COMPLIANT"
                                compulsoryCompliant += 1
                            Case "EXPIRED"
                                compulsoryExpired += 1
                            Case "REQUIRED"
                                compulsoryRequired += 1
                        End Select
                    End If
                Next

                result("CourseStatuses") = courseStatuses
                result("CompulsoryTotal") = compulsoryTotal
                result("CompulsoryCompliant") = compulsoryCompliant
                result("CompulsoryExpired") = compulsoryExpired
                result("CompulsoryRequired") = compulsoryRequired
                result("CompliancePercentage") = If(compulsoryTotal > 0, CDec(compulsoryCompliant) / CDec(compulsoryTotal) * 100, 0D)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetTrainingComplianceStatus ({employeeId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve training compliance status for employee {employeeId}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return result
        End Function

        Public Function GetDepartments() As List(Of String)
            Dim departments As New List(Of String)()
            Try
                Dim sql As String = "SELECT DISTINCT Department FROM Employees WHERE IsActive = 1 ORDER BY Department"

                Using reader As SqlDataReader = _db.ExecuteReader(sql, Nothing)
                    While reader.Read()
                        Dim dept As String = DatabaseHelper.SafeGetString(reader, "Department")
                        If Not String.IsNullOrEmpty(dept) Then
                            departments.Add(dept)
                        End If
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetDepartments: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve departments", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return departments
        End Function

        Public Function GetLocations() As List(Of String)
            Dim locations As New List(Of String)()
            Try
                Dim sql As String = "SELECT DISTINCT Location FROM Employees WHERE IsActive = 1 ORDER BY Location"

                Using reader As SqlDataReader = _db.ExecuteReader(sql, Nothing)
                    While reader.Read()
                        Dim loc As String = DatabaseHelper.SafeGetString(reader, "Location")
                        If Not String.IsNullOrEmpty(loc) Then
                            locations.Add(loc)
                        End If
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetLocations: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve locations", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return locations
        End Function

        Public Function GetNewStarters(days As Integer) As List(Of Employee)
            Dim employees As New List(Of Employee)()
            Try
                Dim sql As String = "SELECT e.*, m.FirstName + ' ' + m.LastName AS ManagerFullName " &
                                    "FROM Employees e " &
                                    "LEFT JOIN Employees m ON e.ManagerId = m.EmployeeId " &
                                    "WHERE e.IsActive = 1 AND e.HireDate >= DATEADD(day, -@Days, GETDATE()) " &
                                    "ORDER BY e.HireDate DESC"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@Days", SqlDbType.Int, days)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    While reader.Read()
                        employees.Add(MapEmployee(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetNewStarters: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve new starters", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return employees
        End Function

        Public Function GetEmployeesWithUpcomingAnniversaries(days As Integer) As List(Of Employee)
            Dim employees As New List(Of Employee)()
            Try
                ' Find employees whose hire date anniversary falls within the next N days
                Dim sql As String = "SELECT e.*, m.FirstName + ' ' + m.LastName AS ManagerFullName " &
                                    "FROM Employees e " &
                                    "LEFT JOIN Employees m ON e.ManagerId = m.EmployeeId " &
                                    "WHERE e.IsActive = 1 " &
                                    "AND DATEDIFF(day, GETDATE(), " &
                                    "    DATEADD(year, DATEDIFF(year, e.HireDate, GETDATE()) + " &
                                    "    CASE WHEN DATEADD(year, DATEDIFF(year, e.HireDate, GETDATE()), e.HireDate) < GETDATE() THEN 1 ELSE 0 END, " &
                                    "    e.HireDate)) BETWEEN 0 AND @Days " &
                                    "ORDER BY DATEADD(year, DATEDIFF(year, e.HireDate, GETDATE()) + " &
                                    "    CASE WHEN DATEADD(year, DATEDIFF(year, e.HireDate, GETDATE()), e.HireDate) < GETDATE() THEN 1 ELSE 0 END, " &
                                    "    e.HireDate)"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@Days", SqlDbType.Int, days)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    While reader.Read()
                        employees.Add(MapEmployee(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetEmployeesWithUpcomingAnniversaries: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve employees with upcoming anniversaries", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return employees
        End Function

        Public Function UpdateLastLoginDate(employeeId As Integer, loginDate As DateTime) As Boolean
            Try
                Dim sql As String = "UPDATE Employees SET LastLoginDate = @LastLoginDate WHERE EmployeeId = @EmployeeId"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@EmployeeId", SqlDbType.Int, employeeId),
                    DatabaseHelper.CreateParameter("@LastLoginDate", SqlDbType.DateTime, loginDate)
                }

                Dim rowsAffected As Integer = _db.ExecuteNonQuery(sql, parameters)
                Return rowsAffected > 0
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in UpdateLastLoginDate ({employeeId}): {ex.Message}", EventLogEntryType.Warning)
                ' Don't throw for login tracking - not critical
                Return False
            Finally
                _db.CloseConnection()
            End Try
        End Function

        Public Function ValidateEmployeePermissions(employeeId As Integer, requiredPermissions As List(Of String)) As Boolean
            Try
                If requiredPermissions Is Nothing OrElse requiredPermissions.Count = 0 Then
                    Return True
                End If

                Dim sql As String = "SELECT COUNT(*) FROM UserPermissions " &
                                    "WHERE EmployeeId = @EmployeeId AND IsActive = 1 AND Permission IN ("

                Dim paramList As New List(Of SqlParameter)()
                paramList.Add(DatabaseHelper.CreateParameter("@EmployeeId", SqlDbType.Int, employeeId))

                For i As Integer = 0 To requiredPermissions.Count - 1
                    If i > 0 Then
                        sql &= ", "
                    End If
                    Dim paramName As String = $"@Perm{i}"
                    sql &= paramName
                    paramList.Add(DatabaseHelper.CreateParameter(paramName, SqlDbType.NVarChar, 100, requiredPermissions(i)))
                Next

                sql &= ")"

                Dim result As Object = _db.ExecuteScalar(sql, paramList.ToArray())
                If result IsNot Nothing AndAlso result IsNot DBNull.Value Then
                    Dim matchCount As Integer = Convert.ToInt32(result)
                    Return matchCount >= requiredPermissions.Count
                End If
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in ValidateEmployeePermissions ({employeeId}): {ex.Message}", EventLogEntryType.Error)
                ' For security, deny access on error
                Return False
            Finally
                _db.CloseConnection()
            End Try
            Return False
        End Function

        Private Function MapEmployee(reader As SqlDataReader) As Employee
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

            ' Try to get the joined manager full name
            Dim managerFullName As String = DatabaseHelper.SafeGetString(reader, "ManagerFullName")
            If Not String.IsNullOrEmpty(managerFullName) Then
                employee.ManagerName = managerFullName
            End If

            Return employee
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            Try
                If _db IsNot Nothing Then
                    _db.Dispose()
                    _db = Nothing
                End If
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error disposing EmployeeRepository: {ex.Message}", EventLogEntryType.Warning)
            End Try
        End Sub

    End Class
End Namespace
Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports Defra.TrainTrack.Models

Namespace Defra.TrainTrack.DataAccess
    Public Class CourseRepository
        Implements IDisposable

        Private _db As DatabaseHelper

        Public Sub New()
            _db = New DatabaseHelper()
        End Sub

        Public Function GetAllCourses() As List(Of Course)
            Dim courses As New List(Of Course)()
            Try
                Dim sql As String = "SELECT * FROM Courses ORDER BY CourseCode"

                Using reader As SqlDataReader = _db.ExecuteReader(sql, Nothing)
                    While reader.Read()
                        courses.Add(MapCourse(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetAllCourses: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve courses", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return courses
        End Function

        Public Function GetActiveCourses() As List(Of Course)
            Dim courses As New List(Of Course)()
            Try
                Dim sql As String = "SELECT * FROM Courses WHERE IsActive = 1 ORDER BY CourseCode"

                Using reader As SqlDataReader = _db.ExecuteReader(sql, Nothing)
                    While reader.Read()
                        courses.Add(MapCourse(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetActiveCourses: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve active courses", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return courses
        End Function

        Public Function GetCourseById(courseId As Integer) As Course
            Try
                Dim sql As String = "SELECT * FROM Courses WHERE CourseId = @CourseId"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@CourseId", SqlDbType.Int, courseId)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    If reader.Read() Then
                        Return MapCourse(reader)
                    End If
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetCourseById ({courseId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve course with ID {courseId}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return Nothing
        End Function

        Public Function GetCourseByCode(courseCode As String) As Course
            Try
                Dim sql As String = "SELECT * FROM Courses WHERE CourseCode = @CourseCode"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@CourseCode", SqlDbType.NVarChar, 20, courseCode)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    If reader.Read() Then
                        Return MapCourse(reader)
                    End If
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetCourseByCode ({courseCode}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve course with code {courseCode}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return Nothing
        End Function

        Public Function SearchCourses(searchTerm As String, category As String, deliveryMethod As String, isActive As Boolean?) As List(Of Course)
            Dim courses As New List(Of Course)()
            Try
                Dim conditions As New List(Of String)()
                Dim paramList As New List(Of SqlParameter)()

                If Not String.IsNullOrWhiteSpace(searchTerm) Then
                    conditions.Add("(CourseCode LIKE @SearchTerm OR Title LIKE @SearchTerm OR Description LIKE @SearchTerm)")
                    paramList.Add(DatabaseHelper.CreateParameter("@SearchTerm", SqlDbType.NVarChar, 200, "%" & searchTerm & "%"))
                End If

                If Not String.IsNullOrWhiteSpace(category) Then
                    conditions.Add("Category = @Category")
                    paramList.Add(DatabaseHelper.CreateParameter("@Category", SqlDbType.NVarChar, 100, category))
                End If

                If Not String.IsNullOrWhiteSpace(deliveryMethod) Then
                    conditions.Add("DeliveryMethod = @DeliveryMethod")
                    paramList.Add(DatabaseHelper.CreateParameter("@DeliveryMethod", SqlDbType.NVarChar, 50, deliveryMethod))
                End If

                If isActive.HasValue Then
                    conditions.Add("IsActive = @IsActive")
                    paramList.Add(DatabaseHelper.CreateParameter("@IsActive", SqlDbType.Bit, isActive.Value))
                End If

                Dim sql As String = "SELECT * FROM Courses" &
                                    DatabaseHelper.BuildWhereClause(conditions) &
                                    " ORDER BY CourseCode"

                Dim parameters As SqlParameter() = Nothing
                If paramList.Count > 0 Then
                    parameters = paramList.ToArray()
                End If

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    While reader.Read()
                        courses.Add(MapCourse(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in SearchCourses: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to search courses", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return courses
        End Function

        Public Function GetCoursesByCategory(category As String) As List(Of Course)
            Dim courses As New List(Of Course)()
            Try
                Dim sql As String = "SELECT * FROM Courses WHERE Category = @Category AND IsActive = 1 ORDER BY CourseCode"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@Category", SqlDbType.NVarChar, 100, category)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    While reader.Read()
                        courses.Add(MapCourse(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetCoursesByCategory ({category}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve courses for category {category}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return courses
        End Function

        Public Function GetCompulsoryCourses() As List(Of Course)
            Dim courses As New List(Of Course)()
            Try
                Dim sql As String = "SELECT * FROM Courses WHERE IsCompulsory = 1 AND IsActive = 1 ORDER BY CourseCode"

                Using reader As SqlDataReader = _db.ExecuteReader(sql, Nothing)
                    While reader.Read()
                        courses.Add(MapCourse(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetCompulsoryCourses: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve compulsory courses", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return courses
        End Function

        Public Function GetCoursesNeedingRenewal() As List(Of Course)
            Dim courses As New List(Of Course)()
            Try
                ' Return active courses that have training records with expiry dates approaching or past
                Dim sql As String = "SELECT DISTINCT c.* " &
                                    "FROM Courses c " &
                                    "INNER JOIN TrainingRecords tr ON c.CourseId = tr.CourseId " &
                                    "WHERE c.IsActive = 1 " &
                                    "AND c.ValidityPeriodMonths > 0 " &
                                    "AND tr.CompletionStatus = 'COMPLETED' " &
                                    "AND tr.ExpiryDate IS NOT NULL " &
                                    "AND tr.ExpiryDate <= DATEADD(day, 90, GETDATE()) " &
                                    "ORDER BY c.CourseCode"

                Using reader As SqlDataReader = _db.ExecuteReader(sql, Nothing)
                    While reader.Read()
                        courses.Add(MapCourse(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetCoursesNeedingRenewal: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve courses needing renewal", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return courses
        End Function

        Public Function CreateCourse(course As Course) As Integer
            Try
                Dim sql As String = "INSERT INTO Courses " &
                                    "(CourseCode, Title, Description, Category, DurationHours, DeliveryMethod, " &
                                    "IsActive, IsCompulsory, MaxParticipants, MinParticipants, Prerequisites, " &
                                    "LearningObjectives, CourseContent, AssessmentMethod, CertificateTemplate, " &
                                    "ValidityPeriodMonths, CostPerParticipant, ApprovalRequired, CreatedDate, CreatedBy, " &
                                    "Version, SkillsFrameworkLevel, ComplianceCategory, ExternalProvider, " &
                                    "ProviderContactEmail, CourseMaterials, EquipmentRequired) " &
                                    "VALUES " &
                                    "(@CourseCode, @Title, @Description, @Category, @DurationHours, @DeliveryMethod, " &
                                    "@IsActive, @IsCompulsory, @MaxParticipants, @MinParticipants, @Prerequisites, " &
                                    "@LearningObjectives, @CourseContent, @AssessmentMethod, @CertificateTemplate, " &
                                    "@ValidityPeriodMonths, @CostPerParticipant, @ApprovalRequired, @CreatedDate, @CreatedBy, " &
                                    "@Version, @SkillsFrameworkLevel, @ComplianceCategory, @ExternalProvider, " &
                                    "@ProviderContactEmail, @CourseMaterials, @EquipmentRequired); " &
                                    "SELECT SCOPE_IDENTITY();"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@CourseCode", SqlDbType.NVarChar, 20, course.CourseCode),
                    DatabaseHelper.CreateParameter("@Title", SqlDbType.NVarChar, 200, course.Title),
                    DatabaseHelper.CreateParameter("@Description", course.Description),
                    DatabaseHelper.CreateParameter("@Category", SqlDbType.NVarChar, 100, course.Category),
                    DatabaseHelper.CreateParameter("@DurationHours", SqlDbType.Decimal, course.DurationHours),
                    DatabaseHelper.CreateParameter("@DeliveryMethod", SqlDbType.NVarChar, 50, course.DeliveryMethod),
                    DatabaseHelper.CreateParameter("@IsActive", SqlDbType.Bit, course.IsActive),
                    DatabaseHelper.CreateParameter("@IsCompulsory", SqlDbType.Bit, course.IsCompulsory),
                    DatabaseHelper.CreateParameter("@MaxParticipants", SqlDbType.Int, course.MaxParticipants),
                    DatabaseHelper.CreateParameter("@MinParticipants", SqlDbType.Int, course.MinParticipants),
                    DatabaseHelper.CreateParameter("@Prerequisites", course.Prerequisites),
                    DatabaseHelper.CreateParameter("@LearningObjectives", course.LearningObjectives),
                    DatabaseHelper.CreateParameter("@CourseContent", course.CourseContent),
                    DatabaseHelper.CreateParameter("@AssessmentMethod", SqlDbType.NVarChar, 200, course.AssessmentMethod),
                    DatabaseHelper.CreateParameter("@CertificateTemplate", SqlDbType.NVarChar, 200, course.CertificateTemplate),
                    DatabaseHelper.CreateParameter("@ValidityPeriodMonths", SqlDbType.Int, course.ValidityPeriodMonths),
                    DatabaseHelper.CreateParameter("@CostPerParticipant", SqlDbType.Decimal, course.CostPerParticipant),
                    DatabaseHelper.CreateParameter("@ApprovalRequired", SqlDbType.Bit, course.ApprovalRequired),
                    DatabaseHelper.CreateParameter("@CreatedDate", SqlDbType.DateTime, course.CreatedDate),
                    DatabaseHelper.CreateParameter("@CreatedBy", SqlDbType.NVarChar, 50, course.CreatedBy),
                    DatabaseHelper.CreateParameter("@Version", SqlDbType.NVarChar, 10, course.Version),
                    DatabaseHelper.CreateParameter("@SkillsFrameworkLevel", SqlDbType.NVarChar, 20, course.SkillsFrameworkLevel),
                    DatabaseHelper.CreateParameter("@ComplianceCategory", SqlDbType.NVarChar, 100, course.ComplianceCategory),
                    DatabaseHelper.CreateParameter("@ExternalProvider", SqlDbType.NVarChar, 200, course.ExternalProvider),
                    DatabaseHelper.CreateParameter("@ProviderContactEmail", SqlDbType.NVarChar, 100, course.ProviderContactEmail),
                    DatabaseHelper.CreateParameter("@CourseMaterials", course.CourseMaterials),
                    DatabaseHelper.CreateParameter("@EquipmentRequired", course.EquipmentRequired)
                }

                Dim result As Object = _db.ExecuteScalar(sql, parameters)
                If result IsNot Nothing AndAlso result IsNot DBNull.Value Then
                    Return Convert.ToInt32(result)
                End If
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in CreateCourse: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to create course", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return 0
        End Function

        Public Function UpdateCourse(course As Course) As Boolean
            Try
                Dim sql As String = "UPDATE Courses SET " &
                                    "CourseCode = @CourseCode, " &
                                    "Title = @Title, " &
                                    "Description = @Description, " &
                                    "Category = @Category, " &
                                    "DurationHours = @DurationHours, " &
                                    "DeliveryMethod = @DeliveryMethod, " &
                                    "IsActive = @IsActive, " &
                                    "IsCompulsory = @IsCompulsory, " &
                                    "MaxParticipants = @MaxParticipants, " &
                                    "MinParticipants = @MinParticipants, " &
                                    "Prerequisites = @Prerequisites, " &
                                    "LearningObjectives = @LearningObjectives, " &
                                    "CourseContent = @CourseContent, " &
                                    "AssessmentMethod = @AssessmentMethod, " &
                                    "CertificateTemplate = @CertificateTemplate, " &
                                    "ValidityPeriodMonths = @ValidityPeriodMonths, " &
                                    "CostPerParticipant = @CostPerParticipant, " &
                                    "ApprovalRequired = @ApprovalRequired, " &
                                    "ModifiedDate = @ModifiedDate, " &
                                    "ModifiedBy = @ModifiedBy, " &
                                    "Version = @Version, " &
                                    "SkillsFrameworkLevel = @SkillsFrameworkLevel, " &
                                    "ComplianceCategory = @ComplianceCategory, " &
                                    "ExternalProvider = @ExternalProvider, " &
                                    "ProviderContactEmail = @ProviderContactEmail, " &
                                    "CourseMaterials = @CourseMaterials, " &
                                    "EquipmentRequired = @EquipmentRequired " &
                                    "WHERE CourseId = @CourseId"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@CourseId", SqlDbType.Int, course.CourseId),
                    DatabaseHelper.CreateParameter("@CourseCode", SqlDbType.NVarChar, 20, course.CourseCode),
                    DatabaseHelper.CreateParameter("@Title", SqlDbType.NVarChar, 200, course.Title),
                    DatabaseHelper.CreateParameter("@Description", course.Description),
                    DatabaseHelper.CreateParameter("@Category", SqlDbType.NVarChar, 100, course.Category),
                    DatabaseHelper.CreateParameter("@DurationHours", SqlDbType.Decimal, course.DurationHours),
                    DatabaseHelper.CreateParameter("@DeliveryMethod", SqlDbType.NVarChar, 50, course.DeliveryMethod),
                    DatabaseHelper.CreateParameter("@IsActive", SqlDbType.Bit, course.IsActive),
                    DatabaseHelper.CreateParameter("@IsCompulsory", SqlDbType.Bit, course.IsCompulsory),
                    DatabaseHelper.CreateParameter("@MaxParticipants", SqlDbType.Int, course.MaxParticipants),
                    DatabaseHelper.CreateParameter("@MinParticipants", SqlDbType.Int, course.MinParticipants),
                    DatabaseHelper.CreateParameter("@Prerequisites", course.Prerequisites),
                    DatabaseHelper.CreateParameter("@LearningObjectives", course.LearningObjectives),
                    DatabaseHelper.CreateParameter("@CourseContent", course.CourseContent),
                    DatabaseHelper.CreateParameter("@AssessmentMethod", SqlDbType.NVarChar, 200, course.AssessmentMethod),
                    DatabaseHelper.CreateParameter("@CertificateTemplate", SqlDbType.NVarChar, 200, course.CertificateTemplate),
                    DatabaseHelper.CreateParameter("@ValidityPeriodMonths", SqlDbType.Int, course.ValidityPeriodMonths),
                    DatabaseHelper.CreateParameter("@CostPerParticipant", SqlDbType.Decimal, course.CostPerParticipant),
                    DatabaseHelper.CreateParameter("@ApprovalRequired", SqlDbType.Bit, course.ApprovalRequired),
                    DatabaseHelper.CreateParameter("@ModifiedDate", SqlDbType.DateTime, course.ModifiedDate),
                    DatabaseHelper.CreateParameter("@ModifiedBy", SqlDbType.NVarChar, 50, course.ModifiedBy),
                    DatabaseHelper.CreateParameter("@Version", SqlDbType.NVarChar, 10, course.Version),
                    DatabaseHelper.CreateParameter("@SkillsFrameworkLevel", SqlDbType.NVarChar, 20, course.SkillsFrameworkLevel),
                    DatabaseHelper.CreateParameter("@ComplianceCategory", SqlDbType.NVarChar, 100, course.ComplianceCategory),
                    DatabaseHelper.CreateParameter("@ExternalProvider", SqlDbType.NVarChar, 200, course.ExternalProvider),
                    DatabaseHelper.CreateParameter("@ProviderContactEmail", SqlDbType.NVarChar, 100, course.ProviderContactEmail),
                    DatabaseHelper.CreateParameter("@CourseMaterials", course.CourseMaterials),
                    DatabaseHelper.CreateParameter("@EquipmentRequired", course.EquipmentRequired)
                }

                Dim rowsAffected As Integer = _db.ExecuteNonQuery(sql, parameters)
                Return rowsAffected > 0
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in UpdateCourse ({course.CourseId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to update course {course.CourseId}", ex)
            Finally
                _db.CloseConnection()
            End Try
        End Function

        Public Function DeactivateCourse(courseId As Integer, modifiedBy As String) As Boolean
            Try
                Dim sql As String = "UPDATE Courses SET IsActive = 0, ModifiedDate = @ModifiedDate, ModifiedBy = @ModifiedBy " &
                                    "WHERE CourseId = @CourseId"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@CourseId", SqlDbType.Int, courseId),
                    DatabaseHelper.CreateParameter("@ModifiedDate", SqlDbType.DateTime, DateTime.Now),
                    DatabaseHelper.CreateParameter("@ModifiedBy", SqlDbType.NVarChar, 50, modifiedBy)
                }

                Dim rowsAffected As Integer = _db.ExecuteNonQuery(sql, parameters)
                Return rowsAffected > 0
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in DeactivateCourse ({courseId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to deactivate course {courseId}", ex)
            Finally
                _db.CloseConnection()
            End Try
        End Function

        Public Function GetActiveSessionCount(courseId As Integer) As Integer
            Try
                Dim sql As String = "SELECT COUNT(*) FROM TrainingSessions " &
                                    "WHERE CourseId = @CourseId " &
                                    "AND SessionStatus IN ('SCHEDULED', 'CONFIRMED') " &
                                    "AND SessionDate >= GETDATE()"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@CourseId", SqlDbType.Int, courseId)
                }

                Dim result As Object = _db.ExecuteScalar(sql, parameters)
                If result IsNot Nothing AndAlso result IsNot DBNull.Value Then
                    Return Convert.ToInt32(result)
                End If
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetActiveSessionCount ({courseId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to get active session count for course {courseId}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return 0
        End Function

        Public Function GetCourseStatistics(courseId As Integer) As Dictionary(Of String, Object)
            Dim stats As New Dictionary(Of String, Object)()
            Try
                Dim sql As String = "SELECT " &
                                    "COUNT(DISTINCT ts.SessionId) AS TotalSessions, " &
                                    "COUNT(DISTINCT CASE WHEN ts.SessionStatus IN ('SCHEDULED', 'CONFIRMED') AND ts.SessionDate >= GETDATE() THEN ts.SessionId END) AS UpcomingSessions, " &
                                    "COUNT(DISTINCT CASE WHEN ts.SessionStatus = 'COMPLETED' THEN ts.SessionId END) AS CompletedSessions, " &
                                    "COUNT(DISTINCT CASE WHEN ts.SessionStatus = 'CANCELLED' THEN ts.SessionId END) AS CancelledSessions, " &
                                    "COUNT(DISTINCT tr.EmployeeId) AS TotalParticipants, " &
                                    "COUNT(DISTINCT CASE WHEN tr.CompletionStatus = 'COMPLETED' THEN tr.EmployeeId END) AS CompletedParticipants, " &
                                    "COUNT(DISTINCT CASE WHEN tr.CompletionStatus = 'FAILED' THEN tr.EmployeeId END) AS FailedParticipants, " &
                                    "ISNULL(AVG(CAST(tr.Score AS DECIMAL)), 0) AS AverageScore " &
                                    "FROM Courses c " &
                                    "LEFT JOIN TrainingSessions ts ON c.CourseId = ts.CourseId " &
                                    "LEFT JOIN TrainingRecords tr ON c.CourseId = tr.CourseId " &
                                    "WHERE c.CourseId = @CourseId"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@CourseId", SqlDbType.Int, courseId)
                }

                Dim table As DataTable = _db.ExecuteDataTable(sql, parameters)

                If table.Rows.Count > 0 Then
                    Dim row As DataRow = table.Rows(0)
                    stats("TotalSessions") = If(row("TotalSessions") Is DBNull.Value, 0, Convert.ToInt32(row("TotalSessions")))
                    stats("UpcomingSessions") = If(row("UpcomingSessions") Is DBNull.Value, 0, Convert.ToInt32(row("UpcomingSessions")))
                    stats("CompletedSessions") = If(row("CompletedSessions") Is DBNull.Value, 0, Convert.ToInt32(row("CompletedSessions")))
                    stats("CancelledSessions") = If(row("CancelledSessions") Is DBNull.Value, 0, Convert.ToInt32(row("CancelledSessions")))
                    stats("TotalParticipants") = If(row("TotalParticipants") Is DBNull.Value, 0, Convert.ToInt32(row("TotalParticipants")))
                    stats("CompletedParticipants") = If(row("CompletedParticipants") Is DBNull.Value, 0, Convert.ToInt32(row("CompletedParticipants")))
                    stats("FailedParticipants") = If(row("FailedParticipants") Is DBNull.Value, 0, Convert.ToInt32(row("FailedParticipants")))
                    stats("AverageScore") = If(row("AverageScore") Is DBNull.Value, 0D, Convert.ToDecimal(row("AverageScore")))
                End If
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetCourseStatistics ({courseId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to get statistics for course {courseId}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return stats
        End Function

        Public Function GetCourseCompletionRate(courseId As Integer, fromDate As DateTime?, toDate As DateTime?) As Decimal
            Try
                Dim conditions As New List(Of String)()
                Dim paramList As New List(Of SqlParameter)()

                conditions.Add("CourseId = @CourseId")
                paramList.Add(DatabaseHelper.CreateParameter("@CourseId", SqlDbType.Int, courseId))

                If fromDate.HasValue Then
                    conditions.Add("EnrollmentDate >= @FromDate")
                    paramList.Add(DatabaseHelper.CreateParameter("@FromDate", SqlDbType.DateTime, fromDate.Value))
                End If

                If toDate.HasValue Then
                    conditions.Add("EnrollmentDate <= @ToDate")
                    paramList.Add(DatabaseHelper.CreateParameter("@ToDate", SqlDbType.DateTime, toDate.Value))
                End If

                Dim whereClause As String = DatabaseHelper.BuildWhereClause(conditions)

                Dim sql As String = "SELECT " &
                                    "CASE WHEN COUNT(*) = 0 THEN 0 " &
                                    "ELSE CAST(SUM(CASE WHEN CompletionStatus = 'COMPLETED' THEN 1 ELSE 0 END) AS DECIMAL) / COUNT(*) * 100 " &
                                    "END AS CompletionRate " &
                                    "FROM TrainingRecords" & whereClause

                Dim result As Object = _db.ExecuteScalar(sql, paramList.ToArray())
                If result IsNot Nothing AndAlso result IsNot DBNull.Value Then
                    Return Convert.ToDecimal(result)
                End If
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetCourseCompletionRate ({courseId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to get completion rate for course {courseId}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return 0D
        End Function

        Public Function GetCoursePrerequisites(courseId As Integer) As List(Of Course)
            Dim courses As New List(Of Course)()
            Try
                Dim sql As String = "SELECT c.* " &
                                    "FROM Courses c " &
                                    "INNER JOIN CoursePrerequisites cp ON c.CourseId = cp.PrerequisiteCourseId " &
                                    "WHERE cp.CourseId = @CourseId AND cp.IsRequired = 1 " &
                                    "ORDER BY c.CourseCode"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@CourseId", SqlDbType.Int, courseId)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    While reader.Read()
                        courses.Add(MapCourse(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetCoursePrerequisites ({courseId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to get prerequisites for course {courseId}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return courses
        End Function

        Public Function GetCoursesRequiringCourse(courseId As Integer) As List(Of Course)
            Dim courses As New List(Of Course)()
            Try
                Dim sql As String = "SELECT c.* " &
                                    "FROM Courses c " &
                                    "INNER JOIN CoursePrerequisites cp ON c.CourseId = cp.CourseId " &
                                    "WHERE cp.PrerequisiteCourseId = @CourseId AND cp.IsRequired = 1 " &
                                    "ORDER BY c.CourseCode"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@CourseId", SqlDbType.Int, courseId)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    While reader.Read()
                        courses.Add(MapCourse(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetCoursesRequiringCourse ({courseId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to get courses requiring course {courseId}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return courses
        End Function

        Public Function GetCourseCategories() As List(Of String)
            Dim categories As New List(Of String)()
            Try
                Dim sql As String = "SELECT DISTINCT Category FROM Courses WHERE IsActive = 1 ORDER BY Category"

                Using reader As SqlDataReader = _db.ExecuteReader(sql, Nothing)
                    While reader.Read()
                        Dim cat As String = DatabaseHelper.SafeGetString(reader, "Category")
                        If Not String.IsNullOrEmpty(cat) Then
                            categories.Add(cat)
                        End If
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetCourseCategories: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve course categories", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return categories
        End Function

        Public Function GetDeliveryMethods() As List(Of String)
            Dim methods As New List(Of String)()
            Try
                Dim sql As String = "SELECT DISTINCT DeliveryMethod FROM Courses WHERE IsActive = 1 ORDER BY DeliveryMethod"

                Using reader As SqlDataReader = _db.ExecuteReader(sql, Nothing)
                    While reader.Read()
                        Dim method As String = DatabaseHelper.SafeGetString(reader, "DeliveryMethod")
                        If Not String.IsNullOrEmpty(method) Then
                            methods.Add(method)
                        End If
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetDeliveryMethods: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve delivery methods", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return methods
        End Function

        Public Function GetPopularCourses(topCount As Integer) As List(Of Course)
            Dim courses As New List(Of Course)()
            Try
                Dim sql As String = "SELECT TOP (@TopCount) c.*, COUNT(tr.RecordId) AS EnrollmentCount " &
                                    "FROM Courses c " &
                                    "INNER JOIN TrainingRecords tr ON c.CourseId = tr.CourseId " &
                                    "WHERE c.IsActive = 1 " &
                                    "GROUP BY c.CourseId, c.CourseCode, c.Title, c.Description, c.Category, " &
                                    "c.DurationHours, c.DeliveryMethod, c.IsActive, c.IsCompulsory, " &
                                    "c.MaxParticipants, c.MinParticipants, c.Prerequisites, c.LearningObjectives, " &
                                    "c.CourseContent, c.AssessmentMethod, c.CertificateTemplate, c.ValidityPeriodMonths, " &
                                    "c.CostPerParticipant, c.ApprovalRequired, c.CreatedDate, c.CreatedBy, " &
                                    "c.ModifiedDate, c.ModifiedBy, c.Version, c.SkillsFrameworkLevel, " &
                                    "c.ComplianceCategory, c.ExternalProvider, c.ProviderContactEmail, " &
                                    "c.CourseMaterials, c.EquipmentRequired " &
                                    "ORDER BY COUNT(tr.RecordId) DESC"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@TopCount", SqlDbType.Int, topCount)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    While reader.Read()
                        courses.Add(MapCourse(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetPopularCourses: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve popular courses", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return courses
        End Function

        Public Function GetRecentlyCreatedCourses(days As Integer) As List(Of Course)
            Dim courses As New List(Of Course)()
            Try
                Dim sql As String = "SELECT * FROM Courses " &
                                    "WHERE CreatedDate >= DATEADD(day, -@Days, GETDATE()) " &
                                    "ORDER BY CreatedDate DESC"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@Days", SqlDbType.Int, days)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    While reader.Read()
                        courses.Add(MapCourse(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetRecentlyCreatedCourses: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve recently created courses", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return courses
        End Function

        Private Function MapCourse(reader As SqlDataReader) As Course
            Dim course As New Course()
            course.CourseId = DatabaseHelper.SafeGetInteger(reader, "CourseId")
            course.CourseCode = DatabaseHelper.SafeGetString(reader, "CourseCode")
            course.Title = DatabaseHelper.SafeGetString(reader, "Title")
            course.Description = DatabaseHelper.SafeGetString(reader, "Description")
            course.Category = DatabaseHelper.SafeGetString(reader, "Category")
            course.DurationHours = DatabaseHelper.SafeGetDecimal(reader, "DurationHours")
            course.DeliveryMethod = DatabaseHelper.SafeGetString(reader, "DeliveryMethod")
            course.IsActive = DatabaseHelper.SafeGetBoolean(reader, "IsActive")
            course.IsCompulsory = DatabaseHelper.SafeGetBoolean(reader, "IsCompulsory")
            course.MaxParticipants = DatabaseHelper.SafeGetInteger(reader, "MaxParticipants")
            course.MinParticipants = DatabaseHelper.SafeGetInteger(reader, "MinParticipants")
            course.Prerequisites = DatabaseHelper.SafeGetString(reader, "Prerequisites")
            course.LearningObjectives = DatabaseHelper.SafeGetString(reader, "LearningObjectives")
            course.CourseContent = DatabaseHelper.SafeGetString(reader, "CourseContent")
            course.AssessmentMethod = DatabaseHelper.SafeGetString(reader, "AssessmentMethod")
            course.CertificateTemplate = DatabaseHelper.SafeGetString(reader, "CertificateTemplate")
            course.ValidityPeriodMonths = DatabaseHelper.SafeGetInteger(reader, "ValidityPeriodMonths")
            course.CostPerParticipant = DatabaseHelper.SafeGetDecimal(reader, "CostPerParticipant")
            course.ApprovalRequired = DatabaseHelper.SafeGetBoolean(reader, "ApprovalRequired")
            course.CreatedDate = DatabaseHelper.SafeGetDateTime(reader, "CreatedDate")
            course.CreatedBy = DatabaseHelper.SafeGetString(reader, "CreatedBy")
            course.ModifiedDate = DatabaseHelper.SafeGetNullableDateTime(reader, "ModifiedDate")
            course.ModifiedBy = DatabaseHelper.SafeGetString(reader, "ModifiedBy")
            course.Version = DatabaseHelper.SafeGetString(reader, "Version")
            course.SkillsFrameworkLevel = DatabaseHelper.SafeGetString(reader, "SkillsFrameworkLevel")
            course.ComplianceCategory = DatabaseHelper.SafeGetString(reader, "ComplianceCategory")
            course.ExternalProvider = DatabaseHelper.SafeGetString(reader, "ExternalProvider")
            course.ProviderContactEmail = DatabaseHelper.SafeGetString(reader, "ProviderContactEmail")
            course.CourseMaterials = DatabaseHelper.SafeGetString(reader, "CourseMaterials")
            course.EquipmentRequired = DatabaseHelper.SafeGetString(reader, "EquipmentRequired")
            Return course
        End Function

        ' ===== CourseModule Methods =====

        Public Function GetModulesByCourseId(courseId As Integer) As List(Of CourseModule)
            Dim modules As New List(Of CourseModule)()
            Try
                Dim sql As String = "SELECT * FROM CourseModules WHERE CourseId = @CourseId ORDER BY ModuleOrder"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@CourseId", SqlDbType.Int, courseId)
                }

                Using reader As SqlDataReader = _db.ExecuteReader(sql, parameters)
                    While reader.Read()
                        modules.Add(MapModuleFromReader(reader))
                    End While
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in GetModulesByCourseId ({courseId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve modules for course {courseId}", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return modules
        End Function

        Public Function CreateModule(courseModule As CourseModule) As Integer
            Try
                Dim sql As String = "INSERT INTO CourseModules " &
                                    "(CourseId, ModuleOrder, Title, ModuleType, DurationMinutes, Description, " &
                                    "CreatedDate, ModifiedDate) " &
                                    "VALUES " &
                                    "(@CourseId, @ModuleOrder, @Title, @ModuleType, @DurationMinutes, @Description, " &
                                    "@CreatedDate, @ModifiedDate); " &
                                    "SELECT SCOPE_IDENTITY();"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@CourseId", SqlDbType.Int, courseModule.CourseId),
                    DatabaseHelper.CreateParameter("@ModuleOrder", SqlDbType.Int, courseModule.ModuleOrder),
                    DatabaseHelper.CreateParameter("@Title", SqlDbType.NVarChar, 200, courseModule.Title),
                    DatabaseHelper.CreateParameter("@ModuleType", SqlDbType.NVarChar, 50, courseModule.ModuleType),
                    DatabaseHelper.CreateParameter("@DurationMinutes", SqlDbType.Int, courseModule.DurationMinutes),
                    DatabaseHelper.CreateParameter("@Description", courseModule.Description),
                    DatabaseHelper.CreateParameter("@CreatedDate", SqlDbType.DateTime, courseModule.CreatedDate),
                    DatabaseHelper.CreateParameter("@ModifiedDate", SqlDbType.DateTime, courseModule.ModifiedDate)
                }

                Dim result As Object = _db.ExecuteScalar(sql, parameters)
                If result IsNot Nothing AndAlso result IsNot DBNull.Value Then
                    Return Convert.ToInt32(result)
                End If
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in CreateModule: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to create course module", ex)
            Finally
                _db.CloseConnection()
            End Try
            Return 0
        End Function

        Public Function UpdateModule(courseModule As CourseModule) As Boolean
            Try
                Dim sql As String = "UPDATE CourseModules SET " &
                                    "CourseId = @CourseId, " &
                                    "ModuleOrder = @ModuleOrder, " &
                                    "Title = @Title, " &
                                    "ModuleType = @ModuleType, " &
                                    "DurationMinutes = @DurationMinutes, " &
                                    "Description = @Description, " &
                                    "ModifiedDate = @ModifiedDate " &
                                    "WHERE ModuleId = @ModuleId"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@ModuleId", SqlDbType.Int, courseModule.ModuleId),
                    DatabaseHelper.CreateParameter("@CourseId", SqlDbType.Int, courseModule.CourseId),
                    DatabaseHelper.CreateParameter("@ModuleOrder", SqlDbType.Int, courseModule.ModuleOrder),
                    DatabaseHelper.CreateParameter("@Title", SqlDbType.NVarChar, 200, courseModule.Title),
                    DatabaseHelper.CreateParameter("@ModuleType", SqlDbType.NVarChar, 50, courseModule.ModuleType),
                    DatabaseHelper.CreateParameter("@DurationMinutes", SqlDbType.Int, courseModule.DurationMinutes),
                    DatabaseHelper.CreateParameter("@Description", courseModule.Description),
                    DatabaseHelper.CreateParameter("@ModifiedDate", SqlDbType.DateTime, courseModule.ModifiedDate)
                }

                Dim rowsAffected As Integer = _db.ExecuteNonQuery(sql, parameters)
                Return rowsAffected > 0
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in UpdateModule ({courseModule.ModuleId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to update course module {courseModule.ModuleId}", ex)
            Finally
                _db.CloseConnection()
            End Try
        End Function

        Public Function DeleteModule(moduleId As Integer) As Boolean
            Try
                Dim sql As String = "DELETE FROM CourseModules WHERE ModuleId = @ModuleId"

                Dim parameters As SqlParameter() = {
                    DatabaseHelper.CreateParameter("@ModuleId", SqlDbType.Int, moduleId)
                }

                Dim rowsAffected As Integer = _db.ExecuteNonQuery(sql, parameters)
                Return rowsAffected > 0
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error in DeleteModule ({moduleId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to delete course module {moduleId}", ex)
            Finally
                _db.CloseConnection()
            End Try
        End Function

        Public Function ReorderModules(courseId As Integer, moduleIds As List(Of Integer)) As Boolean
            Try
                Dim transaction As SqlTransaction = _db.BeginTransaction()
                Try
                    For i As Integer = 0 To moduleIds.Count - 1
                        Dim sql As String = "UPDATE CourseModules SET ModuleOrder = @ModuleOrder, ModifiedDate = @ModifiedDate " &
                                            "WHERE ModuleId = @ModuleId AND CourseId = @CourseId"

                        Using cmd As New SqlCommand(sql, _db.Connection, transaction)
                            cmd.Parameters.Add(DatabaseHelper.CreateParameter("@ModuleOrder", SqlDbType.Int, i + 1))
                            cmd.Parameters.Add(DatabaseHelper.CreateParameter("@ModifiedDate", SqlDbType.DateTime, DateTime.Now))
                            cmd.Parameters.Add(DatabaseHelper.CreateParameter("@ModuleId", SqlDbType.Int, moduleIds(i)))
                            cmd.Parameters.Add(DatabaseHelper.CreateParameter("@CourseId", SqlDbType.Int, courseId))
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
                EventLog.WriteEntry("TrainTrack", $"Error in ReorderModules ({courseId}): {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to reorder modules for course {courseId}", ex)
            Finally
                _db.CloseConnection()
            End Try
        End Function

        Private Function MapModuleFromReader(reader As SqlDataReader) As CourseModule
            Dim courseModule As New CourseModule()
            courseModule.ModuleId = DatabaseHelper.SafeGetInteger(reader, "ModuleId")
            courseModule.CourseId = DatabaseHelper.SafeGetInteger(reader, "CourseId")
            courseModule.ModuleOrder = DatabaseHelper.SafeGetInteger(reader, "ModuleOrder")
            courseModule.Title = DatabaseHelper.SafeGetString(reader, "Title")
            courseModule.ModuleType = DatabaseHelper.SafeGetString(reader, "ModuleType")
            courseModule.DurationMinutes = DatabaseHelper.SafeGetInteger(reader, "DurationMinutes")
            courseModule.Description = DatabaseHelper.SafeGetString(reader, "Description")
            courseModule.CreatedDate = DatabaseHelper.SafeGetDateTime(reader, "CreatedDate")
            courseModule.ModifiedDate = DatabaseHelper.SafeGetNullableDateTime(reader, "ModifiedDate")
            Return courseModule
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            Try
                If _db IsNot Nothing Then
                    _db.Dispose()
                    _db = Nothing
                End If
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error disposing CourseRepository: {ex.Message}", EventLogEntryType.Warning)
            End Try
        End Sub

    End Class
End Namespace
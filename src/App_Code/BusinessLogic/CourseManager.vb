Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports Defra.TrainTrack.Models
Imports Defra.TrainTrack.DataAccess

Namespace Defra.TrainTrack.BusinessLogic
    Public Class CourseManager
        Implements IDisposable

        Private _repository As CourseRepository

        Public Sub New()
            _repository = New CourseRepository()
        End Sub

        Public Function GetAllCourses() As List(Of Course)
            Try
                Return _repository.GetAllCourses()
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting all courses: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve course list", ex)
            End Try
        End Function

        Public Function GetActiveCourses() As List(Of Course)
            Try
                Return _repository.GetActiveCourses()
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting active courses: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve active course list", ex)
            End Try
        End Function

        Public Function GetCourseById(courseId As Integer) As Course
            Try
                If courseId <= 0 Then
                    Throw New ArgumentException("Course ID must be greater than 0", "courseId")
                End If

                Return _repository.GetCourseById(courseId)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting course by ID {courseId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve course with ID {courseId}", ex)
            End Try
        End Function

        Public Function GetCourseByCode(courseCode As String) As Course
            Try
                If String.IsNullOrWhiteSpace(courseCode) Then
                    Throw New ArgumentException("Course code cannot be empty", "courseCode")
                End If

                Return _repository.GetCourseByCode(courseCode)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting course by code {courseCode}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve course with code {courseCode}", ex)
            End Try
        End Function

        Public Function SearchCourses(searchTerm As String, category As String, deliveryMethod As String, isActive As Boolean?) As List(Of Course)
            Try
                Return _repository.SearchCourses(searchTerm, category, deliveryMethod, isActive)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error searching courses: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to search courses", ex)
            End Try
        End Function

        Public Function GetCoursesByCategory(category As String) As List(Of Course)
            Try
                If String.IsNullOrWhiteSpace(category) Then
                    Throw New ArgumentException("Category cannot be empty", "category")
                End If

                Return _repository.GetCoursesByCategory(category)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting courses by category {category}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve courses for category {category}", ex)
            End Try
        End Function

        Public Function GetCompulsoryCourses() As List(Of Course)
            Try
                Return _repository.GetCompulsoryCourses()
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting compulsory courses: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve compulsory courses", ex)
            End Try
        End Function

        Public Function GetCoursesNeedingRenewal() As List(Of Course)
            Try
                Return _repository.GetCoursesNeedingRenewal()
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting courses needing renewal: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve courses needing renewal", ex)
            End Try
        End Function

        Public Function CreateCourse(course As Course, createdBy As String) As Integer
            Try
                If course Is Nothing Then
                    Throw New ArgumentNullException("course", "Course cannot be null")
                End If

                If String.IsNullOrWhiteSpace(createdBy) Then
                    Throw New ArgumentException("Created by cannot be empty", "createdBy")
                End If

                ' Validate course data
                Dim validationErrors = course.GetValidationErrors()
                If validationErrors.Count > 0 Then
                    Throw New ArgumentException($"Course validation failed: {String.Join("; ", validationErrors)}")
                End If

                ' Check for duplicate course code
                Dim existingCourse = _repository.GetCourseByCode(course.CourseCode)
                If existingCourse IsNot Nothing Then
                    Throw New InvalidOperationException($"Course with code {course.CourseCode} already exists")
                End If

                ' Set audit fields
                course.CreatedBy = createdBy
                course.CreatedDate = DateTime.Now

                ' Create course
                Dim newCourseId As Integer = _repository.CreateCourse(course)

                ' Log the action
                EventLog.WriteEntry("TrainTrack", $"Course created: {course.CourseCode} by {createdBy}", EventLogEntryType.Information)

                Return newCourseId
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error creating course: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to create course", ex)
            End Try
        End Function

        Public Function UpdateCourse(course As Course, modifiedBy As String) As Boolean
            Try
                If course Is Nothing Then
                    Throw New ArgumentNullException("course", "Course cannot be null")
                End If

                If course.CourseId <= 0 Then
                    Throw New ArgumentException("Course ID must be greater than 0", "course")
                End If

                If String.IsNullOrWhiteSpace(modifiedBy) Then
                    Throw New ArgumentException("Modified by cannot be empty", "modifiedBy")
                End If

                ' Validate course data
                Dim validationErrors = course.GetValidationErrors()
                If validationErrors.Count > 0 Then
                    Throw New ArgumentException($"Course validation failed: {String.Join("; ", validationErrors)}")
                End If

                ' Check if course exists
                Dim existingCourse = _repository.GetCourseById(course.CourseId)
                If existingCourse Is Nothing Then
                    Throw New InvalidOperationException($"Course with ID {course.CourseId} does not exist")
                End If

                ' Check for duplicate course code (excluding current course)
                If course.CourseCode <> existingCourse.CourseCode Then
                    Dim duplicateCourse = _repository.GetCourseByCode(course.CourseCode)
                    If duplicateCourse IsNot Nothing AndAlso duplicateCourse.CourseId <> course.CourseId Then
                        Throw New InvalidOperationException($"Course with code {course.CourseCode} already exists")
                    End If
                End If

                ' Check if there are any active sessions that would be affected
                If Not course.IsActive AndAlso existingCourse.IsActive Then
                    Dim activeSessionCount = GetActiveSessionCount(course.CourseId)
                    If activeSessionCount > 0 Then
                        Throw New InvalidOperationException($"Cannot deactivate course with {activeSessionCount} active training session(s)")
                    End If
                End If

                ' Set audit fields
                course.ModifiedBy = modifiedBy
                course.ModifiedDate = DateTime.Now

                ' Update course
                Dim success As Boolean = _repository.UpdateCourse(course)

                If success Then
                    ' Log the action
                    EventLog.WriteEntry("TrainTrack", $"Course updated: {course.CourseCode} by {modifiedBy}", EventLogEntryType.Information)
                End If

                Return success
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error updating course: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to update course", ex)
            End Try
        End Function

        Public Function DeactivateCourse(courseId As Integer, modifiedBy As String) As Boolean
            Try
                If courseId <= 0 Then
                    Throw New ArgumentException("Course ID must be greater than 0", "courseId")
                End If

                If String.IsNullOrWhiteSpace(modifiedBy) Then
                    Throw New ArgumentException("Modified by cannot be empty", "modifiedBy")
                End If

                ' Get course to check if exists
                Dim course = _repository.GetCourseById(courseId)
                If course Is Nothing Then
                    Throw New InvalidOperationException($"Course with ID {courseId} does not exist")
                End If

                ' Check for any active training sessions
                Dim activeSessionCount = GetActiveSessionCount(courseId)
                If activeSessionCount > 0 Then
                    Throw New InvalidOperationException($"Cannot deactivate course with {activeSessionCount} active training session(s)")
                End If

                ' Deactivate course
                Dim success As Boolean = _repository.DeactivateCourse(courseId, modifiedBy)

                If success Then
                    ' Log the action
                    EventLog.WriteEntry("TrainTrack", $"Course deactivated: {course.CourseCode} by {modifiedBy}", EventLogEntryType.Information)
                End If

                Return success
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error deactivating course: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to deactivate course", ex)
            End Try
        End Function

        Public Function GetActiveSessionCount(courseId As Integer) As Integer
            Try
                If courseId <= 0 Then
                    Throw New ArgumentException("Course ID must be greater than 0", "courseId")
                End If

                Return _repository.GetActiveSessionCount(courseId)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting active session count for course {courseId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to get active session count for course {courseId}", ex)
            End Try
        End Function

        Public Function GetCourseStatistics(courseId As Integer) As Dictionary(Of String, Object)
            Try
                If courseId <= 0 Then
                    Throw New ArgumentException("Course ID must be greater than 0", "courseId")
                End If

                Return _repository.GetCourseStatistics(courseId)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting course statistics for course {courseId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to get statistics for course {courseId}", ex)
            End Try
        End Function

        Public Function GetCourseCompletionRate(courseId As Integer, fromDate As DateTime?, toDate As DateTime?) As Decimal
            Try
                If courseId <= 0 Then
                    Throw New ArgumentException("Course ID must be greater than 0", "courseId")
                End If

                Return _repository.GetCourseCompletionRate(courseId, fromDate, toDate)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting course completion rate for course {courseId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to get completion rate for course {courseId}", ex)
            End Try
        End Function

        Public Function GetCoursePrerequisites(courseId As Integer) As List(Of Course)
            Try
                If courseId <= 0 Then
                    Throw New ArgumentException("Course ID must be greater than 0", "courseId")
                End If

                Return _repository.GetCoursePrerequisites(courseId)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting course prerequisites for course {courseId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to get prerequisites for course {courseId}", ex)
            End Try
        End Function

        Public Function GetCoursesRequiringCourse(courseId As Integer) As List(Of Course)
            Try
                If courseId <= 0 Then
                    Throw New ArgumentException("Course ID must be greater than 0", "courseId")
                End If

                Return _repository.GetCoursesRequiringCourse(courseId)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting courses requiring course {courseId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to get courses requiring course {courseId}", ex)
            End Try
        End Function

        Public Function GetCourseCategories() As List(Of String)
            Try
                Return _repository.GetCourseCategories()
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting course categories: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve course categories", ex)
            End Try
        End Function

        Public Function GetDeliveryMethods() As List(Of String)
            Try
                Return _repository.GetDeliveryMethods()
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting delivery methods: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve delivery methods", ex)
            End Try
        End Function

        Public Function CanEmployeeTakeCourse(employeeId As Integer, courseId As Integer) As Boolean
            Try
                If employeeId <= 0 Then
                    Throw New ArgumentException("Employee ID must be greater than 0", "employeeId")
                End If

                If courseId <= 0 Then
                    Throw New ArgumentException("Course ID must be greater than 0", "courseId")
                End If

                ' Check if course is active
                Dim course = _repository.GetCourseById(courseId)
                If course Is Nothing OrElse Not course.IsActive Then
                    Return False
                End If

                ' Check if employee has completed prerequisites
                Dim prerequisites = GetCoursePrerequisites(courseId)
                If prerequisites.Count > 0 Then
                    Return HasEmployeeCompletedPrerequisites(employeeId, prerequisites)
                End If

                Return True
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error checking if employee {employeeId} can take course {courseId}: {ex.Message}", EventLogEntryType.Error)
                Return False
            End Try
        End Function

        Private Function HasEmployeeCompletedPrerequisites(employeeId As Integer, prerequisites As List(Of Course)) As Boolean
            Try
                Using trainingRepo As New TrainingRepository()
                    For Each prerequisite In prerequisites
                        Dim hasCompleted = trainingRepo.HasEmployeeCompletedCourse(employeeId, prerequisite.CourseId)
                        If Not hasCompleted Then
                            Return False
                        End If
                    Next
                End Using
                Return True
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error checking prerequisites for employee {employeeId}: {ex.Message}", EventLogEntryType.Error)
                Return False
            End Try
        End Function

        Public Function GetPopularCourses(topCount As Integer) As List(Of Course)
            Try
                If topCount <= 0 Then
                    topCount = 10 ' Default to top 10
                End If

                Return _repository.GetPopularCourses(topCount)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting popular courses: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve popular courses", ex)
            End Try
        End Function

        Public Function GetRecentlyCreatedCourses(days As Integer) As List(Of Course)
            Try
                If days <= 0 Then
                    days = 30 ' Default to 30 days
                End If

                Return _repository.GetRecentlyCreatedCourses(days)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting recently created courses: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to retrieve recently created courses", ex)
            End Try
        End Function

        Public Function DuplicateCourse(courseId As Integer, newCourseCode As String, newTitle As String, createdBy As String) As Integer
            Try
                If courseId <= 0 Then
                    Throw New ArgumentException("Course ID must be greater than 0", "courseId")
                End If

                If String.IsNullOrWhiteSpace(newCourseCode) Then
                    Throw New ArgumentException("New course code cannot be empty", "newCourseCode")
                End If

                If String.IsNullOrWhiteSpace(newTitle) Then
                    Throw New ArgumentException("New course title cannot be empty", "newTitle")
                End If

                If String.IsNullOrWhiteSpace(createdBy) Then
                    Throw New ArgumentException("Created by cannot be empty", "createdBy")
                End If

                ' Get the original course
                Dim originalCourse = _repository.GetCourseById(courseId)
                If originalCourse Is Nothing Then
                    Throw New InvalidOperationException($"Course with ID {courseId} does not exist")
                End If

                ' Check if new course code already exists
                Dim existingCourse = _repository.GetCourseByCode(newCourseCode)
                If existingCourse IsNot Nothing Then
                    Throw New InvalidOperationException($"Course with code {newCourseCode} already exists")
                End If

                ' Create new course based on original
                Dim newCourse As New Course()
                newCourse.CourseCode = newCourseCode
                newCourse.Title = newTitle
                newCourse.Description = originalCourse.Description
                newCourse.Category = originalCourse.Category
                newCourse.DurationHours = originalCourse.DurationHours
                newCourse.DeliveryMethod = originalCourse.DeliveryMethod
                newCourse.IsActive = True
                newCourse.IsCompulsory = originalCourse.IsCompulsory
                newCourse.MaxParticipants = originalCourse.MaxParticipants
                newCourse.MinParticipants = originalCourse.MinParticipants
                newCourse.Prerequisites = originalCourse.Prerequisites
                newCourse.LearningObjectives = originalCourse.LearningObjectives
                newCourse.CourseContent = originalCourse.CourseContent
                newCourse.AssessmentMethod = originalCourse.AssessmentMethod
                newCourse.CertificateTemplate = originalCourse.CertificateTemplate
                newCourse.ValidityPeriodMonths = originalCourse.ValidityPeriodMonths
                newCourse.CostPerParticipant = originalCourse.CostPerParticipant
                newCourse.ApprovalRequired = originalCourse.ApprovalRequired
                newCourse.SkillsFrameworkLevel = originalCourse.SkillsFrameworkLevel
                newCourse.ComplianceCategory = originalCourse.ComplianceCategory
                newCourse.CourseMaterials = originalCourse.CourseMaterials
                newCourse.EquipmentRequired = originalCourse.EquipmentRequired
                newCourse.Version = "1.0"

                ' Create the new course
                Return CreateCourse(newCourse, createdBy)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error duplicating course {courseId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to duplicate course", ex)
            End Try
        End Function

        ' ===== CourseModule Business Logic =====

        Public Function GetCourseModules(courseId As Integer) As List(Of CourseModule)
            Try
                If courseId <= 0 Then
                    Throw New ArgumentException("Course ID must be greater than 0", "courseId")
                End If

                Return _repository.GetModulesByCourseId(courseId)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error getting course modules for course {courseId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to retrieve modules for course {courseId}", ex)
            End Try
        End Function

        Public Function AddModuleToCourse(courseId As Integer, courseModule As CourseModule, createdBy As String) As Integer
            Try
                If courseId <= 0 Then
                    Throw New ArgumentException("Course ID must be greater than 0", "courseId")
                End If

                If courseModule Is Nothing Then
                    Throw New ArgumentNullException("courseModule", "Course module cannot be null")
                End If

                If String.IsNullOrWhiteSpace(createdBy) Then
                    Throw New ArgumentException("Created by cannot be empty", "createdBy")
                End If

                ' Validate module title is required
                If String.IsNullOrWhiteSpace(courseModule.Title) Then
                    Throw New ArgumentException("Module title is required")
                End If

                ' Validate DurationMinutes >= 0
                If courseModule.DurationMinutes < 0 Then
                    Throw New ArgumentException("Duration minutes must be 0 or greater")
                End If

                ' Validate CourseId must exist
                Dim course = _repository.GetCourseById(courseId)
                If course Is Nothing Then
                    Throw New InvalidOperationException($"Course with ID {courseId} does not exist")
                End If

                ' Set the CourseId and audit fields
                courseModule.CourseId = courseId
                courseModule.CreatedDate = DateTime.Now

                ' Create module
                Dim newModuleId As Integer = _repository.CreateModule(courseModule)

                ' Log the action
                EventLog.WriteEntry("TrainTrack", $"Module added to course {courseId}: {courseModule.Title} by {createdBy}", EventLogEntryType.Information)

                Return newModuleId
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error adding module to course {courseId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to add module to course {courseId}", ex)
            End Try
        End Function

        Public Function UpdateCourseModule(courseModule As CourseModule, modifiedBy As String) As Boolean
            Try
                If courseModule Is Nothing Then
                    Throw New ArgumentNullException("courseModule", "Course module cannot be null")
                End If

                If courseModule.ModuleId <= 0 Then
                    Throw New ArgumentException("Module ID must be greater than 0", "courseModule")
                End If

                If String.IsNullOrWhiteSpace(modifiedBy) Then
                    Throw New ArgumentException("Modified by cannot be empty", "modifiedBy")
                End If

                ' Validate module title is required
                If String.IsNullOrWhiteSpace(courseModule.Title) Then
                    Throw New ArgumentException("Module title is required")
                End If

                ' Validate DurationMinutes >= 0
                If courseModule.DurationMinutes < 0 Then
                    Throw New ArgumentException("Duration minutes must be 0 or greater")
                End If

                ' Validate CourseId must exist
                If courseModule.CourseId > 0 Then
                    Dim course = _repository.GetCourseById(courseModule.CourseId)
                    If course Is Nothing Then
                        Throw New InvalidOperationException($"Course with ID {courseModule.CourseId} does not exist")
                    End If
                End If

                ' Set audit fields
                courseModule.ModifiedDate = DateTime.Now

                ' Update module
                Dim success As Boolean = _repository.UpdateModule(courseModule)

                If success Then
                    EventLog.WriteEntry("TrainTrack", $"Module updated: {courseModule.ModuleId} by {modifiedBy}", EventLogEntryType.Information)
                End If

                Return success
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error updating module {courseModule.ModuleId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to update module {courseModule.ModuleId}", ex)
            End Try
        End Function

        Public Function RemoveCourseModule(moduleId As Integer, modifiedBy As String) As Boolean
            Try
                If moduleId <= 0 Then
                    Throw New ArgumentException("Module ID must be greater than 0", "moduleId")
                End If

                If String.IsNullOrWhiteSpace(modifiedBy) Then
                    Throw New ArgumentException("Modified by cannot be empty", "modifiedBy")
                End If

                ' Delete module
                Dim success As Boolean = _repository.DeleteModule(moduleId)

                If success Then
                    EventLog.WriteEntry("TrainTrack", $"Module removed: {moduleId} by {modifiedBy}", EventLogEntryType.Information)
                End If

                Return success
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error removing module {moduleId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to remove module {moduleId}", ex)
            End Try
        End Function

        Public Function ReorderCourseModules(courseId As Integer, moduleIds As List(Of Integer), modifiedBy As String) As Boolean
            Try
                If courseId <= 0 Then
                    Throw New ArgumentException("Course ID must be greater than 0", "courseId")
                End If

                If moduleIds Is Nothing OrElse moduleIds.Count = 0 Then
                    Throw New ArgumentException("Module IDs list cannot be empty", "moduleIds")
                End If

                If String.IsNullOrWhiteSpace(modifiedBy) Then
                    Throw New ArgumentException("Modified by cannot be empty", "modifiedBy")
                End If

                ' Validate CourseId must exist
                Dim course = _repository.GetCourseById(courseId)
                If course Is Nothing Then
                    Throw New InvalidOperationException($"Course with ID {courseId} does not exist")
                End If

                ' Reorder modules
                Dim success As Boolean = _repository.ReorderModules(courseId, moduleIds)

                If success Then
                    EventLog.WriteEntry("TrainTrack", $"Modules reordered for course {courseId} by {modifiedBy}", EventLogEntryType.Information)
                End If

                Return success
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error reordering modules for course {courseId}: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException($"Unable to reorder modules for course {courseId}", ex)
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
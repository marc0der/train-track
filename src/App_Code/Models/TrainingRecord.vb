Imports Microsoft.VisualBasic

Namespace Defra.TrainTrack.Models
    Public Class TrainingRecord
        Private _recordId As Integer
        Private _employeeId As Integer
        Private _sessionId As Integer
        Private _courseId As Integer
        Private _enrollmentDate As DateTime
        Private _completionDate As DateTime?
        Private _attendanceStatus As String
        Private _completionStatus As String
        Private _score As Integer?
        Private _maxScore As Integer?
        Private _passMark As Integer?
        Private _grade As String
        Private _certificateIssued As Boolean
        Private _certificateNumber As String
        Private _certificateIssuedDate As DateTime?
        Private _expiryDate As DateTime?
        Private _isExpired As Boolean
        Private _feedback As String
        Private _instructorComments As String
        Private _createdDate As DateTime
        Private _createdBy As String
        Private _modifiedDate As DateTime?
        Private _modifiedBy As String
        Private _renewalRequired As Boolean
        Private _renewalNotificationSent As Boolean
        Private _costCentre As String
        Private _approvalRequired As Boolean
        Private _approvedBy As String
        Private _approvedDate As DateTime?
        Private _rejectedReason As String

        ' Related objects
        Private _employee As Employee
        Private _session As TrainingSession
        Private _course As Course

        Public Property RecordId() As Integer
            Get
                Return _recordId
            End Get
            Set(ByVal value As Integer)
                _recordId = value
            End Set
        End Property

        Public Property EmployeeId() As Integer
            Get
                Return _employeeId
            End Get
            Set(ByVal value As Integer)
                _employeeId = value
            End Set
        End Property

        Public Property SessionId() As Integer
            Get
                Return _sessionId
            End Get
            Set(ByVal value As Integer)
                _sessionId = value
            End Set
        End Property

        Public Property CourseId() As Integer
            Get
                Return _courseId
            End Get
            Set(ByVal value As Integer)
                _courseId = value
            End Set
        End Property

        Public Property EnrollmentDate() As DateTime
            Get
                Return _enrollmentDate
            End Get
            Set(ByVal value As DateTime)
                _enrollmentDate = value
            End Set
        End Property

        Public Property CompletionDate() As DateTime?
            Get
                Return _completionDate
            End Get
            Set(ByVal value As DateTime?)
                _completionDate = value
            End Set
        End Property

        Public ReadOnly Property IsCompleted() As Boolean
            Get
                Return _completionDate.HasValue AndAlso _completionStatus = "COMPLETED"
            End Get
        End Property

        Public ReadOnly Property DaysSinceCompletion() As Integer?
            Get
                If _completionDate.HasValue Then
                    Return CInt((DateTime.Now - _completionDate.Value).TotalDays)
                End If
                Return Nothing
            End Get
        End Property

        Public Property AttendanceStatus() As String
            Get
                Return _attendanceStatus
            End Get
            Set(ByVal value As String)
                _attendanceStatus = value
            End Set
        End Property

        Public ReadOnly Property AttendanceStatusDisplay() As String
            Get
                Select Case _attendanceStatus.ToUpper()
                    Case "ENROLLED"
                        Return "Enrolled"
                    Case "ATTENDED"
                        Return "Attended"
                    Case "NO_SHOW"
                        Return "No Show"
                    Case "CANCELLED"
                        Return "Cancelled"
                    Case "TRANSFERRED"
                        Return "Transferred"
                    Case "WITHDRAWN"
                        Return "Withdrawn"
                    Case Else
                        Return _attendanceStatus
                End Select
            End Get
        End Property

        Public Property CompletionStatus() As String
            Get
                Return _completionStatus
            End Get
            Set(ByVal value As String)
                _completionStatus = value
            End Set
        End Property

        Public ReadOnly Property CompletionStatusDisplay() As String
            Get
                Select Case _completionStatus.ToUpper()
                    Case "NOT_STARTED"
                        Return "Not Started"
                    Case "IN_PROGRESS"
                        Return "In Progress"
                    Case "COMPLETED"
                        Return "Completed"
                    Case "FAILED"
                        Return "Failed"
                    Case "INCOMPLETE"
                        Return "Incomplete"
                    Case "EXEMPT"
                        Return "Exempt"
                    Case Else
                        Return _completionStatus
                End Select
            End Get
        End Property

        Public Property Score() As Integer?
            Get
                Return _score
            End Get
            Set(ByVal value As Integer?)
                _score = value
            End Set
        End Property

        Public Property MaxScore() As Integer?
            Get
                Return _maxScore
            End Get
            Set(ByVal value As Integer?)
                _maxScore = value
            End Set
        End Property

        Public Property PassMark() As Integer?
            Get
                Return _passMark
            End Get
            Set(ByVal value As Integer?)
                _passMark = value
            End Set
        End Property

        Public ReadOnly Property ScorePercentage() As Integer?
            Get
                If _score.HasValue AndAlso _maxScore.HasValue AndAlso _maxScore.Value > 0 Then
                    Return CInt((_score.Value / _maxScore.Value) * 100)
                End If
                Return Nothing
            End Get
        End Property

        Public ReadOnly Property HasPassed() As Boolean?
            Get
                If _score.HasValue AndAlso _passMark.HasValue Then
                    Return _score.Value >= _passMark.Value
                ElseIf ScorePercentage.HasValue AndAlso _passMark.HasValue Then
                    Return ScorePercentage.Value >= _passMark.Value
                End If
                Return Nothing
            End Get
        End Property

        Public Property Grade() As String
            Get
                Return _grade
            End Get
            Set(ByVal value As String)
                _grade = value
            End Set
        End Property

        Public ReadOnly Property CalculatedGrade() As String
            Get
                If Not String.IsNullOrEmpty(_grade) Then
                    Return _grade
                End If

                If ScorePercentage.HasValue Then
                    Dim percentage As Integer = ScorePercentage.Value
                    If percentage >= 90 Then
                        Return "A"
                    ElseIf percentage >= 80 Then
                        Return "B"
                    ElseIf percentage >= 70 Then
                        Return "C"
                    ElseIf percentage >= 60 Then
                        Return "D"
                    Else
                        Return "F"
                    End If
                End If

                If HasPassed.HasValue Then
                    Return If(HasPassed.Value, "PASS", "FAIL")
                End If

                Return "N/A"
            End Get
        End Property

        Public Property CertificateIssued() As Boolean
            Get
                Return _certificateIssued
            End Get
            Set(ByVal value As Boolean)
                _certificateIssued = value
            End Set
        End Property

        Public Property CertificateNumber() As String
            Get
                Return _certificateNumber
            End Get
            Set(ByVal value As String)
                _certificateNumber = value
            End Set
        End Property

        Public Property CertificateIssuedDate() As DateTime?
            Get
                Return _certificateIssuedDate
            End Get
            Set(ByVal value As DateTime?)
                _certificateIssuedDate = value
            End Set
        End Property

        Public Property ExpiryDate() As DateTime?
            Get
                Return _expiryDate
            End Get
            Set(ByVal value As DateTime?)
                _expiryDate = value
            End Set
        End Property

        Public Property IsExpired() As Boolean
            Get
                Return _isExpired
            End Get
            Set(ByVal value As Boolean)
                _isExpired = value
            End Set
        End Property

        Public ReadOnly Property CalculatedIsExpired() As Boolean
            Get
                If _expiryDate.HasValue Then
                    Return _expiryDate.Value < DateTime.Now.Date
                End If
                Return False
            End Get
        End Property

        Public ReadOnly Property DaysUntilExpiry() As Integer?
            Get
                If _expiryDate.HasValue Then
                    Return CInt((_expiryDate.Value.Date - DateTime.Now.Date).TotalDays)
                End If
                Return Nothing
            End Get
        End Property

        Public ReadOnly Property IsNearExpiry() As Boolean
            Get
                Dim daysUntilExpiry = Me.DaysUntilExpiry
                Return daysUntilExpiry.HasValue AndAlso daysUntilExpiry.Value <= 90 AndAlso daysUntilExpiry.Value > 0
            End Get
        End Property

        Public Property Feedback() As String
            Get
                Return _feedback
            End Get
            Set(ByVal value As String)
                _feedback = value
            End Set
        End Property

        Public Property InstructorComments() As String
            Get
                Return _instructorComments
            End Get
            Set(ByVal value As String)
                _instructorComments = value
            End Set
        End Property

        Public Property CreatedDate() As DateTime
            Get
                Return _createdDate
            End Get
            Set(ByVal value As DateTime)
                _createdDate = value
            End Set
        End Property

        Public Property CreatedBy() As String
            Get
                Return _createdBy
            End Get
            Set(ByVal value As String)
                _createdBy = value
            End Set
        End Property

        Public Property ModifiedDate() As DateTime?
            Get
                Return _modifiedDate
            End Get
            Set(ByVal value As DateTime?)
                _modifiedDate = value
            End Set
        End Property

        Public Property ModifiedBy() As String
            Get
                Return _modifiedBy
            End Get
            Set(ByVal value As String)
                _modifiedBy = value
            End Set
        End Property

        Public Property RenewalRequired() As Boolean
            Get
                Return _renewalRequired
            End Get
            Set(ByVal value As Boolean)
                _renewalRequired = value
            End Set
        End Property

        Public Property RenewalNotificationSent() As Boolean
            Get
                Return _renewalNotificationSent
            End Get
            Set(ByVal value As Boolean)
                _renewalNotificationSent = value
            End Set
        End Property

        Public Property CostCentre() As String
            Get
                Return _costCentre
            End Get
            Set(ByVal value As String)
                _costCentre = value
            End Set
        End Property

        Public Property ApprovalRequired() As Boolean
            Get
                Return _approvalRequired
            End Get
            Set(ByVal value As Boolean)
                _approvalRequired = value
            End Set
        End Property

        Public Property ApprovedBy() As String
            Get
                Return _approvedBy
            End Get
            Set(ByVal value As String)
                _approvedBy = value
            End Set
        End Property

        Public Property ApprovedDate() As DateTime?
            Get
                Return _approvedDate
            End Get
            Set(ByVal value As DateTime?)
                _approvedDate = value
            End Set
        End Property

        Public ReadOnly Property IsApproved() As Boolean
            Get
                Return Not String.IsNullOrEmpty(_approvedBy) AndAlso _approvedDate.HasValue
            End Get
        End Property

        Public Property RejectedReason() As String
            Get
                Return _rejectedReason
            End Get
            Set(ByVal value As String)
                _rejectedReason = value
            End Set
        End Property

        Public ReadOnly Property IsRejected() As Boolean
            Get
                Return Not String.IsNullOrEmpty(_rejectedReason)
            End Get
        End Property

        ' Navigation properties
        Public Property Employee() As Employee
            Get
                Return _employee
            End Get
            Set(ByVal value As Employee)
                _employee = value
            End Set
        End Property

        Public Property Session() As TrainingSession
            Get
                Return _session
            End Get
            Set(ByVal value As TrainingSession)
                _session = value
            End Set
        End Property

        Public Property Course() As Course
            Get
                Return _course
            End Get
            Set(ByVal value As Course)
                _course = value
            End Set
        End Property

        ' Calculated properties
        Public ReadOnly Property ComplianceStatus() As String
            Get
                If Not IsCompleted Then
                    If _attendanceStatus = "NO_SHOW" Then
                        Return "Non-Compliant (No Show)"
                    ElseIf _completionStatus = "FAILED" Then
                        Return "Non-Compliant (Failed)"
                    ElseIf _attendanceStatus = "ENROLLED" Then
                        Return "Pending"
                    Else
                        Return "Non-Compliant"
                    End If
                ElseIf CalculatedIsExpired Then
                    Return "Expired"
                ElseIf IsNearExpiry Then
                    Return "Renewal Required"
                Else
                    Return "Compliant"
                End If
            End Get
        End Property

        Public ReadOnly Property StatusColor() As String
            Get
                Select Case ComplianceStatus
                    Case "Compliant"
                        Return "green"
                    Case "Pending"
                        Return "orange"
                    Case "Renewal Required"
                        Return "yellow"
                    Case "Expired"
                        Return "red"
                    Case Else
                        Return "red"
                End Select
            End Get
        End Property

        ' Constructor
        Public Sub New()
            _enrollmentDate = DateTime.Now
            _attendanceStatus = "ENROLLED"
            _completionStatus = "NOT_STARTED"
            _certificateIssued = False
            _isExpired = False
            _renewalRequired = False
            _renewalNotificationSent = False
            _approvalRequired = False
            _createdDate = DateTime.Now
        End Sub

        Public Sub New(employeeId As Integer, sessionId As Integer, courseId As Integer)
            Me.New()
            _employeeId = employeeId
            _sessionId = sessionId
            _courseId = courseId
        End Sub

        ' Validation methods
        Public Function IsValid() As Boolean
            Return _employeeId > 0 AndAlso
                   _sessionId > 0 AndAlso
                   _courseId > 0 AndAlso
                   Not String.IsNullOrEmpty(_attendanceStatus) AndAlso
                   Not String.IsNullOrEmpty(_completionStatus)
        End Function

        Public Function GetValidationErrors() As List(Of String)
            Dim errors As New List(Of String)()

            If _employeeId <= 0 Then
                errors.Add("Employee ID is required")
            End If

            If _sessionId <= 0 Then
                errors.Add("Session ID is required")
            End If

            If _courseId <= 0 Then
                errors.Add("Course ID is required")
            End If

            If String.IsNullOrEmpty(_attendanceStatus) Then
                errors.Add("Attendance status is required")
            End If

            If String.IsNullOrEmpty(_completionStatus) Then
                errors.Add("Completion status is required")
            End If

            If _score.HasValue AndAlso _maxScore.HasValue AndAlso _score.Value > _maxScore.Value Then
                errors.Add("Score cannot exceed maximum score")
            End If

            If _completionDate.HasValue AndAlso _completionDate.Value > DateTime.Now Then
                errors.Add("Completion date cannot be in the future")
            End If

            Return errors
        End Function

        ' Business methods
        Public Sub MarkAsCompleted(score As Integer?, grade As String, instructorComments As String)
            _completionDate = DateTime.Now
            _completionStatus = "COMPLETED"
            _attendanceStatus = "ATTENDED"
            _score = score
            _grade = grade
            _instructorComments = instructorComments
            _modifiedDate = DateTime.Now

            ' Calculate expiry date if course has validity period
            If _course IsNot Nothing AndAlso _course.ValidityPeriodMonths > 0 Then
                _expiryDate = _completionDate.Value.AddMonths(_course.ValidityPeriodMonths)
            End If
        End Sub

        Public Sub MarkAsNoShow()
            _attendanceStatus = "NO_SHOW"
            _completionStatus = "NOT_STARTED"
            _modifiedDate = DateTime.Now
        End Sub

        Public Sub MarkAsFailed(score As Integer?, comments As String)
            _completionDate = DateTime.Now
            _completionStatus = "FAILED"
            _attendanceStatus = "ATTENDED"
            _score = score
            _instructorComments = comments
            _modifiedDate = DateTime.Now
        End Sub

        Public Sub IssueCertificate(certificateNumber As String)
            If IsCompleted AndAlso (HasPassed Is Nothing OrElse HasPassed.Value) Then
                _certificateIssued = True
                _certificateNumber = certificateNumber
                _certificateIssuedDate = DateTime.Now
                _modifiedDate = DateTime.Now
            End If
        End Sub

        Public Function GenerateCertificateNumber() As String
            If _course IsNot Nothing AndAlso _employee IsNot Nothing Then
                Return String.Format("TT-{0}-{1}-{2:yyyyMMdd}",
                                    _course.CourseCode,
                                    _employee.EmployeeNumber,
                                    DateTime.Now)
            End If
            Return String.Format("TT-{0:yyyyMMdd}-{1}", DateTime.Now, _recordId)
        End Function

        ' Override methods
        Public Overrides Function ToString() As String
            Return String.Format("{0} - {1} ({2})",
                                _employee?.DisplayName,
                                _course?.DisplayTitle,
                                CompletionStatusDisplay)
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            If TypeOf obj Is TrainingRecord Then
                Dim other As TrainingRecord = CType(obj, TrainingRecord)
                Return _recordId = other.RecordId
            End If
            Return False
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return _recordId.GetHashCode()
        End Function

    End Class
End Namespace
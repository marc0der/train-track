Imports Microsoft.VisualBasic

Namespace Defra.TrainTrack.Models
    Public Class TrainingSession
        Private _sessionId As Integer
        Private _courseId As Integer
        Private _sessionDate As DateTime
        Private _startTime As TimeSpan
        Private _endTime As TimeSpan
        Private _location As String
        Private _maxParticipants As Integer
        Private _currentParticipants As Integer
        Private _primaryInstructorId As Integer
        Private _secondaryInstructorId As Integer?
        Private _sessionStatus As String
        Private _registrationDeadline As DateTime
        Private _sessionNotes As String
        Private _instructorNotes As String
        ' NOTE: Field exists but feature not implemented. Equipment/catering/rooms managed manually outside system.
        Private _equipmentRequired As String
        ' NOTE: Field exists but feature not implemented. Equipment/catering/rooms managed manually outside system.
        Private _cateringRequired As Boolean
        Private _materialsPrepared As Boolean
        ' NOTE: Field exists but feature not implemented. Equipment/catering/rooms managed manually outside system.
        Private _roomBooked As Boolean
        Private _notificationsSent As Boolean
        Private _waitingListEnabled As Boolean
        Private _prerequisitesChecked As Boolean
        Private _createdDate As DateTime
        Private _createdBy As String
        Private _modifiedDate As DateTime?
        Private _modifiedBy As String
        ' NOTE: Not used in current system. Costs tracked externally by Finance.
        Private _costPerParticipant As Decimal
        ' NOTE: Not used in current system. Costs tracked externally by Finance.
        Private _totalCost As Decimal
        Private _approvedBy As String
        Private _approvedDate As DateTime?
        Private _cancelledReason As String
        Private _feedbackRequested As Boolean

        ' Related objects
        Private _course As Course
        Private _primaryInstructor As Employee
        Private _secondaryInstructor As Employee
        Private _participants As List(Of Employee)
        Private _waitingList As List(Of WaitingListEntry)

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

        Public Property SessionDate() As DateTime
            Get
                Return _sessionDate
            End Get
            Set(ByVal value As DateTime)
                _sessionDate = value
            End Set
        End Property

        Public Property StartTime() As TimeSpan
            Get
                Return _startTime
            End Get
            Set(ByVal value As TimeSpan)
                _startTime = value
            End Set
        End Property

        Public Property EndTime() As TimeSpan
            Get
                Return _endTime
            End Get
            Set(ByVal value As TimeSpan)
                _endTime = value
            End Set
        End Property

        Public ReadOnly Property StartDateTime() As DateTime
            Get
                Return _sessionDate.Date.Add(_startTime)
            End Get
        End Property

        Public ReadOnly Property EndDateTime() As DateTime
            Get
                Return _sessionDate.Date.Add(_endTime)
            End Get
        End Property

        Public ReadOnly Property Duration() As TimeSpan
            Get
                Return _endTime.Subtract(_startTime)
            End Get
        End Property

        Public ReadOnly Property DurationText() As String
            Get
                Dim hours As Integer = Duration.Hours
                Dim minutes As Integer = Duration.Minutes
                If hours = 0 Then
                    Return String.Format("{0} minutes", minutes)
                ElseIf minutes = 0 Then
                    Return String.Format("{0} hours", hours)
                Else
                    Return String.Format("{0}h {1}m", hours, minutes)
                End If
            End Get
        End Property

        Public Property Location() As String
            Get
                Return _location
            End Get
            Set(ByVal value As String)
                _location = value
            End Set
        End Property

        Public Property MaxParticipants() As Integer
            Get
                Return _maxParticipants
            End Get
            Set(ByVal value As Integer)
                _maxParticipants = value
            End Set
        End Property

        Public Property CurrentParticipants() As Integer
            Get
                Return _currentParticipants
            End Get
            Set(ByVal value As Integer)
                _currentParticipants = value
            End Set
        End Property

        Public ReadOnly Property AvailableSpaces() As Integer
            Get
                Return Math.Max(0, _maxParticipants - _currentParticipants)
            End Get
        End Property

        Public ReadOnly Property IsFullyBooked() As Boolean
            Get
                Return _currentParticipants >= _maxParticipants
            End Get
        End Property

        Public Property PrimaryInstructorId() As Integer
            Get
                Return _primaryInstructorId
            End Get
            Set(ByVal value As Integer)
                _primaryInstructorId = value
            End Set
        End Property

        Public Property SecondaryInstructorId() As Integer?
            Get
                Return _secondaryInstructorId
            End Get
            Set(ByVal value As Integer?)
                _secondaryInstructorId = value
            End Set
        End Property

        Public Property SessionStatus() As String
            Get
                Return _sessionStatus
            End Get
            Set(ByVal value As String)
                _sessionStatus = value
            End Set
        End Property

        Public ReadOnly Property StatusDisplay() As String
            Get
                Select Case _sessionStatus.ToUpper()
                    Case "SCHEDULED"
                        Return "Scheduled"
                    Case "CONFIRMED"
                        Return "Confirmed"
                    Case "CANCELLED"
                        Return "Cancelled"
                    Case "COMPLETED"
                        Return "Completed"
                    Case "IN_PROGRESS"
                        Return "In Progress"
                    Case "PENDING_APPROVAL"
                        Return "Pending Approval"
                    Case Else
                        Return _sessionStatus
                End Select
            End Get
        End Property

        Public Property RegistrationDeadline() As DateTime
            Get
                Return _registrationDeadline
            End Get
            Set(ByVal value As DateTime)
                _registrationDeadline = value
            End Set
        End Property

        Public ReadOnly Property IsRegistrationOpen() As Boolean
            Get
                Return DateTime.Now <= _registrationDeadline AndAlso
                       (_sessionStatus = "SCHEDULED" OrElse _sessionStatus = "CONFIRMED") AndAlso
                       Not IsFullyBooked
            End Get
        End Property

        Public Property SessionNotes() As String
            Get
                Return _sessionNotes
            End Get
            Set(ByVal value As String)
                _sessionNotes = value
            End Set
        End Property

        Public Property InstructorNotes() As String
            Get
                Return _instructorNotes
            End Get
            Set(ByVal value As String)
                _instructorNotes = value
            End Set
        End Property

        Public Property EquipmentRequired() As String
            Get
                Return _equipmentRequired
            End Get
            Set(ByVal value As String)
                _equipmentRequired = value
            End Set
        End Property

        Public Property CateringRequired() As Boolean
            Get
                Return _cateringRequired
            End Get
            Set(ByVal value As Boolean)
                _cateringRequired = value
            End Set
        End Property

        Public Property MaterialsPrepared() As Boolean
            Get
                Return _materialsPrepared
            End Get
            Set(ByVal value As Boolean)
                _materialsPrepared = value
            End Set
        End Property

        Public Property RoomBooked() As Boolean
            Get
                Return _roomBooked
            End Get
            Set(ByVal value As Boolean)
                _roomBooked = value
            End Set
        End Property

        Public Property NotificationsSent() As Boolean
            Get
                Return _notificationsSent
            End Get
            Set(ByVal value As Boolean)
                _notificationsSent = value
            End Set
        End Property

        Public Property WaitingListEnabled() As Boolean
            Get
                Return _waitingListEnabled
            End Get
            Set(ByVal value As Boolean)
                _waitingListEnabled = value
            End Set
        End Property

        Public Property PrerequisitesChecked() As Boolean
            Get
                Return _prerequisitesChecked
            End Get
            Set(ByVal value As Boolean)
                _prerequisitesChecked = value
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

        Public Property CostPerParticipant() As Decimal
            Get
                Return _costPerParticipant
            End Get
            Set(ByVal value As Decimal)
                _costPerParticipant = value
            End Set
        End Property

        Public Property TotalCost() As Decimal
            Get
                Return _totalCost
            End Get
            Set(ByVal value As Decimal)
                _totalCost = value
            End Set
        End Property

        Public ReadOnly Property EstimatedTotalCost() As Decimal
            Get
                Return _costPerParticipant * _currentParticipants
            End Get
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

        Public Property CancelledReason() As String
            Get
                Return _cancelledReason
            End Get
            Set(ByVal value As String)
                _cancelledReason = value
            End Set
        End Property

        Public ReadOnly Property IsCancelled() As Boolean
            Get
                Return _sessionStatus.ToUpper() = "CANCELLED"
            End Get
        End Property

        Public Property FeedbackRequested() As Boolean
            Get
                Return _feedbackRequested
            End Get
            Set(ByVal value As Boolean)
                _feedbackRequested = value
            End Set
        End Property

        ' Navigation properties
        Public Property Course() As Course
            Get
                Return _course
            End Get
            Set(ByVal value As Course)
                _course = value
            End Set
        End Property

        Public Property PrimaryInstructor() As Employee
            Get
                Return _primaryInstructor
            End Get
            Set(ByVal value As Employee)
                _primaryInstructor = value
            End Set
        End Property

        Public Property SecondaryInstructor() As Employee
            Get
                Return _secondaryInstructor
            End Get
            Set(ByVal value As Employee)
                _secondaryInstructor = value
            End Set
        End Property

        Public Property Participants() As List(Of Employee)
            Get
                If _participants Is Nothing Then
                    _participants = New List(Of Employee)()
                End If
                Return _participants
            End Get
            Set(ByVal value As List(Of Employee))
                _participants = value
            End Set
        End Property

        Public Property WaitingList() As List(Of WaitingListEntry)
            Get
                If _waitingList Is Nothing Then
                    _waitingList = New List(Of WaitingListEntry)()
                End If
                Return _waitingList
            End Get
            Set(ByVal value As List(Of WaitingListEntry))
                _waitingList = value
            End Set
        End Property

        ' Calculated properties
        Public ReadOnly Property DaysUntilSession() As Integer
            Get
                Return CInt((_sessionDate.Date - DateTime.Now.Date).TotalDays)
            End Get
        End Property

        Public ReadOnly Property IsUpcoming() As Boolean
            Get
                Return _sessionDate.Date >= DateTime.Now.Date AndAlso Not IsCancelled
            End Get
        End Property

        Public ReadOnly Property IsPast() As Boolean
            Get
                Return _sessionDate.Date < DateTime.Now.Date
            End Get
        End Property

        Public ReadOnly Property IsToday() As Boolean
            Get
                Return _sessionDate.Date = DateTime.Now.Date
            End Get
        End Property

        Public ReadOnly Property ReadinessScore() As Integer
            Get
                Dim score As Integer = 0
                If _roomBooked Then score += 20
                If _materialsPrepared Then score += 20
                If _notificationsSent Then score += 20
                If _prerequisitesChecked Then score += 20
                If _currentParticipants >= 1 Then score += 20
                Return score
            End Get
        End Property

        Public ReadOnly Property ReadinessStatus() As String
            Get
                Dim score As Integer = ReadinessScore
                If score = 100 Then
                    Return "Ready"
                ElseIf score >= 80 Then
                    Return "Nearly Ready"
                ElseIf score >= 60 Then
                    Return "In Progress"
                ElseIf score >= 40 Then
                    Return "Early Stage"
                Else
                    Return "Not Started"
                End If
            End Get
        End Property

        ' Constructor
        Public Sub New()
            _sessionStatus = "SCHEDULED"
            _cateringRequired = False
            _materialsPrepared = False
            _roomBooked = False
            _notificationsSent = False
            _waitingListEnabled = True
            _prerequisitesChecked = False
            _createdDate = DateTime.Now
            _feedbackRequested = False
            _currentParticipants = 0
            _participants = New List(Of Employee)()
        End Sub

        ' Validation methods
        Public Function IsValid() As Boolean
            Return _sessionDate > DateTime.Now AndAlso
                   _startTime < _endTime AndAlso
                   Not String.IsNullOrEmpty(_location) AndAlso
                   _maxParticipants > 0 AndAlso
                   _primaryInstructorId > 0 AndAlso
                   _registrationDeadline <= _sessionDate
        End Function

        Public Function GetValidationErrors() As List(Of String)
            Dim errors As New List(Of String)()

            If _sessionDate <= DateTime.Now Then
                errors.Add("Session date must be in the future")
            End If

            If _startTime >= _endTime Then
                errors.Add("End time must be after start time")
            End If

            If String.IsNullOrEmpty(_location) Then
                errors.Add("Location is required")
            End If

            If _maxParticipants <= 0 Then
                errors.Add("Maximum participants must be greater than 0")
            End If

            If _primaryInstructorId <= 0 Then
                errors.Add("Primary instructor is required")
            End If

            If _registrationDeadline > _sessionDate Then
                errors.Add("Registration deadline cannot be after session date")
            End If

            Return errors
        End Function

        ' Business methods
        Public Function CanAddParticipant() As Boolean
            Return _currentParticipants < _maxParticipants AndAlso
                   IsRegistrationOpen AndAlso
                   Not IsCancelled
        End Function

        Public Function GetPreparationTasks() As List(Of String)
            Dim tasks As New List(Of String)()

            If Not _roomBooked Then
                tasks.Add("Book training room")
            End If

            If Not _materialsPrepared Then
                tasks.Add("Prepare training materials")
            End If

            If Not _notificationsSent Then
                tasks.Add("Send participant notifications")
            End If

            If Not _prerequisitesChecked Then
                tasks.Add("Check participant prerequisites")
            End If

            If _cateringRequired Then
                tasks.Add("Arrange catering")
            End If

            Return tasks
        End Function

        Public Function CalculateCompletionPercentage() As Integer
            If _currentParticipants = 0 Then
                Return 0
            End If

            ' This would typically be calculated based on actual completion records
            ' For now, return a placeholder value
            If _sessionStatus = "COMPLETED" Then
                Return 100
            ElseIf _sessionStatus = "IN_PROGRESS" Then
                Return 50
            Else
                Return 0
            End If
        End Function

        ' Override methods
        Public Overrides Function ToString() As String
            Return String.Format("{0} - {1:dd/MM/yyyy} at {2:hh\:mm}", _course?.DisplayTitle, _sessionDate, _startTime)
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            If TypeOf obj Is TrainingSession Then
                Dim other As TrainingSession = CType(obj, TrainingSession)
                Return _sessionId = other.SessionId
            End If
            Return False
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return _sessionId.GetHashCode()
        End Function

    End Class
End Namespace
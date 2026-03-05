Imports Microsoft.VisualBasic

Namespace Defra.TrainTrack.Models
    Public Class Course
        Private _courseId As Integer
        Private _courseCode As String
        Private _title As String
        Private _description As String
        Private _category As String
        Private _durationHours As Decimal
        Private _deliveryMethod As String
        Private _isActive As Boolean
        Private _isCompulsory As Boolean
        Private _maxParticipants As Integer
        Private _minParticipants As Integer
        Private _prerequisites As String
        Private _learningObjectives As String
        Private _courseContent As String
        Private _modules As List(Of CourseModule)
        Private _assessmentMethod As String
        Private _certificateTemplate As String
        Private _validityPeriodMonths As Integer
        ' NOTE: Field exists in schema but is not populated. Cost tracking is manual via Finance team.
        Private _costPerParticipant As Decimal
        Private _approvalRequired As Boolean
        Private _createdDate As DateTime
        Private _createdBy As String
        Private _modifiedDate As DateTime?
        Private _modifiedBy As String
        Private _version As String
        Private _skillsFrameworkLevel As String
        Private _complianceCategory As String
        Private _externalProvider As String
        Private _providerContactEmail As String
        Private _courseMaterials As String
        Private _equipmentRequired As String

        Public Property CourseId() As Integer
            Get
                Return _courseId
            End Get
            Set(ByVal value As Integer)
                _courseId = value
            End Set
        End Property

        Public Property CourseCode() As String
            Get
                Return _courseCode
            End Get
            Set(ByVal value As String)
                _courseCode = value
            End Set
        End Property

        Public Property Title() As String
            Get
                Return _title
            End Get
            Set(ByVal value As String)
                _title = value
            End Set
        End Property

        Public Property Description() As String
            Get
                Return _description
            End Get
            Set(ByVal value As String)
                _description = value
            End Set
        End Property

        Public Property Category() As String
            Get
                Return _category
            End Get
            Set(ByVal value As String)
                _category = value
            End Set
        End Property

        Public Property DurationHours() As Decimal
            Get
                Return _durationHours
            End Get
            Set(ByVal value As Decimal)
                _durationHours = value
            End Set
        End Property

        Public ReadOnly Property DurationText() As String
            Get
                If _durationHours = 1 Then
                    Return "1 hour"
                ElseIf _durationHours < 1 Then
                    Return String.Format("{0} minutes", _durationHours * 60)
                Else
                    Return String.Format("{0} hours", _durationHours)
                End If
            End Get
        End Property

        Public Property DeliveryMethod() As String
            Get
                Return _deliveryMethod
            End Get
            Set(ByVal value As String)
                _deliveryMethod = value
            End Set
        End Property

        Public Property IsActive() As Boolean
            Get
                Return _isActive
            End Get
            Set(ByVal value As Boolean)
                _isActive = value
            End Set
        End Property

        Public Property IsCompulsory() As Boolean
            Get
                Return _isCompulsory
            End Get
            Set(ByVal value As Boolean)
                _isCompulsory = value
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

        Public Property MinParticipants() As Integer
            Get
                Return _minParticipants
            End Get
            Set(ByVal value As Integer)
                _minParticipants = value
            End Set
        End Property

        Public Property Prerequisites() As String
            Get
                Return _prerequisites
            End Get
            Set(ByVal value As String)
                _prerequisites = value
            End Set
        End Property

        Public Property LearningObjectives() As String
            Get
                Return _learningObjectives
            End Get
            Set(ByVal value As String)
                _learningObjectives = value
            End Set
        End Property

        Public Property CourseContent() As String
            Get
                Return _courseContent
            End Get
            Set(ByVal value As String)
                _courseContent = value
            End Set
        End Property

        Public Property Modules() As List(Of CourseModule)
            Get
                Return _modules
            End Get
            Set(ByVal value As List(Of CourseModule))
                _modules = value
            End Set
        End Property

        Public Property AssessmentMethod() As String
            Get
                Return _assessmentMethod
            End Get
            Set(ByVal value As String)
                _assessmentMethod = value
            End Set
        End Property

        Public Property CertificateTemplate() As String
            Get
                Return _certificateTemplate
            End Get
            Set(ByVal value As String)
                _certificateTemplate = value
            End Set
        End Property

        Public Property ValidityPeriodMonths() As Integer
            Get
                Return _validityPeriodMonths
            End Get
            Set(ByVal value As Integer)
                _validityPeriodMonths = value
            End Set
        End Property

        Public ReadOnly Property ValidityPeriodText() As String
            Get
                If _validityPeriodMonths = 0 Then
                    Return "No expiry"
                ElseIf _validityPeriodMonths = 1 Then
                    Return "1 month"
                ElseIf _validityPeriodMonths = 12 Then
                    Return "1 year"
                ElseIf _validityPeriodMonths Mod 12 = 0 Then
                    Return String.Format("{0} years", _validityPeriodMonths \ 12)
                Else
                    Return String.Format("{0} months", _validityPeriodMonths)
                End If
            End Get
        End Property

        Public Property CostPerParticipant() As Decimal
            Get
                Return _costPerParticipant
            End Get
            Set(ByVal value As Decimal)
                _costPerParticipant = value
            End Set
        End Property

        Public ReadOnly Property CostText() As String
            Get
                If _costPerParticipant = 0 Then
                    Return "Not tracked"
                Else
                    Return String.Format("£{0:N2}", _costPerParticipant)
                End If
            End Get
        End Property

        Public Property ApprovalRequired() As Boolean
            Get
                Return _approvalRequired
            End Get
            Set(ByVal value As Boolean)
                _approvalRequired = value
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

        Public Property Version() As String
            Get
                Return _version
            End Get
            Set(ByVal value As String)
                _version = value
            End Set
        End Property

        Public Property SkillsFrameworkLevel() As String
            Get
                Return _skillsFrameworkLevel
            End Get
            Set(ByVal value As String)
                _skillsFrameworkLevel = value
            End Set
        End Property

        Public Property ComplianceCategory() As String
            Get
                Return _complianceCategory
            End Get
            Set(ByVal value As String)
                _complianceCategory = value
            End Set
        End Property

        Public Property ExternalProvider() As String
            Get
                Return _externalProvider
            End Get
            Set(ByVal value As String)
                _externalProvider = value
            End Set
        End Property

        Public Property ProviderContactEmail() As String
            Get
                Return _providerContactEmail
            End Get
            Set(ByVal value As String)
                _providerContactEmail = value
            End Set
        End Property

        Public Property CourseMaterials() As String
            Get
                Return _courseMaterials
            End Get
            Set(ByVal value As String)
                _courseMaterials = value
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

        ' Calculated properties
        Public ReadOnly Property TotalModuleDuration() As Integer
            Get
                If _modules Is Nothing OrElse _modules.Count = 0 Then
                    Return 0
                End If
                Dim total As Integer = 0
                For Each m As CourseModule In _modules
                    total += m.DurationMinutes
                Next
                Return total
            End Get
        End Property

        Public ReadOnly Property IsExternal() As Boolean
            Get
                Return Not String.IsNullOrEmpty(_externalProvider)
            End Get
        End Property

        Public ReadOnly Property HasPrerequisites() As Boolean
            Get
                Return Not String.IsNullOrEmpty(_prerequisites) AndAlso _prerequisites.ToLower() <> "none"
            End Get
        End Property

        Public ReadOnly Property RequiresCertification() As Boolean
            Get
                Return Not String.IsNullOrEmpty(_certificateTemplate)
            End Get
        End Property

        Public ReadOnly Property DisplayTitle() As String
            Get
                Return String.Format("{0} - {1}", _courseCode, _title)
            End Get
        End Property

        ' Constructor
        Public Sub New()
            _isActive = True
            _isCompulsory = False
            _approvalRequired = False
            _minParticipants = 1
            _maxParticipants = 20
            _validityPeriodMonths = 12
            _createdDate = DateTime.Now
            _version = "1.0"
            _modules = New List(Of CourseModule)()
        End Sub

        Public Sub New(courseCode As String, title As String, category As String)
            Me.New()
            _courseCode = courseCode
            _title = title
            _category = category
        End Sub

        ' Validation methods
        Public Function IsValid() As Boolean
            Return Not String.IsNullOrEmpty(_courseCode) AndAlso
                   Not String.IsNullOrEmpty(_title) AndAlso
                   Not String.IsNullOrEmpty(_category) AndAlso
                   _durationHours > 0 AndAlso
                   _maxParticipants > 0 AndAlso
                   _minParticipants > 0 AndAlso
                   _minParticipants <= _maxParticipants
        End Function

        Public Function GetValidationErrors() As List(Of String)
            Dim errors As New List(Of String)()

            If String.IsNullOrEmpty(_courseCode) Then
                errors.Add("Course code is required")
            End If

            If String.IsNullOrEmpty(_title) Then
                errors.Add("Course title is required")
            End If

            If String.IsNullOrEmpty(_category) Then
                errors.Add("Course category is required")
            End If

            If _durationHours <= 0 Then
                errors.Add("Duration must be greater than 0")
            End If

            If _maxParticipants <= 0 Then
                errors.Add("Maximum participants must be greater than 0")
            End If

            If _minParticipants <= 0 Then
                errors.Add("Minimum participants must be greater than 0")
            End If

            If _minParticipants > _maxParticipants Then
                errors.Add("Minimum participants cannot exceed maximum participants")
            End If

            If String.IsNullOrEmpty(_deliveryMethod) Then
                errors.Add("Delivery method is required")
            End If

            Return errors
        End Function

        ' Business methods
        Public Function CalculateTotalCost(participantCount As Integer) As Decimal
            If participantCount <= 0 Then
                Return 0
            End If
            Return _costPerParticipant * participantCount
        End Function

        Public Function CanAcceptParticipants(currentParticipants As Integer) As Boolean
            Return currentParticipants < _maxParticipants AndAlso _isActive
        End Function

        Public Function GetAvailableSpaces(currentParticipants As Integer) As Integer
            If currentParticipants >= _maxParticipants Then
                Return 0
            End If
            Return _maxParticipants - currentParticipants
        End Function

        ' Override methods
        Public Overrides Function ToString() As String
            Return DisplayTitle
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            If TypeOf obj Is Course Then
                Dim other As Course = CType(obj, Course)
                Return _courseId = other.CourseId
            End If
            Return False
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return _courseId.GetHashCode()
        End Function

    End Class
End Namespace
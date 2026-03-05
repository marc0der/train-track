Imports Microsoft.VisualBasic

Namespace Defra.TrainTrack.Models
    Public Class Employee
        Private _employeeId As Integer
        Private _employeeNumber As String
        Private _firstName As String
        Private _lastName As String
        Private _email As String
        Private _userName As String
        Private _department As String
        Private _position As String
        Private _location As String
        Private _managerId As Integer?
        Private _managerName As String
        Private _hireDate As DateTime
        Private _lastLoginDate As DateTime
        Private _isActive As Boolean
        Private _phoneNumber As String
        Private _lineManagerEmail As String
        Private _costCentre As String
        Private _payBand As String
        Private _contractType As String
        Private _workingPattern As String
        Private _createdDate As DateTime
        Private _createdBy As String
        Private _modifiedDate As DateTime?
        Private _modifiedBy As String

        Public Property EmployeeId() As Integer
            Get
                Return _employeeId
            End Get
            Set(ByVal value As Integer)
                _employeeId = value
            End Set
        End Property

        Public Property EmployeeNumber() As String
            Get
                Return _employeeNumber
            End Get
            Set(ByVal value As String)
                _employeeNumber = value
            End Set
        End Property

        Public Property FirstName() As String
            Get
                Return _firstName
            End Get
            Set(ByVal value As String)
                _firstName = value
            End Set
        End Property

        Public Property LastName() As String
            Get
                Return _lastName
            End Get
            Set(ByVal value As String)
                _lastName = value
            End Set
        End Property

        Public ReadOnly Property FullName() As String
            Get
                Return String.Format("{0}, {1}", _lastName, _firstName)
            End Get
        End Property

        Public ReadOnly Property DisplayName() As String
            Get
                Return String.Format("{0} {1}", _firstName, _lastName)
            End Get
        End Property

        Public Property Email() As String
            Get
                Return _email
            End Get
            Set(ByVal value As String)
                _email = value
            End Set
        End Property

        Public Property UserName() As String
            Get
                Return _userName
            End Get
            Set(ByVal value As String)
                _userName = value
            End Set
        End Property

        Public Property Department() As String
            Get
                Return _department
            End Get
            Set(ByVal value As String)
                _department = value
            End Set
        End Property

        Public Property Position() As String
            Get
                Return _position
            End Get
            Set(ByVal value As String)
                _position = value
            End Set
        End Property

        Public Property Location() As String
            Get
                Return _location
            End Get
            Set(ByVal value As String)
                _location = value
            End Set
        End Property

        Public Property ManagerId() As Integer?
            Get
                Return _managerId
            End Get
            Set(ByVal value As Integer?)
                _managerId = value
            End Set
        End Property

        Public Property ManagerName() As String
            Get
                Return _managerName
            End Get
            Set(ByVal value As String)
                _managerName = value
            End Set
        End Property

        Public Property HireDate() As DateTime
            Get
                Return _hireDate
            End Get
            Set(ByVal value As DateTime)
                _hireDate = value
            End Set
        End Property

        Public Property LastLoginDate() As DateTime
            Get
                Return _lastLoginDate
            End Get
            Set(ByVal value As DateTime)
                _lastLoginDate = value
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

        Public Property PhoneNumber() As String
            Get
                Return _phoneNumber
            End Get
            Set(ByVal value As String)
                _phoneNumber = value
            End Set
        End Property

        Public Property LineManagerEmail() As String
            Get
                Return _lineManagerEmail
            End Get
            Set(ByVal value As String)
                _lineManagerEmail = value
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

        Public Property PayBand() As String
            Get
                Return _payBand
            End Get
            Set(ByVal value As String)
                _payBand = value
            End Set
        End Property

        Public Property ContractType() As String
            Get
                Return _contractType
            End Get
            Set(ByVal value As String)
                _contractType = value
            End Set
        End Property

        Public Property WorkingPattern() As String
            Get
                Return _workingPattern
            End Get
            Set(ByVal value As String)
                _workingPattern = value
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

        ' Calculated properties
        Public ReadOnly Property YearsOfService() As Integer
            Get
                Return DateTime.Now.Year - _hireDate.Year
            End Get
        End Property

        Public ReadOnly Property IsNewStarter() As Boolean
            Get
                Return (DateTime.Now - _hireDate).TotalDays <= 90
            End Get
        End Property

        Public ReadOnly Property DaysUntilHireAnniversary() As Integer
            Get
                Dim thisYearAnniversary As New DateTime(DateTime.Now.Year, _hireDate.Month, _hireDate.Day)
                If thisYearAnniversary < DateTime.Now Then
                    thisYearAnniversary = thisYearAnniversary.AddYears(1)
                End If
                Return CInt((thisYearAnniversary - DateTime.Now).TotalDays)
            End Get
        End Property

        ' Constructor
        Public Sub New()
            _isActive = True
            _createdDate = DateTime.Now
            _hireDate = DateTime.Now
            _lastLoginDate = DateTime.MinValue
        End Sub

        Public Sub New(employeeNumber As String, firstName As String, lastName As String, email As String)
            Me.New()
            _employeeNumber = employeeNumber
            _firstName = firstName
            _lastName = lastName
            _email = email
        End Sub

        ' Validation methods
        Public Function IsValid() As Boolean
            Return Not String.IsNullOrEmpty(_employeeNumber) AndAlso
                   Not String.IsNullOrEmpty(_firstName) AndAlso
                   Not String.IsNullOrEmpty(_lastName) AndAlso
                   Not String.IsNullOrEmpty(_email) AndAlso
                   IsValidEmail(_email)
        End Function

        Private Function IsValidEmail(email As String) As Boolean
            Try
                Dim addr As New System.Net.Mail.MailAddress(email)
                Return addr.Address = email
            Catch
                Return False
            End Try
        End Function

        Public Function GetValidationErrors() As List(Of String)
            Dim errors As New List(Of String)()

            If String.IsNullOrEmpty(_employeeNumber) Then
                errors.Add("Employee number is required")
            End If

            If String.IsNullOrEmpty(_firstName) Then
                errors.Add("First name is required")
            End If

            If String.IsNullOrEmpty(_lastName) Then
                errors.Add("Last name is required")
            End If

            If String.IsNullOrEmpty(_email) Then
                errors.Add("Email address is required")
            ElseIf Not IsValidEmail(_email) Then
                errors.Add("Email address is not valid")
            End If

            If String.IsNullOrEmpty(_department) Then
                errors.Add("Department is required")
            End If

            If String.IsNullOrEmpty(_position) Then
                errors.Add("Position is required")
            End If

            Return errors
        End Function

        ' Override methods
        Public Overrides Function ToString() As String
            Return String.Format("{0} ({1})", FullName, _employeeNumber)
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            If TypeOf obj Is Employee Then
                Dim other As Employee = CType(obj, Employee)
                Return _employeeId = other.EmployeeId
            End If
            Return False
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return _employeeId.GetHashCode()
        End Function

    End Class
End Namespace
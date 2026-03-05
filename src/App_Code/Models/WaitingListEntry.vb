Imports Microsoft.VisualBasic

Namespace Defra.TrainTrack.Models
    Public Class WaitingListEntry
        Private _entryId As Integer
        Private _sessionId As Integer
        Private _employeeId As Integer
        Private _requestDate As DateTime
        Private _position As Integer
        Private _status As String
        Private _offeredDate As DateTime?
        Private _responseDeadline As DateTime?
        Private _notes As String
        Private _createdBy As String

        Public Property EntryId() As Integer
            Get
                Return _entryId
            End Get
            Set(ByVal value As Integer)
                _entryId = value
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

        Public Property EmployeeId() As Integer
            Get
                Return _employeeId
            End Get
            Set(ByVal value As Integer)
                _employeeId = value
            End Set
        End Property

        Public Property RequestDate() As DateTime
            Get
                Return _requestDate
            End Get
            Set(ByVal value As DateTime)
                _requestDate = value
            End Set
        End Property

        Public Property Position() As Integer
            Get
                Return _position
            End Get
            Set(ByVal value As Integer)
                _position = value
            End Set
        End Property

        Public Property Status() As String
            Get
                Return _status
            End Get
            Set(ByVal value As String)
                _status = value
            End Set
        End Property

        Public Property OfferedDate() As DateTime?
            Get
                Return _offeredDate
            End Get
            Set(ByVal value As DateTime?)
                _offeredDate = value
            End Set
        End Property

        Public Property ResponseDeadline() As DateTime?
            Get
                Return _responseDeadline
            End Get
            Set(ByVal value As DateTime?)
                _responseDeadline = value
            End Set
        End Property

        Public Property Notes() As String
            Get
                Return _notes
            End Get
            Set(ByVal value As String)
                _notes = value
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

        ' Calculated properties
        Public ReadOnly Property IsActive() As Boolean
            Get
                Return _status = "Waiting" OrElse _status = "Offered"
            End Get
        End Property

        Public ReadOnly Property DaysWaiting() As Integer
            Get
                Return CInt((DateTime.Now - _requestDate).TotalDays)
            End Get
        End Property

        ' Constructor
        Public Sub New()
            _status = "Waiting"
            _position = 1
            _requestDate = DateTime.Now
        End Sub

        Public Sub New(sessionId As Integer, employeeId As Integer, createdBy As String)
            Me.New()
            _sessionId = sessionId
            _employeeId = employeeId
            _createdBy = createdBy
        End Sub

        ' Validation methods
        Public Function IsValid() As Boolean
            Return _sessionId > 0 AndAlso
                   _employeeId > 0 AndAlso
                   _position > 0 AndAlso
                   Not String.IsNullOrEmpty(_status) AndAlso
                   Not String.IsNullOrEmpty(_createdBy)
        End Function

        Public Function GetValidationErrors() As List(Of String)
            Dim errors As New List(Of String)()

            If _sessionId <= 0 Then
                errors.Add("Session ID is required")
            End If

            If _employeeId <= 0 Then
                errors.Add("Employee ID is required")
            End If

            If _position <= 0 Then
                errors.Add("Position must be greater than 0")
            End If

            If String.IsNullOrEmpty(_status) Then
                errors.Add("Status is required")
            End If

            If String.IsNullOrEmpty(_createdBy) Then
                errors.Add("Created by is required")
            End If

            Return errors
        End Function

        ' Override methods
        Public Overrides Function ToString() As String
            Return String.Format("Waiting list #{0}: Position {1} ({2})", _entryId, _position, _status)
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            If TypeOf obj Is WaitingListEntry Then
                Dim other As WaitingListEntry = CType(obj, WaitingListEntry)
                Return _entryId = other.EntryId
            End If
            Return False
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return _entryId.GetHashCode()
        End Function

    End Class
End Namespace

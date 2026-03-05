Imports Microsoft.VisualBasic

Namespace Defra.TrainTrack.Models
    Public Class EmployeeNote
        Private _noteId As Integer
        Private _employeeId As Integer
        Private _noteText As String
        Private _noteType As String
        Private _createdDate As DateTime
        Private _createdBy As String
        Private _isResolved As Boolean
        Private _resolvedDate As DateTime?
        Private _resolvedBy As String

        Public Property NoteId() As Integer
            Get
                Return _noteId
            End Get
            Set(ByVal value As Integer)
                _noteId = value
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

        Public Property NoteText() As String
            Get
                Return _noteText
            End Get
            Set(ByVal value As String)
                _noteText = value
            End Set
        End Property

        Public Property NoteType() As String
            Get
                Return _noteType
            End Get
            Set(ByVal value As String)
                _noteType = value
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

        Public Property IsResolved() As Boolean
            Get
                Return _isResolved
            End Get
            Set(ByVal value As Boolean)
                _isResolved = value
            End Set
        End Property

        Public Property ResolvedDate() As DateTime?
            Get
                Return _resolvedDate
            End Get
            Set(ByVal value As DateTime?)
                _resolvedDate = value
            End Set
        End Property

        Public Property ResolvedBy() As String
            Get
                Return _resolvedBy
            End Get
            Set(ByVal value As String)
                _resolvedBy = value
            End Set
        End Property

        ' Calculated properties
        Public ReadOnly Property DaysSinceCreated() As Integer
            Get
                Return CInt((DateTime.Now - _createdDate).TotalDays)
            End Get
        End Property

        ' Constructor
        Public Sub New()
            _noteType = "General"
            _isResolved = False
            _createdDate = DateTime.Now
        End Sub

        Public Sub New(noteText As String, noteType As String, createdBy As String)
            Me.New()
            _noteText = noteText
            _noteType = noteType
            _createdBy = createdBy
        End Sub

        ' Validation methods
        Public Function IsValid() As Boolean
            Return Not String.IsNullOrEmpty(_noteText) AndAlso
                   Not String.IsNullOrEmpty(_noteType) AndAlso
                   Not String.IsNullOrEmpty(_createdBy) AndAlso
                   _employeeId > 0
        End Function

        Public Function GetValidationErrors() As List(Of String)
            Dim errors As New List(Of String)()

            If String.IsNullOrEmpty(_noteText) Then
                errors.Add("Note text is required")
            End If

            If String.IsNullOrEmpty(_noteType) Then
                errors.Add("Note type is required")
            End If

            If String.IsNullOrEmpty(_createdBy) Then
                errors.Add("Created by is required")
            End If

            If _employeeId <= 0 Then
                errors.Add("Employee ID is required")
            End If

            Return errors
        End Function

        ' Override methods
        Public Overrides Function ToString() As String
            Return String.Format("{0} note ({1})", _noteType, _createdDate.ToString("dd/MM/yyyy"))
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            If TypeOf obj Is EmployeeNote Then
                Dim other As EmployeeNote = CType(obj, EmployeeNote)
                Return _noteId = other.NoteId
            End If
            Return False
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return _noteId.GetHashCode()
        End Function

    End Class
End Namespace

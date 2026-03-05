Imports Microsoft.VisualBasic

Namespace Defra.TrainTrack.Models
    Public Class CourseModule
        Private _moduleId As Integer
        Private _courseId As Integer
        Private _moduleOrder As Integer
        Private _title As String
        Private _moduleType As String
        Private _durationMinutes As Integer
        Private _description As String
        Private _createdDate As DateTime
        Private _modifiedDate As DateTime?

        Public Property ModuleId() As Integer
            Get
                Return _moduleId
            End Get
            Set(ByVal value As Integer)
                _moduleId = value
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

        Public Property ModuleOrder() As Integer
            Get
                Return _moduleOrder
            End Get
            Set(ByVal value As Integer)
                _moduleOrder = value
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

        Public Property ModuleType() As String
            Get
                Return _moduleType
            End Get
            Set(ByVal value As String)
                _moduleType = value
            End Set
        End Property

        Public Property DurationMinutes() As Integer
            Get
                Return _durationMinutes
            End Get
            Set(ByVal value As Integer)
                _durationMinutes = value
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

        Public Property CreatedDate() As DateTime
            Get
                Return _createdDate
            End Get
            Set(ByVal value As DateTime)
                _createdDate = value
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

        ' Calculated properties
        Public ReadOnly Property DurationText() As String
            Get
                If _durationMinutes <= 0 Then
                    Return "0 mins"
                ElseIf _durationMinutes < 60 Then
                    Return String.Format("{0} mins", _durationMinutes)
                ElseIf _durationMinutes = 60 Then
                    Return "1 hr"
                ElseIf _durationMinutes Mod 60 = 0 Then
                    Return String.Format("{0} hrs", _durationMinutes \ 60)
                Else
                    Return String.Format("{0} hrs {1} mins", _durationMinutes \ 60, _durationMinutes Mod 60)
                End If
            End Get
        End Property

        ' Constructor
        Public Sub New()
            _moduleOrder = 1
            _durationMinutes = 0
            _createdDate = DateTime.Now
        End Sub

        Public Sub New(title As String, moduleType As String, durationMinutes As Integer)
            Me.New()
            _title = title
            _moduleType = moduleType
            _durationMinutes = durationMinutes
        End Sub

        ' Validation methods
        Public Function IsValid() As Boolean
            Return Not String.IsNullOrEmpty(_title) AndAlso
                   _moduleOrder > 0
        End Function

        Public Function GetValidationErrors() As List(Of String)
            Dim errors As New List(Of String)()

            If String.IsNullOrEmpty(_title) Then
                errors.Add("Module title is required")
            End If

            If _moduleOrder <= 0 Then
                errors.Add("Module order must be greater than 0")
            End If

            Return errors
        End Function

        ' Override methods
        Public Overrides Function ToString() As String
            Return String.Format("Module {0}: {1}", _moduleOrder, _title)
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            If TypeOf obj Is CourseModule Then
                Dim other As CourseModule = CType(obj, CourseModule)
                Return _moduleId = other.ModuleId
            End If
            Return False
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return _moduleId.GetHashCode()
        End Function

    End Class
End Namespace

Imports Microsoft.VisualBasic

Namespace Defra.TrainTrack.Models
    Public Class ELearningSyncRecord
        Private _syncId As Integer
        Private _syncDate As DateTime
        Private _syncStatus As String
        Private _totalRecords As Integer
        Private _processedRecords As Integer
        Private _failedRecords As Integer
        Private _errorDetails As String
        Private _batchFilePath As String
        Private _startTime As DateTime
        Private _endTime As DateTime?

        Public Property SyncId() As Integer
            Get
                Return _syncId
            End Get
            Set(ByVal value As Integer)
                _syncId = value
            End Set
        End Property

        Public Property SyncDate() As DateTime
            Get
                Return _syncDate
            End Get
            Set(ByVal value As DateTime)
                _syncDate = value
            End Set
        End Property

        Public Property SyncStatus() As String
            Get
                Return _syncStatus
            End Get
            Set(ByVal value As String)
                _syncStatus = value
            End Set
        End Property

        Public Property TotalRecords() As Integer
            Get
                Return _totalRecords
            End Get
            Set(ByVal value As Integer)
                _totalRecords = value
            End Set
        End Property

        Public Property ProcessedRecords() As Integer
            Get
                Return _processedRecords
            End Get
            Set(ByVal value As Integer)
                _processedRecords = value
            End Set
        End Property

        Public Property FailedRecords() As Integer
            Get
                Return _failedRecords
            End Get
            Set(ByVal value As Integer)
                _failedRecords = value
            End Set
        End Property

        Public Property ErrorDetails() As String
            Get
                Return _errorDetails
            End Get
            Set(ByVal value As String)
                _errorDetails = value
            End Set
        End Property

        Public Property BatchFilePath() As String
            Get
                Return _batchFilePath
            End Get
            Set(ByVal value As String)
                _batchFilePath = value
            End Set
        End Property

        Public Property StartTime() As DateTime
            Get
                Return _startTime
            End Get
            Set(ByVal value As DateTime)
                _startTime = value
            End Set
        End Property

        Public Property EndTime() As DateTime?
            Get
                Return _endTime
            End Get
            Set(ByVal value As DateTime?)
                _endTime = value
            End Set
        End Property

        ' Calculated properties
        Public ReadOnly Property SuccessRate() As Double
            Get
                If _totalRecords = 0 Then Return 0.0
                Return (_processedRecords / CDbl(_totalRecords)) * 100.0
            End Get
        End Property

        Public ReadOnly Property Duration() As TimeSpan?
            Get
                If _endTime.HasValue Then
                    Return _endTime.Value - _startTime
                End If
                Return Nothing
            End Get
        End Property

        ' Constructor
        Public Sub New()
            _syncDate = DateTime.Now
            _startTime = DateTime.Now
            _syncStatus = "Failed"
            _totalRecords = 0
            _processedRecords = 0
            _failedRecords = 0
        End Sub

        Public Sub New(batchFilePath As String)
            Me.New()
            _batchFilePath = batchFilePath
        End Sub

        ' Validation methods
        Public Function IsValid() As Boolean
            Return Not String.IsNullOrEmpty(_syncStatus) AndAlso
                   _totalRecords >= 0 AndAlso
                   _processedRecords >= 0 AndAlso
                   _failedRecords >= 0
        End Function

        Public Function GetValidationErrors() As List(Of String)
            Dim errors As New List(Of String)()

            If String.IsNullOrEmpty(_syncStatus) Then
                errors.Add("Sync status is required")
            End If

            If _totalRecords < 0 Then
                errors.Add("Total records cannot be negative")
            End If

            If _processedRecords < 0 Then
                errors.Add("Processed records cannot be negative")
            End If

            If _failedRecords < 0 Then
                errors.Add("Failed records cannot be negative")
            End If

            Return errors
        End Function

        ' Override methods
        Public Overrides Function ToString() As String
            Return String.Format("Sync #{0}: {1} on {2} ({3}/{4} records)", _syncId, _syncStatus, _syncDate.ToString("dd/MM/yyyy"), _processedRecords, _totalRecords)
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            If TypeOf obj Is ELearningSyncRecord Then
                Dim other As ELearningSyncRecord = CType(obj, ELearningSyncRecord)
                Return _syncId = other.SyncId
            End If
            Return False
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return _syncId.GetHashCode()
        End Function

    End Class
End Namespace

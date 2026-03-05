Imports Microsoft.VisualBasic

Namespace Defra.TrainTrack.Models
    Public Class EmployeeImportResult
        Private _totalRows As Integer
        Private _newEmployees As Integer
        Private _updatedEmployees As Integer
        Private _skippedRows As Integer
        Private _errors As List(Of String)
        Private _warnings As List(Of String)
        Private _importDate As DateTime
        Private _importedBy As String

        Public Property TotalRows() As Integer
            Get
                Return _totalRows
            End Get
            Set(ByVal value As Integer)
                _totalRows = value
            End Set
        End Property

        Public Property NewEmployees() As Integer
            Get
                Return _newEmployees
            End Get
            Set(ByVal value As Integer)
                _newEmployees = value
            End Set
        End Property

        Public Property UpdatedEmployees() As Integer
            Get
                Return _updatedEmployees
            End Get
            Set(ByVal value As Integer)
                _updatedEmployees = value
            End Set
        End Property

        Public Property SkippedRows() As Integer
            Get
                Return _skippedRows
            End Get
            Set(ByVal value As Integer)
                _skippedRows = value
            End Set
        End Property

        Public Property Errors() As List(Of String)
            Get
                Return _errors
            End Get
            Set(ByVal value As List(Of String))
                _errors = value
            End Set
        End Property

        Public Property Warnings() As List(Of String)
            Get
                Return _warnings
            End Get
            Set(ByVal value As List(Of String))
                _warnings = value
            End Set
        End Property

        Public Property ImportDate() As DateTime
            Get
                Return _importDate
            End Get
            Set(ByVal value As DateTime)
                _importDate = value
            End Set
        End Property

        Public Property ImportedBy() As String
            Get
                Return _importedBy
            End Get
            Set(ByVal value As String)
                _importedBy = value
            End Set
        End Property

        ' Calculated properties
        Public ReadOnly Property IsSuccess() As Boolean
            Get
                Return _errors.Count = 0
            End Get
        End Property

        ' Constructor
        Public Sub New()
            _totalRows = 0
            _newEmployees = 0
            _updatedEmployees = 0
            _skippedRows = 0
            _errors = New List(Of String)()
            _warnings = New List(Of String)()
            _importDate = DateTime.Now
        End Sub

        Public Sub New(importedBy As String)
            Me.New()
            _importedBy = importedBy
        End Sub

        ' Override methods
        Public Overrides Function ToString() As String
            Return String.Format("Import on {0:dd/MM/yyyy HH:mm} by {1}: {2} total, {3} new, {4} updated, {5} skipped, {6} errors",
                _importDate, _importedBy, _totalRows, _newEmployees, _updatedEmployees, _skippedRows, _errors.Count)
        End Function

    End Class
End Namespace

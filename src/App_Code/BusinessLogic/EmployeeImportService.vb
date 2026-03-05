Imports Microsoft.VisualBasic
Imports System.Data
Imports System.IO
Imports Defra.TrainTrack.Models
Imports Defra.TrainTrack.DataAccess

Namespace Defra.TrainTrack.BusinessLogic

    ''' <summary>
    ''' Represents a single row in the import preview, showing what action would be taken.
    ''' </summary>
    Public Class EmployeeImportPreviewRow
        Private _rowNumber As Integer
        Private _employeeNumber As String
        Private _firstName As String
        Private _lastName As String
        Private _email As String
        Private _department As String
        Private _action As String
        Private _details As String
        Private _errors As List(Of String)

        Public Property RowNumber() As Integer
            Get
                Return _rowNumber
            End Get
            Set(ByVal value As Integer)
                _rowNumber = value
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

        Public Property Email() As String
            Get
                Return _email
            End Get
            Set(ByVal value As String)
                _email = value
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

        ''' <summary>
        ''' The action that would be taken: "NEW", "UPDATE", or "ERROR"
        ''' </summary>
        Public Property Action() As String
            Get
                Return _action
            End Get
            Set(ByVal value As String)
                _action = value
            End Set
        End Property

        Public Property Details() As String
            Get
                Return _details
            End Get
            Set(ByVal value As String)
                _details = value
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

        Public Sub New()
            _errors = New List(Of String)()
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format("Row {0}: {1} - {2} {3} ({4})", _rowNumber, _action, _firstName, _lastName, _employeeNumber)
        End Function
    End Class

    ''' <summary>
    ''' Core import logic for CSV employee data uploads.
    ''' Supports validation, preview/staging, and execution of employee imports.
    ''' </summary>
    Public Class EmployeeImportService
        Implements IDisposable

        Private _repository As EmployeeRepository

        ' Expected CSV column names matching the Employee model fields
        Private Shared ReadOnly RequiredColumns As String() = {
            "EmployeeNumber", "FirstName", "LastName", "Email"
        }

        Private Shared ReadOnly OptionalColumns As String() = {
            "Department", "Position", "Location", "ManagerId", "PhoneNumber",
            "LineManagerEmail", "CostCentre", "PayBand", "ContractType",
            "WorkingPattern", "HireDate"
        }

        Private Shared ReadOnly AllExpectedColumns As String() = {
            "EmployeeNumber", "FirstName", "LastName", "Email",
            "Department", "Position", "Location", "ManagerId", "PhoneNumber",
            "LineManagerEmail", "CostCentre", "PayBand", "ContractType",
            "WorkingPattern", "HireDate"
        }

        Public Sub New()
            _repository = New EmployeeRepository()
        End Sub

        ''' <summary>
        ''' Validates a CSV file by checking that expected column headers exist.
        ''' Returns errors for missing required columns and warnings for unrecognised extra columns.
        ''' </summary>
        Public Function ValidateCsvFile(filePath As String) As EmployeeImportResult
            Dim result As New EmployeeImportResult()

            Try
                If String.IsNullOrWhiteSpace(filePath) Then
                    result.Errors.Add("File path cannot be empty")
                    Return result
                End If

                If Not File.Exists(filePath) Then
                    result.Errors.Add($"File not found: {filePath}")
                    Return result
                End If

                ' Read the CSV file
                Dim csvData As DataTable = ParseCsvFile(filePath)

                If csvData Is Nothing OrElse csvData.Columns.Count = 0 Then
                    result.Errors.Add("CSV file is empty or could not be parsed")
                    Return result
                End If

                result.TotalRows = csvData.Rows.Count

                ' Get actual column names from the CSV
                Dim actualColumns As New List(Of String)()
                For Each col As DataColumn In csvData.Columns
                    actualColumns.Add(col.ColumnName.Trim())
                Next

                ' Check for missing required columns
                For Each requiredCol As String In RequiredColumns
                    Dim found As Boolean = False
                    For Each actualCol As String In actualColumns
                        If String.Equals(actualCol, requiredCol, StringComparison.OrdinalIgnoreCase) Then
                            found = True
                            Exit For
                        End If
                    Next
                    If Not found Then
                        result.Errors.Add($"Missing required column: {requiredCol}")
                    End If
                Next

                ' Check for unrecognised extra columns (warn, don't error)
                For Each actualCol As String In actualColumns
                    Dim recognised As Boolean = False
                    For Each expectedCol As String In AllExpectedColumns
                        If String.Equals(actualCol, expectedCol, StringComparison.OrdinalIgnoreCase) Then
                            recognised = True
                            Exit For
                        End If
                    Next
                    If Not recognised Then
                        result.Warnings.Add($"Unrecognised column will be ignored: {actualCol}")
                    End If
                Next

                If result.Errors.Count = 0 Then
                    EventLog.WriteEntry("TrainTrack", $"CSV validation passed: {filePath} ({result.TotalRows} rows)", EventLogEntryType.Information)
                Else
                    EventLog.WriteEntry("TrainTrack", $"CSV validation failed: {filePath} ({result.Errors.Count} errors)", EventLogEntryType.Warning)
                End If

            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error validating CSV file {filePath}: {ex.Message}", EventLogEntryType.Error)
                result.Errors.Add($"Error reading CSV file: {ex.Message}")
            End Try

            Return result
        End Function

        ''' <summary>
        ''' Parses the CSV and for each row determines what action would be taken:
        ''' NEW (no matching EmployeeNumber), UPDATE (matching EmployeeNumber found),
        ''' or ERROR (validation failure). This is the staging/preview step.
        ''' </summary>
        Public Function PreviewImport(filePath As String) As List(Of EmployeeImportPreviewRow)
            Dim previewRows As New List(Of EmployeeImportPreviewRow)()

            Try
                ' First validate the file
                Dim validation As EmployeeImportResult = ValidateCsvFile(filePath)
                If Not validation.IsSuccess Then
                    ' Return a single error row indicating validation failure
                    Dim errorRow As New EmployeeImportPreviewRow()
                    errorRow.RowNumber = 0
                    errorRow.Action = "ERROR"
                    errorRow.Details = "CSV validation failed"
                    errorRow.Errors = validation.Errors
                    previewRows.Add(errorRow)
                    Return previewRows
                End If

                Dim csvData As DataTable = ParseCsvFile(filePath)

                For rowIndex As Integer = 0 To csvData.Rows.Count - 1
                    Dim row As DataRow = csvData.Rows(rowIndex)
                    Dim preview As New EmployeeImportPreviewRow()
                    preview.RowNumber = rowIndex + 1

                    ' Extract key fields for display
                    preview.EmployeeNumber = GetColumnValue(csvData, row, "EmployeeNumber")
                    preview.FirstName = GetColumnValue(csvData, row, "FirstName")
                    preview.LastName = GetColumnValue(csvData, row, "LastName")
                    preview.Email = GetColumnValue(csvData, row, "Email")
                    preview.Department = GetColumnValue(csvData, row, "Department")

                    ' Validate required fields for this row
                    Dim rowErrors As New List(Of String)()

                    If String.IsNullOrWhiteSpace(preview.EmployeeNumber) Then
                        rowErrors.Add("EmployeeNumber is required")
                    End If

                    If String.IsNullOrWhiteSpace(preview.FirstName) Then
                        rowErrors.Add("FirstName is required")
                    End If

                    If String.IsNullOrWhiteSpace(preview.LastName) Then
                        rowErrors.Add("LastName is required")
                    End If

                    If String.IsNullOrWhiteSpace(preview.Email) Then
                        rowErrors.Add("Email is required")
                    End If

                    If rowErrors.Count > 0 Then
                        preview.Action = "ERROR"
                        preview.Details = String.Join("; ", rowErrors)
                        preview.Errors = rowErrors
                    Else
                        ' Check if employee already exists by EmployeeNumber
                        Dim existingEmployee As Employee = _repository.GetEmployeeByNumber(preview.EmployeeNumber)

                        If existingEmployee IsNot Nothing Then
                            preview.Action = "UPDATE"
                            preview.Details = $"Existing employee found (ID: {existingEmployee.EmployeeId})"
                        Else
                            preview.Action = "NEW"
                            preview.Details = "New employee will be created"
                        End If
                    End If

                    previewRows.Add(preview)
                Next

                EventLog.WriteEntry("TrainTrack", $"Import preview generated: {filePath} ({previewRows.Count} rows)", EventLogEntryType.Information)

            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error generating import preview for {filePath}: {ex.Message}", EventLogEntryType.Error)
                Dim errorRow As New EmployeeImportPreviewRow()
                errorRow.RowNumber = 0
                errorRow.Action = "ERROR"
                errorRow.Details = $"Error generating preview: {ex.Message}"
                errorRow.Errors.Add(ex.Message)
                previewRows.Add(errorRow)
            End Try

            Return previewRows
        End Function

        ''' <summary>
        ''' Executes the import: for each row, matches on EmployeeNumber.
        ''' If match found: updates fields (but does NOT blank out fields if CSV value is empty).
        ''' If no match: creates new employee.
        ''' Logs each action to the audit log.
        ''' </summary>
        Public Function ExecuteImport(filePath As String, importedBy As String) As EmployeeImportResult
            Dim result As New EmployeeImportResult(importedBy)

            Try
                If String.IsNullOrWhiteSpace(filePath) Then
                    result.Errors.Add("File path cannot be empty")
                    Return result
                End If

                If String.IsNullOrWhiteSpace(importedBy) Then
                    result.Errors.Add("Imported by cannot be empty")
                    Return result
                End If

                ' First validate the file
                Dim validation As EmployeeImportResult = ValidateCsvFile(filePath)
                If Not validation.IsSuccess Then
                    result.Errors = validation.Errors
                    result.Warnings = validation.Warnings
                    Return result
                End If

                Dim csvData As DataTable = ParseCsvFile(filePath)
                result.TotalRows = csvData.Rows.Count

                EventLog.WriteEntry("TrainTrack", $"Employee import started by {importedBy}: {filePath} ({result.TotalRows} rows)", EventLogEntryType.Information)

                For rowIndex As Integer = 0 To csvData.Rows.Count - 1
                    Dim row As DataRow = csvData.Rows(rowIndex)
                    Dim rowNumber As Integer = rowIndex + 1

                    Try
                        Dim employeeNumber As String = GetColumnValue(csvData, row, "EmployeeNumber")

                        ' Validate required fields
                        If String.IsNullOrWhiteSpace(employeeNumber) Then
                            result.Errors.Add($"Row {rowNumber}: EmployeeNumber is required")
                            result.SkippedRows += 1
                            Continue For
                        End If

                        Dim firstName As String = GetColumnValue(csvData, row, "FirstName")
                        Dim lastName As String = GetColumnValue(csvData, row, "LastName")
                        Dim email As String = GetColumnValue(csvData, row, "Email")

                        If String.IsNullOrWhiteSpace(firstName) OrElse
                           String.IsNullOrWhiteSpace(lastName) OrElse
                           String.IsNullOrWhiteSpace(email) Then
                            result.Errors.Add($"Row {rowNumber}: FirstName, LastName, and Email are required")
                            result.SkippedRows += 1
                            Continue For
                        End If

                        ' Check if employee already exists by EmployeeNumber
                        Dim existingEmployee As Employee = _repository.GetEmployeeByNumber(employeeNumber)

                        If existingEmployee IsNot Nothing Then
                            ' UPDATE existing employee -- do NOT blank out fields if CSV value is empty
                            UpdateEmployeeFromCsvRow(existingEmployee, csvData, row, importedBy)
                            Dim success As Boolean = _repository.UpdateEmployee(existingEmployee)

                            If success Then
                                result.UpdatedEmployees += 1
                                EventLog.WriteEntry("TrainTrack", $"Import: Updated employee {employeeNumber} (row {rowNumber}) by {importedBy}", EventLogEntryType.Information)
                            Else
                                result.Errors.Add($"Row {rowNumber}: Failed to update employee {employeeNumber}")
                                result.SkippedRows += 1
                            End If
                        Else
                            ' CREATE new employee
                            Dim newEmployee As New Employee()
                            PopulateEmployeeFromCsvRow(newEmployee, csvData, row, importedBy)
                            newEmployee.CreatedBy = importedBy
                            newEmployee.CreatedDate = DateTime.Now

                            Dim newId As Integer = _repository.CreateEmployee(newEmployee)

                            If newId > 0 Then
                                result.NewEmployees += 1
                                EventLog.WriteEntry("TrainTrack", $"Import: Created employee {employeeNumber} (ID: {newId}, row {rowNumber}) by {importedBy}", EventLogEntryType.Information)
                            Else
                                result.Errors.Add($"Row {rowNumber}: Failed to create employee {employeeNumber}")
                                result.SkippedRows += 1
                            End If
                        End If

                    Catch rowEx As Exception
                        result.Errors.Add($"Row {rowNumber}: {rowEx.Message}")
                        result.SkippedRows += 1
                        EventLog.WriteEntry("TrainTrack", $"Import error at row {rowNumber}: {rowEx.Message}", EventLogEntryType.Warning)
                    End Try
                Next

                EventLog.WriteEntry("TrainTrack",
                    $"Employee import completed by {importedBy}: {result.NewEmployees} new, {result.UpdatedEmployees} updated, {result.SkippedRows} skipped, {result.Errors.Count} errors",
                    EventLogEntryType.Information)

            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error executing employee import: {ex.Message}", EventLogEntryType.Error)
                result.Errors.Add($"Import failed: {ex.Message}")
            End Try

            Return result
        End Function

        ''' <summary>
        ''' Formats a human-readable summary of the import result.
        ''' </summary>
        Public Function GenerateImportReport(result As EmployeeImportResult) As String
            If result Is Nothing Then
                Return "No import result available."
            End If

            Dim report As New System.Text.StringBuilder()

            report.AppendLine("========================================")
            report.AppendLine("  Employee Import Report")
            report.AppendLine("========================================")
            report.AppendLine()
            report.AppendLine($"Import Date:   {result.ImportDate:dd/MM/yyyy HH:mm:ss}")
            report.AppendLine($"Imported By:   {result.ImportedBy}")
            report.AppendLine()
            report.AppendLine("--- Summary ---")
            report.AppendLine($"Total Rows:       {result.TotalRows}")
            report.AppendLine($"New Employees:    {result.NewEmployees}")
            report.AppendLine($"Updated:          {result.UpdatedEmployees}")
            report.AppendLine($"Skipped:          {result.SkippedRows}")
            report.AppendLine($"Status:           {If(result.IsSuccess, "SUCCESS", "COMPLETED WITH ERRORS")}")
            report.AppendLine()

            If result.Warnings.Count > 0 Then
                report.AppendLine("--- Warnings ---")
                For Each warning As String In result.Warnings
                    report.AppendLine($"  WARNING: {warning}")
                Next
                report.AppendLine()
            End If

            If result.Errors.Count > 0 Then
                report.AppendLine("--- Errors ---")
                For Each errorMsg As String In result.Errors
                    report.AppendLine($"  ERROR: {errorMsg}")
                Next
                report.AppendLine()
            End If

            report.AppendLine("========================================")

            Return report.ToString()
        End Function

        ' ===== Private Helper Methods =====

        ''' <summary>
        ''' Updates an existing employee with CSV row data.
        ''' Does NOT overwrite non-empty fields with blank CSV values --
        ''' this addresses the incident where blank columns wiped existing data.
        ''' </summary>
        Private Sub UpdateEmployeeFromCsvRow(employee As Employee, csvData As DataTable, row As DataRow, modifiedBy As String)
            ' Only update fields if the CSV value is not empty (blank-field protection)
            Dim value As String

            value = GetColumnValue(csvData, row, "FirstName")
            If Not String.IsNullOrWhiteSpace(value) Then employee.FirstName = value

            value = GetColumnValue(csvData, row, "LastName")
            If Not String.IsNullOrWhiteSpace(value) Then employee.LastName = value

            value = GetColumnValue(csvData, row, "Email")
            If Not String.IsNullOrWhiteSpace(value) Then employee.Email = value

            value = GetColumnValue(csvData, row, "Department")
            If Not String.IsNullOrWhiteSpace(value) Then employee.Department = value

            value = GetColumnValue(csvData, row, "Position")
            If Not String.IsNullOrWhiteSpace(value) Then employee.Position = value

            value = GetColumnValue(csvData, row, "Location")
            If Not String.IsNullOrWhiteSpace(value) Then employee.Location = value

            value = GetColumnValue(csvData, row, "ManagerId")
            If Not String.IsNullOrWhiteSpace(value) Then
                Dim managerId As Integer
                If Integer.TryParse(value, managerId) Then
                    employee.ManagerId = managerId
                End If
            End If

            value = GetColumnValue(csvData, row, "PhoneNumber")
            If Not String.IsNullOrWhiteSpace(value) Then employee.PhoneNumber = value

            value = GetColumnValue(csvData, row, "LineManagerEmail")
            If Not String.IsNullOrWhiteSpace(value) Then employee.LineManagerEmail = value

            value = GetColumnValue(csvData, row, "CostCentre")
            If Not String.IsNullOrWhiteSpace(value) Then employee.CostCentre = value

            value = GetColumnValue(csvData, row, "PayBand")
            If Not String.IsNullOrWhiteSpace(value) Then employee.PayBand = value

            value = GetColumnValue(csvData, row, "ContractType")
            If Not String.IsNullOrWhiteSpace(value) Then employee.ContractType = value

            value = GetColumnValue(csvData, row, "WorkingPattern")
            If Not String.IsNullOrWhiteSpace(value) Then employee.WorkingPattern = value

            value = GetColumnValue(csvData, row, "HireDate")
            If Not String.IsNullOrWhiteSpace(value) Then
                Dim hireDate As DateTime
                If DateTime.TryParse(value, hireDate) Then
                    employee.HireDate = hireDate
                End If
            End If

            ' Set audit fields
            employee.ModifiedBy = modifiedBy
            employee.ModifiedDate = DateTime.Now
        End Sub

        ''' <summary>
        ''' Populates a new employee from CSV row data.
        ''' All values are set including blanks (since this is a new record).
        ''' </summary>
        Private Sub PopulateEmployeeFromCsvRow(employee As Employee, csvData As DataTable, row As DataRow, createdBy As String)
            employee.EmployeeNumber = GetColumnValue(csvData, row, "EmployeeNumber")
            employee.FirstName = GetColumnValue(csvData, row, "FirstName")
            employee.LastName = GetColumnValue(csvData, row, "LastName")
            employee.Email = GetColumnValue(csvData, row, "Email")
            employee.Department = GetColumnValue(csvData, row, "Department")
            employee.Position = GetColumnValue(csvData, row, "Position")
            employee.Location = GetColumnValue(csvData, row, "Location")

            Dim managerIdStr As String = GetColumnValue(csvData, row, "ManagerId")
            If Not String.IsNullOrWhiteSpace(managerIdStr) Then
                Dim managerId As Integer
                If Integer.TryParse(managerIdStr, managerId) Then
                    employee.ManagerId = managerId
                End If
            End If

            employee.PhoneNumber = GetColumnValue(csvData, row, "PhoneNumber")
            employee.LineManagerEmail = GetColumnValue(csvData, row, "LineManagerEmail")
            employee.CostCentre = GetColumnValue(csvData, row, "CostCentre")
            employee.PayBand = GetColumnValue(csvData, row, "PayBand")
            employee.ContractType = GetColumnValue(csvData, row, "ContractType")
            employee.WorkingPattern = GetColumnValue(csvData, row, "WorkingPattern")

            Dim hireDateStr As String = GetColumnValue(csvData, row, "HireDate")
            If Not String.IsNullOrWhiteSpace(hireDateStr) Then
                Dim hireDate As DateTime
                If DateTime.TryParse(hireDateStr, hireDate) Then
                    employee.HireDate = hireDate
                End If
            End If

            employee.IsActive = True
            employee.CreatedBy = createdBy
            employee.CreatedDate = DateTime.Now
        End Sub

        ''' <summary>
        ''' Gets a column value from a DataRow, performing case-insensitive column name matching.
        ''' Returns empty string if the column does not exist.
        ''' </summary>
        Private Function GetColumnValue(csvData As DataTable, row As DataRow, columnName As String) As String
            For Each col As DataColumn In csvData.Columns
                If String.Equals(col.ColumnName.Trim(), columnName, StringComparison.OrdinalIgnoreCase) Then
                    Dim value As Object = row(col)
                    If value IsNot Nothing AndAlso value IsNot DBNull.Value Then
                        Return value.ToString().Trim()
                    End If
                    Return String.Empty
                End If
            Next
            Return String.Empty
        End Function

        ''' <summary>
        ''' Basic CSV file parser. Reads a CSV file and returns a DataTable with columns
        ''' matching CSV headers. Handles quoted fields and commas within quoted values.
        ''' </summary>
        Private Function ParseCsvFile(filePath As String) As DataTable
            Dim table As New DataTable()

            Using reader As New StreamReader(filePath)
                ' Read the header line
                Dim headerLine As String = reader.ReadLine()
                If String.IsNullOrWhiteSpace(headerLine) Then
                    Return table
                End If

                ' Parse header columns
                Dim headers As String() = ParseCsvLine(headerLine)
                For Each header As String In headers
                    Dim cleanHeader As String = header.Trim().Trim(""""c)
                    If Not String.IsNullOrWhiteSpace(cleanHeader) Then
                        table.Columns.Add(cleanHeader)
                    End If
                Next

                ' Read data rows
                Dim line As String = reader.ReadLine()
                While line IsNot Nothing
                    If Not String.IsNullOrWhiteSpace(line) Then
                        Dim values As String() = ParseCsvLine(line)
                        Dim dataRow As DataRow = table.NewRow()

                        For i As Integer = 0 To Math.Min(values.Length, table.Columns.Count) - 1
                            dataRow(i) = values(i).Trim().Trim(""""c)
                        Next

                        table.Rows.Add(dataRow)
                    End If

                    line = reader.ReadLine()
                End While
            End Using

            Return table
        End Function

        ''' <summary>
        ''' Parses a single CSV line, handling quoted fields that may contain commas.
        ''' </summary>
        Private Function ParseCsvLine(line As String) As String()
            Dim fields As New List(Of String)()
            Dim currentField As New System.Text.StringBuilder()
            Dim inQuotes As Boolean = False

            For i As Integer = 0 To line.Length - 1
                Dim c As Char = line(i)

                If c = """"c Then
                    If inQuotes AndAlso i + 1 < line.Length AndAlso line(i + 1) = """"c Then
                        ' Escaped quote (double quote inside quotes)
                        currentField.Append(""""c)
                        i += 1
                    Else
                        ' Toggle quote mode
                        inQuotes = Not inQuotes
                    End If
                ElseIf c = ","c AndAlso Not inQuotes Then
                    ' Field separator
                    fields.Add(currentField.ToString())
                    currentField.Clear()
                Else
                    currentField.Append(c)
                End If
            Next

            ' Add the last field
            fields.Add(currentField.ToString())

            Return fields.ToArray()
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            If _repository IsNot Nothing Then
                _repository.Dispose()
                _repository = Nothing
            End If
        End Sub

    End Class
End Namespace

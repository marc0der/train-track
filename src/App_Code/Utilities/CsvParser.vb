Imports Microsoft.VisualBasic
Imports System.Data
Imports System.IO
Imports System.Text

Namespace Defra.TrainTrack.Utilities

    ''' <summary>
    ''' Simple CSV file parser. Reads CSV files into DataTable format.
    ''' Handles quoted fields, commas within quoted values, and header rows.
    ''' No third-party dependencies -- consistent with legacy codebase style.
    ''' </summary>
    Public Class CsvParser

        ''' <summary>
        ''' Parses a CSV file and returns a DataTable with columns matching the CSV headers.
        ''' Handles quoted fields, commas in values, and escaped quotes.
        ''' </summary>
        ''' <param name="filePath">Full path to the CSV file to parse.</param>
        ''' <returns>DataTable with columns matching CSV headers and rows of parsed data.</returns>
        Public Shared Function ParseCsvFile(filePath As String) As DataTable
            Dim table As New DataTable()

            If String.IsNullOrWhiteSpace(filePath) Then
                Throw New ArgumentException("File path cannot be empty.", "filePath")
            End If

            If Not File.Exists(filePath) Then
                Throw New FileNotFoundException($"CSV file not found: {filePath}", filePath)
            End If

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
        ''' Parses a single CSV line, handling quoted fields that may contain commas
        ''' and escaped quotes (doubled double-quotes).
        ''' </summary>
        ''' <param name="line">A single line from a CSV file.</param>
        ''' <returns>Array of field values parsed from the line.</returns>
        Public Shared Function ParseCsvLine(line As String) As String()
            Dim fields As New List(Of String)()

            If String.IsNullOrEmpty(line) Then
                Return fields.ToArray()
            End If

            Dim currentField As New StringBuilder()
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

    End Class
End Namespace

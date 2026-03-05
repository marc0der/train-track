Imports Microsoft.VisualBasic
Imports System.Net.Mail
Imports System.Net
Imports System.Configuration
Imports Defra.TrainTrack.Models

Namespace Defra.TrainTrack.Utilities
    Public Class EmailHelper
        Private _smtpClient As SmtpClient
        Private _fromAddress As String
        Private _fromDisplayName As String

        Public Sub New()
            InitializeSmtpClient()
        End Sub

        Private Sub InitializeSmtpClient()
            Try
                Dim smtpServer As String = ConfigurationManager.AppSettings("EmailSMTPServer")
                Dim smtpPort As String = ConfigurationManager.AppSettings("EmailSMTPPort")
                _fromAddress = ConfigurationManager.AppSettings("EmailFromAddress")
                _fromDisplayName = ConfigurationManager.AppSettings("EmailFromDisplayName")

                _smtpClient = New SmtpClient()
                If Not String.IsNullOrEmpty(smtpServer) Then
                    _smtpClient.Host = smtpServer
                End If
                If Not String.IsNullOrEmpty(smtpPort) Then
                    Integer.TryParse(smtpPort, _smtpClient.Port)
                End If

                ' Use integrated authentication for government network
                _smtpClient.UseDefaultCredentials = True
                _smtpClient.EnableSsl = True
                _smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network

            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error initializing SMTP client: {ex.Message}", EventLogEntryType.Error)
                Throw New ApplicationException("Unable to initialize email service", ex)
            End Try
        End Sub

        Public Function SendEmail(toAddress As String, subject As String, body As String) As Boolean
            Return SendEmail(New List(Of String) From {toAddress}, subject, body, Nothing, Nothing)
        End Function

        Public Function SendEmail(toAddress As String, subject As String, body As String, isHtml As Boolean) As Boolean
            Return SendEmail(New List(Of String) From {toAddress}, subject, body, Nothing, Nothing, isHtml)
        End Function

        Public Function SendEmail(toAddresses As List(Of String), subject As String, body As String,
                                Optional ccAddresses As List(Of String) = Nothing,
                                Optional bccAddresses As List(Of String) = Nothing,
                                Optional isHtml As Boolean = True) As Boolean
            Try
                If toAddresses Is Nothing OrElse toAddresses.Count = 0 Then
                    Throw New ArgumentException("At least one recipient email address is required")
                End If

                If String.IsNullOrEmpty(subject) Then
                    Throw New ArgumentException("Email subject is required")
                End If

                If String.IsNullOrEmpty(body) Then
                    Throw New ArgumentException("Email body is required")
                End If

                Using message As New MailMessage()
                    ' Set from address
                    message.From = New MailAddress(_fromAddress, _fromDisplayName)

                    ' Add recipients
                    For Each toAddress In toAddresses
                        If IsValidEmailAddress(toAddress) Then
                            message.To.Add(toAddress)
                        End If
                    Next

                    If message.To.Count = 0 Then
                        Throw New ArgumentException("No valid recipient email addresses provided")
                    End If

                    ' Add CC addresses
                    If ccAddresses IsNot Nothing Then
                        For Each ccAddress In ccAddresses
                            If IsValidEmailAddress(ccAddress) Then
                                message.CC.Add(ccAddress)
                            End If
                        Next
                    End If

                    ' Add BCC addresses
                    If bccAddresses IsNot Nothing Then
                        For Each bccAddress In bccAddresses
                            If IsValidEmailAddress(bccAddress) Then
                                message.Bcc.Add(bccAddress)
                            End If
                        Next
                    End If

                    ' Set message properties
                    message.Subject = subject
                    message.Body = body
                    message.IsBodyHtml = isHtml
                    message.Priority = MailPriority.Normal

                    ' Add standard headers
                    message.Headers.Add("X-Mailer", "TrainTrack Training Management System")
                    message.Headers.Add("X-Application", "TrainTrack v2.1")

                    ' Send the email
                    _smtpClient.Send(message)

                    EventLog.WriteEntry("TrainTrack", $"Email sent successfully to {String.Join(", ", toAddresses)}", EventLogEntryType.Information)
                    Return True
                End Using

            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error sending email: {ex.Message}", EventLogEntryType.Error)
                Return False
            End Try
        End Function

        Public Function SendTrainingInvitation(employee As Employee, session As TrainingSession, course As Course) As Boolean
            Try
                If employee Is Nothing OrElse String.IsNullOrEmpty(employee.Email) Then
                    Throw New ArgumentException("Employee and email address are required")
                End If

                If session Is Nothing OrElse course Is Nothing Then
                    Throw New ArgumentException("Training session and course information are required")
                End If

                Dim subject As String = $"Training Invitation: {course.Title}"
                Dim body As String = BuildTrainingInvitationBody(employee, session, course)

                Return SendEmail(employee.Email, subject, body, True)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error sending training invitation to {employee?.Email}: {ex.Message}", EventLogEntryType.Error)
                Return False
            End Try
        End Function

        Public Function SendTrainingReminder(employee As Employee, session As TrainingSession, course As Course, daysUntilSession As Integer) As Boolean
            Try
                If employee Is Nothing OrElse String.IsNullOrEmpty(employee.Email) Then
                    Throw New ArgumentException("Employee and email address are required")
                End If

                If session Is Nothing OrElse course Is Nothing Then
                    Throw New ArgumentException("Training session and course information are required")
                End If

                Dim subject As String = $"Training Reminder: {course.Title} - {daysUntilSession} day(s) remaining"
                Dim body As String = BuildTrainingReminderBody(employee, session, course, daysUntilSession)

                Return SendEmail(employee.Email, subject, body, True)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error sending training reminder to {employee?.Email}: {ex.Message}", EventLogEntryType.Error)
                Return False
            End Try
        End Function

        Public Function SendCertificateExpiration(employee As Employee, course As Course, expiryDate As DateTime) As Boolean
            Try
                If employee Is Nothing OrElse String.IsNullOrEmpty(employee.Email) Then
                    Throw New ArgumentException("Employee and email address are required")
                End If

                If course Is Nothing Then
                    Throw New ArgumentException("Course information is required")
                End If

                Dim daysUntilExpiry As Integer = CInt((expiryDate - DateTime.Now).TotalDays)
                Dim subject As String = $"Certificate Expiration Notice: {course.Title}"
                Dim body As String = BuildCertificateExpirationBody(employee, course, expiryDate, daysUntilExpiry)

                Return SendEmail(employee.Email, subject, body, True)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error sending certificate expiration notice to {employee?.Email}: {ex.Message}", EventLogEntryType.Error)
                Return False
            End Try
        End Function

        Public Function SendErrorNotification(exception As Exception) As Boolean
            Try
                Dim errorEmail As String = ConfigurationManager.AppSettings("ErrorReportingEmail")
                If String.IsNullOrEmpty(errorEmail) Then
                    Return False ' No error reporting email configured
                End If

                Dim subject As String = $"TrainTrack System Error - {DateTime.Now:yyyy-MM-dd HH:mm}"
                Dim body As String = BuildErrorNotificationBody(exception)

                Return SendEmail(errorEmail, subject, body, False)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error sending error notification: {ex.Message}", EventLogEntryType.Error)
                Return False
            End Try
        End Function

        Public Function SendPasswordExpirationWarning(employee As Employee, daysUntilExpiry As Integer) As Boolean
            Try
                If employee Is Nothing OrElse String.IsNullOrEmpty(employee.Email) Then
                    Throw New ArgumentException("Employee and email address are required")
                End If

                Dim subject As String = $"Password Expiration Warning - {daysUntilExpiry} day(s) remaining"
                Dim body As String = BuildPasswordExpirationBody(employee, daysUntilExpiry)

                Return SendEmail(employee.Email, subject, body, True)
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error sending password expiration warning to {employee?.Email}: {ex.Message}", EventLogEntryType.Error)
                Return False
            End Try
        End Function

        Private Function BuildTrainingInvitationBody(employee As Employee, session As TrainingSession, course As Course) As String
            Dim sb As New System.Text.StringBuilder()

            sb.AppendLine("<html><body style='font-family: Arial, sans-serif;'>")
            sb.AppendLine($"<h2>Training Invitation</h2>")
            sb.AppendLine($"<p>Dear {employee.DisplayName},</p>")
            sb.AppendLine("<p>You have been enrolled in the following training session:</p>")
            sb.AppendLine("<table border='1' style='border-collapse: collapse; width: 100%;'>")
            sb.AppendLine($"<tr><td><strong>Course:</strong></td><td>{course.DisplayTitle}</td></tr>")
            sb.AppendLine($"<tr><td><strong>Date:</strong></td><td>{session.SessionDate:dddd, dd MMMM yyyy}</td></tr>")
            sb.AppendLine($"<tr><td><strong>Time:</strong></td><td>{session.StartTime:hh\:mm} - {session.EndTime:hh\:mm}</td></tr>")
            sb.AppendLine($"<tr><td><strong>Duration:</strong></td><td>{session.DurationText}</td></tr>")
            sb.AppendLine($"<tr><td><strong>Location:</strong></td><td>{session.Location}</td></tr>")
            sb.AppendLine($"<tr><td><strong>Instructor:</strong></td><td>{session.PrimaryInstructor?.DisplayName}</td></tr>")
            sb.AppendLine("</table>")

            If Not String.IsNullOrEmpty(course.Description) Then
                sb.AppendLine($"<h3>Course Description</h3>")
                sb.AppendLine($"<p>{course.Description}</p>")
            End If

            If Not String.IsNullOrEmpty(course.Prerequisites) AndAlso course.Prerequisites.ToLower() <> "none" Then
                sb.AppendLine($"<h3>Prerequisites</h3>")
                sb.AppendLine($"<p>{course.Prerequisites}</p>")
            End If

            If Not String.IsNullOrEmpty(session.SessionNotes) Then
                sb.AppendLine($"<h3>Session Notes</h3>")
                sb.AppendLine($"<p>{session.SessionNotes}</p>")
            End If

            sb.AppendLine($"<p><strong>Registration Deadline:</strong> {session.RegistrationDeadline:dddd, dd MMMM yyyy}</p>")
            sb.AppendLine("<p>Please ensure you attend this training session. If you are unable to attend, please contact your line manager or the Training team as soon as possible.</p>")
            sb.AppendLine("<p>Best regards,<br/>TrainTrack Training Management System<br/>Department for Environment, Food and Rural Affairs</p>")
            sb.AppendLine("</body></html>")

            Return sb.ToString()
        End Function

        Private Function BuildTrainingReminderBody(employee As Employee, session As TrainingSession, course As Course, daysUntilSession As Integer) As String
            Dim sb As New System.Text.StringBuilder()

            sb.AppendLine("<html><body style='font-family: Arial, sans-serif;'>")
            sb.AppendLine($"<h2>Training Reminder</h2>")
            sb.AppendLine($"<p>Dear {employee.DisplayName},</p>")
            sb.AppendLine($"<p>This is a reminder that you have a training session scheduled in {daysUntilSession} day(s):</p>")
            sb.AppendLine("<table border='1' style='border-collapse: collapse; width: 100%;'>")
            sb.AppendLine($"<tr><td><strong>Course:</strong></td><td>{course.DisplayTitle}</td></tr>")
            sb.AppendLine($"<tr><td><strong>Date:</strong></td><td>{session.SessionDate:dddd, dd MMMM yyyy}</td></tr>")
            sb.AppendLine($"<tr><td><strong>Time:</strong></td><td>{session.StartTime:hh\:mm} - {session.EndTime:hh\:mm}</td></tr>")
            sb.AppendLine($"<tr><td><strong>Location:</strong></td><td>{session.Location}</td></tr>")
            sb.AppendLine("</table>")
            sb.AppendLine("<p>Please ensure you arrive on time and bring any required materials.</p>")
            sb.AppendLine("<p>Best regards,<br/>TrainTrack Training Management System<br/>Department for Environment, Food and Rural Affairs</p>")
            sb.AppendLine("</body></html>")

            Return sb.ToString()
        End Function

        Private Function BuildCertificateExpirationBody(employee As Employee, course As Course, expiryDate As DateTime, daysUntilExpiry As Integer) As String
            Dim sb As New System.Text.StringBuilder()

            sb.AppendLine("<html><body style='font-family: Arial, sans-serif;'>")
            sb.AppendLine($"<h2>Certificate Expiration Notice</h2>")
            sb.AppendLine($"<p>Dear {employee.DisplayName},</p>")

            If daysUntilExpiry > 0 Then
                sb.AppendLine($"<p>Your training certificate for <strong>{course.DisplayTitle}</strong> will expire in {daysUntilExpiry} day(s) on {expiryDate:dddd, dd MMMM yyyy}.</p>")
                sb.AppendLine("<p>Please arrange to retake this training to maintain your compliance.</p>")
            Else
                sb.AppendLine($"<p>Your training certificate for <strong>{course.DisplayTitle}</strong> has expired as of {expiryDate:dddd, dd MMMM yyyy}.</p>")
                sb.AppendLine("<p>Please contact the Training team immediately to schedule a renewal session.</p>")
            End If

            sb.AppendLine("<p>For more information or to book a training session, please contact the Training team.</p>")
            sb.AppendLine("<p>Best regards,<br/>TrainTrack Training Management System<br/>Department for Environment, Food and Rural Affairs</p>")
            sb.AppendLine("</body></html>")

            Return sb.ToString()
        End Function

        Private Function BuildPasswordExpirationBody(employee As Employee, daysUntilExpiry As Integer) As String
            Dim sb As New System.Text.StringBuilder()

            sb.AppendLine("<html><body style='font-family: Arial, sans-serif;'>")
            sb.AppendLine($"<h2>Password Expiration Warning</h2>")
            sb.AppendLine($"<p>Dear {employee.DisplayName},</p>")
            sb.AppendLine($"<p>Your TrainTrack system password will expire in {daysUntilExpiry} day(s).</p>")
            sb.AppendLine("<p>Please log in to the system and change your password to avoid any disruption to your access.</p>")
            sb.AppendLine("<p>If you need assistance with changing your password, please contact the IT Support team.</p>")
            sb.AppendLine("<p>Best regards,<br/>TrainTrack Training Management System<br/>Department for Environment, Food and Rural Affairs</p>")
            sb.AppendLine("</body></html>")

            Return sb.ToString()
        End Function

        Private Function BuildErrorNotificationBody(exception As Exception) As String
            Dim sb As New System.Text.StringBuilder()

            sb.AppendLine("TrainTrack System Error Report")
            sb.AppendLine("====================================")
            sb.AppendLine($"Timestamp: {DateTime.Now}")
            sb.AppendLine($"Exception Type: {exception.GetType().Name}")
            sb.AppendLine($"Message: {exception.Message}")
            sb.AppendLine("")
            sb.AppendLine("Stack Trace:")
            sb.AppendLine(exception.StackTrace)

            If exception.InnerException IsNot Nothing Then
                sb.AppendLine("")
                sb.AppendLine("Inner Exception:")
                sb.AppendLine($"Type: {exception.InnerException.GetType().Name}")
                sb.AppendLine($"Message: {exception.InnerException.Message}")
                sb.AppendLine($"Stack Trace: {exception.InnerException.StackTrace}")
            End If

            sb.AppendLine("")
            sb.AppendLine("System Information:")
            sb.AppendLine($"Machine Name: {Environment.MachineName}")
            sb.AppendLine($"User Name: {Environment.UserName}")
            sb.AppendLine($"OS Version: {Environment.OSVersion}")
            sb.AppendLine($".NET Version: {Environment.Version}")

            If HttpContext.Current IsNot Nothing Then
                sb.AppendLine("")
                sb.AppendLine("Web Request Information:")
                sb.AppendLine($"URL: {HttpContext.Current.Request.Url}")
                sb.AppendLine($"User Agent: {HttpContext.Current.Request.UserAgent}")
                sb.AppendLine($"IP Address: {HttpContext.Current.Request.UserHostAddress}")
                If HttpContext.Current.User IsNot Nothing Then
                    sb.AppendLine($"User: {HttpContext.Current.User.Identity.Name}")
                End If
            End If

            Return sb.ToString()
        End Function

        Private Function IsValidEmailAddress(email As String) As Boolean
            Try
                If String.IsNullOrWhiteSpace(email) Then
                    Return False
                End If

                Dim addr As New MailAddress(email)
                Return addr.Address = email
            Catch
                Return False
            End Try
        End Function

        Public Sub Dispose()
            If _smtpClient IsNot Nothing Then
                _smtpClient.Dispose()
                _smtpClient = Nothing
            End If
        End Sub

    End Class
End Namespace
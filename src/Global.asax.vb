Imports System.Web
Imports System.Web.Security
Imports System.Web.SessionState
Imports System.Configuration
Imports System.Data.SqlClient

Namespace Defra.TrainTrack
    Public Class Global
        Inherits System.Web.HttpApplication

        Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
            ' Initialize application
            LogEvent("TrainTrack Application Started", EventLogEntryType.Information)

            ' Cache application settings
            CacheApplicationSettings()

            ' Initialize security roles
            InitializeSecurityRoles()

            ' Schedule maintenance tasks
            ScheduleMaintenanceTasks()

            ' Verify database connectivity
            VerifyDatabaseConnections()
        End Sub

        Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
            LogEvent("TrainTrack Application Ended", EventLogEntryType.Information)
        End Sub

        Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
            Dim ex As Exception = Server.GetLastError()

            If ex IsNot Nothing Then
                ' Log the error
                LogError(ex)

                ' Send error notification if critical
                If IsCriticalError(ex) Then
                    SendErrorNotification(ex)
                End If

                ' Clear the error
                Server.ClearError()

                ' Redirect to error page
                If Not Response.IsRequestBeingRedirected Then
                    Response.Redirect("~/Error.aspx")
                End If
            End If
        End Sub

        Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
            ' Initialize session data
            Session("UserStartTime") = DateTime.Now
            Session("PageViewCount") = 0

            ' Log session start for audit
            If HttpContext.Current.User.Identity.IsAuthenticated Then
                LogEvent($"Session started for user: {HttpContext.Current.User.Identity.Name}", EventLogEntryType.Information)
            End If
        End Sub

        Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
            ' Clean up session data
            If Session("UserStartTime") IsNot Nothing Then
                Dim sessionDuration As TimeSpan = DateTime.Now - CType(Session("UserStartTime"), DateTime)
                LogEvent($"Session ended. Duration: {sessionDuration.TotalMinutes:F1} minutes", EventLogEntryType.Information)
            End If
        End Sub

        Protected Sub Application_PreSendRequestHeaders()
            ' Remove server header for security
            Response.Headers.Remove("Server")
        End Sub

        Protected Sub Application_AuthenticateRequest(ByVal sender As Object, ByVal e As EventArgs)
            ' Custom authentication logic if needed
            If HttpContext.Current.User.Identity.IsAuthenticated Then
                ' Load user roles and permissions
                LoadUserRolesAndPermissions()
            End If
        End Sub

        Private Sub CacheApplicationSettings()
            Try
                ' Cache frequently used settings for performance
                Application("ApplicationName") = ConfigurationManager.AppSettings("ApplicationName")
                Application("ApplicationVersion") = ConfigurationManager.AppSettings("ApplicationVersion")
                Application("DepartmentCode") = ConfigurationManager.AppSettings("DepartmentCode")
                Application("SupportEmail") = ConfigurationManager.AppSettings("SupportEmail")
                Application("EnableAuditLogging") = Boolean.Parse(ConfigurationManager.AppSettings("EnableAuditLogging"))
                Application("EnableEmailNotifications") = Boolean.Parse(ConfigurationManager.AppSettings("EnableEmailNotifications"))

                LogEvent("Application settings cached successfully", EventLogEntryType.Information)
            Catch ex As Exception
                LogError(ex, "Error caching application settings")
            End Try
        End Sub

        Private Sub InitializeSecurityRoles()
            Try
                ' Initialize role-based security
                Dim roles As String() = {
                    "DEFRA\TrainTrack_Admins",
                    "DEFRA\TrainTrack_Managers",
                    "DEFRA\TrainTrack_Instructors",
                    "DEFRA\TrainTrack_Users",
                    "DEFRA\TrainTrack_Reports"
                }

                Application("SecurityRoles") = roles
                LogEvent("Security roles initialized", EventLogEntryType.Information)
            Catch ex As Exception
                LogError(ex, "Error initializing security roles")
            End Try
        End Sub

        Private Sub ScheduleMaintenanceTasks()
            Try
                ' Schedule daily maintenance tasks
                Dim timer As New System.Threading.Timer(
                    AddressOf PerformMaintenanceTasks,
                    Nothing,
                    TimeSpan.Zero,
                    TimeSpan.FromHours(24)
                )

                Application("MaintenanceTimer") = timer
                LogEvent("Maintenance tasks scheduled", EventLogEntryType.Information)
            Catch ex As Exception
                LogError(ex, "Error scheduling maintenance tasks")
            End Try
        End Sub

        Private Sub PerformMaintenanceTasks(ByVal state As Object)
            Try
                ' Clear old session data
                ClearExpiredSessions()

                ' Clean up temporary files
                CleanupTempFiles()

                ' Archive old audit logs
                ArchiveOldAuditLogs()

                ' Update training statistics cache
                RefreshTrainingStatistics()

                LogEvent("Daily maintenance tasks completed", EventLogEntryType.Information)
            Catch ex As Exception
                LogError(ex, "Error performing maintenance tasks")
            End Try
        End Sub

        Private Sub VerifyDatabaseConnections()
            Try
                ' Test main database connection
                Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("TrainTrackDatabase").ConnectionString)
                    conn.Open()
                    conn.Close()
                End Using

                ' Test audit database connection
                Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("AuditDatabase").ConnectionString)
                    conn.Open()
                    conn.Close()
                End Using

                LogEvent("Database connections verified successfully", EventLogEntryType.Information)
            Catch ex As Exception
                LogError(ex, "Critical: Database connection failure")
                SendErrorNotification(ex)
            End Try
        End Sub

        Private Sub LoadUserRolesAndPermissions()
            Try
                If HttpContext.Current.User IsNot Nothing AndAlso HttpContext.Current.User.Identity.IsAuthenticated Then
                    Dim userName As String = HttpContext.Current.User.Identity.Name

                    ' Load user permissions from database
                    Dim permissions As List(Of String) = GetUserPermissions(userName)
                    HttpContext.Current.Session("UserPermissions") = permissions

                    ' Load user profile information
                    Dim userProfile As Object = GetUserProfile(userName)
                    HttpContext.Current.Session("UserProfile") = userProfile
                End If
            Catch ex As Exception
                LogError(ex, "Error loading user roles and permissions")
            End Try
        End Sub

        Private Function GetUserPermissions(userName As String) As List(Of String)
            Dim permissions As New List(Of String)()

            Try
                Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("TrainTrackDatabase").ConnectionString)
                    Using cmd As New SqlCommand("SELECT Permission FROM UserPermissions WHERE UserName = @UserName", conn)
                        cmd.Parameters.AddWithValue("@UserName", userName)
                        conn.Open()

                        Using reader As SqlDataReader = cmd.ExecuteReader()
                            While reader.Read()
                                permissions.Add(reader("Permission").ToString())
                            End While
                        End Using
                    End Using
                End Using
            Catch ex As Exception
                LogError(ex, $"Error loading permissions for user: {userName}")
            End Try

            Return permissions
        End Function

        Private Function GetUserProfile(userName As String) As Object
            Try
                Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("TrainTrackDatabase").ConnectionString)
                    Using cmd As New SqlCommand("SELECT * FROM UserProfiles WHERE UserName = @UserName", conn)
                        cmd.Parameters.AddWithValue("@UserName", userName)
                        conn.Open()

                        Using reader As SqlDataReader = cmd.ExecuteReader()
                            If reader.Read() Then
                                Return New With {
                                    .UserName = reader("UserName").ToString(),
                                    .DisplayName = reader("DisplayName").ToString(),
                                    .Email = reader("Email").ToString(),
                                    .Department = reader("Department").ToString(),
                                    .LastLogin = CType(reader("LastLogin"), DateTime)
                                }
                            End If
                        End Using
                    End Using
                End Using
            Catch ex As Exception
                LogError(ex, $"Error loading profile for user: {userName}")
            End Try

            Return Nothing
        End Function

        Private Sub ClearExpiredSessions()
            ' Implementation would clear expired session data from database
        End Sub

        Private Sub CleanupTempFiles()
            ' Implementation would clean up temporary files
        End Sub

        Private Sub ArchiveOldAuditLogs()
            ' Implementation would archive old audit logs
        End Sub

        Private Sub RefreshTrainingStatistics()
            ' Implementation would refresh cached training statistics
        End Sub

        Private Function IsCriticalError(ex As Exception) As Boolean
            ' Determine if error is critical (database connection, security, etc.)
            Return TypeOf ex Is SqlException OrElse
                   TypeOf ex Is SecurityException OrElse
                   ex.Message.ToLower().Contains("database") OrElse
                   ex.Message.ToLower().Contains("security")
        End Function

        Private Sub LogEvent(message As String, entryType As EventLogEntryType)
            Try
                ' Write to Windows Event Log
                EventLog.WriteEntry("TrainTrack", message, entryType)

                ' Also write to file log
                WriteToFileLog($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{entryType}] {message}")
            Catch
                ' Fail silently if logging fails
            End Try
        End Sub

        Private Sub LogError(ex As Exception, Optional additionalInfo As String = "")
            Try
                Dim errorMessage As String = $"{additionalInfo} Exception: {ex.Message} StackTrace: {ex.StackTrace}"
                LogEvent(errorMessage, EventLogEntryType.Error)
            Catch
                ' Fail silently if logging fails
            End Try
        End Sub

        Private Sub WriteToFileLog(message As String)
            Try
                Dim logPath As String = Server.MapPath("~/App_Data/Logs/")
                If Not System.IO.Directory.Exists(logPath) Then
                    System.IO.Directory.CreateDirectory(logPath)
                End If

                Dim logFile As String = System.IO.Path.Combine(logPath, $"TrainTrack-{DateTime.Now:yyyy-MM-dd}.log")
                System.IO.File.AppendAllText(logFile, message & Environment.NewLine)
            Catch
                ' Fail silently if file logging fails
            End Try
        End Sub

        Private Sub SendErrorNotification(ex As Exception)
            Try
                If ConfigurationManager.AppSettings("EnableErrorReporting") = "true" Then
                    Dim emailHelper As New Utilities.EmailHelper()
                    emailHelper.SendErrorNotification(ex)
                End If
            Catch
                ' Fail silently if email notification fails
            End Try
        End Sub

    End Class
End Namespace
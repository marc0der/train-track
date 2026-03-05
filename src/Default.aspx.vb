Imports System.Configuration
Imports System.Web.Security

Namespace Defra.TrainTrack
    Partial Public Class DefaultPage
        Inherits System.Web.UI.Page

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
            If Not IsPostBack Then
                InitializePage()
                CheckMaintenanceMode()
                LoadUserInformation()
                CheckUserPermissions()
            End If
        End Sub

        Private Sub InitializePage()
            Try
                ' Set page title and version information
                Page.Title = $"{ConfigurationManager.AppSettings("ApplicationName")} - Welcome"
                lblVersion.Text = ConfigurationManager.AppSettings("ApplicationVersion")
                lblFooterVersion.Text = ConfigurationManager.AppSettings("ApplicationVersion")

                ' Load application settings
                If Application("EnableAuditLogging") IsNot Nothing AndAlso CBool(Application("EnableAuditLogging")) Then
                    ' Log page access
                    Utilities.AuditHelper.LogPageAccess("Default.aspx", User.Identity.Name)
                End If

            Catch ex As Exception
                ' Handle initialization errors gracefully
                Response.Redirect("~/Error.aspx")
            End Try
        End Sub

        Private Sub CheckMaintenanceMode()
            Try
                Dim maintenanceMode As Boolean = False
                Boolean.TryParse(ConfigurationManager.AppSettings("EnableMaintenanceMode"), maintenanceMode)

                If maintenanceMode Then
                    pnlMaintenanceMode.Visible = True
                    lblMaintenanceMessage.Text = ConfigurationManager.AppSettings("MaintenanceMessage")

                    ' Disable navigation buttons during maintenance
                    btnDashboard.Enabled = False
                    btnEmployeeSearch.Enabled = False
                    btnCourseCatalog.Enabled = False
                    btnScheduleTraining.Enabled = False
                End If

            Catch ex As Exception
                ' Log error but don't prevent page load
                EventLog.WriteEntry("TrainTrack", $"Error checking maintenance mode: {ex.Message}", EventLogEntryType.Warning)
            End Try
        End Sub

        Private Sub LoadUserInformation()
            Try
                If User.Identity.IsAuthenticated Then
                    ' Set username display
                    lblUserName.Text = GetDisplayName(User.Identity.Name)

                    ' Load user profile information from session or database
                    Dim userProfile As Object = Session("UserProfile")
                    If userProfile IsNot Nothing Then
                        ' Use reflection to access properties from anonymous type
                        Dim profileType = userProfile.GetType()

                        Dim emailProperty = profileType.GetProperty("Email")
                        Dim departmentProperty = profileType.GetProperty("Department")
                        Dim lastLoginProperty = profileType.GetProperty("LastLogin")

                        If departmentProperty IsNot Nothing Then
                            lblDepartment.Text = departmentProperty.GetValue(userProfile)?.ToString()
                        End If

                        If lastLoginProperty IsNot Nothing Then
                            Dim lastLogin As DateTime = CType(lastLoginProperty.GetValue(userProfile), DateTime)
                            lblLastLogin.Text = lastLogin.ToString("dd/MM/yyyy HH:mm")
                        End If
                    Else
                        ' Load from database if not in session
                        LoadUserProfileFromDatabase()
                    End If

                    ' Determine user role
                    lblUserRole.Text = GetUserRole()

                Else
                    ' User not authenticated, redirect to login
                    FormsAuthentication.RedirectToLoginPage()
                End If

            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error loading user information: {ex.Message}", EventLogEntryType.Error)
                lblUserName.Text = "Unknown User"
                lblDepartment.Text = "Unknown"
                lblLastLogin.Text = "Unknown"
                lblUserRole.Text = "Unknown"
            End Try
        End Sub

        Private Sub LoadUserProfileFromDatabase()
            Try
                Using empManager As New BusinessLogic.EmployeeManager()
                    Dim employee = empManager.GetEmployeeByUserName(User.Identity.Name)
                    If employee IsNot Nothing Then
                        lblDepartment.Text = employee.Department
                        lblLastLogin.Text = employee.LastLoginDate.ToString("dd/MM/yyyy HH:mm")
                    End If
                End Using
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error loading user profile from database: {ex.Message}", EventLogEntryType.Warning)
            End Try
        End Sub

        Private Function GetDisplayName(userName As String) As String
            Try
                ' Extract display name from domain\username format
                If userName.Contains("\") Then
                    Return userName.Split("\"c)(1)
                Else
                    Return userName
                End If
            Catch
                Return userName
            End Try
        End Function

        Private Function GetUserRole() As String
            Try
                ' Check user roles in order of precedence
                If User.IsInRole("DEFRA\TrainTrack_Admins") Then
                    Return "Administrator"
                ElseIf User.IsInRole("DEFRA\TrainTrack_Managers") Then
                    Return "Manager"
                ElseIf User.IsInRole("DEFRA\TrainTrack_Instructors") Then
                    Return "Instructor"
                ElseIf User.IsInRole("DEFRA\TrainTrack_Reports") Then
                    Return "Reports User"
                ElseIf User.IsInRole("DEFRA\TrainTrack_Users") Then
                    Return "Standard User"
                Else
                    Return "Guest"
                End If
            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error determining user role: {ex.Message}", EventLogEntryType.Warning)
                Return "Unknown"
            End Try
        End Function

        Private Sub CheckUserPermissions()
            Try
                ' Enable/disable buttons based on user permissions
                If Not User.IsInRole("DEFRA\TrainTrack_Users") AndAlso
                   Not User.IsInRole("DEFRA\TrainTrack_Managers") AndAlso
                   Not User.IsInRole("DEFRA\TrainTrack_Admins") Then
                    ' User has no standard permissions, disable most functionality
                    btnEmployeeSearch.Enabled = False
                    btnScheduleTraining.Enabled = False
                End If

                ' Instructors and above can schedule training
                If Not User.IsInRole("DEFRA\TrainTrack_Instructors") AndAlso
                   Not User.IsInRole("DEFRA\TrainTrack_Managers") AndAlso
                   Not User.IsInRole("DEFRA\TrainTrack_Admins") Then
                    btnScheduleTraining.Enabled = False
                End If

            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error checking user permissions: {ex.Message}", EventLogEntryType.Warning)
            End Try
        End Sub

        Protected Sub btnDashboard_Click(sender As Object, e As EventArgs) Handles btnDashboard.Click
            Response.Redirect("~/Dashboard.aspx")
        End Sub

        Protected Sub btnEmployeeSearch_Click(sender As Object, e As EventArgs) Handles btnEmployeeSearch.Click
            Response.Redirect("~/EmployeeSearch.aspx")
        End Sub

        Protected Sub btnCourseCatalog_Click(sender As Object, e As EventArgs) Handles btnCourseCatalog.Click
            Response.Redirect("~/CourseCatalog.aspx")
        End Sub

        Protected Sub btnScheduleTraining_Click(sender As Object, e As EventArgs) Handles btnScheduleTraining.Click
            Response.Redirect("~/ScheduleTraining.aspx")
        End Sub

        Protected Sub lnkLogout_Click(sender As Object, e As EventArgs) Handles lnkLogout.Click
            Try
                ' Log the logout for audit purposes
                If Application("EnableAuditLogging") IsNot Nothing AndAlso CBool(Application("EnableAuditLogging")) Then
                    Utilities.AuditHelper.LogUserLogout(User.Identity.Name)
                End If

                ' Clear session data
                Session.Clear()
                Session.Abandon()

                ' Sign out and redirect
                FormsAuthentication.SignOut()
                FormsAuthentication.RedirectToLoginPage()

            Catch ex As Exception
                EventLog.WriteEntry("TrainTrack", $"Error during logout: {ex.Message}", EventLogEntryType.Warning)
                ' Still try to redirect to login even if logging fails
                FormsAuthentication.RedirectToLoginPage()
            End Try
        End Sub

        Protected Sub Page_PreRender(sender As Object, e As EventArgs) Handles Me.PreRender
            Try
                ' Set focus to first enabled button for accessibility
                If btnDashboard.Enabled Then
                    btnDashboard.Focus()
                End If

                ' Update session page view count
                If Session("PageViewCount") IsNot Nothing Then
                    Session("PageViewCount") = CInt(Session("PageViewCount")) + 1
                Else
                    Session("PageViewCount") = 1
                End If

            Catch ex As Exception
                ' Non-critical error, log but don't interfere with page rendering
                EventLog.WriteEntry("TrainTrack", $"Error in Page_PreRender: {ex.Message}", EventLogEntryType.Information)
            End Try
        End Sub

    End Class
End Namespace
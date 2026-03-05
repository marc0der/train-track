&lt;%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="Defra.TrainTrack.DefaultPage" %&gt;

&lt;!DOCTYPE html&gt;

&lt;html xmlns="http://www.w3.org/1999/xhtml"&gt;
&lt;head runat="server"&gt;
    &lt;title&gt;TrainTrack - Training Management System&lt;/title&gt;
    &lt;link rel="stylesheet" href="Styles/TrainTrack.css" type="text/css" /&gt;
    &lt;script src="Scripts/TrainTrack.js" type="text/javascript"&gt;&lt;/script&gt;
    &lt;meta http-equiv="X-UA-Compatible" content="IE=edge" /&gt;
&lt;/head&gt;
&lt;body&gt;
    &lt;form id="form1" runat="server"&gt;
        &lt;div class="main-container"&gt;
            &lt;!-- Header --&gt;
            &lt;div class="header"&gt;
                &lt;div class="app-title"&gt;TrainTrack Training Management System&lt;/div&gt;
                &lt;div class="user-info"&gt;
                    Welcome, &lt;asp:Label ID="lblUserName" runat="server"&gt;&lt;/asp:Label&gt; |
                    &lt;asp:LinkButton ID="lnkLogout" runat="server" Text="Logout" OnClick="lnkLogout_Click" /&gt;
                &lt;/div&gt;
            &lt;/div&gt;

            &lt;!-- Content --&gt;
            &lt;div class="content"&gt;
                &lt;div class="content-header"&gt;
                    &lt;h1&gt;Welcome to TrainTrack&lt;/h1&gt;
                &lt;/div&gt;

                &lt;div class="welcome-panel"&gt;
                    &lt;h2&gt;Training Management System&lt;/h2&gt;
                    &lt;p&gt;Welcome to the Department for Environment, Food and Rural Affairs Training Management System.&lt;/p&gt;
                    &lt;p&gt;This system allows you to manage employee training records, schedule training sessions, and generate compliance reports.&lt;/p&gt;
                &lt;/div&gt;

                &lt;div class="quick-access"&gt;
                    &lt;h3&gt;Quick Access&lt;/h3&gt;
                    &lt;div class="button-row"&gt;
                        &lt;asp:Button ID="btnDashboard" runat="server" Text="View Dashboard" CssClass="button button-primary" OnClick="btnDashboard_Click" /&gt;
                        &lt;asp:Button ID="btnEmployeeSearch" runat="server" Text="Search Employees" CssClass="button" OnClick="btnEmployeeSearch_Click" /&gt;
                        &lt;asp:Button ID="btnCourseCatalog" runat="server" Text="Course Catalog" CssClass="button" OnClick="btnCourseCatalog_Click" /&gt;
                        &lt;asp:Button ID="btnScheduleTraining" runat="server" Text="Schedule Training" CssClass="button" OnClick="btnScheduleTraining_Click" /&gt;
                    &lt;/div&gt;
                &lt;/div&gt;

                &lt;div class="system-info"&gt;
                    &lt;h3&gt;System Information&lt;/h3&gt;
                    &lt;table class="info-table"&gt;
                        &lt;tr&gt;
                            &lt;td&gt;Version:&lt;/td&gt;
                            &lt;td&gt;&lt;asp:Label ID="lblVersion" runat="server"&gt;&lt;/asp:Label&gt;&lt;/td&gt;
                        &lt;/tr&gt;
                        &lt;tr&gt;
                            &lt;td&gt;Last Login:&lt;/td&gt;
                            &lt;td&gt;&lt;asp:Label ID="lblLastLogin" runat="server"&gt;&lt;/asp:Label&gt;&lt;/td&gt;
                        &lt;/tr&gt;
                        &lt;tr&gt;
                            &lt;td&gt;User Role:&lt;/td&gt;
                            &lt;td&gt;&lt;asp:Label ID="lblUserRole" runat="server"&gt;&lt;/asp:Label&gt;&lt;/td&gt;
                        &lt;/tr&gt;
                        &lt;tr&gt;
                            &lt;td&gt;Department:&lt;/td&gt;
                            &lt;td&gt;&lt;asp:Label ID="lblDepartment" runat="server"&gt;&lt;/asp:Label&gt;&lt;/td&gt;
                        &lt;/tr&gt;
                    &lt;/table&gt;
                &lt;/div&gt;

                &lt;asp:Panel ID="pnlMaintenanceMode" runat="server" Visible="false" CssClass="maintenance-panel"&gt;
                    &lt;h3&gt;System Maintenance&lt;/h3&gt;
                    &lt;p&gt;&lt;asp:Label ID="lblMaintenanceMessage" runat="server"&gt;&lt;/asp:Label&gt;&lt;/p&gt;
                &lt;/asp:Panel&gt;
            &lt;/div&gt;

            &lt;!-- Footer --&gt;
            &lt;div class="footer"&gt;
                TrainTrack Training Management System v&lt;asp:Label ID="lblFooterVersion" runat="server"&gt;&lt;/asp:Label&gt; |
                © 2019 Department for Environment, Food and Rural Affairs
            &lt;/div&gt;
        &lt;/div&gt;
    &lt;/form&gt;
&lt;/body&gt;
&lt;/html&gt;
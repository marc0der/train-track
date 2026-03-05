-- TrainTrack Training Management System Database Schema
-- Department for Environment, Food and Rural Affairs
-- Created: 2016-01-15
-- Version: 2.1

USE [master]
GO

-- Create database if it doesn't exist
IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = N'TrainTrack_Production')
BEGIN
    CREATE DATABASE [TrainTrack_Production]
END
GO

USE [TrainTrack_Production]
GO

-- Drop existing tables in dependency order
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[TrainingRecords]') AND type in (N'U'))
    DROP TABLE [dbo].[TrainingRecords]
GO

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[SessionParticipants]') AND type in (N'U'))
    DROP TABLE [dbo].[SessionParticipants]
GO

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[TrainingSessions]') AND type in (N'U'))
    DROP TABLE [dbo].[TrainingSessions]
GO

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[CoursePrerequisites]') AND type in (N'U'))
    DROP TABLE [dbo].[CoursePrerequisites]
GO

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Courses]') AND type in (N'U'))
    DROP TABLE [dbo].[Courses]
GO

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UserPermissions]') AND type in (N'U'))
    DROP TABLE [dbo].[UserPermissions]
GO

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UserProfiles]') AND type in (N'U'))
    DROP TABLE [dbo].[UserProfiles]
GO

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Employees]') AND type in (N'U'))
    DROP TABLE [dbo].[Employees]
GO

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AuditLog]') AND type in (N'U'))
    DROP TABLE [dbo].[AuditLog]
GO

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[SystemSettings]') AND type in (N'U'))
    DROP TABLE [dbo].[SystemSettings]
GO

-- Create Employees table
CREATE TABLE [dbo].[Employees] (
    [EmployeeId] INT IDENTITY(1,1) NOT NULL,
    [EmployeeNumber] NVARCHAR(20) NOT NULL,
    [FirstName] NVARCHAR(50) NOT NULL,
    [LastName] NVARCHAR(50) NOT NULL,
    [Email] NVARCHAR(100) NOT NULL,
    [UserName] NVARCHAR(50) NULL,
    [Department] NVARCHAR(100) NOT NULL,
    [Position] NVARCHAR(100) NOT NULL,
    [Location] NVARCHAR(100) NOT NULL,
    [ManagerId] INT NULL,
    [ManagerName] NVARCHAR(100) NULL,
    [HireDate] DATETIME NOT NULL,
    [LastLoginDate] DATETIME NULL,
    [IsActive] BIT NOT NULL DEFAULT 1,
    [PhoneNumber] NVARCHAR(20) NULL,
    [LineManagerEmail] NVARCHAR(100) NULL,
    [CostCentre] NVARCHAR(20) NULL,
    [PayBand] NVARCHAR(10) NULL,
    [ContractType] NVARCHAR(20) NULL,
    [WorkingPattern] NVARCHAR(50) NULL,
    [CreatedDate] DATETIME NOT NULL DEFAULT GETDATE(),
    [CreatedBy] NVARCHAR(50) NOT NULL,
    [ModifiedDate] DATETIME NULL,
    [ModifiedBy] NVARCHAR(50) NULL,
    CONSTRAINT [PK_Employees] PRIMARY KEY CLUSTERED ([EmployeeId] ASC),
    CONSTRAINT [FK_Employees_Manager] FOREIGN KEY([ManagerId]) REFERENCES [dbo].[Employees] ([EmployeeId])
)
GO

-- Create indexes on Employees
CREATE UNIQUE NONCLUSTERED INDEX [IX_Employees_EmployeeNumber] ON [dbo].[Employees] ([EmployeeNumber] ASC)
GO

CREATE UNIQUE NONCLUSTERED INDEX [IX_Employees_Email] ON [dbo].[Employees] ([Email] ASC)
GO

CREATE NONCLUSTERED INDEX [IX_Employees_Department] ON [dbo].[Employees] ([Department] ASC)
GO

CREATE NONCLUSTERED INDEX [IX_Employees_IsActive] ON [dbo].[Employees] ([IsActive] ASC)
GO

-- Create UserProfiles table
CREATE TABLE [dbo].[UserProfiles] (
    [ProfileId] INT IDENTITY(1,1) NOT NULL,
    [EmployeeId] INT NOT NULL,
    [UserName] NVARCHAR(50) NOT NULL,
    [DisplayName] NVARCHAR(100) NOT NULL,
    [Email] NVARCHAR(100) NOT NULL,
    [Department] NVARCHAR(100) NOT NULL,
    [LastLogin] DATETIME NULL,
    [LoginCount] INT NOT NULL DEFAULT 0,
    [Preferences] NVARCHAR(MAX) NULL,
    [CreatedDate] DATETIME NOT NULL DEFAULT GETDATE(),
    [ModifiedDate] DATETIME NULL,
    CONSTRAINT [PK_UserProfiles] PRIMARY KEY CLUSTERED ([ProfileId] ASC),
    CONSTRAINT [FK_UserProfiles_Employees] FOREIGN KEY([EmployeeId]) REFERENCES [dbo].[Employees] ([EmployeeId])
)
GO

CREATE UNIQUE NONCLUSTERED INDEX [IX_UserProfiles_UserName] ON [dbo].[UserProfiles] ([UserName] ASC)
GO

-- Create UserPermissions table
CREATE TABLE [dbo].[UserPermissions] (
    [PermissionId] INT IDENTITY(1,1) NOT NULL,
    [EmployeeId] INT NOT NULL,
    [UserName] NVARCHAR(50) NOT NULL,
    [Permission] NVARCHAR(100) NOT NULL,
    [GrantedBy] NVARCHAR(50) NOT NULL,
    [GrantedDate] DATETIME NOT NULL DEFAULT GETDATE(),
    [IsActive] BIT NOT NULL DEFAULT 1,
    CONSTRAINT [PK_UserPermissions] PRIMARY KEY CLUSTERED ([PermissionId] ASC),
    CONSTRAINT [FK_UserPermissions_Employees] FOREIGN KEY([EmployeeId]) REFERENCES [dbo].[Employees] ([EmployeeId])
)
GO

CREATE NONCLUSTERED INDEX [IX_UserPermissions_UserName] ON [dbo].[UserPermissions] ([UserName] ASC)
GO

-- Create Courses table
CREATE TABLE [dbo].[Courses] (
    [CourseId] INT IDENTITY(1,1) NOT NULL,
    [CourseCode] NVARCHAR(20) NOT NULL,
    [Title] NVARCHAR(200) NOT NULL,
    [Description] NVARCHAR(MAX) NULL,
    [Category] NVARCHAR(100) NOT NULL,
    [DurationHours] DECIMAL(5,2) NOT NULL,
    [DeliveryMethod] NVARCHAR(50) NOT NULL,
    [IsActive] BIT NOT NULL DEFAULT 1,
    [IsCompulsory] BIT NOT NULL DEFAULT 0,
    [MaxParticipants] INT NOT NULL DEFAULT 20,
    [MinParticipants] INT NOT NULL DEFAULT 1,
    [Prerequisites] NVARCHAR(MAX) NULL,
    [LearningObjectives] NVARCHAR(MAX) NULL,
    [CourseContent] NVARCHAR(MAX) NULL,
    [AssessmentMethod] NVARCHAR(200) NULL,
    [CertificateTemplate] NVARCHAR(200) NULL,
    [ValidityPeriodMonths] INT NOT NULL DEFAULT 12,
    [CostPerParticipant] DECIMAL(10,2) NOT NULL DEFAULT 0,
    [ApprovalRequired] BIT NOT NULL DEFAULT 0,
    [CreatedDate] DATETIME NOT NULL DEFAULT GETDATE(),
    [CreatedBy] NVARCHAR(50) NOT NULL,
    [ModifiedDate] DATETIME NULL,
    [ModifiedBy] NVARCHAR(50) NULL,
    [Version] NVARCHAR(10) NOT NULL DEFAULT '1.0',
    [SkillsFrameworkLevel] NVARCHAR(20) NULL,
    [ComplianceCategory] NVARCHAR(100) NULL,
    [ExternalProvider] NVARCHAR(200) NULL,
    [ProviderContactEmail] NVARCHAR(100) NULL,
    [CourseMaterials] NVARCHAR(MAX) NULL,
    [EquipmentRequired] NVARCHAR(MAX) NULL,
    CONSTRAINT [PK_Courses] PRIMARY KEY CLUSTERED ([CourseId] ASC)
)
GO

CREATE UNIQUE NONCLUSTERED INDEX [IX_Courses_CourseCode] ON [dbo].[Courses] ([CourseCode] ASC)
GO

CREATE NONCLUSTERED INDEX [IX_Courses_Category] ON [dbo].[Courses] ([Category] ASC)
GO

CREATE NONCLUSTERED INDEX [IX_Courses_IsActive] ON [dbo].[Courses] ([IsActive] ASC)
GO

-- Create CoursePrerequisites table
CREATE TABLE [dbo].[CoursePrerequisites] (
    [PrerequisiteId] INT IDENTITY(1,1) NOT NULL,
    [CourseId] INT NOT NULL,
    [PrerequisiteCourseId] INT NOT NULL,
    [IsRequired] BIT NOT NULL DEFAULT 1,
    [CreatedDate] DATETIME NOT NULL DEFAULT GETDATE(),
    [CreatedBy] NVARCHAR(50) NOT NULL,
    CONSTRAINT [PK_CoursePrerequisites] PRIMARY KEY CLUSTERED ([PrerequisiteId] ASC),
    CONSTRAINT [FK_CoursePrerequisites_Course] FOREIGN KEY([CourseId]) REFERENCES [dbo].[Courses] ([CourseId]),
    CONSTRAINT [FK_CoursePrerequisites_PrerequisiteCourse] FOREIGN KEY([PrerequisiteCourseId]) REFERENCES [dbo].[Courses] ([CourseId])
)
GO

-- Create TrainingSessions table
CREATE TABLE [dbo].[TrainingSessions] (
    [SessionId] INT IDENTITY(1,1) NOT NULL,
    [CourseId] INT NOT NULL,
    [SessionDate] DATETIME NOT NULL,
    [StartTime] TIME NOT NULL,
    [EndTime] TIME NOT NULL,
    [Location] NVARCHAR(200) NOT NULL,
    [MaxParticipants] INT NOT NULL,
    [CurrentParticipants] INT NOT NULL DEFAULT 0,
    [PrimaryInstructorId] INT NOT NULL,
    [SecondaryInstructorId] INT NULL,
    [SessionStatus] NVARCHAR(20) NOT NULL DEFAULT 'SCHEDULED',
    [RegistrationDeadline] DATETIME NOT NULL,
    [SessionNotes] NVARCHAR(MAX) NULL,
    [InstructorNotes] NVARCHAR(MAX) NULL,
    [EquipmentRequired] NVARCHAR(MAX) NULL,
    [CateringRequired] BIT NOT NULL DEFAULT 0,
    [MaterialsPrepared] BIT NOT NULL DEFAULT 0,
    [RoomBooked] BIT NOT NULL DEFAULT 0,
    [NotificationsSent] BIT NOT NULL DEFAULT 0,
    [WaitingListEnabled] BIT NOT NULL DEFAULT 1,
    [PrerequisitesChecked] BIT NOT NULL DEFAULT 0,
    [CreatedDate] DATETIME NOT NULL DEFAULT GETDATE(),
    [CreatedBy] NVARCHAR(50) NOT NULL,
    [ModifiedDate] DATETIME NULL,
    [ModifiedBy] NVARCHAR(50) NULL,
    [CostPerParticipant] DECIMAL(10,2) NOT NULL DEFAULT 0,
    [TotalCost] DECIMAL(10,2) NOT NULL DEFAULT 0,
    [ApprovedBy] NVARCHAR(50) NULL,
    [ApprovedDate] DATETIME NULL,
    [CancelledReason] NVARCHAR(MAX) NULL,
    [FeedbackRequested] BIT NOT NULL DEFAULT 0,
    CONSTRAINT [PK_TrainingSessions] PRIMARY KEY CLUSTERED ([SessionId] ASC),
    CONSTRAINT [FK_TrainingSessions_Course] FOREIGN KEY([CourseId]) REFERENCES [dbo].[Courses] ([CourseId]),
    CONSTRAINT [FK_TrainingSessions_PrimaryInstructor] FOREIGN KEY([PrimaryInstructorId]) REFERENCES [dbo].[Employees] ([EmployeeId]),
    CONSTRAINT [FK_TrainingSessions_SecondaryInstructor] FOREIGN KEY([SecondaryInstructorId]) REFERENCES [dbo].[Employees] ([EmployeeId])
)
GO

CREATE NONCLUSTERED INDEX [IX_TrainingSessions_SessionDate] ON [dbo].[TrainingSessions] ([SessionDate] ASC)
GO

CREATE NONCLUSTERED INDEX [IX_TrainingSessions_CourseId] ON [dbo].[TrainingSessions] ([CourseId] ASC)
GO

CREATE NONCLUSTERED INDEX [IX_TrainingSessions_Status] ON [dbo].[TrainingSessions] ([SessionStatus] ASC)
GO

-- Create SessionParticipants table
CREATE TABLE [dbo].[SessionParticipants] (
    [ParticipantId] INT IDENTITY(1,1) NOT NULL,
    [SessionId] INT NOT NULL,
    [EmployeeId] INT NOT NULL,
    [EnrollmentDate] DATETIME NOT NULL DEFAULT GETDATE(),
    [EnrollmentStatus] NVARCHAR(20) NOT NULL DEFAULT 'ENROLLED',
    [WaitingListPosition] INT NULL,
    [NotificationsSent] BIT NOT NULL DEFAULT 0,
    [CreatedDate] DATETIME NOT NULL DEFAULT GETDATE(),
    [CreatedBy] NVARCHAR(50) NOT NULL,
    CONSTRAINT [PK_SessionParticipants] PRIMARY KEY CLUSTERED ([ParticipantId] ASC),
    CONSTRAINT [FK_SessionParticipants_Session] FOREIGN KEY([SessionId]) REFERENCES [dbo].[TrainingSessions] ([SessionId]),
    CONSTRAINT [FK_SessionParticipants_Employee] FOREIGN KEY([EmployeeId]) REFERENCES [dbo].[Employees] ([EmployeeId])
)
GO

CREATE UNIQUE NONCLUSTERED INDEX [IX_SessionParticipants_SessionEmployee] ON [dbo].[SessionParticipants] ([SessionId] ASC, [EmployeeId] ASC)
GO

-- Create TrainingRecords table
CREATE TABLE [dbo].[TrainingRecords] (
    [RecordId] INT IDENTITY(1,1) NOT NULL,
    [EmployeeId] INT NOT NULL,
    [SessionId] INT NOT NULL,
    [CourseId] INT NOT NULL,
    [EnrollmentDate] DATETIME NOT NULL,
    [CompletionDate] DATETIME NULL,
    [AttendanceStatus] NVARCHAR(20) NOT NULL DEFAULT 'ENROLLED',
    [CompletionStatus] NVARCHAR(20) NOT NULL DEFAULT 'NOT_STARTED',
    [Score] INT NULL,
    [MaxScore] INT NULL,
    [PassMark] INT NULL,
    [Grade] NVARCHAR(10) NULL,
    [CertificateIssued] BIT NOT NULL DEFAULT 0,
    [CertificateNumber] NVARCHAR(50) NULL,
    [CertificateIssuedDate] DATETIME NULL,
    [ExpiryDate] DATETIME NULL,
    [IsExpired] BIT NOT NULL DEFAULT 0,
    [Feedback] NVARCHAR(MAX) NULL,
    [InstructorComments] NVARCHAR(MAX) NULL,
    [CreatedDate] DATETIME NOT NULL DEFAULT GETDATE(),
    [CreatedBy] NVARCHAR(50) NOT NULL,
    [ModifiedDate] DATETIME NULL,
    [ModifiedBy] NVARCHAR(50) NULL,
    [RenewalRequired] BIT NOT NULL DEFAULT 0,
    [RenewalNotificationSent] BIT NOT NULL DEFAULT 0,
    [CostCentre] NVARCHAR(20) NULL,
    [ApprovalRequired] BIT NOT NULL DEFAULT 0,
    [ApprovedBy] NVARCHAR(50) NULL,
    [ApprovedDate] DATETIME NULL,
    [RejectedReason] NVARCHAR(MAX) NULL,
    CONSTRAINT [PK_TrainingRecords] PRIMARY KEY CLUSTERED ([RecordId] ASC),
    CONSTRAINT [FK_TrainingRecords_Employee] FOREIGN KEY([EmployeeId]) REFERENCES [dbo].[Employees] ([EmployeeId]),
    CONSTRAINT [FK_TrainingRecords_Session] FOREIGN KEY([SessionId]) REFERENCES [dbo].[TrainingSessions] ([SessionId]),
    CONSTRAINT [FK_TrainingRecords_Course] FOREIGN KEY([CourseId]) REFERENCES [dbo].[Courses] ([CourseId])
)
GO

CREATE NONCLUSTERED INDEX [IX_TrainingRecords_EmployeeId] ON [dbo].[TrainingRecords] ([EmployeeId] ASC)
GO

CREATE NONCLUSTERED INDEX [IX_TrainingRecords_CourseId] ON [dbo].[TrainingRecords] ([CourseId] ASC)
GO

CREATE NONCLUSTERED INDEX [IX_TrainingRecords_CompletionStatus] ON [dbo].[TrainingRecords] ([CompletionStatus] ASC)
GO

CREATE NONCLUSTERED INDEX [IX_TrainingRecords_ExpiryDate] ON [dbo].[TrainingRecords] ([ExpiryDate] ASC)
GO

-- Create AuditLog table
CREATE TABLE [dbo].[AuditLog] (
    [AuditId] BIGINT IDENTITY(1,1) NOT NULL,
    [TableName] NVARCHAR(100) NOT NULL,
    [RecordId] INT NOT NULL,
    [Action] NVARCHAR(20) NOT NULL,
    [UserName] NVARCHAR(50) NOT NULL,
    [Timestamp] DATETIME NOT NULL DEFAULT GETDATE(),
    [OldValues] NVARCHAR(MAX) NULL,
    [NewValues] NVARCHAR(MAX) NULL,
    [IPAddress] NVARCHAR(45) NULL,
    [UserAgent] NVARCHAR(500) NULL,
    CONSTRAINT [PK_AuditLog] PRIMARY KEY CLUSTERED ([AuditId] ASC)
)
GO

CREATE NONCLUSTERED INDEX [IX_AuditLog_TableName] ON [dbo].[AuditLog] ([TableName] ASC)
GO

CREATE NONCLUSTERED INDEX [IX_AuditLog_Timestamp] ON [dbo].[AuditLog] ([Timestamp] ASC)
GO

CREATE NONCLUSTERED INDEX [IX_AuditLog_UserName] ON [dbo].[AuditLog] ([UserName] ASC)
GO

-- Create SystemSettings table
CREATE TABLE [dbo].[SystemSettings] (
    [SettingId] INT IDENTITY(1,1) NOT NULL,
    [Category] NVARCHAR(100) NOT NULL,
    [SettingName] NVARCHAR(100) NOT NULL,
    [SettingValue] NVARCHAR(MAX) NOT NULL,
    [Description] NVARCHAR(500) NULL,
    [IsActive] BIT NOT NULL DEFAULT 1,
    [CreatedDate] DATETIME NOT NULL DEFAULT GETDATE(),
    [CreatedBy] NVARCHAR(50) NOT NULL,
    [ModifiedDate] DATETIME NULL,
    [ModifiedBy] NVARCHAR(50) NULL,
    CONSTRAINT [PK_SystemSettings] PRIMARY KEY CLUSTERED ([SettingId] ASC)
)
GO

CREATE UNIQUE NONCLUSTERED INDEX [IX_SystemSettings_CategoryName] ON [dbo].[SystemSettings] ([Category] ASC, [SettingName] ASC)
GO

-- Insert sample system settings
INSERT INTO [dbo].[SystemSettings] ([Category], [SettingName], [SettingValue], [Description], [CreatedBy])
VALUES
('Email', 'SMTPServer', 'smtp.defra.gov.uk', 'SMTP server for sending emails', 'SYSTEM'),
('Email', 'SMTPPort', '587', 'SMTP server port', 'SYSTEM'),
('Email', 'FromAddress', 'noreply@defra.gov.uk', 'Default from email address', 'SYSTEM'),
('Training', 'DefaultSessionDuration', '4', 'Default training session duration in hours', 'SYSTEM'),
('Training', 'MaxParticipants', '20', 'Default maximum participants per session', 'SYSTEM'),
('Training', 'RegistrationDeadlineDays', '5', 'Default registration deadline in days before session', 'SYSTEM'),
('Notifications', 'SendWelcomeEmail', 'true', 'Send welcome email to new participants', 'SYSTEM'),
('Notifications', 'SendReminderEmail', 'true', 'Send reminder email before training', 'SYSTEM'),
('Certificates', 'AutoIssue', 'true', 'Automatically issue certificates on completion', 'SYSTEM'),
('Certificates', 'ExpiryWarningDays', '90', 'Days before expiry to send warning', 'SYSTEM'),
('Security', 'SessionTimeoutMinutes', '60', 'User session timeout in minutes', 'SYSTEM'),
('Security', 'PasswordExpiryDays', '90', 'Password expiry period in days', 'SYSTEM')
GO

-- Insert sample data
-- Sample Employees
INSERT INTO [dbo].[Employees] (
    [EmployeeNumber], [FirstName], [LastName], [Email], [UserName], [Department],
    [Position], [Location], [HireDate], [CreatedBy], [PhoneNumber], [ContractType]
)
VALUES
('EMP001234', 'Sarah', 'Mitchell', 'sarah.mitchell@defra.gov.uk', 'smitchell', 'Human Resources',
 'Training Manager', 'London', '2015-03-15', 'SYSTEM', '020 7946 0123', 'Permanent'),
('EMP001235', 'Michael', 'Brown', 'michael.brown@defra.gov.uk', 'mbrown', 'Information Technology',
 'Senior Developer', 'Bristol', '2016-01-20', 'SYSTEM', '0117 946 0124', 'Permanent'),
('EMP001236', 'Emma', 'Clark', 'emma.clark@defra.gov.uk', 'eclark', 'Finance',
 'Financial Analyst', 'London', '2017-06-10', 'SYSTEM', '020 7946 0125', 'Permanent'),
('EMP001237', 'James', 'Davis', 'james.davis@defra.gov.uk', 'jdavis', 'Operations',
 'Operations Manager', 'York', '2014-09-05', 'SYSTEM', '01904 946 0126', 'Permanent'),
('EMP001238', 'Lisa', 'Evans', 'lisa.evans@defra.gov.uk', 'levans', 'Legal',
 'Legal Advisor', 'London', '2018-02-14', 'SYSTEM', '020 7946 0127', 'Permanent')
GO

-- Sample Courses
INSERT INTO [dbo].[Courses] (
    [CourseCode], [Title], [Description], [Category], [DurationHours], [DeliveryMethod],
    [IsCompulsory], [MaxParticipants], [Prerequisites], [CreatedBy]
)
VALUES
('HS001', 'Health & Safety Fundamentals', 'Basic health and safety training for all employees',
 'Health & Safety', 4.0, 'Blended Learning', 1, 20, 'None', 'SYSTEM'),
('DP001', 'Data Protection Fundamentals', 'Introduction to GDPR and data protection principles',
 'Data Protection', 3.0, 'Online', 1, 25, 'None', 'SYSTEM'),
('IT001', 'IT Security Awareness', 'Cybersecurity awareness and best practices',
 'Information Security', 2.0, 'Online', 1, 30, 'None', 'SYSTEM'),
('CS001', 'Customer Service Excellence', 'Delivering excellent customer service',
 'Customer Service', 6.0, 'Classroom', 0, 15, 'None', 'SYSTEM'),
('LM001', 'Leadership Fundamentals', 'Basic leadership and management skills',
 'Leadership', 8.0, 'Classroom', 0, 12, 'CS001', 'SYSTEM')
GO

-- Sample Training Sessions
INSERT INTO [dbo].[TrainingSessions] (
    [CourseId], [SessionDate], [StartTime], [EndTime], [Location], [MaxParticipants],
    [PrimaryInstructorId], [RegistrationDeadline], [CreatedBy]
)
SELECT
    c.CourseId,
    DATEADD(day, 30, GETDATE()),
    '09:00:00',
    '13:00:00',
    'Training Room A',
    20,
    1, -- Sarah Mitchell as instructor
    DATEADD(day, 25, GETDATE()),
    'SYSTEM'
FROM [dbo].[Courses] c
WHERE c.CourseCode = 'HS001'

INSERT INTO [dbo].[TrainingSessions] (
    [CourseId], [SessionDate], [StartTime], [EndTime], [Location], [MaxParticipants],
    [PrimaryInstructorId], [RegistrationDeadline], [CreatedBy]
)
SELECT
    c.CourseId,
    DATEADD(day, 45, GETDATE()),
    '10:00:00',
    '13:00:00',
    'Online Session',
    25,
    2, -- Michael Brown as instructor
    DATEADD(day, 40, GETDATE()),
    'SYSTEM'
FROM [dbo].[Courses] c
WHERE c.CourseCode = 'DP001'
GO

-- Create views for reporting
CREATE VIEW [dbo].[vw_EmployeeTrainingStatus] AS
SELECT
    e.EmployeeId,
    e.EmployeeNumber,
    e.FirstName + ' ' + e.LastName AS FullName,
    e.Department,
    e.Position,
    COUNT(DISTINCT tr.CourseId) AS CoursesCompleted,
    COUNT(DISTINCT CASE WHEN tr.CompletionStatus = 'COMPLETED' THEN tr.CourseId END) AS CoursesPassedCount,
    COUNT(DISTINCT CASE WHEN tr.ExpiryDate < GETDATE() AND tr.CompletionStatus = 'COMPLETED' THEN tr.CourseId END) AS ExpiredCertificatesCount,
    COUNT(DISTINCT CASE WHEN tr.ExpiryDate BETWEEN GETDATE() AND DATEADD(day, 90, GETDATE()) AND tr.CompletionStatus = 'COMPLETED' THEN tr.CourseId END) AS ExpiringCertificatesCount
FROM [dbo].[Employees] e
LEFT JOIN [dbo].[TrainingRecords] tr ON e.EmployeeId = tr.EmployeeId
WHERE e.IsActive = 1
GROUP BY e.EmployeeId, e.EmployeeNumber, e.FirstName, e.LastName, e.Department, e.Position
GO

CREATE VIEW [dbo].[vw_CourseStatistics] AS
SELECT
    c.CourseId,
    c.CourseCode,
    c.Title,
    c.Category,
    COUNT(DISTINCT ts.SessionId) AS TotalSessions,
    COUNT(DISTINCT CASE WHEN ts.SessionStatus IN ('SCHEDULED', 'CONFIRMED') AND ts.SessionDate >= GETDATE() THEN ts.SessionId END) AS UpcomingSessions,
    COUNT(DISTINCT tr.EmployeeId) AS TotalParticipants,
    COUNT(DISTINCT CASE WHEN tr.CompletionStatus = 'COMPLETED' THEN tr.EmployeeId END) AS CompletedParticipants,
    CASE
        WHEN COUNT(DISTINCT tr.EmployeeId) > 0
        THEN CAST(COUNT(DISTINCT CASE WHEN tr.CompletionStatus = 'COMPLETED' THEN tr.EmployeeId END) AS DECIMAL) / COUNT(DISTINCT tr.EmployeeId) * 100
        ELSE 0
    END AS CompletionRate
FROM [dbo].[Courses] c
LEFT JOIN [dbo].[TrainingSessions] ts ON c.CourseId = ts.CourseId
LEFT JOIN [dbo].[TrainingRecords] tr ON c.CourseId = tr.CourseId
WHERE c.IsActive = 1
GROUP BY c.CourseId, c.CourseCode, c.Title, c.Category
GO

-- Create stored procedures for common operations
CREATE PROCEDURE [dbo].[sp_GetEmployeeComplianceStatus]
    @EmployeeId INT
AS
BEGIN
    SET NOCOUNT ON;

    SELECT
        c.CourseCode,
        c.Title,
        c.IsCompulsory,
        tr.CompletionStatus,
        tr.CompletionDate,
        tr.ExpiryDate,
        CASE
            WHEN c.IsCompulsory = 1 AND tr.CompletionStatus IS NULL THEN 'REQUIRED'
            WHEN c.IsCompulsory = 1 AND tr.CompletionStatus = 'COMPLETED' AND tr.ExpiryDate < GETDATE() THEN 'EXPIRED'
            WHEN c.IsCompulsory = 1 AND tr.CompletionStatus = 'COMPLETED' AND tr.ExpiryDate >= GETDATE() THEN 'COMPLIANT'
            WHEN c.IsCompulsory = 1 AND tr.CompletionStatus <> 'COMPLETED' THEN 'NON_COMPLIANT'
            ELSE 'OPTIONAL'
        END AS ComplianceStatus
    FROM [dbo].[Courses] c
    LEFT JOIN [dbo].[TrainingRecords] tr ON c.CourseId = tr.CourseId AND tr.EmployeeId = @EmployeeId
    WHERE c.IsActive = 1
    ORDER BY c.IsCompulsory DESC, c.CourseCode
END
GO

CREATE PROCEDURE [dbo].[sp_GetDashboardMetrics]
AS
BEGIN
    SET NOCOUNT ON;

    -- Active employees
    DECLARE @ActiveEmployees INT = (SELECT COUNT(*) FROM [dbo].[Employees] WHERE IsActive = 1)

    -- Total courses
    DECLARE @TotalCourses INT = (SELECT COUNT(*) FROM [dbo].[Courses] WHERE IsActive = 1)

    -- Upcoming sessions
    DECLARE @UpcomingSessions INT = (SELECT COUNT(*) FROM [dbo].[TrainingSessions] WHERE SessionDate >= GETDATE() AND SessionStatus IN ('SCHEDULED', 'CONFIRMED'))

    -- Training records completed this month
    DECLARE @CompletedThisMonth INT = (SELECT COUNT(*) FROM [dbo].[TrainingRecords] WHERE CompletionStatus = 'COMPLETED' AND CompletionDate >= DATEADD(month, DATEDIFF(month, 0, GETDATE()), 0))

    -- Overall compliance rate
    DECLARE @ComplianceRate DECIMAL(5,2) = (
        SELECT
            CASE
                WHEN COUNT(*) = 0 THEN 0
                ELSE CAST(SUM(CASE WHEN tr.CompletionStatus = 'COMPLETED' AND tr.ExpiryDate >= GETDATE() THEN 1 ELSE 0 END) AS DECIMAL) / COUNT(*) * 100
            END
        FROM [dbo].[Employees] e
        CROSS JOIN [dbo].[Courses] c
        LEFT JOIN [dbo].[TrainingRecords] tr ON e.EmployeeId = tr.EmployeeId AND c.CourseId = tr.CourseId
        WHERE e.IsActive = 1 AND c.IsActive = 1 AND c.IsCompulsory = 1
    )

    SELECT
        @ActiveEmployees AS ActiveEmployees,
        @TotalCourses AS TotalCourses,
        @UpcomingSessions AS UpcomingSessions,
        @CompletedThisMonth AS CompletedThisMonth,
        @ComplianceRate AS ComplianceRate
END
GO

-- Grant permissions to application roles
-- Note: In production, these would be actual Windows groups
IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = N'TrainTrack_Users')
    CREATE ROLE [TrainTrack_Users]
GO

IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = N'TrainTrack_Admins')
    CREATE ROLE [TrainTrack_Admins]
GO

-- Grant basic permissions to users
GRANT SELECT ON [dbo].[vw_EmployeeTrainingStatus] TO [TrainTrack_Users]
GRANT SELECT ON [dbo].[vw_CourseStatistics] TO [TrainTrack_Users]
GRANT EXECUTE ON [dbo].[sp_GetEmployeeComplianceStatus] TO [TrainTrack_Users]
GRANT EXECUTE ON [dbo].[sp_GetDashboardMetrics] TO [TrainTrack_Users]

-- Grant full permissions to admins
GRANT SELECT, INSERT, UPDATE ON [dbo].[Employees] TO [TrainTrack_Admins]
GRANT SELECT, INSERT, UPDATE ON [dbo].[Courses] TO [TrainTrack_Admins]
GRANT SELECT, INSERT, UPDATE, DELETE ON [dbo].[TrainingSessions] TO [TrainTrack_Admins]
GRANT SELECT, INSERT, UPDATE ON [dbo].[TrainingRecords] TO [TrainTrack_Admins]
GRANT SELECT ON [dbo].[AuditLog] TO [TrainTrack_Admins]
GRANT SELECT, INSERT, UPDATE ON [dbo].[SystemSettings] TO [TrainTrack_Admins]

PRINT 'TrainTrack database schema created successfully.'
PRINT 'Schema version: 2.1'
PRINT 'Total tables created: 9'
PRINT 'Total views created: 2'
PRINT 'Total stored procedures created: 2'
GO
# TrainTrack - Training Management System v2.1

**Purpose**: This is a fabricated legacy application created for testing Defra's AI-enabled modernisation playbook. It serves as realistic source material for user research and validation of the reverse engineering process.

## Application Overview

TrainTrack is a legacy training records management system that handles:
- Employee training records and compliance tracking
- Course catalog and scheduling management
- Training completion tracking and certification management
- Reporting and analytics on training effectiveness
- Instructor and venue management

## Technology Stack

- **Frontend**: ASP.NET Web Forms (VB.NET)
- **Backend**: VB.NET with SQL Server 2012
- **Server**: Windows Server 2008 R2, Defra data centre
- **Auth**: Windows Authentication (Active Directory)
- **Architecture**: Traditional 3-tier web application

## Directory Structure

```
TrainTrack/
├── transcripts/          # Demo walkthrough transcripts
├── src/                  # VB.NET source code
├── html/                 # HTML mockups (for screenshot generation)
└── README.md            # This file
```

### Transcripts Details

The `transcripts/` directory contains a realistic demonstration recording:

- **traintrack_demo.txt** (57 minutes) - Meeting recording of system demonstration by Training Manager Sarah Mitchell showing business functionality, workflows, and pain points to the modernization committee

This transcript simulates a real stakeholder demo session that would be conducted during modernization discovery phases.

## Usage for Playbook Testing

1. **HTML Mockups**: Open the styled HTML files in `html/` in a browser - they're designed to look like realistic legacy application screens
2. **Screenshots**: Take screenshots of each HTML mockup for use as playbook inputs
3. **Transcripts**: Use the demo transcript files in `transcripts/` as stakeholder interview inputs
4. **Source Code**: The `src/` directory contains a complete VB.NET Web Forms application
5. **Testing**: Run this through the AI modernisation playbook to validate the process

### Screenshot Generation Process

The HTML mockups in `html/` are styled to look like authentic legacy Web Forms applications with:
- Realistic form layouts and controls
- Legacy styling (Web Forms aesthetic)
- Sample data that looks believable
- Proper navigation and branding

Simply open each HTML file in a browser and take full-window screenshots for the best results.

### Source Code Structure

The `src/` directory contains a complete VB.NET Web Forms application:

```
src/
├── TrainTrack.vbproj        # Visual Studio project file
├── Web.config               # Application configuration
├── Global.asax              # Application startup
├── Global.asax.vb           # Application startup code-behind
├── Default.aspx             # Main entry point
├── Default.aspx.vb          # Main entry point code-behind
├── App_Code/                # Business logic and utilities
│   ├── Models/              # Data models
│   │   ├── Employee.vb
│   │   ├── Course.vb
│   │   ├── TrainingSession.vb
│   │   └── TrainingRecord.vb
│   ├── BusinessLogic/       # Managers and business rules
│   │   ├── EmployeeManager.vb
│   │   └── CourseManager.vb
│   ├── DataAccess/          # Database helper and repositories
│   │   ├── DatabaseHelper.vb
│   │   ├── EmployeeRepository.vb
│   │   ├── CourseRepository.vb
│   │   └── TrainingRepository.vb
│   └── Utilities/           # Email and helper classes
│       └── EmailHelper.vb
└── Database/                # Database schema and sample data
    └── TrainTrack_Schema.sql
```

The application demonstrates typical legacy patterns including:
- Three-tier architecture with Web Forms presentation layer
- VB.NET business logic with extensive error handling
- SQL Server database with stored procedures
- Email notifications via SMTP
- Windows Authentication integration
- Manual database connection management

## Application Screens

The TrainTrack application contains 12 main screens covering all aspects of training management:

1. **Dashboard** (`01-dashboard.html`) - Overview of training metrics and pending items
2. **Employee Search** (`02-employee-search.html`) - Find and filter employee records
3. **Employee Profile** (`03-employee-profile.html`) - Individual employee training history
4. **Course Catalog** (`04-course-catalog.html`) - Browse available training courses
5. **Course Details** (`05-course-details.html`) - Detailed course information and enrollment
6. **Training Schedule** (`06-training-schedule.html`) - Calendar view of scheduled training
7. **Reports** (`07-reports.html`) - Training effectiveness and compliance reports
8. **User Management** (`08-user-management.html`) - Manage system users and permissions
9. **Schedule Training** (`09-schedule-training.html`) - Book training sessions
10. **Training Records** (`10-training-records.html`) - View/edit training completion records
11. **Course Creation** (`11-course-creation.html`) - Add/edit training courses
12. **System Settings** (`12-system-settings.html`) - Configure application settings

## Fabricated Nature

This application is completely fictional but designed to be realistic enough to test the modernisation playbook effectively. All data, screenshots, transcripts, and source code are fabricated for research purposes.
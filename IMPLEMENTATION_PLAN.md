# TrainTrack Coherence Fixes - Implementation Plan

This plan resolves all mismatches between the HTML mockups, transcript, and source code.
Each step is self-contained and can be executed in a separate chat session.

---

## Step 1: Add CourseModule Model and Refactor Course Content

### Problem
The HTML mockups (05-course-details.html, 11-course-creation.html) show structured course
modules with individual Title, Type, Duration, and Description fields. The code only has a
flat `CourseContent` text field on the `Course` model (src/App_Code/Models/Course.vb, line 18).
The database schema has no modules table.

### What to Do

- [x] **1a.** Create a new `CourseModule.vb` model

Create file: `src/App_Code/Models/CourseModule.vb`

Fields to include (matching HTML table columns in 05-course-details.html lines 114-148
and 11-course-creation.html lines 152-291):

| Field | Type | Description |
|-------|------|-------------|
| ModuleId | Integer | Primary key |
| CourseId | Integer | FK to Courses table |
| ModuleOrder | Integer | Display order (1, 2, 3...) |
| Title | String | Module title, e.g. "Introduction to Health & Safety" |
| ModuleType | String | One of: "Online", "Classroom", "Video", "Reading", "Quiz", "Interactive" |
| DurationMinutes | Integer | Duration in minutes |
| Description | String | Module description text |
| CreatedDate | DateTime | Audit field |
| ModifiedDate | DateTime | Audit field |

Include:
- A read-only `DurationText` property that formats minutes as "X hrs Y mins" or "X mins"
- An `IsValid()` method requiring Title and ModuleOrder
- Follow the same VB.NET patterns used in the existing models (private backing fields,
  Public Property with Get/Set, constructor overloads)

- [x] **1b.** Add a `Modules` navigation property to `Course.vb`

In `src/App_Code/Models/Course.vb`:
- Add a `Private _modules As List(Of CourseModule)` field (after line 18)
- Add a `Public Property Modules() As List(Of CourseModule)` property
- Add a read-only `TotalModuleDuration` computed property that sums all module DurationMinutes
- Keep the existing `CourseContent` field for backward compatibility (legacy data)

- [x] **1c.** Add `CourseModules` table to the database schema

In `src/TrainTrack_Schema.sql`, add:

```sql
CREATE TABLE [dbo].[CourseModules] (
    [ModuleId]        INT IDENTITY(1,1) PRIMARY KEY,
    [CourseId]        INT NOT NULL,
    [ModuleOrder]     INT NOT NULL,
    [Title]           NVARCHAR(200) NOT NULL,
    [ModuleType]      NVARCHAR(50) NOT NULL,
    [DurationMinutes] INT NOT NULL DEFAULT 0,
    [Description]     NVARCHAR(MAX) NULL,
    [CreatedDate]     DATETIME NOT NULL DEFAULT GETDATE(),
    [ModifiedDate]    DATETIME NULL,
    CONSTRAINT [FK_CourseModules_Courses] FOREIGN KEY ([CourseId])
        REFERENCES [dbo].[Courses] ([CourseId])
);

CREATE INDEX [IX_CourseModules_CourseId] ON [dbo].[CourseModules] ([CourseId]);
```

- [x] **1d.** Add repository methods

In `src/App_Code/DataAccess/CourseRepository.vb`, add:
- `GetModulesByCourseId(courseId As Integer) As List(Of CourseModule)`
- `CreateModule(module As CourseModule) As Integer`
- `UpdateModule(module As CourseModule) As Boolean`
- `DeleteModule(moduleId As Integer) As Boolean`
- `ReorderModules(courseId As Integer, moduleIds As List(Of Integer)) As Boolean`

Follow the same patterns as existing repository methods (use `DatabaseHelper`, parameterised
queries, `MapModuleFromReader` helper).

- [x] **1e.** Add business logic methods

In `src/App_Code/BusinessLogic/CourseManager.vb`, add:
- `GetCourseModules(courseId As Integer) As List(Of CourseModule)`
- `AddModuleToCourse(courseId As Integer, module As CourseModule, createdBy As String) As Integer`
- `UpdateCourseModule(module As CourseModule, modifiedBy As String) As Boolean`
- `RemoveCourseModule(moduleId As Integer, modifiedBy As String) As Boolean`
- `ReorderCourseModules(courseId As Integer, moduleIds As List(Of Integer), modifiedBy As String) As Boolean`

Include validation: module Title required, DurationMinutes >= 0, CourseId must exist.

### Verification
- [ ] The `CourseModule` model fields match the HTML table columns exactly
- [ ] The 5 sample modules in 05-course-details.html (lines 114-148) are representable
- [ ] The module types in the dropdown in 11-course-creation.html (VIDEO, READING, QUIZ, INTERACTIVE at line 156) are all valid `ModuleType` values

---

## Step 2: Add CSV Import Functionality for Employee Data

### Problem
The transcript (lines 349, 535-667, 1075) describes a critical monthly workflow where HR
exports employee data as CSV, and it's manually uploaded into TrainTrack. The current code
has no import functionality -- only individual CRUD via `EmployeeManager`.

### What to Do

- [x] **2a.** Create an `EmployeeImportResult.vb` model

Create file: `src/App_Code/Models/EmployeeImportResult.vb`

Fields:

| Field | Type | Description |
|-------|------|-------------|
| TotalRows | Integer | Total rows in CSV |
| NewEmployees | Integer | Count of new records created |
| UpdatedEmployees | Integer | Count of existing records updated |
| SkippedRows | Integer | Count of rows skipped (errors/duplicates) |
| Errors | List(Of String) | Detailed error messages per row |
| Warnings | List(Of String) | Non-fatal warnings (e.g. blank optional fields) |
| ImportDate | DateTime | When import was executed |
| ImportedBy | String | Who ran the import |

Include a read-only `IsSuccess` property (True if Errors.Count = 0).

- [x] **2b.** Create an `EmployeeImportService.vb` business logic class

Create file: `src/App_Code/BusinessLogic/EmployeeImportService.vb`

This is the core import logic. Methods:

- `ValidateCsvFile(filePath As String) As EmployeeImportResult`
  - Parse CSV headers and validate expected columns exist
  - Expected columns (matching existing Employee model fields):
    EmployeeNumber, FirstName, LastName, Email, Department, Position, Location,
    ManagerId, PhoneNumber, LineManagerEmail, CostCentre, PayBand, ContractType,
    WorkingPattern, HireDate
  - Return errors for missing required columns
  - Return warnings for unrecognised extra columns

- `PreviewImport(filePath As String) As List(Of EmployeeImportPreviewRow)`
  - Parse CSV and for each row, determine: NEW (no matching EmployeeNumber),
    UPDATE (matching EmployeeNumber found), or ERROR (validation failure)
  - This is the **staging/preview** the transcript says is missing (line 607-631)
  - Return a list showing what WOULD happen without actually doing it

- `ExecuteImport(filePath As String, importedBy As String) As EmployeeImportResult`
  - For each row: match on EmployeeNumber using `EmployeeRepository.GetEmployeeByNumber()`
  - If match found: update fields (but NOT blank out fields if CSV value is empty --
    this addresses the transcript's complaint at lines 646-667 about blank columns
    wiping data)
  - If no match: create new employee via `EmployeeRepository.CreateEmployee()`
  - Log each action to audit log
  - Return comprehensive result

- `GenerateImportReport(result As EmployeeImportResult) As String`
  - Format a human-readable summary of the import

- [x] **2c.** Add CSV parsing utility

Create file: `src/App_Code/Utilities/CsvParser.vb`

A simple CSV parser:
- `ParseCsvFile(filePath As String) As DataTable`
- Handle quoted fields, commas in values, header row
- Return DataTable with columns matching CSV headers

Keep it simple -- no third-party dependencies (consistent with the legacy codebase style).

- [x] **2d.** Add import repository methods

In `src/App_Code/DataAccess/EmployeeRepository.vb`, add:
- `BulkCreateEmployees(employees As List(Of Employee)) As Integer` -- returns count created
- `BulkUpdateEmployees(employees As List(Of Employee)) As Integer` -- returns count updated

These wrap the existing `CreateEmployee` and `UpdateEmployee` in a transaction for atomicity.

### Key Business Rules (from transcript)
- Match on EmployeeNumber (transcript line 577: "It matches on employee number")
- Do NOT overwrite non-empty fields with blank CSV values (fix for transcript line 646-667 incident)
- Log all changes for audit trail
- The preview step addresses transcript lines 607-631 ("no staging area", "no preview")

### Verification
- [ ] The import workflow matches what the transcript describes: CSV in, match on employee number, create or update
- [ ] The preview/staging capability addresses the gap the transcript identifies
- [ ] The blank-field protection addresses the specific incident described at line 646

---

## Step 3: Align Role Names Across All Artefacts

### Problem
Role names are inconsistent across three artefacts:

| Transcript (lines 2650-2671) | HTML 08-user-management.html | Code Default.aspx.vb (lines 134-146) |
|------------------------------|------------------------------|--------------------------------------|
| Administrator | System Administrator | "Administrator" (from TrainTrack_Admins) |
| Training Manager | Training Administrator | "Manager" (from TrainTrack_Managers) |
| -- | HR Manager | -- |
| Department Manager | Department Manager | -- |
| Employee | Standard User | "Standard User" (from TrainTrack_Users) |
| -- | -- | "Instructor" (from TrainTrack_Instructors) |
| -- | -- | "Reports User" (from TrainTrack_Reports) |

### What to Do

Adopt a **canonical set of 5 roles** that reconciles all three sources. The code's AD groups
are the most authoritative (they reflect what's actually deployed), so use those as the base
and update the HTML and transcript to match.

#### Canonical Role Names

| AD Group | Display Name | Abbrev (for tables) |
|----------|-------------|---------------------|
| DEFRA\TrainTrack_Admins | System Administrator | Sys Admin |
| DEFRA\TrainTrack_Managers | Training Manager | Training Mgr |
| DEFRA\TrainTrack_Instructors | Instructor | Instructor |
| DEFRA\TrainTrack_Reports | Reports User | Reports |
| DEFRA\TrainTrack_Users | Standard User | Std User |

Note: The transcript's "Department Manager" and the HTML's "HR Manager" roles do not exist
as AD groups. They should be removed from the HTML or added to the code. Since the transcript
says there are only 4 roles and the code has 5 AD groups, the simplest fix is to update the
HTML to match the 5 code roles.

- [x] **3a.** Update `Default.aspx.vb` role display names

In `src/Default.aspx.vb`, lines 134-146, update the `GetUserRole()` return values:

| Line | Current Return Value | New Return Value |
|------|---------------------|-----------------|
| 135 | "Administrator" | "System Administrator" |
| 137 | "Manager" | "Training Manager" |
| 139 | "Instructor" | "Instructor" (no change) |
| 141 | "Reports User" | "Reports User" (no change) |
| 143 | "Standard User" | "Standard User" (no change) |

- [x] **3b.** Update `08-user-management.html` role names

In `html/08-user-management.html`:

- Lines 42-49 (Role filter dropdown): Replace the 5 options with:
  System Administrator, Training Manager, Instructor, Reports User, Standard User

- Lines 210-275 (Role Permissions Matrix): Replace headers with the same 5 roles

- Lines 53-102 (Users table): Update the Role column values for sample users to use
  the canonical names

- Remove "HR Manager" and "Department Manager" -- these don't exist as AD groups

- [x] **3c.** Update the transcript

In `transcripts/traintrack_demo.txt`, search for the role descriptions around lines 2650-2671
and update "Administrator, Training Manager, Department Manager, and Employee" to match
the canonical 5 roles: "System Administrator, Training Manager, Instructor, Reports User,
and Standard User".

Also update any other role references throughout the transcript to use consistent names.

### Verification
- [ ] Searching for old role names ("HR Manager", "Dept Manager", "Department Manager") returns zero results across all files
- [ ] All three artefacts reference the same 5 roles with identical display names

---

## Step 4: Align Enhanced Features with Transcript (Remove or Document)

### Problem
The transcript explicitly says several features DON'T exist in the current system, but the
HTML mockups and code include them. This is the most significant coherence issue.

Affected features:
- **Cost tracking**: Transcript says "no cost fields" but `Course.CostPerParticipant` (line 22)
  and `TrainingSession.CostPerParticipant` (line 30) / `TotalCost` (line 31) exist
- **Equipment tracking**: Transcript says "no equipment tracking" but
  `TrainingSession.EquipmentRequired` exists (line 19)
- **Room booking**: Transcript says "no room booking" but `TrainingSession.RoomBooked`
  exists (line 22), and 06-training-schedule.html shows room utilization (lines 385-431)
- **Catering**: Not a transcript feature but `TrainingSession.CateringRequired` exists (line 20)
- **Calendar integration**: Transcript says "no calendar integration" but Web.config has
  `CalendarIntegrationEnabled=true` (line 36) and `ExchangeServerURL` (line 37)

### Decision: Align Everything to "As-Is" State

The transcript is the authoritative record of the actual system. The HTML mockups and code
should represent what the system ACTUALLY does. Remove or neutralise the features the
transcript says don't exist.

- [x] **4a.** Neutralise cost fields in the code

In `src/App_Code/Models/Course.vb`:
- Keep the `CostPerParticipant` field (line 22) but add a comment:
  `' NOTE: Field exists in schema but is not populated. Cost tracking is manual via Finance team.`
- Update `CostText` property (lines 226-234) to return "Not tracked" when value is 0

In `src/App_Code/Models/TrainingSession.vb`:
- Add comments to `CostPerParticipant` (line 30) and `TotalCost` (line 31):
  `' NOTE: Not used in current system. Costs tracked externally by Finance.`

- [x] **4b.** Neutralise equipment/catering/room fields in the code

In `src/App_Code/Models/TrainingSession.vb`:
- Add comments to `EquipmentRequired` (line 19), `CateringRequired` (line 20),
  `RoomBooked` (line 22):
  `' NOTE: Field exists but feature not implemented. Equipment/catering/rooms managed manually outside system.`

- [ ] **4c.** Remove Resource Requirements section from HTML

In `html/09-schedule-training.html`:
- Remove or comment out the entire Resource Requirements section (lines 291-342)
  covering Equipment Needed, Catering, and Materials checkboxes
- Replace with a simple note: "Equipment, catering, and room booking are managed
  outside this system."

- [ ] **4d.** Remove Room Utilization section from HTML

In `html/06-training-schedule.html`:
- Remove or comment out the Training Room Utilization section (lines 385-431)
- This feature doesn't exist per the transcript

- [ ] **4e.** Disable calendar integration in config

In `src/Web.config`:
- Change line 36: `CalendarIntegrationEnabled` from `true` to `false`
- Add XML comment: `<!-- Calendar integration not yet implemented. Rooms and calendars managed via Outlook manually. -->`

- [ ] **4f.** Update the transcript to acknowledge schema fields

In `transcripts/traintrack_demo.txt`, where cost tracking and room booking are discussed,
add parenthetical notes clarifying that while database fields exist for these features,
they are not populated or used in the current system. This explains why the schema has
them but the system doesn't use them (likely added during development but never wired up).

### Verification
- [ ] HTML 09 no longer shows equipment/catering/materials checkboxes
- [ ] HTML 06 no longer shows room utilization table
- [ ] Web.config shows CalendarIntegrationEnabled=false
- [ ] Code fields have explanatory comments
- [ ] No functional code was removed (fields stay for forward compatibility)

---

## Step 5: Add Employee Notes Model

### Problem
The transcript describes notes on employee profiles for tracking follow-up actions
(transcript discussion of notes section and Outlook reminders). HTML 03-employee-profile.html
shows a "Training Notes" section with notes and an "Add Note" capability. But the code has
no Notes model -- only `SessionNotes` and `InstructorNotes` strings on `TrainingSession`.

### What to Do

- [ ] **5a.** Create `EmployeeNote.vb` model

Create file: `src/App_Code/Models/EmployeeNote.vb`

Fields:

| Field | Type | Description |
|-------|------|-------------|
| NoteId | Integer | Primary key |
| EmployeeId | Integer | FK to Employees table |
| NoteText | String | The note content |
| NoteType | String | "General", "Follow-up", "Compliance", "Training" |
| CreatedDate | DateTime | When note was created |
| CreatedBy | String | Who created the note |
| IsResolved | Boolean | For follow-up notes: has the action been taken? |
| ResolvedDate | DateTime (nullable) | When follow-up was resolved |
| ResolvedBy | String | Who resolved it |

Follow the same VB.NET patterns as other models.

- [ ] **5b.** Add `EmployeeNotes` table to schema

In `src/TrainTrack_Schema.sql`:

```sql
CREATE TABLE [dbo].[EmployeeNotes] (
    [NoteId]       INT IDENTITY(1,1) PRIMARY KEY,
    [EmployeeId]   INT NOT NULL,
    [NoteText]     NVARCHAR(MAX) NOT NULL,
    [NoteType]     NVARCHAR(50) NOT NULL DEFAULT 'General',
    [CreatedDate]  DATETIME NOT NULL DEFAULT GETDATE(),
    [CreatedBy]    NVARCHAR(100) NOT NULL,
    [IsResolved]   BIT NOT NULL DEFAULT 0,
    [ResolvedDate] DATETIME NULL,
    [ResolvedBy]   NVARCHAR(100) NULL,
    CONSTRAINT [FK_EmployeeNotes_Employees] FOREIGN KEY ([EmployeeId])
        REFERENCES [dbo].[Employees] ([EmployeeId])
);

CREATE INDEX [IX_EmployeeNotes_EmployeeId] ON [dbo].[EmployeeNotes] ([EmployeeId]);
```

- [ ] **5c.** Add repository methods

In `src/App_Code/DataAccess/EmployeeRepository.vb`, add:
- `GetNotesByEmployeeId(employeeId As Integer) As List(Of EmployeeNote)`
- `CreateNote(note As EmployeeNote) As Integer`
- `UpdateNote(note As EmployeeNote) As Boolean`
- `ResolveNote(noteId As Integer, resolvedBy As String) As Boolean`
- `GetUnresolvedNotes() As List(Of EmployeeNote)` -- for Sarah's follow-up tracking

- [ ] **5d.** Add business logic methods

In `src/App_Code/BusinessLogic/EmployeeManager.vb`, add:
- `GetEmployeeNotes(employeeId As Integer) As List(Of EmployeeNote)`
- `AddEmployeeNote(employeeId As Integer, noteText As String, noteType As String, createdBy As String) As Integer`
- `ResolveEmployeeNote(noteId As Integer, resolvedBy As String) As Boolean`
- `GetAllUnresolvedNotes() As List(Of EmployeeNote)` -- cross-employee view

- [ ] **5e.** Add a `Notes` property to the `Employee` model

In `src/App_Code/Models/Employee.vb`:
- Add `Private _notes As List(Of EmployeeNote)` field
- Add `Public Property Notes() As List(Of EmployeeNote)` property
- Add read-only `UnresolvedNoteCount As Integer` computed property

### Verification
- [ ] The note structure matches what HTML 03-employee-profile.html shows in the Training Notes section
- [ ] The NoteType values cover the use cases described in the transcript (follow-up actions, compliance notes, general observations)

---

## Step 6: Add Waiting List Model and Basic Logic

### Problem
The transcript (lines 1672-1753) describes waiting list management as a critical workflow
that's currently handled via separate spreadsheets. The code has only a
`WaitingListEnabled` boolean flag on `TrainingSession` (line 24) but no actual waiting
list data model or management logic.

### What to Do

- [ ] **6a.** Create `WaitingListEntry.vb` model

Create file: `src/App_Code/Models/WaitingListEntry.vb`

Fields:

| Field | Type | Description |
|-------|------|-------------|
| EntryId | Integer | Primary key |
| SessionId | Integer | FK to TrainingSessions |
| EmployeeId | Integer | FK to Employees |
| RequestDate | DateTime | When they joined the waiting list |
| Position | Integer | Queue position (1 = next in line) |
| Status | String | "Waiting", "Offered", "Accepted", "Declined", "Expired", "Enrolled" |
| OfferedDate | DateTime (nullable) | When a space was offered |
| ResponseDeadline | DateTime (nullable) | Deadline to accept the offer |
| Notes | String | Free text |
| CreatedBy | String | Who added them to the list |

Include:
- Read-only `IsActive` property (Status = "Waiting" or "Offered")
- Read-only `DaysWaiting` computed property

- [ ] **6b.** Add `WaitingList` table to schema

In `src/TrainTrack_Schema.sql`:

```sql
CREATE TABLE [dbo].[WaitingList] (
    [EntryId]           INT IDENTITY(1,1) PRIMARY KEY,
    [SessionId]         INT NOT NULL,
    [EmployeeId]        INT NOT NULL,
    [RequestDate]       DATETIME NOT NULL DEFAULT GETDATE(),
    [Position]          INT NOT NULL,
    [Status]            NVARCHAR(50) NOT NULL DEFAULT 'Waiting',
    [OfferedDate]       DATETIME NULL,
    [ResponseDeadline]  DATETIME NULL,
    [Notes]             NVARCHAR(MAX) NULL,
    [CreatedBy]         NVARCHAR(100) NOT NULL,
    CONSTRAINT [FK_WaitingList_Sessions] FOREIGN KEY ([SessionId])
        REFERENCES [dbo].[TrainingSessions] ([SessionId]),
    CONSTRAINT [FK_WaitingList_Employees] FOREIGN KEY ([EmployeeId])
        REFERENCES [dbo].[Employees] ([EmployeeId])
);

CREATE INDEX [IX_WaitingList_SessionId] ON [dbo].[WaitingList] ([SessionId]);
CREATE INDEX [IX_WaitingList_EmployeeId] ON [dbo].[WaitingList] ([EmployeeId]);
```

- [ ] **6c.** Add repository methods

Create a new file: `src/App_Code/DataAccess/WaitingListRepository.vb`

Methods:
- `GetWaitingListBySession(sessionId As Integer) As List(Of WaitingListEntry)`
- `GetWaitingListByEmployee(employeeId As Integer) As List(Of WaitingListEntry)`
- `AddToWaitingList(entry As WaitingListEntry) As Integer`
- `UpdateWaitingListEntry(entry As WaitingListEntry) As Boolean`
- `RemoveFromWaitingList(entryId As Integer) As Boolean`
- `GetNextInLine(sessionId As Integer) As WaitingListEntry`
- `GetAllActiveWaitingListEntries() As List(Of WaitingListEntry)` -- cross-session overview
- `ReorderWaitingList(sessionId As Integer, entryIds As List(Of Integer)) As Boolean`

- [ ] **6d.** Add business logic

In `src/App_Code/BusinessLogic/TrainingManager.vb`, add:
- `AddToWaitingList(sessionId As Integer, employeeId As Integer, createdBy As String) As Integer`
  - Validate session exists, is full, and employee isn't already enrolled or on the list
  - Auto-calculate position as max(position) + 1 for that session
- `RemoveFromWaitingList(entryId As Integer) As Boolean`
- `PromoteFromWaitingList(sessionId As Integer, promotedBy As String) As Boolean`
  - Get next in line, change status to "Offered", set response deadline (e.g. 3 business days)
  - This is the manual promotion the transcript describes Sarah doing
- `GetSessionWaitingList(sessionId As Integer) As List(Of WaitingListEntry)`
- `GetAllWaitingListOverview() As List(Of WaitingListEntry)`
  - Cross-session view that the transcript says Sarah needs (line 1735-1738)
- `GetWaitingListCountBySession(sessionId As Integer) As Integer`

### Verification
- [ ] The model supports the manual workflow described in the transcript (Sarah checks list, contacts person, promotes them)
- [ ] The cross-session overview addresses the specific gap mentioned at transcript line 1735-1738
- [ ] No automatic promotion -- the transcript describes this as a manual process in the current system

---

## Step 7: Add E-Learning Batch Sync Stub

### Problem
The transcript (lines 1222-1249, 2569-2608) describes a third-party e-learning platform
with an overnight batch sync process that frequently fails. There is zero integration
code in the codebase.

### What to Do

- [ ] **7a.** Create `ELearningSyncRecord.vb` model

Create file: `src/App_Code/Models/ELearningSyncRecord.vb`

Fields:

| Field | Type | Description |
|-------|------|-------------|
| SyncId | Integer | Primary key |
| SyncDate | DateTime | When the sync ran |
| SyncStatus | String | "Success", "Partial", "Failed" |
| TotalRecords | Integer | Records in the batch file |
| ProcessedRecords | Integer | Successfully processed |
| FailedRecords | Integer | Failed to process |
| ErrorDetails | String | Error messages |
| BatchFilePath | String | Path to the batch file processed |
| StartTime | DateTime | When sync started |
| EndTime | DateTime (nullable) | When sync finished |

- [ ] **7b.** Create `ELearningSyncService.vb`

Create file: `src/App_Code/BusinessLogic/ELearningSyncService.vb`

This represents the overnight batch process. Methods:

- `ProcessBatchFile(filePath As String, processedBy As String) As ELearningSyncRecord`
  - Parse the batch file from the e-learning platform
  - For each completion record: find employee by EmployeeNumber, find course by CourseCode
  - Create or update `TrainingRecord` via `TrainingRepository`
  - Log successes and failures
  - Return a sync record with statistics

- `GetSyncHistory(fromDate As DateTime, toDate As DateTime) As List(Of ELearningSyncRecord)`
  - View past sync runs and their success/failure status

- `GetLastSyncStatus() As ELearningSyncRecord`
  - Quick check: did last night's sync succeed?

- `GetUnmatchedRecords(syncId As Integer) As List(Of String)`
  - Records from batch file that couldn't be matched to employees/courses
  - Addresses the reliability issues described in transcript

- [ ] **7c.** Add sync log table to schema

In `src/TrainTrack_Schema.sql`:

```sql
CREATE TABLE [dbo].[ELearningSyncLog] (
    [SyncId]           INT IDENTITY(1,1) PRIMARY KEY,
    [SyncDate]         DATETIME NOT NULL DEFAULT GETDATE(),
    [SyncStatus]       NVARCHAR(50) NOT NULL,
    [TotalRecords]     INT NOT NULL DEFAULT 0,
    [ProcessedRecords] INT NOT NULL DEFAULT 0,
    [FailedRecords]    INT NOT NULL DEFAULT 0,
    [ErrorDetails]     NVARCHAR(MAX) NULL,
    [BatchFilePath]    NVARCHAR(500) NULL,
    [StartTime]        DATETIME NOT NULL,
    [EndTime]          DATETIME NULL
);
```

- [ ] **7d.** Add to system settings HTML

In `html/12-system-settings.html`, in the System Integration section, add an
"E-Learning Platform" subsection showing:
- Platform name (read-only)
- Last sync date and status
- Batch file location path
- "Run Manual Sync" button

This reflects the system's actual integration as described in the transcript.

### Verification
- [ ] The batch process model matches the transcript's description of overnight sync
- [ ] The error tracking addresses the frequent failures mentioned at transcript line 2587
- [ ] This is intentionally a basic batch implementation (not real-time API) because the transcript describes the current system as batch-based

---

## Step 8: Add Training Records Page Reference to Transcript

### Problem
HTML page `10-training-records.html` shows a dedicated training records management page
with bulk operations, detailed record panels, and statistics. The transcript doesn't
specifically walk through or mention this page during the demo.

### What to Do

- [ ] **8a.** Add a section to the transcript

In `transcripts/traintrack_demo.txt`, find an appropriate location after the course
management discussion (around the section where training completion and certificates are
discussed) and insert a passage where the demonstrator shows the Training Records page.

The added content should describe:
- Navigating to the Training Records section from the main menu
- The ability to search/filter training completion records by employee, course, status,
  and date range
- The record detail panel showing assessment results and certification info
- The bulk operations for generating certificates and sending completion emails
- The statistics showing completion rates and average scores
- Mention that there are 1,456 records in the system (matching the HTML)

Keep the writing style consistent with the rest of the transcript (conversational,
demo-walkthrough tone with questions from other attendees).

- [ ] **8b.** Cross-reference from other transcript sections

Where the transcript discusses certificate management and training completions, add brief
references to "the Training Records page" to establish it as part of the system's navigation.

### Verification
- [ ] The transcript now mentions all 12 HTML pages that exist as mockups
- [ ] The description of the Training Records page matches what HTML 10 shows

---

## Step 9: Final Consistency Pass

### Problem
After all previous steps, do a final sweep to ensure no new inconsistencies were introduced.

### What to Do

- [ ] **9a.** Verify navigation menus

Check that all 12 HTML files have consistent navigation menus listing the same pages.
Ensure the nav items match what the code and transcript describe as available sections.

- [ ] **9b.** Verify data counts

Ensure these numbers are consistent across all artefacts:
- 1,247 total employees
- 87.3% compliance rate
- 156 pending completions
- 23 overdue training items
- 45 active courses
- 142 completions this month
- 1,456 training records (HTML 10)

Check that no step accidentally changed these figures.

- [ ] **9c.** Verify new models are referenced in existing code

Ensure that:
- `CourseModule` is referenced in `Course.vb` (Modules property)
- `EmployeeNote` is referenced in `Employee.vb` (Notes property)
- `WaitingListEntry` is referenced in `TrainingSession.vb` (if a WaitingList property
  is added) or at minimum in `TrainingManager.vb`
- `ELearningSyncRecord` is referenced in system settings or admin code

- [ ] **9d.** Check for orphaned references

Search all files for references to removed role names ("HR Manager", "Dept Manager"
as roles), removed HTML sections (room utilization, resource requirements), and ensure
no dead links or references remain.

- [ ] **9e.** Verify schema completeness

Ensure `TrainTrack_Schema.sql` contains CREATE TABLE statements for all tables referenced
by all repository classes, including the new ones:
- CourseModules
- EmployeeNotes
- WaitingList
- ELearningSyncLog

---

## Summary

| | Step | Description | Files Created | Files Modified |
|---|------|-------------|--------------|----------------|
| [ ] | 1 | Course modules model | CourseModule.vb | Course.vb, CourseRepository.vb, CourseManager.vb, TrainTrack_Schema.sql |
| [ ] | 2 | CSV import | EmployeeImportResult.vb, EmployeeImportService.vb, CsvParser.vb | EmployeeRepository.vb |
| [ ] | 3 | Role name alignment | (none) | Default.aspx.vb, 08-user-management.html, traintrack_demo.txt |
| [ ] | 4 | Enhanced features alignment | (none) | Course.vb, TrainingSession.vb, Web.config, 09-schedule-training.html, 06-training-schedule.html, traintrack_demo.txt |
| [ ] | 5 | Employee notes model | EmployeeNote.vb | Employee.vb, EmployeeRepository.vb, EmployeeManager.vb, TrainTrack_Schema.sql |
| [ ] | 6 | Waiting list model | WaitingListEntry.vb, WaitingListRepository.vb | TrainingManager.vb, TrainTrack_Schema.sql |
| [ ] | 7 | E-learning sync stub | ELearningSyncRecord.vb, ELearningSyncService.vb | TrainTrack_Schema.sql, 12-system-settings.html |
| [ ] | 8 | Training records in transcript | (none) | traintrack_demo.txt |
| [ ] | 9 | Final consistency pass | (none) | Various (verification only) |

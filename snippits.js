// Define individual VBA snippets
const sub_template = `Sub MySubroutine()
    ' Description: Replace with a short description of what this code does
    On Error GoTo ErrorHandler
    
    ' Your code here
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description
End Sub
`;

const function_template = `Function MyFunction(ByVal inputValue As Variant) As Variant
    ' Description: Replace with a short description of the function's purpose
    On Error GoTo ErrorHandler
    
    ' Your code here
    
    MyFunction = inputValue ' Replace with your return value
    Exit Function
ErrorHandler:
    MsgBox "Error: " & Err.Description
    MyFunction = CVErr(xlErrValue) ' Return an error value
End Function

`;

const add_sheet = `Sub AddSheet(sheetName As String)
    Dim ws As Worksheet
    Set ws = Worksheets.Add
    ws.Name = sheetName
End Sub

`;

const delete_sheet = `Sub DeleteSheet(sheetName As String)
    Application.DisplayAlerts = False
    Worksheets(sheetName).Delete
    Application.DisplayAlerts = True
End Sub

`;

const loop_worksheets = `Sub LoopThroughSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Debug.Print ws.Name ' Replace with your desired operation
    Next ws
End Sub

`;

const set_cell_value = `Sub SetCellValue(sheetName As String, cellAddress As String, value As Variant)
    Worksheets(sheetName).Range(cellAddress).Value = value
End Sub

`;

const loop_range = `Sub LoopThroughRange()
    Dim cell As Range
    For Each cell In Range("A1:A10")
        Debug.Print cell.Value ' Replace with your desired operation
    Next cell
End Sub

`;

const find_and_replace = `Sub FindReplace(sheetName As String, findText As String, replaceText As String)
    Worksheets(sheetName).Cells.Replace What:=findText, Replacement:=replaceText, LookAt:=xlPart
End Sub

`;

const sort_range = `Sub SortRange(sheetName As String, sortRange As String, sortKey As String)
    With Worksheets(sheetName).Sort
        .SetRange Range(sortRange)
        .SortFields.Clear
        .SortFields.Add Key:=Range(sortKey), Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
End Sub
`

const ClassManager = `
' Private variables
Private pSchedules As Collection
Private pLastLoadTime As Date
Private pMinutesRequiredToReload As Long
Private pSheetName As String

' Initialize the ScheduleManager with default values
Private Sub Class_Initialize()
    Set pSchedules = New Collection
    pLastLoadTime = Now
    pMinutesRequiredToReload = 5
    pSheetName = "Sheet1"
End Sub

' Clean up resources on termination
Private Sub Class_Terminate()
    Set pSchedules = Nothing
End Sub

' --- Properties ---

' Get or Set the sheet name for data loading
Public Property Get SheetName() As String
    SheetName = pSheetName
End Property

Public Property Let SheetName(value As String)
    pSheetName = value
End Property

' Get or Set the reload interval in minutes
Public Property Get MinutesRequiredToReload() As Long
    MinutesRequiredToReload = pMinutesRequiredToReload
End Property

Public Property Let MinutesRequiredToReload(value As Long)
    pMinutesRequiredToReload = value
End Property

' Get the count of schedule entries
Public Property Get Count() As Long
    Count = pSchedules.Count
End Property

' Get the timestamp of the last load
Public Property Get LastLoaded() As Date
    LastLoaded = pLastLoadTime
End Property

' --- Methods ---

' Add a new schedule entry to the collection
Public Sub Add(ByVal entry As clsEntry)
    pSchedules.Add entry
End Sub

' Retrieve a schedule entry by index (1-based)
Public Function GetSchedule(ByVal index As Long) As clsEntry
    If index > 0 And index <= pSchedules.Count Then
        Set GetSchedule = pSchedules(index)
    Else
        Set GetSchedule = Nothing
    End If
End Function


' Find a schedule entry by job number and op number
Public Function FindEntry(ByVal JobNumber As String, ByVal OpNumber As String) As clsEntry
    Dim entry As clsEntry
    For Each entry In pSchedules
        If entry.JobNumber = JobNumber And entry.OpNumber = OpNumber Then
            Set FindEntry = entry
            Exit Function
        End If
    Next entry
    Set FindEntry = Nothing
End Function

' Clear all entries from the collection
Public Sub Clear()
    Set pSchedules = New Collection
    Me.UpdateLastLoadTime
End Sub

' Check if reloading is required based on the time interval
Public Function IsReloadRequired() As Boolean
    IsReloadRequired = (Now - pLastLoadTime) * 1440 >= pMinutesRequiredToReload
End Function

' Update the timestamp for last load time
Public Sub UpdateLastLoadTime()
    pLastLoadTime = Now
End Sub


' Load data from the specified sheet into the collection
Public Sub LoadData()
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim row As Range
    Dim entry As clsEntry
    
    ' Clear existing schedules before loading new data
    Me.Clear
    
    ' Set the worksheet and find the last row
    Set ws = ThisWorkbook.Sheets(Me.SheetName)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' Load each row of data into a new clsEntry and add to collection
    For Each row In ws.Range("A2:C" & lastRow).Rows
        Set entry = New clsEntry
        entry.AddFromRow row   ' Assumes clsEntry has AddFromRow method
        Me.Add entry
    Next row
    
    ' Update last load time after successful load
    Me.UpdateLastLoadTime
    Exit Sub

ErrorHandler:
    Debug.Print "Error in LoadData: " & Err.Description
End Sub

' List all entries (for debugging purposes)
Public Sub ListSchedules()
    Dim i As Long
    Dim schedule As clsEntry
    For i = 1 To pSchedules.Count
        Set schedule = pSchedules(i)
        Debug.Print "Entry " & i & ":"
        schedule.PrintAll   ' Assumes clsEntry has a PrintAll method
    Next i
End Sub
`;

const SetupMain = `
Option Explicit

Sub LoadScheduleData()
    ' Check if ScheduleManager is initialized
    If ScheduleManager Is Nothing Then
        ' Initialize the ScheduleManager
        Debug.Print "Adding Data to ScheduleManager cache"
        Set ScheduleManager = New clsEntryManager
        ' Load the data initially
        ScheduleManager.LoadData
    Else
        ' If ScheduleManager is already initialized, check if reload is required
        If ScheduleManager.IsReloadRequired Then
            Set ScheduleManager = New clsEntryManager
            
            Debug.Print ScheduleManager.MinutesRequiredToReload & " minutes have passed. Reloading data..."
            ScheduleManager.LoadData
        Else
            Debug.Print "Data retrieved from existing ScheduleManager cache"
        End If
    End If
    
    ' Call function to check entry count
    Call CheckEntryCount
End Sub


' Subroutine to display the count of entries in ScheduleManager
Sub CheckEntryCount()
    If ScheduleManager Is Nothing Then
        Debug.Print "ScheduleManager is not initialized."
    Else
        Debug.Print "Total entries in ScheduleManager:", ScheduleManager.Count
    End If
End Sub

' Subroutine to list all entries currently stored in ScheduleManager
Sub ListAllEntries()
    If ScheduleManager Is Nothing Then
        Debug.Print "ScheduleManager is not initialized."
    Else
        ScheduleManager.ListSchedules
    End If
End Sub
`;

const filter_range = `Sub FilterRange(sheetName As String, column As Integer, criteria As String)
    Worksheets(sheetName).UsedRange.AutoFilter Field:=column, Criteria1:=criteria
End Sub
`

const user_input_box = `Sub ShowInputBox()
    Dim userInput As String
    userInput = InputBox("Enter a value:", "Input")
    MsgBox "You entered: " & userInput
End Sub
`

const protect_workbook = `Sub ProtectWorkbook(password As String)
    ThisWorkbook.Protect Password:=password
End Sub
`

const unprotect_workbook = `Sub UnprotectWorkbook(password As String)
    ThisWorkbook.Unprotect Password:=password
End Sub
`

const error_handling_template = `Sub SimpleErrorHandling()
    On Error Resume Next
    ' Your code here
    If Err.Number <> 0 Then
        MsgBox "Error occurred: " & Err.Description
    End If
    On Error GoTo 0
End Sub
`

const pause = `Sub PauseCode(seconds As Double)
    Dim endTime As Double
    endTime = Timer + seconds
    Do While Timer < endTime
        DoEvents
    Loop
End Sub
`

const measure_execution_time = `Sub MeasureExecutionTime()
    Dim startTime As Double, endTime As Double
    startTime = Timer
    
    ' Your code here
    
    endTime = Timer
    MsgBox "Execution Time: " & (endTime - startTime) & " seconds"
End Sub
`

const open_workbook_with_dialoge = `Sub OpenWorkbook()
    Dim wb As Workbook
    Dim filePath As String

    On Error GoTo ErrorHandler

    ' Open File Dialog to Select Workbook
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            MsgBox "No file selected.", vbExclamation, "Action Cancelled"
            Exit Sub
        End If
    End With

    ' Open Selected Workbook
    Set wb = Workbooks.Open(filePath)
    MsgBox "Workbook successfully opened: " & wb.Name, vbInformation, "Success"
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Failed to Open Workbook"
End Sub
`

const load_data_from_seperate_sheet = `Sub TestLoopThroughSheetData()
    Dim filePath As String
    Dim sheetName As String

    filePath = "C:\Path\To\Your\File.xlsx" ' Replace with your file path
    sheetName = "Sheet1"                   ' Replace with your sheet name

    LoopThroughSheetData filePath, sheetName
End Sub


Sub LoopThroughSheetData(filePath As String, sheetName As String)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range

    On Error GoTo ErrorHandler

    ' Check if File Exists
    If Dir(filePath) = "" Then
        MsgBox "The specified file does not exist: " & filePath, vbCritical, "File Not Found"
        Exit Sub
    End If

    ' Open Workbook
    Set wb = Workbooks.Open(filePath)

    ' Check if Sheet Exists
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo ErrorHandler
    If ws Is Nothing Then
        MsgBox "The specified sheet does not exist in the workbook: " & sheetName, vbCritical, "Sheet Not Found"
        wb.Close SaveChanges:=False
        Exit Sub
    End If

    ' Loop Through Used Range
    Set rng = ws.UsedRange
    For Each cell In rng
        Debug.Print "Row: " & cell.Row & ", Column: " & cell.Column & ", Value: " & cell.Value
    Next cell

    MsgBox "Data looped successfully in sheet: " & sheetName, vbInformation, "Success"
    wb.Close SaveChanges:=False
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Operation Failed"
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub
`

// Group the snippets in an object
const snippets = {
    "Subroutine Template": sub_template,
    "Function Template": function_template,
    "Add Sheet": add_sheet,
    "Delete Sheet": delete_sheet,
    "Loop Worksheets": loop_worksheets,
    "Set Cell Value": set_cell_value,
    "Loop Range": loop_range,
    "Find & Replace": find_and_replace,
    "Sort Range": sort_range,
    "Filter Range": filter_range,
    "Class Manager": ClassManager,
    "Setup Main": SetupMain,
    "User Input": user_input_box,
    "Protect Workbook": protect_workbook,
    "Unprotect Workbook": unprotect_workbook,
    "Error Handling Template": error_handling_template,
    "Pause" : pause,
    "Measure Execution Time": measure_execution_time, 
    "Open Workbook W/ Dialoge": open_workbook_with_dialoge,
    "Load Data from Seperate File": load_data_from_seperate_sheet,
    
};

// Initialize the snippet manager
function initSnippetManager() {
    const dropdown = document.getElementById('snippet-dropdown');
    const codeDisplay = document.getElementById('vbaCode');
    const copyButton = document.getElementById('copy-button');

    // Populate dropdown menu with snippet titles
    Object.keys(snippets).forEach(title => {
        const option = document.createElement('option');
        option.value = title;
        option.textContent = title;
        dropdown.appendChild(option);
    });

    // Add event listener for dropdown selection
    dropdown.addEventListener('change', function() {
        const selectedSnippet = this.value;
        if (selectedSnippet) {
            codeDisplay.textContent = snippets[selectedSnippet];
            copyButton.disabled = false; // Enable the copy button
        } else {
            codeDisplay.textContent = '';
            copyButton.disabled = true; // Disable the copy button
        }
    });

    // Copy button functionality
    copyButton.addEventListener('click', () => {
        navigator.clipboard.writeText(codeDisplay.textContent)
            .then(() => {
                const copyButton = document.getElementById('copy-button');
                copyButton.textContent = "Copied";
                copyButton.style.backgroundColor = "green"; // Change background to green
    
                // Reset text and background after 5 seconds
                setTimeout(() => {
                    copyButton.textContent = "Copy to Clipboard"; // Reset button text
                    copyButton.style.backgroundColor = ""; // Reset background
                }, 5000);
            })
            .catch(err => {
                console.error("Failed to copy code: ", err); // Log error to console
            });
    });
}

initSnippetManager();

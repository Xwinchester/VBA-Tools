// Define individual VBA snippets
const snippet1 = `Sub HelloWorld()
    MsgBox "Hello, World!"
End Sub
`;

const snippet2 = `Function AddNumbers(a As Integer, b As Integer) As Integer
    AddNumbers = a + b
End Function
`;

const snippet3 = `Sub LoopThroughCells()
    Dim cell As Range
    For Each cell In Range("A1:A10")
        cell.Value = cell.Value * 2
    Next cell
End Sub
`;

const snippet4 = `Sub FormatCells()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    With ws.Range("A1:A10")
        .Font.Bold = True
        .Interior.Color = RGB(255, 255, 0) ' Yellow background
    End With
End Sub
`;

const snippet5 = `Sub AutoFitColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    ws.Columns.AutoFit
End Sub
`;

const snippet6 = `Sub CopyRange()
    Dim sourceRange As Range
    Dim targetRange As Range
    Set sourceRange = ThisWorkbook.Sheets("Sheet1").Range("A1:A10")
    Set targetRange = ThisWorkbook.Sheets("Sheet2").Range("A1")
    sourceRange.Copy Destination:=targetRange
End Sub
`;

const snippet7 = `Sub DeleteEmptyRows()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim i As Long
    For i = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row To 1 Step -1
        If Application.WorksheetFunction.CountA(ws.Rows(i)) = 0 Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub
`;

const snippet8 = `Sub CreateChart()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=100, Width:=375, Top:=50, Height:=225)
    With chartObj.Chart
        .SetSourceData Source:=ws.Range("A1:B10")
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Sample Chart"
    End With
End Sub
`;

const snippet9 = `Sub FindAndReplace()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    ws.Cells.Replace What:="OldValue", Replacement:="NewValue", LookAt:=xlPart
End Sub
`;

const snippet10 = `Function GetSumOfRange(rng As Range) As Double
    GetSumOfRange = Application.WorksheetFunction.Sum(rng)
End Function
`;

// Additional useful VBA snippets
const getLastRowIndex = `Function GetLastRowIndex(sheetName As String, col As String) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    GetLastRowIndex = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function
`;

const loopThroughRange = `Sub LoopThroughRange(startCell As String, endCell As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change to your sheet name
    Dim cell As Range

    For Each cell In ws.Range(startCell & ":" & endCell)
        Debug.Print cell.Address & " - Value: " & cell.Value
    Next cell
End Sub
`;

const loopThroughColumn = `Sub LoopThroughColumn(sheetName As String, col As String)
    Dim lastRow As Long
    lastRow = GetLastRowIndex(sheetName, col)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    Dim i As Long

    For i = 1 To lastRow
        Debug.Print ws.Cells(i, col).Address & " - Value: " & ws.Cells(i, col).Value
    Next i
End Sub
`;

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

// Group the snippets in an object
const snippets = {
    "Hello World": snippet1,
    "Add Numbers": snippet2,
    "Loop Through Cells": snippet3,
    "Format Cells": snippet4,
    "Auto Fit Columns": snippet5,
    "Copy Range": snippet6,
    "Delete Empty Rows": snippet7,
    "Create Chart": snippet8,
    "Find and Replace": snippet9,
    "Get Sum of Range": snippet10,
    "Get Last Row Index": getLastRowIndex,
    "Loop Through Range": loopThroughRange,
    "Loop Through Column": loopThroughColumn,
    "Class Manager": ClassManager,
    "Setup Main": SetupMain,
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

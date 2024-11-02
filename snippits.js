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

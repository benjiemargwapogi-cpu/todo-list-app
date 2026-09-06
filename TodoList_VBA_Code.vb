'====================================================================================================
'TO-DO LIST APPLICATION - VBA MACRO CODE FOR EXCEL
'====================================================================================================
'Features:
'1. Add new tasks
'2. Mark tasks as complete
'3. Delete tasks
'4. Local storage using Excel sheets
'5. Task priority levels
'6. Due dates
'7. Category/Tags
'8. Search functionality
'9. Task statistics
'====================================================================================================

Option Explicit

'====================================================================================================
'ADD NEW TASK MACRO
'====================================================================================================
Sub AddTask()
    Dim ws As Worksheet
    Dim taskName As String
    Dim priority As String
    Dim dueDate As String
    Dim category As String
    Dim lastRow As Long
    Dim taskID As Long
    
    Set ws = ThisWorkbook.Sheets("Tasks")
    
    'Get input from input cells
    taskName = ws.Range("B2").Value
    priority = ws.Range("B3").Value
    dueDate = ws.Range("B4").Value
    category = ws.Range("B5").Value
    
    'Validate input
    If taskName = "" Then
        MsgBox "Please enter a task name!", vbExclamation, "Missing Task Name"
        Exit Sub
    End If
    
    If priority = "" Then
        MsgBox "Please select a priority level!", vbExclamation, "Missing Priority"
        Exit Sub
    End If
    
    'Find last row in task list
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    'Generate Task ID
    If lastRow = 1 Then
        taskID = 1
    Else
        taskID = ws.Cells(lastRow, 1).Value + 1
    End If
    
    'Add new task
    With ws.Cells(lastRow + 1, 1)
        .Value = taskID
        .Font.Color = RGB(0, 0, 0)
    End With
    
    ws.Cells(lastRow + 1, 2).Value = taskName
    ws.Cells(lastRow + 1, 3).Value = priority
    
    'Color code by priority
    Select Case priority
        Case "High"
            ws.Cells(lastRow + 1, 3).Interior.Color = RGB(255, 0, 0) 'Red
            ws.Cells(lastRow + 1, 3).Font.Color = RGB(255, 255, 255)
        Case "Medium"
            ws.Cells(lastRow + 1, 3).Interior.Color = RGB(255, 192, 0) 'Orange
            ws.Cells(lastRow + 1, 3).Font.Color = RGB(255, 255, 255)
        Case "Low"
            ws.Cells(lastRow + 1, 3).Interior.Color = RGB(0, 176, 80) 'Green
            ws.Cells(lastRow + 1, 3).Font.Color = RGB(255, 255, 255)
    End Select
    
    ws.Cells(lastRow + 1, 4).Value = dueDate
    ws.Cells(lastRow + 1, 5).Value = category
    ws.Cells(lastRow + 1, 6).Value = "Pending"
    ws.Cells(lastRow + 1, 7).Value = Now() 'Created date
    
    'Clear input fields
    ws.Range("B2:B5").ClearContents
    
    MsgBox "Task added successfully! ID: " & taskID, vbInformation, "Task Added"
    
    'Update statistics
    UpdateTaskStatistics
End Sub

'====================================================================================================
'MARK TASK AS COMPLETE MACRO
'====================================================================================================
Sub MarkTaskComplete()
    Dim ws As Worksheet
    Dim taskID As Long
    Dim found As Boolean
    Dim i As Long
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Sheets("Tasks")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    'Get task ID to complete
    taskID = CLng(InputBox("Enter Task ID to mark as complete:", "Complete Task"))
    
    If taskID = 0 Then Exit Sub
    
    'Search for task
    found = False
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = taskID Then
            ws.Cells(i, 6).Value = "Completed"
            ws.Cells(i, 6).Interior.Color = RGB(192, 192, 192)
            ws.Cells(i, 6).Font.Color = RGB(0, 0, 0)
            
            'Strike through the task name
            ws.Cells(i, 2).Font.Strikethrough = True
            
            found = True
            MsgBox "Task " & taskID & " marked as complete!", vbInformation, "Task Completed"
            Exit For
        End If
    Next i
    
    If Not found Then
        MsgBox "Task ID " & taskID & " not found!", vbExclamation, "Task Not Found"
    End If
    
    UpdateTaskStatistics
End Sub

'====================================================================================================
'DELETE TASK MACRO
'====================================================================================================
Sub DeleteTask()
    Dim ws As Worksheet
    Dim taskID As Long
    Dim i As Long
    Dim lastRow As Long
    Dim response As VbMsgBoxResult
    
    Set ws = ThisWorkbook.Sheets("Tasks")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    'Get task ID to delete
    taskID = CLng(InputBox("Enter Task ID to delete:", "Delete Task"))
    
    If taskID = 0 Then Exit Sub
    
    'Search for task
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = taskID Then
            response = MsgBox("Are you sure you want to delete Task " & taskID & "?", vbYesNo, "Confirm Delete")
            
            If response = vbYes Then
                ws.Rows(i).Delete
                MsgBox "Task deleted successfully!", vbInformation, "Task Deleted"
            End If
            
            UpdateTaskStatistics
            Exit Sub
        End If
    Next i
    
    MsgBox "Task ID " & taskID & " not found!", vbExclamation, "Task Not Found"
End Sub

'====================================================================================================
'SEARCH TASKS MACRO
'====================================================================================================
Sub SearchTasks()
    Dim ws As Worksheet
    Dim searchTerm As String
    Dim i As Long
    Dim lastRow As Long
    Dim resultsWs As Worksheet
    Dim resultRow As Long
    
    Set ws = ThisWorkbook.Sheets("Tasks")
    searchTerm = InputBox("Enter search term:", "Search Tasks")
    
    If searchTerm = "" Then Exit Sub
    
    'Clear previous search results
    On Error Resume Next
    ThisWorkbook.Sheets("Search Results").Delete
    On Error GoTo 0
    
    'Create search results sheet
    Set resultsWs = ThisWorkbook.Sheets.Add
    resultsWs.Name = "Search Results"
    
    'Add headers
    resultsWs.Range("A1:G1").Value = Array("ID", "Task", "Priority", "Due Date", "Category", "Status", "Created")
    resultRow = 2
    
    'Search and display results
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        If InStr(1, ws.Cells(i, 2).Value, searchTerm, vbTextCompare) > 0 Or _
           InStr(1, ws.Cells(i, 5).Value, searchTerm, vbTextCompare) > 0 Then
            
            resultsWs.Cells(resultRow, 1).Value = ws.Cells(i, 1).Value
            resultsWs.Cells(resultRow, 2).Value = ws.Cells(i, 2).Value
            resultsWs.Cells(resultRow, 3).Value = ws.Cells(i, 3).Value
            resultsWs.Cells(resultRow, 4).Value = ws.Cells(i, 4).Value
            resultsWs.Cells(resultRow, 5).Value = ws.Cells(i, 5).Value
            resultsWs.Cells(resultRow, 6).Value = ws.Cells(i, 6).Value
            resultsWs.Cells(resultRow, 7).Value = ws.Cells(i, 7).Value
            
            resultRow = resultRow + 1
        End If
    Next i
    
    MsgBox "Search complete! Found " & (resultRow - 2) & " results.", vbInformation, "Search Results"
End Sub

'====================================================================================================
'CLEAR ALL TASKS MACRO
'====================================================================================================
Sub ClearAllTasks()
    Dim ws As Worksheet
    Dim response As VbMsgBoxResult
    
    Set ws = ThisWorkbook.Sheets("Tasks")
    
    response = MsgBox("Are you sure you want to clear all tasks? This cannot be undone!", vbYesNo, "Clear All Tasks")
    
    If response = vbYes Then
        ws.Range("A2:G1000").Delete
        MsgBox "All tasks cleared!", vbInformation, "Tasks Cleared"
        UpdateTaskStatistics
    End If
End Sub

'====================================================================================================
'UPDATE TASK STATISTICS MACRO
'====================================================================================================
Sub UpdateTaskStatistics()
    Dim ws As Worksheet
    Dim statsWs As Worksheet
    Dim lastRow As Long
    Dim totalTasks As Long
    Dim completedTasks As Long
    Dim pendingTasks As Long
    Dim highPriority As Long
    Dim i As Long
    Dim completionRate As Double
    
    Set ws = ThisWorkbook.Sheets("Tasks")
    Set statsWs = ThisWorkbook.Sheets("Dashboard")
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    'Initialize counters
    totalTasks = 0
    completedTasks = 0
    pendingTasks = 0
    highPriority = 0
    
    'Count statistics
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value <> "" Then
            totalTasks = totalTasks + 1
            
            If ws.Cells(i, 6).Value = "Completed" Then
                completedTasks = completedTasks + 1
            Else
                pendingTasks = pendingTasks + 1
            End If
            
            If ws.Cells(i, 3).Value = "High" Then
                highPriority = highPriority + 1
            End If
        End If
    Next i
    
    'Calculate completion rate
    If totalTasks > 0 Then
        completionRate = (completedTasks / totalTasks) * 100
    Else
        completionRate = 0
    End If
    
    'Update dashboard
    statsWs.Range("B2").Value = totalTasks
    statsWs.Range("B3").Value = completedTasks
    statsWs.Range("B4").Value = pendingTasks
    statsWs.Range("B5").Value = highPriority
    statsWs.Range("B6").Value = Format(completionRate, "0.00") & "%"
    statsWs.Range("B7").Value = Now()
End Sub

'====================================================================================================
'SORT TASKS BY PRIORITY MACRO
'====================================================================================================
Sub SortByPriority()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Sheets("Tasks")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    'Sort by priority (Column C)
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("C2:C" & lastRow), Order:=xlAscending
        .SetRange ws.Range("A1:G" & lastRow)
        .Header = xlYes
        .Apply
    End With
    
    MsgBox "Tasks sorted by priority!", vbInformation, "Sorted"
End Sub

'====================================================================================================
'SORT TASKS BY DUE DATE MACRO
'====================================================================================================
Sub SortByDueDate()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Sheets("Tasks")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    'Sort by due date (Column D)
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("D2:D" & lastRow), Order:=xlAscending
        .SetRange ws.Range("A1:G" & lastRow)
        .Header = xlYes
        .Apply
    End With
    
    MsgBox "Tasks sorted by due date!", vbInformation, "Sorted"
End Sub

'====================================================================================================
'EXPORT TASKS TO PDF MACRO
'====================================================================================================
Sub ExportToPDF()
    Dim ws As Worksheet
    Dim filePath As String
    
    Set ws = ThisWorkbook.Sheets("Tasks")
    
    filePath = Application.GetSaveAsFilename(FileFilter:="PDF Files (*.pdf),*.pdf")
    
    If filePath <> "False" Then
        ws.ExportAsFixedFormat xlTypePDF, filePath
        MsgBox "Tasks exported to PDF successfully!", vbInformation, "Export Complete"
    End If
End Sub

'====================================================================================================
'INITIALIZE APPLICATION MACRO
'====================================================================================================
Sub InitializeApp()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim dashWs As Worksheet
    
    Set wb = ThisWorkbook
    
    'Create Tasks sheet if it doesn't exist
    On Error Resume Next
    Set ws = wb.Sheets("Tasks")
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add
        ws.Name = "Tasks"
        
        'Add headers
        ws.Range("A1:G1").Value = Array("ID", "Task Name", "Priority", "Due Date", "Category", "Status", "Created Date")
        ws.Range("A1:G1").Font.Bold = True
        ws.Range("A1:G1").Interior.Color = RGB(0, 102, 204)
        ws.Range("A1:G1").Font.Color = RGB(255, 255, 255)
        
        'Set column widths
        ws.Columns("A").ColumnWidth = 5
        ws.Columns("B").ColumnWidth = 25
        ws.Columns("C").ColumnWidth = 10
        ws.Columns("D").ColumnWidth = 12
        ws.Columns("E").ColumnWidth = 15
        ws.Columns("F").ColumnWidth = 12
        ws.Columns("G").ColumnWidth = 18
    End If
    On Error GoTo 0
    
    'Create Dashboard sheet if it doesn't exist
    On Error Resume Next
    Set dashWs = wb.Sheets("Dashboard")
    If dashWs Is Nothing Then
        Set dashWs = wb.Sheets.Add
        dashWs.Name = "Dashboard"
        dashWs.Move Before:=ws
        
        'Add dashboard headers
        dashWs.Range("A1").Value = "TO-DO LIST DASHBOARD"
        dashWs.Range("A1").Font.Bold = True
        dashWs.Range("A1").Font.Size = 16
        
        dashWs.Range("A3").Value = "Total Tasks:"
        dashWs.Range("A4").Value = "Completed Tasks:"
        dashWs.Range("A5").Value = "Pending Tasks:"
        dashWs.Range("A6").Value = "High Priority Tasks:"
        dashWs.Range("A7").Value = "Completion Rate:"
        dashWs.Range("A8").Value = "Last Updated:"
        
        'Format labels as bold
        dashWs.Range("A3:A8").Font.Bold = True
        
        'Set column widths
        dashWs.Columns("A").ColumnWidth = 25
        dashWs.Columns("B").ColumnWidth = 15
    End If
    On Error GoTo 0
    
    MsgBox "Application initialized successfully!", vbInformation, "Ready to Use"
End Sub

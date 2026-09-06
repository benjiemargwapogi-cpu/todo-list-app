# To-Do List Application - Excel VBA Setup Guide

## 📋 Overview

A fully functional **To-Do List Application** built entirely in Microsoft Excel with VBA macros and local storage. Manage your tasks, track priorities, set due dates, and monitor productivity—all without leaving Excel!

---

## ✨ Features

### Core Functionality
✅ **Add Tasks** - Create new tasks with details (name, priority, due date, category)
✅ **Mark Complete** - Track task completion status
✅ **Delete Tasks** - Remove completed or unwanted tasks
✅ **Local Storage** - All data saved in Excel sheets (no internet required)
✅ **Search Tasks** - Find tasks by name or category
✅ **Sort Tasks** - Organize by priority or due date
✅ **Task Statistics** - Dashboard with completion metrics
✅ **Priority Levels** - High, Medium, Low with color coding
✅ **Categories/Tags** - Organize tasks by type
✅ **Export to PDF** - Share task lists as PDF documents

### Dashboard
- Total Tasks Count
- Completed Tasks Count
- Pending Tasks Count
- High Priority Tasks Count
- Completion Rate Percentage
- Last Updated Timestamp

---

## 🚀 Quick Start

### Step 1: Create the Excel File

1. Open **Microsoft Excel** (2010 or later)
2. Create a new blank workbook
3. Save it as `TodoList.xlsm` (Excel Macro-Enabled)

### Step 2: Set Up Sheets

#### Sheet 1: Dashboard
Create a sheet named **"Dashboard"** with the following layout:

```
A1: TO-DO LIST DASHBOARD (Bold, Size 16)

A3: Total Tasks:           | B3: [statistics will auto-populate]
A4: Completed Tasks:       | B4: [statistics will auto-populate]
A5: Pending Tasks:         | B5: [statistics will auto-populate]
A6: High Priority Tasks:   | B6: [statistics will auto-populate]
A7: Completion Rate:       | B7: [statistics will auto-populate]
A8: Last Updated:          | B8: [statistics will auto-populate]
```

#### Sheet 2: Tasks
Create a sheet named **"Tasks"** with headers in row 1:

```
A1: ID
B1: Task Name
C1: Priority
D1: Due Date
E1: Category
F1: Status
G1: Created Date
```

**Format Headers:**
- Make Bold
- Background Color: Dark Blue (RGB: 0, 102, 204)
- Text Color: White
- Set Column Widths: A=5, B=25, C=10, D=12, E=15, F=12, G=18

#### Sheet 3: Input Form (Optional but Recommended)
Create an input form sheet with these labels:

```
A2: Task Name:    | B2: [Input field]
A3: Priority:     | B3: [Input field - Dropdown: High/Medium/Low]
A4: Due Date:     | B4: [Input field - Date format]
A5: Category:     | B5: [Input field]
```

---

## 📝 Step 3: Add VBA Macros

1. Open **Visual Basic Editor** (Alt+F11)
2. Right-click on your workbook name → **Insert Module**
3. Copy the entire code from `TodoList_VBA_Code.vb`
4. Paste it into the module
5. Save the file (Ctrl+S)

---

## 🎯 Step 4: Create Control Buttons

### Button 1: Initialize Application
**Purpose**: Set up sheets and structure

1. Go to **Developer Tab** → **Insert** → **Button (Form Control)**
2. Draw a button
3. Right-click → **Assign Macro** → Select `InitializeApp`
4. Type: **"⚙️ Initialize App"**
5. Click **Run** first time to set up structure

### Button 2: Add Task
1. Draw a button
2. Assign Macro: `AddTask`
3. Label: **"➕ Add Task"**

### Button 3: Mark Complete
1. Draw a button
2. Assign Macro: `MarkTaskComplete`
3. Label: **"✅ Mark Complete"**

### Button 4: Delete Task
1. Draw a button
2. Assign Macro: `DeleteTask`
3. Label: **"🗑️ Delete Task"**

### Button 5: Search Tasks
1. Draw a button
2. Assign Macro: `SearchTasks`
3. Label: **"🔍 Search"**

### Button 6: Sort by Priority
1. Draw a button
2. Assign Macro: `SortByPriority`
3. Label: **"📊 Sort by Priority"**

### Button 7: Sort by Due Date
1. Draw a button
2. Assign Macro: `SortByDueDate`
3. Label: **"📅 Sort by Date"**

### Button 8: Export to PDF
1. Draw a button
2. Assign Macro: `ExportToPDF`
3. Label: **"📄 Export PDF"**

### Button 9: Clear All Tasks
1. Draw a button
2. Assign Macro: `ClearAllTasks`
3. Label: **"🧹 Clear All"**

---

## 💾 How to Use

### Adding a Task
1. Go to the **Input Form** sheet (or use Input cells in Dashboard)
2. Enter the **Task Name** in B2
3. Select **Priority** from dropdown (High/Medium/Low)
4. Enter **Due Date** in B4 (format: MM/DD/YYYY)
5. Enter **Category** in B5 (e.g., Work, Personal, Health)
6. Click **"➕ Add Task"** button
7. Task automatically gets an ID and appears in Tasks sheet

### Marking a Task as Complete
1. Click **"✅ Mark Complete"** button
2. Enter the **Task ID** when prompted
3. Task status changes to "Completed"
4. Task name shows strikethrough
5. Statistics update automatically

### Deleting a Task
1. Click **"🗑️ Delete Task"** button
2. Enter the **Task ID** when prompted
3. Confirm deletion
4. Task is permanently removed
5. Statistics update automatically

### Searching for Tasks
1. Click **"🔍 Search"** button
2. Enter search term (task name or category)
3. Results display in new "Search Results" sheet
4. Search can find partial matches

### Viewing Statistics
1. Go to **Dashboard** sheet
2. View real-time statistics:
   - Total number of tasks
   - Completed vs. Pending
   - High priority count
   - Completion percentage
   - Last update time

### Sorting Tasks
- **By Priority**: Click "📊 Sort by Priority" - High → Medium → Low
- **By Due Date**: Click "📅 Sort by Date" - Earliest → Latest

### Exporting to PDF
1. Click **"📄 Export PDF"** button
2. Choose save location
3. PDF file created with current task list

---

## 🎨 Priority Color Coding

| Priority | Color | Meaning |
|----------|-------|----------|
| High | 🔴 Red | Urgent - Do First |
| Medium | 🟠 Orange | Important - Do Soon |
| Low | 🟢 Green | Can Wait - Low Priority |

---

## 📊 Task Status Options

- **Pending** - Task not yet completed (Yellow highlight)
- **Completed** - Task finished (Gray highlight)

---

## 🔒 Local Storage

All data is stored in Excel sheets:
- **Tasks Sheet**: Contains all task data
- **Dashboard Sheet**: Contains calculated statistics
- **Search Results Sheet**: Auto-generated when searching

**No internet required** - Everything is saved locally in your Excel file!

---

## 💡 Tips & Tricks

### Backup Your Data
- Save your Excel file regularly (Ctrl+S)
- Create copies before clearing large amounts of data
- Export to PDF for permanent records

### Organizing Tasks
- Use consistent category names (Work, Personal, Health, Finance, etc.)
- Set realistic due dates
- Review dashboard weekly for productivity insights

### Keyboard Shortcuts
- Alt+F11: Open VBA Editor
- Ctrl+S: Save file
- Ctrl+Z: Undo last action

---

## 🐛 Troubleshooting

### Macros Not Running?
1. Ensure file is saved as `.xlsm`
2. Check Excel Trust Center settings
3. Enable macros when prompted
4. Try Alt+F11 → Run InitializeApp first

### Data Not Updating?
1. Make sure you're using the correct sheet names ("Tasks", "Dashboard")
2. Verify Task IDs are unique
3. Check that input cells have values

### Can't Find a Task?
1. Use the Search function
2. Verify Task ID or task name spelling
3. Check if task was deleted

---

## 📋 File Structure

```
TodoList.xlsm
├── Dashboard (Sheet 1)
│   ├── Statistics
│   └── Quick Links
├── Tasks (Sheet 2)
│   ├── Task ID (Column A)
│   ├── Task Name (Column B)
│   ├── Priority (Column C)
│   ├── Due Date (Column D)
│   ├── Category (Column E)
│   ├── Status (Column F)
│   └── Created Date (Column G)
└── VBA Modules
    ├── AddTask
    ├── MarkTaskComplete
    ├── DeleteTask
    ├── SearchTasks
    ├── SortByPriority
    ├── SortByDueDate
    ├── ExportToPDF
    ├── ClearAllTasks
    ├── UpdateTaskStatistics
    └── InitializeApp
```

---

## 📞 Support

For issues or questions:
1. Review this guide completely
2. Check VBA code for comments
3. Verify all sheet names match exactly
4. Ensure macros are enabled

---

## 📈 Future Enhancements

Consider adding:
- Recurring tasks
- Task reminders
- Subtasks
- Time tracking
- Collaborative sharing
- Integration with Outlook
- Mobile sync

---

**Version**: 1.0  
**Last Updated**: September 6, 2026  
**Author**: Copilot - GitHub Assistant  
**License**: Open Source

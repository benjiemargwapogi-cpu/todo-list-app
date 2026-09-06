# Frequently Asked Questions (FAQ)

## General Questions

### Q: Can I use this on Mac and Windows?
**A:** Yes! This Excel VBA application works on both Windows and macOS (Excel 2010+). Macros are supported on both platforms.

### Q: Do I need to be online?
**A:** No! All data is stored locally in your Excel file. Internet is not required.

### Q: How many tasks can I add?
**A:** Excel supports over 1 million rows, but for practical use, 10,000+ tasks work fine. Performance may slow with extremely large datasets.

### Q: Can I share this with my team?
**A:** Yes! Simply send the .xlsm file to team members. They must enable macros to use it.

### Q: How do I backup my data?
**A:** Save the Excel file to cloud storage (OneDrive, Google Drive, Dropbox) or create copies manually. The file itself is your database.

---

## Technical Questions

### Q: Why must I save as .xlsm?
**A:** .xlsm is Excel Macro-Enabled format. Macros won't work in .xlsx (standard Excel) files.

### Q: Where is my data stored?
**A:** All data is in the Excel file itself, in the "Tasks" and "Dashboard" sheets. No external database or cloud storage used.

### Q: Can I edit the macro code?
**A:** Yes, if you know VBA! Open Visual Basic Editor (Alt+F11) and modify the code. Always backup first.

### Q: Why aren't macros running?
**A:**
- File must be .xlsm (not .xlsx)
- Macros must be enabled in Trust Center
- Try: File → Options → Trust Center → Enable all macros

### Q: Can I add more columns?
**A:** Yes, add them after column G. Don't modify existing columns A-G as macros reference them.

---

## Feature Questions

### Q: Can I mark a task as "In Progress"?
**A:** Currently supports Pending and Completed. Modify the macro code to add more statuses.

### Q: Can I add subtasks?
**A:** Not built-in. You can create subtask naming convention (e.g., "1.1", "1.2") and filter manually.

### Q: Can tasks repeat automatically?
**A:** Not built-in. You can manually re-add recurring tasks or modify the macro code.

### Q: Can I set reminders/notifications?
**A:** Not built-in. You could integrate with Outlook reminders (advanced VBA).

### Q: Can I filter by category?
**A:** Use Excel's built-in filter (Data → AutoFilter). Search function also filters by category.

### Q: Can I change priority levels (e.g., Urgent, Important, Low)?
**A:** Yes! Modify the macro code and the dropdown in the Priority column. Search for "High", "Medium", "Low" and replace.

---

## Usage Questions

### Q: How do I add a task?
**A:**
1. Enter task details in input cells (B2-B5)
2. Click "➕ Add Task" button
3. Task automatically appears in Tasks sheet with auto-generated ID

### Q: How do I mark a task complete?
**A:**
1. Click "✅ Mark Complete" button
2. Enter Task ID when prompted
3. Task status changes to "Completed" with strikethrough

### Q: How do I delete a task?
**A:**
1. Click "🗑️ Delete Task" button
2. Enter Task ID
3. Confirm deletion
4. Task permanently removed

### Q: How do I search for tasks?
**A:**
1. Click "🔍 Search" button
2. Enter search term (task name or category)
3. Results display in new sheet
4. Supports partial matches

### Q: How do I export to PDF?
**A:**
1. Click "📄 Export PDF" button
2. Choose save location
3. PDF created with current task list

### Q: How do I clear all tasks?
**A:**
1. Click "🧹 Clear All" button
2. Confirm you want to delete all
3. All tasks removed (backup first!)

---

## Dashboard Questions

### Q: Why isn't the dashboard updating?
**A:** Dashboard updates after each action. If not, click any macro button to refresh. Try: Dashboard → Recalculate (F9).

### Q: What does Completion Rate mean?
**A:** (Completed Tasks / Total Tasks) × 100 = Completion Rate %. Shows productivity percentage.

### Q: Can I see which user made changes?
**A:** Not built-in. You can modify the macro to add user tracking.

### Q: Can I export dashboard statistics?
**A:** Yes! Click "📄 Export PDF" to save dashboard and all tasks as PDF.

---

## Customization Questions

### Q: How do I add custom categories?
**A:** 
1. Go to Input Form sheet
2. In cell B5, you can type any category
3. Use consistent names for filtering
4. No code changes needed

### Q: How do I change colors?
**A:** 
1. Open Visual Basic Editor (Alt+F11)
2. Find color codes: RGB(255, 0, 0) = Red
3. Modify RGB values as desired
4. Test and save

### Q: How do I add a new button?
**A:**
1. Developer tab → Insert → Button
2. Draw button
3. Assign existing macro or create new one
4. Label the button

### Q: Can I add a custom macro?
**A:** Yes! If you know VBA, open Visual Basic Editor and create new Sub. Assign to buttons.

---

## Performance Questions

### Q: Why is it slow with many tasks?
**A:** 
- Sort operations slow with 10,000+ tasks
- Macros run slower on older computers
- Solution: Archive old tasks to separate sheet

### Q: How do I speed up macros?
**A:**
1. Reduce total tasks (archive old ones)
2. Upgrade to faster computer
3. Close other Excel files
4. Disable Excel add-ins

### Q: Can I archive completed tasks?
**A:** Create an "Archive" sheet and manually move completed tasks there. Keep Tasks sheet lean.

---

## Troubleshooting Questions

### Q: Error: "Macro not found"
**A:**
- Verify macro name spelling in button assignment
- Check macro exists in Visual Basic Editor
- Reload file and try again

### Q: Error: "Cannot find task ID"
**A:**
- Verify task ID is correct
- Check Tasks sheet directly
- Use Search function to find task
- Task may have been deleted

### Q: Error: "This operation requires Microsoft Excel to be in edit mode"
**A:**
- Click elsewhere to deselect cells
- Click button again
- Restart Excel

### Q: Formulas showing as text, not values
**A:**
- Select cells
- Right-click → Format Cells → Number
- Choose appropriate format

---

## Integration Questions

### Q: Can I import from Outlook?
**A:** Not built-in. You can copy/paste email subjects as tasks.

### Q: Can I sync with other apps?
**A:** Not built-in. Advanced users could add web API integration via VBA.

### Q: Can I use this with OneDrive/Google Drive?
**A:** Yes! Save .xlsm file to cloud storage and open from there.

### Q: Can I collaborate with team members?
**A:** Not in real-time (would need VBA multi-user code). You can take turns editing and email file.

---

## Training Questions

### Q: How do I learn VBA?
**A:**
- Microsoft Official Documentation
- Excel VBA Tutorials Online
- "Excel VBA for Dummies" book
- Udemy/Coursera courses

### Q: Can I hire someone to customize this?
**A:** Yes! Any VBA programmer can modify it. Provide requirements clearly.

### Q: How do I train my team to use this?
**A:**
1. Share QUICK_REFERENCE.md
2. Do live demo
3. Have them create a test task
4. Answer questions
5. Provide documentation link

---

## Support Questions

### Q: Where do I report bugs?
**A:** Open issue on GitHub with:
- Excel version
- What went wrong
- Steps to reproduce
- Error message (if any)

### Q: Where do I request features?
**A:** GitHub Issues → New Issue → Feature Request. Describe the feature and why you need it.

### Q: Is there technical support?
**A:** Community support via GitHub. Professional support available for hire.

### Q: Can I modify this for my business?
**A:** Yes! It's open source. Modify as needed for your use case.

---

## Did we miss something?

**Have more questions?**
- Check `TROUBLESHOOTING.md` for common issues
- Review `TODO_LIST_SETUP_GUIDE.md` for detailed help
- Search GitHub issues for similar questions
- Create a new GitHub issue with your question

---

**Last Updated:** September 6, 2026  
**Version:** 1.0

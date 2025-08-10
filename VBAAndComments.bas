'============================
' EXCEL CHANGE TRACKING SYSTEM - FULLY COMMENTED VERSION
' This VBA code creates an automated change tracking system for Excel spreadsheets
' It monitors changes to budget amounts and creates audit logs and variance reports
'============================

'============================
' MODULE CODE (Standard Module - Insert > Module in VBA Editor)
' This section contains the main functions and configuration
'============================
Option Explicit  ' Forces declaration of all variables - good programming practice

'== CONFIGURATION SECTION ==
' These constants define which sheets and columns the system will monitor
' Change these values to match your specific spreadsheet structure

Public Const DATA_SHEET As String = "Sheet1"          ' Name of sheet containing your main data
Public Const LOG_SHEET As String = "Change_Log"       ' Name of sheet where changes will be logged
Public Const REPORT_SHEET As String = "Variance_Report" ' Name of sheet for variance reports
Public Const COL_EVENT As Long = 1    ' Column A: Event/Sponsorship/Organization Name
Public Const COL_NEW As Long = 2      ' Column B: New (Yes/No) indicator
Public Const COL_PLANNED As Long = 7  ' Column G: Planned Amount (monitored for changes)
Public Const COL_APPROVED As Long = 8 ' Column H: Approved Amount (monitored for changes)
Public Const COL_VARIANCE As Long = 9 ' Column I: Amount Variance (calculated automatically)

'============================
' UTILITY FUNCTIONS
'============================

' Function: EnsureSheet
' Purpose: Creates a worksheet if it doesn't exist, returns reference to the sheet
' Parameters: sheetName - the name of the sheet to create/find
' Returns: Worksheet object
Public Function EnsureSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next  ' Ignore errors temporarily
    ' Try to find existing sheet with the specified name
    Set EnsureSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0      ' Resume normal error handling
    
    ' If sheet doesn't exist, create it
    If EnsureSheet Is Nothing Then
        ' Add new sheet at the end of existing sheets
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ' Set the name of the new sheet
        EnsureSheet.Name = sheetName
    End If
End Function

' Function: ParseMoney
' Purpose: Safely converts currency text (like "$15,000", "TBD", "N/A") to a number
' Parameters: v - the value to convert (can be text or number)
' Returns: Double value (0 if conversion fails)
Public Function ParseMoney(ByVal v As Variant) As Double
    Dim s As String
    ' Convert input to string and remove spaces
    s = Trim$(CStr(v))
    
    ' Exit if empty string
    If s = "" Then Exit Function
    
    ' Exit if value is "TBD" or "N/A" (return 0)
    If UCase$(s) = "TBD" Or UCase$(s) = "N/A" Then Exit Function
    
    ' Remove dollar signs and commas
    s = Replace$(s, "$", "")
    s = Replace$(s, ",", "")
    
    ' Convert to number if possible
    If IsNumeric(s) Then ParseMoney = CDbl(s)
End Function

' Subroutine: AppendChangeLog
' Purpose: Adds a new entry to the change log with details about what was modified
' Parameters: 
'   - rowNum: which row was changed
'   - colNum: which column was changed
'   - oldVal: the previous value
'   - newVal: the new value
'   - plannedVal: current planned amount for this row
'   - approvedVal: current approved amount for this row
Public Sub AppendChangeLog(ByVal rowNum As Long, ByVal colNum As Long, _
                           ByVal oldVal As Variant, ByVal newVal As Variant, _
                           ByVal plannedVal As Variant, ByVal approvedVal As Variant)

    Dim wsLog As Worksheet, nextRow As Long, wsData As Worksheet
    
    ' Get reference to log sheet (create if doesn't exist)
    Set wsLog = EnsureSheet(LOG_SHEET)
    ' Get reference to main data sheet
    Set wsData = ThisWorkbook.Worksheets(DATA_SHEET)

    ' If log sheet is empty, create headers
    If wsLog.Range("A1").Value = "" Then
        ' Create header row with column descriptions
        wsLog.Range("A1:I1").Value = Array("When", "User", "Row", "Changed Column", "Event/Org", "Old Value", "New Value", "Planned (row)", "Approved (row)")
        ' Make headers bold for better readability
        wsLog.Rows(1).Font.Bold = True
        ' Auto-fit columns to content
        wsLog.Columns.AutoFit
    End If

    ' Find the next empty row in the log
    nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Fill in the log entry details
    wsLog.Cells(nextRow, 1).Value = Now                    ' Current date/time
    wsLog.Cells(nextRow, 2).Value = Environ$("Username")   ' Windows username
    wsLog.Cells(nextRow, 3).Value = rowNum                 ' Row that was changed
    ' Determine which column was changed (Planned or Approved)
    wsLog.Cells(nextRow, 4).Value = IIf(colNum = COL_PLANNED, "Planned (G)", "Approved (H)")
    ' Get the event/organization name from the changed row
    wsLog.Cells(nextRow, 5).Value = wsData.Cells(rowNum, COL_EVENT).Value
    wsLog.Cells(nextRow, 6).Value = oldVal                 ' Previous value
    wsLog.Cells(nextRow, 7).Value = newVal                 ' New value
    wsLog.Cells(nextRow, 8).Value = plannedVal             ' Current planned amount
    wsLog.Cells(nextRow, 9).Value = approvedVal            ' Current approved amount
End Sub

' Subroutine: RebuildVarianceReport
' Purpose: Creates a summary report of all rows where Planned amount differs from Approved amount
' This helps identify budget discrepancies at a glance
Public Sub RebuildVarianceReport()
    Dim wsData As Worksheet, wsRep As Worksheet, r As Long, lastRow As Long, outRow As Long
    
    ' Get references to data sheet and report sheet
    Set wsData = ThisWorkbook.Worksheets(DATA_SHEET)
    Set wsRep = EnsureSheet(REPORT_SHEET)

    ' Clear existing report content
    wsRep.Cells.Clear
    
    ' Create report headers
    wsRep.Range("A1:F1").Value = Array("Event/Org", "New?", "Planned", "Approved", "Variance", "Row #")
    wsRep.Rows(1).Font.Bold = True

    ' Find the last row with data in the main sheet
    lastRow = wsData.Cells(wsData.Rows.Count, COL_EVENT).End(xlUp).Row
    outRow = 2  ' Start output on row 2 (after headers)

    ' Loop through all data rows (assuming headers are on row 1)
    For r = 2 To lastRow
        Dim p As Double, a As Double
        ' Convert planned and approved amounts to numbers
        p = ParseMoney(wsData.Cells(r, COL_PLANNED).Value)
        a = ParseMoney(wsData.Cells(r, COL_APPROVED).Value)
        
        ' If planned and approved amounts differ, add to report
        If p <> a Then
            wsRep.Cells(outRow, 1).Value = wsData.Cells(r, COL_EVENT).Value  ' Event name
            wsRep.Cells(outRow, 2).Value = wsData.Cells(r, COL_NEW).Value    ' New indicator
            wsRep.Cells(outRow, 3).Value = p                                 ' Planned amount
            wsRep.Cells(outRow, 4).Value = a                                 ' Approved amount
            wsRep.Cells(outRow, 5).Value = p - a                            ' Variance (difference)
            wsRep.Cells(outRow, 6).Value = r                                ' Original row number
            outRow = outRow + 1  ' Move to next output row
        End If
    Next r

    ' Format currency columns if there are any results
    If outRow > 2 Then
        wsRep.Range("C2:E" & wsRep.Cells(wsRep.Rows.Count, "C").End(xlUp).Row).NumberFormat = "$#,##0"
    End If
    
    ' Auto-fit all columns for better display
    wsRep.Columns.AutoFit
End Sub


'============================
' WORKSHEET EVENT CODE (Goes in the data sheet's code module)
' To access: Right-click the data sheet tab > "View Code"
' This section handles real-time change detection
'============================
Option Explicit

' Module-level variables to track the original value before changes
Dim OldValue As Variant    ' Stores the original value before editing
Dim OldAddress As String   ' Stores the cell address that was originally selected

' Event: Worksheet_SelectionChange
' Purpose: Captures the original value when user selects a cell in columns G or H
' This runs automatically whenever the user clicks on a different cell
' Parameters: Target - the range that was just selected
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next  ' Ignore any errors
    
    ' Only capture value if exactly one cell is selected
    If Target.Count = 1 Then
        ' Check if the selected cell is in columns G or H (Planned/Approved)
        If Not Intersect(Target, Me.Range("G:H")) Is Nothing Then
            ' Store the current value and address for later comparison
            OldValue = Target.Value
            OldAddress = Target.Address
        Else
            ' If not in monitored columns, clear stored values
            OldValue = vbNullString
            OldAddress = vbNullString
        End If
    Else
        ' If multiple cells selected, clear stored values
        OldValue = vbNullString
        OldAddress = vbNullString
    End If
End Sub

' Event: Worksheet_Change
' Purpose: Logs changes when Planned (G) or Approved (H) columns are modified
' This runs automatically whenever any cell value changes on the worksheet
' Parameters: Target - the range of cells that were just changed
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo CleanExit  ' Jump to cleanup if any errors occur

    ' Define which columns to monitor (G and H)
    Dim rngWatch As Range
    Set rngWatch = Union(Me.Columns(7), Me.Columns(8)) ' Columns G:H

    ' Exit if the changed cells aren't in our monitored columns
    If Intersect(Target, rngWatch) Is Nothing Then Exit Sub

    ' Temporarily disable events to prevent infinite loops during our updates
    Application.EnableEvents = False

    ' Process each changed cell individually
    Dim c As Range
    For Each c In Intersect(Target, rngWatch).Cells
        Dim r As Long, plannedNow As Variant, approvedNow As Variant
        r = c.Row  ' Get the row number of the changed cell
        
        ' Get current values for both planned and approved in this row
        plannedNow = Me.Cells(r, 7).Value   ' Column G
        approvedNow = Me.Cells(r, 8).Value  ' Column H

        ' OPTIONAL FEATURE: Calculate and display variance in column I
        Dim p As Double, a As Double
        p = ParseMoney(plannedNow)   ' Convert planned amount to number
        a = ParseMoney(approvedNow)  ' Convert approved amount to number
        
        ' If both amounts are zero, clear the variance cell
        If p = 0 And a = 0 Then
            Me.Cells(r, 9).ClearContents
        Else
            ' Calculate and display the variance (planned - approved)
            Me.Cells(r, 9).Value = p - a
            Me.Cells(r, 9).NumberFormat = "$#,##0"  ' Format as currency
        End If

        ' Determine what the old value was (best effort for single-cell edits)
        Dim oldValToLog As Variant
        If c.Address = OldAddress Then
            ' We have the exact old value
            oldValToLog = OldValue
        Else
            ' This happens with multi-cell paste or programmatic changes
            oldValToLog = "(unknown)"
        End If

        ' Log this change to the change tracking sheet
        AppendChangeLog r, c.Column, oldValToLog, c.Value, plannedNow, approvedNow
    Next c

CleanExit:
    ' Always re-enable events and clear temporary variables
    Application.EnableEvents = True
    OldValue = vbNullString
    OldAddress = vbNullString
End Sub

'============================
' EXCEL CHANGE TRACKING SYSTEM - TEMPLATE VERSION
' Generic template that can be customized for any Excel change tracking needs
' 
' SETUP INSTRUCTIONS:
' 1. Copy this code into your Excel VBA editor
' 2. Modify the configuration constants below to match your spreadsheet
' 3. Copy the worksheet event code to your data sheet's code module
' 4. Enable macros when opening the file
'============================

'============================
' MODULE CODE (Insert > Module in VBA Editor)
'============================
Option Explicit

'== CONFIGURATION SECTION - CUSTOMIZE THESE VALUES ==
' TODO: Change these constants to match your specific spreadsheet structure

Public Const DATA_SHEET As String = "YourDataSheetName"     ' TODO: Replace with your main data sheet name
Public Const LOG_SHEET As String = "Change_Log"             ' Name for the change tracking sheet
Public Const REPORT_SHEET As String = "Variance_Report"     ' Name for the variance report sheet

' TODO: Update these column numbers to match your spreadsheet layout
Public Const COL_IDENTIFIER As Long = 1    ' Column containing unique identifier (e.g., ID, Name, etc.)
Public Const COL_CATEGORY As Long = 2      ' Column containing category or type information
Public Const COL_VALUE1 As Long = 7        ' First column to monitor for changes
Public Const COL_VALUE2 As Long = 8        ' Second column to monitor for changes
Public Const COL_VARIANCE As Long = 9      ' Column where variance will be calculated

'============================
' UTILITY FUNCTIONS - NO CHANGES NEEDED
'============================

' Creates a worksheet if it doesn't exist
Public Function EnsureSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        EnsureSheet.Name = sheetName
    End If
End Function

' Safely converts currency text to numbers
Public Function ParseMoney(ByVal v As Variant) As Double
    Dim s As String
    s = Trim$(CStr(v))
    If s = "" Then Exit Function
    If UCase$(s) = "TBD" Or UCase$(s) = "N/A" Then Exit Function
    s = Replace$(s, "$", "")
    s = Replace$(s, ",", "")
    If IsNumeric(s) Then ParseMoney = CDbl(s)
End Function

' Logs changes to the tracking sheet
Public Sub AppendChangeLog(ByVal rowNum As Long, ByVal colNum As Long, _
                           ByVal oldVal As Variant, ByVal newVal As Variant, _
                           ByVal value1 As Variant, ByVal value2 As Variant)

    Dim wsLog As Worksheet, nextRow As Long, wsData As Worksheet
    Set wsLog = EnsureSheet(LOG_SHEET)
    Set wsData = ThisWorkbook.Worksheets(DATA_SHEET)

    ' Create headers if log sheet is empty
    If wsLog.Range("A1").Value = "" Then
        ' TODO: Customize these headers to match your data
        wsLog.Range("A1:I1").Value = Array("When", "User", "Row", "Changed Column", "Identifier", "Old Value", "New Value", "Value1 (row)", "Value2 (row)")
        wsLog.Rows(1).Font.Bold = True
        wsLog.Columns.AutoFit
    End If

    nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    wsLog.Cells(nextRow, 1).Value = Now
    wsLog.Cells(nextRow, 2).Value = Environ$("Username")
    wsLog.Cells(nextRow, 3).Value = rowNum
    
    ' TODO: Customize these column descriptions
    wsLog.Cells(nextRow, 4).Value = IIf(colNum = COL_VALUE1, "Value1", "Value2")
    wsLog.Cells(nextRow, 5).Value = wsData.Cells(rowNum, COL_IDENTIFIER).Value
    wsLog.Cells(nextRow, 6).Value = oldVal
    wsLog.Cells(nextRow, 7).Value = newVal
    wsLog.Cells(nextRow, 8).Value = value1
    wsLog.Cells(nextRow, 9).Value = value2
End Sub

' Creates a variance report
Public Sub RebuildVarianceReport()
    Dim wsData As Worksheet, wsRep As Worksheet, r As Long, lastRow As Long, outRow As Long
    Set wsData = ThisWorkbook.Worksheets(DATA_SHEET)
    Set wsRep = EnsureSheet(REPORT_SHEET)

    wsRep.Cells.Clear
    
    ' TODO: Customize these report headers
    wsRep.Range("A1:F1").Value = Array("Identifier", "Category", "Value1", "Value2", "Variance", "Row #")
    wsRep.Rows(1).Font.Bold = True

    lastRow = wsData.Cells(wsData.Rows.Count, COL_IDENTIFIER).End(xlUp).Row
    outRow = 2

    For r = 2 To lastRow ' Assumes headers on row 1
        Dim v1 As Double, v2 As Double
        v1 = ParseMoney(wsData.Cells(r, COL_VALUE1).Value)
        v2 = ParseMoney(wsData.Cells(r, COL_VALUE2).Value)
        If v1 <> v2 Then
            wsRep.Cells(outRow, 1).Value = wsData.Cells(r, COL_IDENTIFIER).Value
            wsRep.Cells(outRow, 2).Value = wsData.Cells(r, COL_CATEGORY).Value
            wsRep.Cells(outRow, 3).Value = v1
            wsRep.Cells(outRow, 4).Value = v2
            wsRep.Cells(outRow, 5).Value = v1 - v2
            wsRep.Cells(outRow, 6).Value = r
            outRow = outRow + 1
        End If
    Next r

    If outRow > 2 Then
        wsRep.Range("C2:E" & wsRep.Cells(wsRep.Rows.Count, "C").End(xlUp).Row).NumberFormat = "$#,##0"
    End If
    wsRep.Columns.AutoFit
End Sub

'============================
' WORKSHEET EVENT CODE - COPY THIS TO YOUR DATA SHEET'S CODE MODULE
' To access: Right-click your data sheet tab > "View Code"
' Then paste this code in that window
'============================

' Option Explicit

' Dim OldValue As Variant
' Dim OldAddress As String

' Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'     On Error Resume Next
'     If Target.Count = 1 Then
'         ' TODO: Update this range to match your monitored columns
'         If Not Intersect(Target, Me.Range("G:H")) Is Nothing Then
'             OldValue = Target.Value
'             OldAddress = Target.Address
'         Else
'             OldValue = vbNullString
'             OldAddress = vbNullString
'         End If
'     Else
'         OldValue = vbNullString
'         OldAddress = vbNullString
'     End If
' End Sub

' Private Sub Worksheet_Change(ByVal Target As Range)
'     On Error GoTo CleanExit

'     ' TODO: Update these column numbers to match COL_VALUE1 and COL_VALUE2
'     Dim rngWatch As Range
'     Set rngWatch = Union(Me.Columns(7), Me.Columns(8))

'     If Intersect(Target, rngWatch) Is Nothing Then Exit Sub

'     Application.EnableEvents = False

'     Dim c As Range
'     For Each c In Intersect(Target, rngWatch).Cells
'         Dim r As Long, value1Now As Variant, value2Now As Variant
'         r = c.Row
'         ' TODO: Update these column numbers
'         value1Now = Me.Cells(r, 7).Value
'         value2Now = Me.Cells(r, 8).Value

'         ' Optional: Calculate variance
'         Dim v1 As Double, v2 As Double
'         v1 = ParseMoney(value1Now)
'         v2 = ParseMoney(value2Now)
'         If v1 = 0 And v2 = 0 Then
'             ' TODO: Update this column number
'             Me.Cells(r, 9).ClearContents
'         Else
'             Me.Cells(r, 9).Value = v1 - v2
'             Me.Cells(r, 9).NumberFormat = "$#,##0"
'         End If

'         Dim oldValToLog As Variant
'         If c.Address = OldAddress Then
'             oldValToLog = OldValue
'         Else
'             oldValToLog = "(unknown)"
'         End If

'         AppendChangeLog r, c.Column, oldValToLog, c.Value, value1Now, value2Now
'     Next c

' CleanExit:
'     Application.EnableEvents = True
'     OldValue = vbNullString
'     OldAddress = vbNullString
' End Sub

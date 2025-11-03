Attribute VB_Name = "Module3"
Sub DeleteRowsBasedOnCriteria()
    Dim criteria As String
    Dim headerName As String
    Dim dataRange As Range
    Dim colField As Long
    
    ' --- 1. Get User Input ---
    headerName = InputBox("Enter the exact header name of the column to filter:", "Column Header")
    If headerName = "" Then Exit Sub
    
    criteria = InputBox("Enter the value of rows you want to DELETE:", "Criteria to Delete")
    If criteria = "" Then Exit Sub

    ' --- 2. Set up Speed Mode & Error Handling ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False ' Turn off "Delete" warnings
    On Error GoTo Cleanup

    ' 3. Find the data and the column
    Set dataRange = ActiveSheet.UsedRange
    
    ' Find the column number from the header name
    On Error Resume Next ' In case header isn't found
    colField = WorksheetFunction.Match(headerName, dataRange.Rows(1), 0)
    On Error GoTo Cleanup
    
    If colField = 0 Then
        MsgBox "Header '" & headerName & "' not found. Macro stopped.", vbCritical
        GoTo Cleanup
    End If
    
    ' 4. Apply AutoFilter
    dataRange.AutoFilter Field:=colField, Criteria1:=criteria
    
    ' 5. Delete all visible rows (except the header)
    dataRange.Offset(1).Resize(dataRange.Rows.Count - 1).SpecialCells(xlCellTypeVisible).Delete

Cleanup:
    ' 6. Clean up
    On Error Resume Next ' Clear filter even if no rows were found
    ActiveSheet.AutoFilterMode = False
    ActiveSheet.ShowAllData
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    MsgBox "Deletion complete."
End Sub

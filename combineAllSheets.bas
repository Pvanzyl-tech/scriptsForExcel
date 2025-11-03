Attribute VB_Name = "Module2"
Sub CombineAllSheets()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo Cleanup

    Dim ws As Worksheet
    Dim masterSheet As Worksheet
    Dim lastRowMaster As Long
    Dim sourceRange As Range
    Dim headerCopied As Boolean
    
    ' 1. Create or clear the Master sheet
    If SheetExists("Master") Then
        Sheets("Master").Cells.Clear
        Set masterSheet = Sheets("Master")
    Else
        Set masterSheet = Worksheets.Add(After:=Sheets(Sheets.Count))
        masterSheet.Name = "Master"
    End If
    
    headerCopied = False
    
    ' 2. Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        ' 3. Ignore the Master sheet itself
        If ws.Name <> "Master" Then
            
            ' 4. Find the data on the source sheet
            Set sourceRange = ws.UsedRange
            
            ' 5. Handle Headers
            If Not headerCopied Then
                ' Copy header row from the first sheet
                sourceRange.Rows(1).Copy masterSheet.Range("A1")
                headerCopied = True
                
                ' Copy data *without* the header
                sourceRange.Offset(1).Resize(sourceRange.Rows.Count - 1).Copy
            Else
                ' If headers are already copied, just copy data
                ' (Assumes all sheets have a header row)
                sourceRange.Offset(1).Resize(sourceRange.Rows.Count - 1).Copy
            End If
            
            ' 6. Find the next empty row on the Master sheet
            lastRowMaster = masterSheet.Cells(masterSheet.Rows.Count, "A").End(xlUp).Row
            
            ' 7. Paste the data
            masterSheet.Cells(lastRowMaster + 1, "A").PasteSpecial xlPasteValues
            Application.CutCopyMode = False
        End If
    Next ws

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    masterSheet.Columns.AutoFit
    masterSheet.Activate
    masterSheet.Range("A1").Select
    MsgBox "All sheets combined."
End Sub

' Helper function for the "CombineAllSheets" macro
Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    SheetExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next ws
End Function

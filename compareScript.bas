Attribute VB_Name = "Module1"
Sub CompareAndHighlight()
    Dim col1 As Range, col2 As Range
    Dim lastRow As Long, i As Long
    Dim lastRow1 As Long, lastRow2 As Long
    Dim val1 As String, val2 As String
    
    ' --- 1. Get User Input (Error handling for 'Cancel') ---
    On Error Resume Next
    
    ' Get first column
    Set col1 = Application.InputBox(Prompt:="Please select the first column to compare (click anywhere in it).", _
                                    Title:="Select Column 1", Type:=8)
    If col1 Is Nothing Then Exit Sub ' User pressed Cancel
    
    ' Get second column
    Set col2 = Application.InputBox(Prompt:="Please select the second column to compare (click anywhere in it).", _
                                    Title:="Select Column 2", Type:=8)
    If col2 Is Nothing Then Exit Sub ' User pressed Cancel
    
    On Error GoTo 0 ' Re-enable error handling
    
    ' --- 2. Use the Entire Column and Clear Old Formatting ---
    Set col1 = col1.EntireColumn
    Set col2 = col2.EntireColumn
    
    col1.Interior.ColorIndex = xlNone
    col2.Interior.ColorIndex = xlNone
    
    ' --- 3. Find the Last Used Row (checking both columns) ---
    lastRow1 = Cells(Rows.Count, col1.Column).End(xlUp).Row
    lastRow2 = Cells(Rows.Count, col2.Column).End(xlUp).Row
    
    ' Use the larger of the two last rows
    If lastRow1 > lastRow2 Then
        lastRow = lastRow1
    Else
        lastRow = lastRow2
    End If
    
    ' --- 4. Loop, Compare, and Highlight (NEW LOGIC) ---
    ' Loop from the first row to the last used row
    For i = 1 To lastRow
        ' Trim() removes extra spaces. Use .Value2 for faster comparison.
        val1 = Trim(Cells(i, col1.Column).Value2)
        val2 = Trim(Cells(i, col2.Column).Value2)
        
        ' Skip rows where both cells are blank
        If val1 = "" And val2 = "" Then
            ' Do nothing, leave blank
        ElseIf val1 = val2 Then
            ' MATCH: Color both cells Green
            Cells(i, col1.Column).Interior.Color = vbGreen
            Cells(i, col2.Column).Interior.Color = vbGreen
        Else
            ' NO MATCH: Color both cells Red
            Cells(i, col1.Column).Interior.Color = vbRed
            Cells(i, col2.Column).Interior.Color = vbRed
        End If
    Next i
    
    MsgBox "Comparison complete. Matches are green, non-matches are red."
End Sub


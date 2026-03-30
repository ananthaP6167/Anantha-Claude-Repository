Attribute VB_Name = "TransformCusip"
'=============================================================
' MACRO: TransformCusip
'
' Opens a file picker, loads a CUSIP file, applies LEFT(3)
' to column A, writes results to column B, and saves the file.
'
' HOW TO SET UP (one-time):
'   1. Open a new blank workbook in Excel
'   2. Press Alt+F11 to open the VBA Editor
'   3. Go to File > Import File... and select this .bas file
'   4. Close the VBA Editor
'   5. Save the workbook as .xlsm (macro-enabled)
'
' HOW TO RUN (each time):
'   1. Open the .xlsm file
'   2. Press Alt+F8, select "TransformCusip", click Run
'   3. Pick your CUSIP input file (.xlsx, .xls, or .csv)
'   4. The macro transforms and saves the file automatically
'=============================================================

Sub TransformCusip()

    ' --- File Picker ---
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Select the CUSIP input file"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xls;*.xlsm;*.csv"
        .Filters.Add "All Files", "*.*"
        .AllowMultiSelect = False
    End With
    
    If fd.Show <> -1 Then
        MsgBox "No file selected. Macro cancelled.", vbExclamation, "Cusip Transform"
        Exit Sub
    End If
    
    Dim sFilePath As String
    sFilePath = fd.SelectedItems(1)
    
    ' --- Open the file ---
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks.Open(sFilePath)
    On Error GoTo 0
    
    If wb Is Nothing Then
        MsgBox "Could not open file:" & vbCrLf & sFilePath, vbCritical, "Cusip Transform"
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = wb.Sheets(1)
    
    ' --- Find last row in column A ---
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "No data found in column A (need header + at least 1 row).", _
               vbExclamation, "Cusip Transform"
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' --- Clear column B and write header ---
    ws.Range("B1").Value = "transformed cusip"
    If lastRow > 1 Then
        ws.Range("B2:B" & lastRow).ClearContents
    End If
    
    ' --- Apply LEFT(,3) formula ---
    Dim i As Long
    For i = 2 To lastRow
        ws.Cells(i, 2).Formula = "=LEFT(A" & i & ",3)"
    Next i
    
    ' --- Auto-fit and save ---
    ws.Columns("B").AutoFit
    wb.Save
    
    MsgBox "Transformation complete!" & vbCrLf & _
           "Processed " & (lastRow - 1) & " rows." & vbCrLf & vbCrLf & _
           "File saved: " & sFilePath, _
           vbInformation, "Cusip Transform"

End Sub

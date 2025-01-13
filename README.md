# CloseSpreadsheet
VBA Script to Close Multiple SAP-Opened Spreadsheets

Public Function CloseSpreadsheet(sheetName As String) As Boolean
    
    Dim xlApp As Object ' Excel Application object
    Dim wb As Object ' Workbook object
    Dim OpenedSheetName As String
    Dim i As Integer
    
    ' Attempt to connect to an open instance of Excel
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    On Error GoTo 0
    
    ' Check if Excel is open
    If xlApp Is Nothing Then
        MsgBox "No instance of Excel is currently open"
        Exit Function
    End If
    
    ' Loop through all open workbooks
    For Each wb In xlApp.Workbooks
        OpenedSheetName = wb.Name
        ' If the workbook name matches the target name, close it
        If OpenedSheetName = sheetName Then
            wb.Close SaveChanges:=False ' Close without saving changes
        End If
    Next wb
    
    ' Release the Excel application object
    Set xlApp = Nothing
    
End Function

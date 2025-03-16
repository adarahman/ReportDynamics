Attribute VB_Name = "BVariation"
' --- Module1.bas ---
Option Explicit

Public Sub AccumulateValues()
    On Error GoTo ErrorHandler
    
    ' Get date range from user
    Dim startDate As String
    Dim endDate As String
    
    ' Show input form for dates
    If Not GetDateRange(startDate, endDate) Then
        MsgBox "Operation cancelled by user.", vbInformation
        Exit Sub
    End If
    
    ' Initialize objects
    Set FileProcessor = New FileProcessor
    Set DataAccumulator = New DataAccumulator
    
    Dim outputWs As Worksheet
    
    ' Validate output sheet
    If Not SheetExists("Accumulated Values") Then
        Err.Raise vbObjectError + 1, "AccumulateValues", "Sheet 'Accumulated Values' not found!"
    End If
    Set outputWs = ThisWorkbook.Sheets("Accumulated Values")
    
    ' Initialize file processor
    If Not FileProcessor.Initialize("C:\EDW_RPT\", startDate, endDate) Then
        Exit Sub
    End If
    
    ' Configure Excel
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .StatusBar = "Processing files..."
    End With
    
    ' Process files
    If Not FileProcessor.ProcessFiles(DataAccumulator) Then
        GoTo CleanExit
    End If
    
    ' Output results
    DataAccumulator.OutputToWorksheet outputWs
    
    ' Add date range to the output sheet
    With outputWs
        .Range("A1").EntireRow.Insert
        .Range("A1:F1").Merge
        .Range("A1").value = "Daily Variation CR/DR Transaction >= 50 Lacs for Date Range: " & startDate & " to " & endDate
        .Range("A1").Font.Bold = True
        .Range("A1").HorizontalAlignment = xlCenter
    End With
    
    ' Show results
    ShowMsgBox "Business Variation Accumulation complete!" & vbNewLine & _
           "Processed " & FileProcessor.ProcessedCount & " files." & vbNewLine & _
           "Found " & DataAccumulator.GetCount & " unique combinations." ', vbInformation
           
CleanExit:
    ' Reset Excel
    With Application
        .StatusBar = False
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume CleanExit
End Sub



Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function




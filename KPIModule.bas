Option Explicit
' V2.0 Standard Module KPIConsolidator 06/30/2025
' Attribute VB_Name = "KPIModule"

' Global variables
Public msg As String
Public rDate As String, pDate As String
Public wsRef As Worksheet
Public FILE_PATH As String, FILE_KPI_PATH As String
Public KPIcolArray As Variant
Public START_SET As Long, DERIVED_SET As Long
Public branchArray As Variant
Public FILE_PATTERNS As String
Public FILE_INFO As Range
Public dateVals As Variant
Public BRANCH_CODE_COL As Long
Public BRANCH_NAME_COL As Long

Public Sub BusinessKPI()
    Dim branchData As Variant
    Dim wsBrFigure As Worksheet, wsBranchVar As Worksheet, wsBVar As Worksheet
    Dim KPIConsolidator As KPIConsolidator ' Use early binding, assuming class module exists
    Dim budgetDict As Object
    Dim savePath As String
    Dim startTime As Double
    Dim eFlag As Boolean
    Dim totalRows As Long, totalCols As Long
    Dim i As Long, j As Long
    Dim lastRow As Long

    On Error GoTo ErrorHandler
    startTime = Timer
    msg = ""

    ' Optimize Excel for speed
    OptimizeExcel True

    ' Load configuration and validate
    If Not LoadConfiguration(dateVals) Then
        MsgBox "Failed to load configuration.", vbCritical
        GoTo CleanUp
    End If
    If IsEmpty(dateVals) Or UBound(dateVals) < LBound(dateVals) Then
        MsgBox "No valid dates provided.", vbExclamation
        GoTo CleanUp
    End If

    ' Instantiate consolidator and process files
    Set KPIConsolidator = New KPIConsolidator
    KPIConsolidator.ProcessFiles FILE_PATH, dateVals

    ' Populate budget data
    Set budgetDict = CreateObject("Scripting.Dictionary")
    KPIConsolidator.budgetPopulate budgetDict
    If budgetDict.Count = 0 Then
        MsgBox "No budget data loaded.", vbExclamation
        GoTo CleanUp
    End If

    ' Generate reports & update branch data
    If Not KPIConsolidator.HasProcessedFiles Then
        MsgBox "No files processed for the given dates.", vbExclamation
        GoTo CleanUp
    End If
    KPIConsolidator.GenerateAllKPIReports FILE_KPI_PATH
    KPIConsolidator.UpdateBranchDataAndMetrics branchData, dateVals, budgetDict, KPIcolArray

    ' Get worksheet references
    Set wsBranchVar = GetWorksheetSafe(ThisWorkbook, "BRVAR1CR")
    Set wsBVar = GetWorksheetSafe(ThisWorkbook, "ACVAR50L")

    ' Process net business variations
    NetBusinessVar wsBVar

    ' Apply filter to branch variations
    OptimizeExcel False
    With wsBranchVar
        If .AutoFilterMode Then .AutoFilterMode = False
        lastRow = .Cells(.rows.Count, "A").End(xlUp).row
        If lastRow > 2 Then .Range("A2:N" & lastRow).AutoFilter Field:=14, Criteria1:="<>0"
    End With
    OptimizeExcel True

    ' Achieved calculation
'    KPIConsolidator.ACHIEVED branchData, budgetDict, KPIcolArray, dateVals

    ' Save workbook - DATA ONLY METHOD (No Macros)
    SaveAsDataOnlyWorkbook rDate, pDate

    ' Success message
    msg = msg & vbNewLine & "Business Daily Report Updated for Date: " & rDate & vbNewLine & _
          "Processed " & UBound(dateVals) - LBound(dateVals) + 1 & " dates in " & _
          Format(Timer - startTime, "0.00") & " seconds." & vbNewLine & _
          "File saved as: " & ThisWorkbook.path & "\" & rDate & "-" & pDate & "eDAILY.xlsx"

CleanUp:
    OptimizeExcel False
    Set KPIConsolidator = Nothing
    Set budgetDict = Nothing
    Set wsBrFigure = Nothing
    Set wsBranchVar = Nothing
    Set wsBVar = Nothing
    showMsgBox msg
    Exit Sub

ErrorHandler:
    msg = msg & vbNewLine & "Error in BusinessKPI: " & Err.Description
    Resume CleanUp
End Sub


' DATA-ONLY SAVE METHOD - Separate Procedure
Private Sub SaveAsDataOnlyWorkbook(ByVal reportDate As String, ByVal processDate As String)
    Dim newWb As Workbook
    Dim savePath As String
    Dim sheetsToSave As Variant
    Dim i As Long
    Dim originalWs As Worksheet, newWs As Worksheet
    
    On Error GoTo SaveDataError
    
    ' Define which sheets to save (adjust as needed)
    sheetsToSave = Array("1Pager", "BRANCH", "BRVAR1CR", "ACVAR50L")
    
    Application.DisplayAlerts = False
    ' Create new workbook
    Set newWb = Workbooks.Add
    
    ' Remove default sheets except one
    Do While newWb.Worksheets.Count > 1
        newWb.Worksheets(newWb.Worksheets.Count).Delete
    Loop
    
    ' Copy each required sheet
    For i = LBound(sheetsToSave) To UBound(sheetsToSave)
        Set originalWs = GetWorksheetSafe(ThisWorkbook, CStr(sheetsToSave(i)))
        If Not originalWs Is Nothing Then
            ' Copy the worksheet
            originalWs.Copy After:=newWb.Worksheets(newWb.Worksheets.Count)
            Set newWs = newWb.Worksheets(newWb.Worksheets.Count)
            newWs.Name = CStr(sheetsToSave(i))
            
            ' Clear any VBA code ranges or formula links if needed
            ClearVBAReferences newWs
        End If
    Next i

    ' Delete the original default sheet
    newWb.Worksheets(1).Delete
'    ApplyFormulasWithVBA
    ' Set the first data sheet as active
    If newWb.Worksheets.Count > 0 Then
        newWb.Worksheets(1).Activate
    End If
    
    ' Save the new workbook
    savePath = ThisWorkbook.path & "\" & reportDate & "-" & processDate & "eDAILY.xlsx"
    newWb.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
    newWb.Close SaveChanges:=False
    
    Set newWb = Nothing
    Exit Sub
    
SaveDataError:
    Application.DisplayAlerts = True
    If Not newWb Is Nothing Then
        newWb.Close SaveChanges:=False
        Set newWb = Nothing
    End If
    msg = msg & vbNewLine & "Error saving data-only workbook: " & Err.Description
    On Error GoTo 0
End Sub

' Helper procedure to clear VBA references from copied sheets
Private Sub ClearVBAReferences(ByRef ws As Worksheet)
    On Error Resume Next
    
    ' Clear any potential VBA references or external links
    Dim cell As Range
    
    ' Remove any worksheet-level names that might reference VBA
    Dim nm As Name
    For Each nm In ws.Parent.Names
        If InStr(nm.RefersTo, ws.Name) > 0 Then
            nm.Delete
        End If
    Next nm
    
'     Convert any remaining formulas to values if they reference external sources
'     (Optional - uncomment if needed)
     For Each cell In ws.UsedRange
         If cell.HasFormula Then
             If InStr(cell.Formula, "!") > 0 Or InStr(cell.Formula, "[") > 0 Then
                 cell.value = cell.value ' Convert to value
             End If
         End If
     Next cell
    
    On Error GoTo 0
End Sub

' Enhanced version with sheet selection dialog (Optional)
Private Sub SaveAsDataOnlyWorkbookWithSelection(ByVal reportDate As String, ByVal processDate As String)
    Dim newWb As Workbook
    Dim savePath As String
    Dim selectedSheets As String
    Dim sheetNames As Variant
    Dim i As Long
    Dim originalWs As Worksheet
    
    On Error GoTo SaveDataError
    
    ' Let user select which sheets to save (optional enhancement)
    selectedSheets = InputBox("Enter sheet names to save (comma-separated):" & vbNewLine & _
                             "Available: BRANCH, BRVAR1CR, ACVAR50L, Reference", _
                             "Select Sheets to Save", "BRANCH,BRVAR1CR,ACVAR50L")
    
    If selectedSheets = "" Then Exit Sub
    
    ' Parse selected sheets
    sheetNames = Split(selectedSheets, ",")
    
    ' Create new workbook
    Application.DisplayAlerts = False
    Set newWb = Workbooks.Add
    
    ' Remove default sheets except one
    Do While newWb.Worksheets.Count > 1
        newWb.Worksheets(newWb.Worksheets.Count).Delete
    Loop
    
    ' Copy selected sheets
    For i = LBound(sheetNames) To UBound(sheetNames)
        Set originalWs = GetWorksheet(ThisWorkbook, Trim(CStr(sheetNames(i))))
        If Not originalWs Is Nothing Then
            originalWs.Copy After:=newWb.Worksheets(newWb.Worksheets.Count)
            newWb.Worksheets(newWb.Worksheets.Count).Name = Trim(CStr(sheetNames(i)))
        End If
    Next i
    
    ' Delete the original default sheet
    newWb.Worksheets(1).Delete
    
    ' Save and close
    savePath = ThisWorkbook.path & "\" & reportDate & "-" & processDate & "eDAILY.xlsx"
    newWb.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
    newWb.Close SaveChanges:=False
    
    Application.DisplayAlerts = True
    Set newWb = Nothing
    Exit Sub
    
SaveDataError:
    Application.DisplayAlerts = True
    If Not newWb Is Nothing Then
        newWb.Close SaveChanges:=False
        Set newWb = Nothing
    End If
    msg = msg & vbNewLine & "Error saving data-only workbook: " & Err.Description
    On Error GoTo 0
End Sub


Private Sub FileOpen(ByVal filePath As String, ByRef wb As Workbook, ByRef errorFlag As Boolean)
    On Error Resume Next
    Set wb = Workbooks.Open(filePath, ReadOnly:=False, UpdateLinks:=False)
    errorFlag = (Err.Number <> 0)
    If errorFlag Then
        MsgBox "Error opening file: " & filePath & vbCrLf & Err.Description, vbExclamation
    End If
    On Error GoTo 0
End Sub
Public Sub NetBusinessVar(ByRef outputWs As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim startDate As String, endDate As String
    Dim startTime As Double: startTime = Timer
    
    If Not GetDateRange(startDate, endDate) Then
        MsgBox "Operation cancelled by user.", vbInformation
        Exit Sub
    End If
    
    Dim FileProcessor As FileProcessor
    Dim DataAccumulator As DataAccumulator
    Set FileProcessor = New FileProcessor
    Set DataAccumulator = New DataAccumulator
    
    If Not FileProcessor.Initialize(wsRef.Range("RptPath"), startDate, endDate) Then
        GoTo CleanExit
    End If
    
    OptimizeExcel True
    If Not FileProcessor.ProcessFiles(DataAccumulator) Then
        GoTo CleanExit
    End If
    
    DataAccumulator.OutputToWorksheet outputWs
    With outputWs
        .Range("A1").EntireRow.Insert
        .Range("A1:F1").Merge
        .Range("A1").value = "Daily Variation CR/DR Transaction >= 50 Lacs for Date Range: After " & startDate & " to " & endDate
        .Range("A1").Font.Bold = True
        .Range("A1").HorizontalAlignment = xlCenter
    End With
    
    msg = msg & vbNewLine & "Business Variation Accumulation complete!" & vbNewLine & _
          "Processed " & FileProcessor.ProcessedCount & " Files. " & _
          DataAccumulator.GetCount & " Unique Combinations in " & _
          Format(Timer - startTime, "0.00") & " seconds"

CleanExit:
    OptimizeExcel False
    Set FileProcessor = Nothing
    Set DataAccumulator = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    On Error Resume Next  ' Prevent further errors during cleanup
    GoTo CleanExit
End Sub

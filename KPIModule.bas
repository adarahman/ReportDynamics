Attribute VB_Name = "KPIModule"
Option Explicit

' KPIModule.bas Module Level Variables
Public branchDict As Object
Public MsgString As String
Public rowCache As Object
Public ColumnCache As Object
Public FILE_PATH As String
Public FILE_KPI_PATH As String
Public Sub KPI(DateVals As Variant)
    On Error GoTo ErrorHandler
    ' Initialize the reference worksheet
    Set wsRefer = ThisWorkbook.Worksheets("Reference") ' Replace "Settings" with your actual worksheet name
    
    ' Optimize Excel settings
    OptimizeExcel True
    FILE_PATH = wsRefer.Range("RptPath").value
    FILE_KPI_PATH = wsRefer.Range("KPIpath").value
    
' Create and initialize KPIConsolidator
    Dim KPIConsolidator As KPIConsolidator
    Set KPIConsolidator = New KPIConsolidator
    KPIConsolidator.LoadCommonHeaders
    KPIConsolidator.LoadLabelMaps
    
'''     'Get date range from user
'''    Dim startDate As String, endDate As String
'''    ' Show input form for dates
'''    If Not GetDateRange(startDate, endDate) Then
'''        MsgBox "Operation cancelled by user.", vbInformation
'''        Exit Sub
'''    End If
  
Dim dateVal As Variant
For Each dateVal In DateVals
    KPIConsolidator.ProcessFiles FILE_PATH, dateVal
    KPIConsolidator.GenerateReport FILE_KPI_PATH, dateVal
Next dateVal

CleanUp:
    OptimizeExcel False
    Set KPIConsolidator = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanUp
End Sub
Sub BusinessFigure()
    ' Purpose: Process business data, calculate variances, and update reports
    ' Last Updated: March 2025

    ' Local variables
    Dim DateVals As Variant, KPIColArray As Variant
    Dim startSet As Long, DataSet As Long
    Dim LocalFilePath As String
    Dim LR As Long, LC As Long, k As Long
    Dim branchData As Variant
    Dim Budget_Dict As Object
    Dim DateSrl As Long
    Dim KPIFilePath As String, wbKPI As Workbook, wsKPI As Worksheet
    Dim MissingCols As String
    Dim KPIdb As Variant, KPIndex() As Long, BranchArray As Variant
    Dim KPI_Dict As Object
'    Dim LookupValue As String, KPIRow As Variant
    Dim BudgetLookup As String, savePath As String, BudgetCol As Long, lastRow As Long
    Dim eFlag As Boolean
    Dim rDate As String, pDate As String, fDate As String, qDate As String, pfyDate As String
'    Dim CDateVal As Long
    Dim wbDaily As Workbook, wsBrFigure As Worksheet, wsBranchVar As Worksheet, WSBVAR As Worksheet
'    On Error GoTo ErrorHandler

    ' Disable UI updates for performance
    Call OptimizeExcel(True)

    ' Initialize message string
    MsgString = ""
    Set wsMain = ThisWorkbook.Sheets("RepGenT")
    Set wsRefer = ThisWorkbook.Sheets("Reference")

    ' Set dates from wsRefer sheet
    rDate = wsRefer.Range("cDate").Value2
    pDate = wsRefer.Range("pDate").Value2
    fDate = wsRefer.Range("fDate").Value2
    qDate = wsRefer.Range("qDate").Value2
    pfyDate = wsRefer.Range("pfyDate").Value2
    DateVals = Application.Transpose(wsRefer.Range("DateArray").Value2)
    
    ' Set up data range
    startSet = wsRefer.Range("SetBegin").Value2
    DataSet = wsRefer.Range("DataSet").Value2 + UBound(DateVals) + 1

    ' Initialize budget dictionary
    Call BudgetDict(Budget_Dict)
    MsgString = MsgString & vbCrLf & "Budget Updated For Quarter : " & wsRefer.Range("BudgetPeriod")
    
   ' Retrieve KPI columns
    KPIColArray = Application.Transpose(wsRefer.Range("KPICols").Columns(2).Value2)
    ReDim KPIndex(UBound(KPIColArray))
    
    ' Add all branch data in one operation
    BranchArray = wsRefer.Range("CircleSet").Value2
    ' Get data from BranchFigure sheet in one operation
    LR = UBound(BranchArray) + 2
    LC = UBound(KPIColArray) * DataSet + startSet
    ReDim branchData(1 To LR, 1 To LC)
    
    Dim rowIndex As Long, colIndex As Long, key As Variant, BrannchInfo As Variant, Col As Variant
    ' Assuming BranchArray is a range with 4 columns
    For rowIndex = 1 To UBound(BranchArray, 1)
        ' Assign values from BranchArray to branchData
        branchData(rowIndex + 1, 1) = rowIndex - 1
        For colIndex = 1 To UBound(BranchArray, 2)
            branchData(rowIndex + 1, colIndex + 1) = BranchArray(rowIndex, colIndex)
        Next colIndex
    Next rowIndex

' Process each date
Dim dateVal As Object
Set dateVal = CreateObject("Scripting.Dictionary")

For DateSrl = LBound(DateVals) To UBound(DateVals)
    ' Get KPI file path
    KPIFilePath = wsRefer.Range("KPIpath").Value2 & DateVals(DateSrl) & "KPI.xlsx"
    
    ' Check if KPI file exists, and if the date is not already in the dictionary
    If Dir(KPIFilePath) = "" Then
        If Not dateVal.Exists(DateVals(DateSrl)) Then
            dateVal.Add DateVals(DateSrl), DateVals(DateSrl)
        End If
    End If
Next DateSrl
    
    Call KPI(dateVal)
    
    ' Process KPI data efficiently
    Dim keyDate As Variant
For DateSrl = LBound(DateVals) To UBound(DateVals)
    Call ProcessKPIData(KPIFilePath, KPIdb, KPI_Dict, KPIColArray, KPIndex, MissingCols)

    ' Update branch data with KPI values
    Call UpdateBranchData(branchData, KPI_Dict, KPIndex, KPIColArray, DateVals, DateSrl, startSet, DataSet, Budget_Dict)
Next DateSrl
    Set KPIdb = Nothing
    
    ' Open daily file
    LocalFilePath = wsRefer.Range("DailyPath").Value2 & "DAILY-v10.xlsx"
    Call FileOpen(LocalFilePath, wbDaily, eFlag)
    If eFlag Then GoTo CleanupAndExit

    ' Set worksheet references
    Set wsBrFigure = wbDaily.Sheets("BranchFigure")
    Set wsBranchVar = wbDaily.Sheets("Branch Var>1Cr")
    Set WSBVAR = wbDaily.Sheets("Variation Actwise =>50L")
    
    ' Process business variations
    Call BusinessVar(WSBVAR, rDate)
    Set KPI_Dict = Nothing

    ' Report missing columns
    If Len(MissingCols) > 0 Then
        MsgString = MsgString & vbCrLf & "Following Business Figure Columns for Date are Missing : " & vbCrLf _
        & Left(MissingCols, Len(MissingCols) - 2)
    End If

    ' Update BranchFigure with calculated data
    With wsBrFigure
        .Range(.Cells(1, 1), .Cells(UBound(branchData, 1), UBound(branchData, 2))) = branchData
    End With
    Set branchData = Nothing

    ' Filter branch variation
    With wsBranchVar
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).row
        .Range("A2:I" & lastRow).AutoFilter field:=9, Criteria1:="<>0", Operator:=xlAnd
    End With

    ' Save workbook
    On Error Resume Next
    savePath = wbDaily.Path & "\" & rDate & "-" & pDate & "DAILY-v10.xlsx"
    wbDaily.SaveAs fileName:=savePath, FileFormat:=wbDaily.FileFormat

    ' Handle save errors
    MsgString = MsgString & vbCrLf & "Business Daily Report for Date : " & rDate & " Updated Successfully!"

CleanupAndExit:
    ' Re-enable UI updates
    Call OptimizeExcel(False)
    wbDaily.Close SaveChanges:=False
    ShowMsgBox MsgString
    Exit Sub

ErrorHandler:
    MsgString = "Error in BusinessFigure: " & Err.Number & ": " & Err.Description & vbCrLf & MsgString
    Resume CleanupAndExit
End Sub
' Update branch data with KPI values and calculate metrics
Private Sub UpdateBranchData(ByRef branchData As Variant, ByVal KPI_Dict As Object, ByVal KPIndex As Variant, _
                           ByRef KPIColArray As Variant, ByVal DateVals As Variant, ByVal DateSrl As Long, _
                           ByVal startSet As Long, ByVal DataSet As Long, ByVal Budget_Dict As Object)
    Dim i As Long, k As Long
    Dim LookupValue As String, KPIRow As Variant
    Dim colIndex As Long, CDateVal As Long, rowIndex As Long
    Dim BudgetLookup As String, BudgetCol As Long
    Dim key As Variant
    Dim CircleSet As Variant, ColArray As Variant
    Dim LR As Long, LC As Long
    
    
    ' Update branch data with KPI values
    For i = 3 To UBound(branchData, 1)
        LookupValue = CStr(branchData(i, 5))
        If KPI_Dict.Exists(LookupValue) Then
            KPIRow = KPI_Dict(LookupValue)
            ' Process each KPI column
            For k = 1 To UBound(KPIndex) - 1
                colIndex = KPIndex(k)
                CDateVal = DataSet * (k - 1) + startSet + DateSrl

                ' Update date in header row
                branchData(2, CDateVal) = DateVals(DateSrl)

                ' Update KPI values
                If colIndex > 0 And colIndex <= UBound(KPIRow) Then
                    branchData(i, CDateVal) = IIf(IsEmpty(KPIRow(colIndex)) Or KPIRow(colIndex) = "", 0, Format(KPIRow(colIndex), "0.00"))
                End If

                ' Calculate variances, growth etc. at the last date
                If DateSrl = UBound(DateVals) Then
                    Call CalculateMetrics(branchData, i, k, CDateVal, KPIColArray, Budget_Dict)
                End If
            Next k
        End If
    Next i
End Sub
' Calculate metrics for a specific KPI and branch
Private Sub CalculateMetrics(ByRef branchData As Variant, ByVal rowIndex As Long, ByVal KPIndex As Long, _
                           ByVal CDateVal As Long, ByVal KPIColArray As Variant, ByVal Budget_Dict As Object)
    Dim BudgetKey As String, BudgetCol As Long

    ' Set column headers
    branchData(1, CDateVal) = KPIColArray(KPIndex)
    branchData(2, CDateVal + 1) = "Var. @ " & branchData(2, CDateVal - 1)
    branchData(2, CDateVal + 2) = "Growth @ " & branchData(2, CDateVal - 3)
    branchData(2, CDateVal + 3) = "%Grow.@ " & branchData(2, CDateVal - 3)
    branchData(2, CDateVal + 4) = "Budget " & CStr(wsRefer.Range("BudgetPeriod"))
    branchData(2, CDateVal + 5) = "Gap to " & branchData(2, CDateVal + 4)

    ' Calculate metrics
    branchData(rowIndex, CDateVal + 1) = branchData(rowIndex, CDateVal) - branchData(rowIndex, CDateVal - 1)
    branchData(rowIndex, CDateVal + 2) = branchData(rowIndex, CDateVal) - branchData(rowIndex, CDateVal - 3)

    ' Calculate percentage growth
    If CDbl(branchData(rowIndex, CDateVal - 3)) <> 0 Then
        branchData(rowIndex, CDateVal + 3) = Format(branchData(rowIndex, CDateVal + 2) / branchData(rowIndex, CDateVal - 3), "0.00%")
    Else
        branchData(rowIndex, CDateVal + 3) = "N/A"
    End If

    ' Update budget data using dictionary lookup
    BudgetKey = CStr(branchData(rowIndex, 5)) & "|" & UCase(KPIColArray(KPIndex))
    If Budget_Dict.Exists(BudgetKey) Then
        BudgetCol = CDateVal + 4
        branchData(rowIndex, BudgetCol) = Budget_Dict(BudgetKey)
    End If

    ' Calculate gap
    branchData(rowIndex, CDateVal + 5) = branchData(rowIndex, CDateVal) - branchData(rowIndex, CDateVal + 4)
End Sub

Sub BudgetDict(ByRef Budget_Dict As Object)
Dim j As Long
Dim BudgetPath As String
Dim wbBudget As Workbook
Dim wsBudget As Worksheet
Dim BudgetQtrCol As Long
Dim lastRow As Long
Dim BudgetPeriod As String
Dim headerRow As Variant
Dim budgetData As Variant

On Error GoTo ErrorHandler

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

' Define file path and validate it exists
BudgetPath = wsRefer.Range("Budgetpath").value
If Len(BudgetPath) = 0 Then
    MsgBox "Budget path is empty.", vbExclamation, "Invalid Path"
    Exit Sub
End If

If Dir(BudgetPath) = "" Then
    MsgBox "Budget file not found: " & BudgetPath, vbExclamation, "File Not Found"
    Exit Sub
End If

' Get BudgetPeriod value as string
BudgetPeriod = CStr(wsRefer.Range("BudgetPeriod").value)
If Len(BudgetPeriod) = 0 Then
    MsgBox "Budget period is empty.", vbExclamation, "Invalid Period"
    Exit Sub
End If

' Open workbook and set worksheet
Set wbBudget = Workbooks.Open(BudgetPath, ReadOnly:=True)

' Check if "Budgets" sheet exists
On Error Resume Next
Set wsBudget = wbBudget.Sheets("Budgets")
On Error GoTo ErrorHandler

If wsBudget Is Nothing Then
    MsgBox "The 'Budgets' worksheet was not found in the budget file.", vbExclamation, "Sheet Not Found"
    wbBudget.Close SaveChanges:=False
    Exit Sub
End If

' Check if sheet has data
If Application.WorksheetFunction.CountA(wsBudget.Cells) = 0 Then
    MsgBox "The 'Budgets' worksheet is empty.", vbExclamation, "No Data"
    wbBudget.Close SaveChanges:=False
    Exit Sub
End If

' Store header row in array for faster access
headerRow = Application.Transpose(wsBudget.Range("1:1").value)

' Find Budget Quarter column
BudgetQtrCol = Application.Match(BudgetPeriod, headerRow, 0)

' Handle case where BudgetPeriod is not found
If IsError(BudgetQtrCol) Then
    MsgBox "BudgetPeriod '" & BudgetPeriod & "' not found in the header row.", vbCritical, "Period Not Found"
    wbBudget.Close SaveChanges:=False
    Exit Sub
End If

' Create dictionary for efficient key-based lookup
Set Budget_Dict = CreateObject("Scripting.Dictionary")
Budget_Dict.CompareMode = vbTextCompare ' Case-insensitive keys

' Get data range into an array for faster processing
lastRow = wsBudget.Cells(wsBudget.Rows.Count, 1).End(xlUp).row

If lastRow <= 1 Then
    MsgBox "No budget data found below the header row.", vbExclamation, "No Data"
    wbBudget.Close SaveChanges:=False
    Exit Sub
End If

' Read all relevant data at once (columns 1, 2, and BudgetQtrCol)
budgetData = wsBudget.Range(wsBudget.Cells(2, 1), wsBudget.Cells(lastRow, BudgetQtrCol)).value

' Populate Budget_Dict
For j = 1 To UBound(budgetData, 1)
    If Not IsEmpty(budgetData(j, 1)) And Not IsEmpty(budgetData(j, 2)) Then
        Dim key As String
        key = budgetData(j, 1) & "|" & UCase(Trim(CStr(budgetData(j, 2))))
        
        ' Check if the key already exists to avoid duplicates
        If Not Budget_Dict.Exists(key) Then
            Budget_Dict(key) = budgetData(j, BudgetQtrCol - 2)
        Else
            ' Optional: handle duplicate keys (uncomment if needed)
            ' Debug.Print "Duplicate key found: " & key
        End If
    End If
Next j

CleanExit:
    ' Close the workbook without saving changes
    If Not wbBudget Is Nothing Then
        wbBudget.Close SaveChanges:=False
    End If

' Restore Excel settings
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
    Resume CleanExit
End Sub
Sub BusinessVar(ByRef WSBVAR As Worksheet, ByRef rDate As String)
    Dim pArray() As Variant, reportData As Object
    Dim i As Long, rowCount As Long
    Dim BVarCol As Variant, colIndex As Variant
    Dim BVardata() As Variant
    Dim rptName As Variant
    Dim BVRange As Range
    
    ' Turn off screen updating and automatic calculation for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Initialize report name
    rptName = "BVariance"
    
    ' Fetch report data
    ProcessFiles reportData, rDate, rptName
    For Each rptName In reportData.Keys
        If Not IsEmpty(reportData(rptName)) Then
            pArray = reportData(rptName)(1)
        Else
            MsgBox "No data returned by reportData. Please check Business Variance.", vbInformation, "Error"
            GoTo CleanUp
        End If
    Next rptName
    
    ' Identify column positions
    ClearCaches
    BVarCol = Array("BRANCH CODE", "BRANCH NAME", "SCHEME TYPE", "VARIATION")
    colIndex = FindColumns(pArray, BVarCol)
    
    ' Check if required columns exist
    If colIndex(1) = 0 Or colIndex(2) = 0 Or colIndex(3) = 0 Then
        MsgBox "The Required Columns not found", vbInformation, "Business Variation"
        GoTo CleanUp
    End If
    
    ' Prepare the sorted array with selected columns
    ReDim BVardata(1 To UBound(pArray, 1), 1 To 7)
    rowCount = 0
    
    For i = LBound(pArray) To UBound(pArray)
        If Not IsEmpty(pArray(i, 1)) Then
            rowCount = rowCount + 1
            BVardata(rowCount, 1) = pArray(i, 1)
            BVardata(rowCount, 2) = pArray(i, colIndex(0))
            BVardata(rowCount, 3) = pArray(i, colIndex(1))
            BVardata(rowCount, 4) = pArray(i, colIndex(1) + 1)
            BVardata(rowCount, 5) = pArray(i, colIndex(1) + 2)
            BVardata(rowCount, 6) = pArray(i, colIndex(2))
            
            ' Handle header row separately
            If i = 1 Then
                BVardata(rowCount, 7) = pArray(i, colIndex(3))
            Else
                If IsNumeric(pArray(i, colIndex(3))) Then
                    BVardata(rowCount, 7) = CDbl(pArray(i, colIndex(3))) / 10000000
                Else
                    BVardata(rowCount, 7) = pArray(i, colIndex(3))
                End If
            End If
        End If
    Next i
    
    ' Paste data to the destination sheet
    If rowCount > 0 Then
        With WSBVAR
            ' Clear existing data
            .Rows("3:" & .Rows.Count).Clear
            
            ' Paste data
            Set BVRange = .Range("A2:G" & rowCount + 1)
            BVRange.NumberFormat = "@"
            BVRange.value = BVardata
            
            ' Apply formatting
            .Cells.Font.name = "Trebuchet MS"
            .Columns("F").NumberFormat = "0.00"
            .Range("A1").value = "Daily Variation CR/DR Transaction >= 50 Lacs for : " & rDate
            
            ' Sort data on Scheme Type column
            With .Sort
                .SortFields.Clear
                .SortFields.Add key:=WSBVAR.Range("F2:F" & rowCount + 1), _
                               SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                .SetRange BVRange.offset(0, 1)
                .header = xlYes
                .Apply
            End With
            
            ' Set print titles
            With .PageSetup
                .PrintTitleRows = "$1:$2"
                .PrintTitleColumns = ""
            End With
            
            ' Apply borders and auto-fit columns
            With BVRange
                .Borders.LineStyle = xlContinuous
                .Columns.AutoFit
            End With
        End With
    Else
        MsgBox "No data available to process.", vbInformation, "Business Variation"
    End If
    
CleanUp:
    ' Turn on screen updating and automatic calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Clear memory
    Erase pArray, BVardata
End Sub
    ' Subroutine to format KPI header
Sub FormatHeader(ws As Worksheet)
    With ws.UsedRange.Rows(1)
        .Font.Color = vbYellow
        .Interior.Color = RGB(180, 0, 0)
        .Font.Bold = True
        .Font.name = "Trebuchet MS"
        .Font.Size = 10
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

' Subroutine to format KPI body
Sub FormatBody(ws As Worksheet)
    Dim borderTypes As Variant, i As Long
    borderTypes = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)

    With ws.UsedRange
        .Borders.LineStyle = xlNone
        .Columns.AutoFit
        For i = LBound(borderTypes) To UBound(borderTypes)
            With .Borders(borderTypes(i))
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
        Next i
    End With
End Sub
Sub AddToDict(Dict As Object, ByRef sArray As Variant, ColArray As Variant)
    Dim DictKey As String, DictVal As String
    Dim colIndex() As Long
    Dim lastRow As Long, k As Long

    ' Find columns for Branch Code and Branch Name
    ReDim colIndex(LBound(ColArray) To UBound(ColArray))
    colIndex = FindColumns(sArray, ColArray)
    lastRow = UBound(sArray, 1)

    ' Ensure colIndex(0) is set correctly
    If colIndex(0) = 0 Then colIndex(0) = colIndex(1)

    ' Extract branch codes and names
    For k = 1 To lastRow
        If k = 1 Then
            DictKey = ColArray(0)
            DictVal = ColArray(1)
        Else
            DictKey = UCase(Left(sArray(k, colIndex(0)), 6))
            DictVal = UCase(sArray(k, colIndex(1)))

            ' Adjust branch name if necessary
            If Left(DictVal, 6) = DictKey And Len(DictVal) > 7 Then
                DictVal = Mid(DictVal, 8)
            End If
        End If

        ' Add to dictionary if key does not exist
        If Not Dict.Exists(DictKey) Then
            Dict.Add DictKey, DictVal
        End If
    Next k
End Sub
Sub ProcessFiles(ByRef reportData As Object, dateVal As String, ByVal FilePrefix As String)
        ' Process report data from multiple Excel files
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Declare variables with appropriate types
    Dim wbSource As Workbook
    Dim ws As Worksheet
    Dim rptInfo As Range, headerInfo As Variant
    Dim rptName As Variant, rptMap As Variant
    Dim SourceFolder As String, fileName As String, MissingFile As String
    Dim BranchArray As Variant
    Dim LabelMap As Object
    Dim sheetData() As Variant
    Dim rptIndex As Long, SheetIndex As Long
    Dim RowStart As Long, ColStart As Long, BrCol As Long, BrNcol As Long
    Dim lastCol As Long, firstCol As Long, lastRow As Long, headerRow As Long
    Dim pArray As Variant, colIndex As Variant

    On Error GoTo ErrorHandler

    ' Initialize dictionaries
    If branchDict Is Nothing Then Set branchDict = CreateObject("Scripting.Dictionary")
    Set LabelMap = CreateObject("Scripting.Dictionary")
    Set reportData = CreateObject("Scripting.Dictionary")

    ' Cache KPI mapping for performance
    Dim CommonHeaders As Range, ReportHeaders As Range, Headers As Range
    Set CommonHeaders = wsRefer.Range("CommonHeaders")

    ' Get configuration values
    SourceFolder = wsRefer.Range("RptPath").Value2
    If Right(SourceFolder, 1) <> "\" Then SourceFolder = SourceFolder & "\"

    Set rptInfo = wsRefer.Range(FilePrefix)

    ' Process each report
    For rptIndex = 1 To rptInfo.Rows.Count
        Set rptInfo = wsRefer.Range(FilePrefix).Rows(rptIndex)
        rptName = rptInfo.Columns(1).value
        rptMap = rptInfo.Columns(2).value
        RowStart = rptInfo.Columns(3).value
        ColStart = rptInfo.Columns(4).value
        BrCol = rptInfo.Columns(5).value
        BrNcol = rptInfo.Columns(6).value
        
        fileName = Dir(SourceFolder & CStr(dateVal) & " " & rptName & "*.xl*")
        If fileName = "" Then
            MissingFile = MissingFile & rptMap & ", "
            reportData(rptMap) = Empty
            GoTo NextReport
        End If

        ' Open workbook with optimized settings
        Set wbSource = Workbooks.Open(SourceFolder & fileName, UpdateLinks:=False, ReadOnly:=True)
        ReDim sheetData(1 To wbSource.Sheets.Count)
        SheetIndex = 1
                
        ' Set ReportHeaders if rptMap value is not empty
        Set ReportHeaders = wsRefer.Range(rptMap)
        Set Headers = Application.Union(CommonHeaders, ReportHeaders)
        If Not LabelMap.Exists(rptMap) Then LabelMap.Add rptMap, Headers

        ' Process each worksheet
        For Each ws In wbSource.Sheets
            If RowStart = 0 Then
                sheetData(SheetIndex) = Empty
                SheetIndex = SheetIndex + 1
                GoTo NextSheet
            End If
            
            headerInfo = FindRowCol(ws)
            If RowStart <> headerInfo(1) Then MsgBox " Ah! Header Row not Validated, Header Start at Rows No   : " & headerInfo(1) & vbCrLf & _
                                            " Report Name : " & rptName & " Date of Report : " & dateVal
            
            
            With ws
                ' Optimize range operations with error checking
                If WorksheetFunction.CountA(.Cells) = 0 Then GoTo NextSheet
                lastCol = .Cells(RowStart, .Columns.Count).End(xlToLeft).Column

                ' Validate column settings
                If ColStart > lastCol Or ColStart = 0 Then
                    ColStart = 1
                    If lastCol = 0 Then lastCol = 1
                End If

                lastRow = .Cells(.Rows.Count, ColStart).End(xlUp).row
                ' Skip if no data rows found
                If lastRow <= RowStart Then
                    sheetData(SheetIndex) = Empty
                    SheetIndex = SheetIndex + 1
                    GoTo NextSheet
                End If

                ' Process headers with error handling
                If IsArray(LabelMap(rptMap)) Then
                    ProcessHeaders ws, RowStart, ColStart, lastCol, LabelMap(rptMap)
                End If

                ' Process branch codes
                If BrCol = 0 Then
                    ProcessBranchCodes ws, RowStart, lastRow, lastCol, BrCol, BrNcol
                End If

                ' Clean up branch names if branch columns exist
                If BrNcol <> 0 Then
                  ProcessBranchNames ws, RowStart, lastRow, BrCol, BrNcol
                End If
                ' Get data range in one operation with bounds checking
                If lastRow >= RowStart And lastCol >= ColStart Then
                    pArray = .Range(.Cells(RowStart, ColStart), .Cells(lastRow, lastCol)).Value2
                    ' Update branch dictionary if we have data and valid branch columns
                    If IsArray(pArray) Then
                        AddToDict branchDict, pArray, Array("BRANCH CODE", "BRANCH NAME")
                    End If
                    sheetData(SheetIndex) = pArray
                Else
                    sheetData(SheetIndex) = Empty
                End If
            End With
            
            SheetIndex = SheetIndex + 1
NextSheet:
        Next ws

        ' Store sheetData in reportData dictionary
        reportData(rptMap) = sheetData
        wbSource.Close SaveChanges:=False
        Set wbSource = Nothing

NextReport:
    Next rptIndex

CleanUp:
    ' Clean up objects to prevent memory leaks
    Set LabelMap = Nothing
    Set CommonHeaders = Nothing
    Set ReportHeaders = Nothing
    Set Headers = Nothing

    ' Restore Excel settings
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    ' Report missing files with better formatting
    If Len(MissingFile) > 1 Then
        MissingFile = Left(MissingFile, Len(MissingFile) - 2) ' Remove trailing comma and space
        MsgBox "Missing Files for Date: " & dateVal & vbCrLf & MissingFile, vbInformation, "Missing Files"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Error in ProcData: " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Report: " & rptMap & ", Sheet: " & IIf(Not ws Is Nothing, ws.name, "Unknown"), _
           vbExclamation, "Error"
    Resume CleanUp
End Sub
Private Sub ProcessHeaders(ws As Worksheet, headerRow As Long, firstCol As Long, lastCol As Long, renArray As Range)
    Dim Col As Long, cell As Variant, area As Range
    Dim FirstColumn As Range

    ' Loop through each area in the renArray
    For Each area In renArray.Areas
        ' Set FirstColumn to the first column of the current area
        Set FirstColumn = area.Columns(1).Cells

        ' Debugging statements
        Debug.Print "Processing Area: " & area.Address
        Debug.Print "FirstColumn Count: " & FirstColumn.Address
        Debug.Print "FirstColumn Count: " & FirstColumn.Count

        ' Loop through each column in the header row
        For Col = firstCol To lastCol
            Dim headervalue As String
            headervalue = UCase(ws.Cells(headerRow, Col).value)
            Debug.Print "Header Value: " & headervalue

            ' Loop through each cell in the FirstColumn of the current area
            For Each cell In FirstColumn
                If UCase(cell.value) = headervalue Then
                    Debug.Print "Cell Value: " & UCase(cell.value)
                    Debug.Print "Match Found: " & cell.offset(0, 1).value
                    ws.Cells(headerRow, Col).value = cell.offset(0, 1).value
                    Exit For
                End If
            Next cell
        Next Col
    Next area
End Sub
' New helper function for branch code processing
Private Sub ProcessBranchCodes(ByRef ws As Worksheet, headerRow As Long, lastRow As Long, ByRef lastCol As Long, ByRef BrC As Long, BrN As Long)
    Dim f As Long: f = 1
    lastCol = lastCol + 1
    BrC = lastCol
    
    With ws
        Dim formulaRange As Range
        Dim formulaValue As String
        .Cells(headerRow, lastCol) = "BRANCH CODE"
        ' Define the range where the formula will be applied
        Set formulaRange = .Range(.Cells(headerRow + 1, lastCol), .Cells(lastRow, lastCol))
        formulaRange.NumberFormat = "@"
        
        For f = headerRow + 1 To lastRow
            .Cells(f, lastCol) = Left(.Cells(f, BrN), 6)
        Next f
    End With
End Sub

' New helper function for cleaning branch names
Private Sub ProcessBranchNames(ByRef ws As Worksheet, headerRow As Long, lastRow As Long, BrC As Long, BrN As Long)
    Dim f As Long
    For f = headerRow + 1 To lastRow
        With ws
            If Left(.Cells(f, BrN), 6) = .Cells(f, BrC) _
            And Len(.Cells(f, BrN)) > 7 Then
                .Cells(f, BrN) = Mid(.Cells(f, BrN), 8)
            End If
        End With
    Next f
End Sub

' Process KPI data from file
Private Sub ProcessKPIData(ByVal FilePath As String, ByRef KPIdb As Variant, ByRef KPI_Dict As Object, _
                                                ByRef KPIColArray As Variant, ByRef KPIndex() As Long, MissingCols As String)
    Dim wbKPI As Workbook, wsKPI As Worksheet
    Dim KPIRows As Long, KPICols As Long
    Dim j As Long

    ' Open KPI file and get data
    Set wbKPI = Workbooks.Open(FilePath, ReadOnly:=True)
    Set wsKPI = wbKPI.Sheets("KPI")

    ' Get KPI data in one operation
    With wsKPI
        KPIRows = .Cells(.Rows.Count, "A").End(xlUp).row
        KPICols = .Cells(1, .Columns.Count).End(xlToLeft).Column
        KPIdb = .Range("A1").Resize(KPIRows, KPICols).value
    End With
    wbKPI.Close SaveChanges:=False
    
        Call ClearCaches
    ' Find column indexes for KPI data
    KPIndex = FindColumns(KPIdb, KPIColArray, MissingCols)


    ' Create dictionary for quick lookup
    Set KPI_Dict = CreateObject("Scripting.Dictionary")

    ' Populate dictionary with branch codes as keys
    For j = 2 To KPIRows  ' Start at 2 to skip header
        If Not KPI_Dict.Exists(CStr(KPIdb(j, 2))) Then
            KPI_Dict.Add CStr(KPIdb(j, 2)), Application.Index(KPIdb, j, 0)
        End If
    Next j
End Sub
Sub KPIpopulate(ByRef pArray As Variant, wsKPI As Worksheet, dateVal As Variant, Optional DataTransform As Boolean = False)
    Dim colIndex() As Long, MissingCols As String
    Dim isSMEAD1 As Boolean
    Dim paramRows As Variant, paramCols As Variant
    Dim KPIRows As Long, KPICols As Long
    Dim i As Long, j As Long, k As Long, n As Long, p As Long, LookupValue As String
    ' Define and locate required column headers
    ClearCaches
    Dim ColArray As Variant
    ColArray = Array("BRANCH CODE", "CATEGORY", "Tot Nos")
    ReDim colIndex(LBound(ColArray) To UBound(ColArray))
    colIndex = FindColumns(pArray, ColArray, MissingCols)
    ' Check for SMEAD1 by the presence of the CATEGORY column
    isSMEAD1 = (colIndex(1) > 0)
    ' Get dimensions of KPI sheet
    KPIRows = wsKPI.Cells(wsKPI.Rows.Count, 1).End(xlUp).row
    KPICols = wsKPI.Cells(1, wsKPI.Columns.Count).End(xlToLeft).Column
    
    If isSMEAD1 Then
        SMEAD1_RUN pArray
    End If
        ' Non-SMEAD1 logic (populate numeric columns based on BRANCH CODE match)
        Dim NumericCols() As Long, numColsCount As Long
        numColsCount = 0
        For i = 1 To UBound(pArray, 2)
            If i <> colIndex(0) And IsNumeric(pArray(2, i)) Then
                numColsCount = numColsCount + 1
                ReDim Preserve NumericCols(1 To numColsCount)
                NumericCols(numColsCount) = i
            End If
        Next i
        For i = 1 To KPIRows - 1
            LookupValue = UCase(wsKPI.Cells(i, 2).value)
            Dim FoundIndex As Long: FoundIndex = 0
            ' Locate matching row in pArray for the branch code
            For k = 1 To UBound(pArray, 1)
                If UCase(pArray(k, colIndex(0))) = LookupValue Then
                    FoundIndex = k
                    Exit For
                End If
            Next k
            ' Populate numeric columns
            If FoundIndex > 0 Then
                For k = LBound(NumericCols) To UBound(NumericCols)
                    Dim valueToAssign As Variant: valueToAssign = pArray(FoundIndex, NumericCols(k))
                    If DataTransform And IsNumeric(valueToAssign) Then
                        valueToAssign = valueToAssign / 100000
                    End If
                    wsKPI.Cells(i, KPICols + k).value = valueToAssign
                Next k
            Else
                ' Clear cells if no match is found
                For k = LBound(NumericCols) To UBound(NumericCols)
                    wsKPI.Cells(i, KPICols + k).value = ""
                Next k
            End If
        Next i

        ' === ADDING THE TOTAL ROW FOR BRANCH CODE "C223" ===
        ' Sum all numeric columns and insert into the total row
        For k = LBound(NumericCols) To UBound(NumericCols)
            Dim sumRange As Range
            Set sumRange = wsKPI.Range(wsKPI.Cells(2, KPICols + k), wsKPI.Cells(KPIRows - 1, KPICols + k)) ' Range to sum
            wsKPI.Cells(KPIRows, KPICols + k).value = Application.WorksheetFunction.Sum(sumRange)
        Next k
End Sub
Sub ProcessRRB(ByRef pArray As Variant, wsKPI As Worksheet, dateVal As Variant)
 
    Dim SchemeTypeCol As Long, BalCol As Long, BrCodeCol As Long, BrR As Long
    Dim branchTotals As Object, BranchCode As String, KPICols As Long, i As Long
    Dim currentTotals As Variant, KPIRows As Long
    Dim schemeType As Variant, balance As Double
 
    ' Check if pArray has data
    If Not IsEmpty(pArray) Then
        ClearCaches
 
        ' Find columns in the data array
        BrCodeCol = FindColumn(pArray, "BRANCH CODE")
        SchemeTypeCol = FindColumn(pArray, "SCHEME TYPE")
        BalCol = FindColumn(pArray, "BALANCE")
 
        ' Find the number of columns in the KPI sheet
        KPICols = wsKPI.Cells(1, wsKPI.Columns.Count).End(xlToLeft).Column
 
        ' Initialize dictionary for branch totals
        Set branchTotals = CreateObject("Scripting.Dictionary")
 
        ' Loop through the data array
        For i = 1 To UBound(pArray, 1)
            BranchCode = UCase(CStr(pArray(i, BrCodeCol)))
            schemeType = CStr(pArray(i, SchemeTypeCol))
           
            ' Skip headers or invalid balances
            If UCase(pArray(i, BalCol)) <> "BALANCE" Then
                balance = pArray(i, BalCol) / 100000
            Else
                balance = 0
            End If
 
            ' Initialize totals if branch code and scheme type are not in the dictionary
            If Not branchTotals.Exists(BranchCode) Then
                branchTotals.Add BranchCode, CreateObject("Scripting.Dictionary")
            End If
 
            If Not branchTotals(BranchCode).Exists(schemeType) Then
                branchTotals(BranchCode).Add schemeType, 0
            End If
 
            ' Update totals based on SchemeType and Balance
            Select Case schemeType
                Case "SBA"
                    If balance > 0 Then
                        branchTotals(BranchCode)("SBA") = branchTotals(BranchCode)("SBA") + balance
                    End If
                Case "TDA", "SCHEME TYPE"
                    ' Skip these Scheme Types
                Case Else
                    If balance > 0 Then
                        If Not branchTotals(BranchCode).Exists("RRBCA") Then
                            branchTotals(BranchCode).Add "RRBCA", 0
                        End If
                        branchTotals(BranchCode)("RRBCA") = branchTotals(BranchCode)("RRBCA") + balance
                    ElseIf balance < 0 Then
                        If Not branchTotals(BranchCode).Exists("RRBADV") Then
                            branchTotals(BranchCode).Add "RRBADV", 0
                        End If
                        branchTotals(BranchCode)("RRBADV") = branchTotals(BranchCode)("RRBADV") - balance
                    End If
            End Select
        Next i
 
        ' Update the wsKPI sheet with the calculated totals
        KPIRows = wsKPI.Cells(wsKPI.Rows.Count, 2).End(xlUp).row
 
        For i = 2 To KPIRows
            BranchCode = UCase(CStr(wsKPI.Cells(i, 2)))
            If branchTotals.Exists(BranchCode) Then
                If branchTotals(BranchCode).Exists("RRBCA") Then
                    wsKPI.Cells(i, KPICols + 1) = branchTotals(BranchCode)("RRBCA")
                End If
                If branchTotals(BranchCode).Exists("RRBADV") Then
                    wsKPI.Cells(i, KPICols + 2) = branchTotals(BranchCode)("RRBADV")
                End If
                If branchTotals(BranchCode).Exists("RRBSB") Then
                    wsKPI.Cells(i, KPICols + 2) = branchTotals(BranchCode)("RRBSB")
                End If
            End If
        Next i
 
        ' Add headers for new columns
        wsKPI.Cells(1, KPICols + 1) = "RRB CA"
        wsKPI.Cells(1, KPICols + 2) = "RRB ADV"
        wsKPI.Cells(1, KPICols + 3) = "RRB SB"
        ' Clear pArray after processing
        Erase pArray
    Else
        MsgString = "RRB data Empty for Date " & dateVal & vbCrLf & MsgString
    End If
End Sub
Sub ShowMsgBox(MsgString)
    Dim frm As UserForm1
    Set frm = New UserForm1

    ' Customize the UserForm's appearance
    With frm
        .Caption = "Daily Business Figure"
        .LabelMsg.Caption = MsgString
        .LabelMsg.Font.Size = 14
    End With

    frm.Show
End Sub
Sub FileOpen(FilePath As String, ByRef wb As Workbook, Optional ByRef eFlag As Boolean = False)
        Dim LocalFilePath As String
 
    On Error Resume Next
        Set wb = Workbooks.Open(FilePath)
    On Error GoTo 0
        If wb Is Nothing Then
            MsgBox "File does not exist: " & FilePath, vbInformation, "Warning"
            eFlag = True
            Exit Sub
        End If
    End Sub
Public Sub OptimizeExcel(ByVal optimize As Boolean)
    With Application
        .ScreenUpdating = Not optimize
        .EnableEvents = Not optimize
        .Calculation = IIf(optimize, xlCalculationManual, xlCalculationAutomatic)
    End With
End Sub


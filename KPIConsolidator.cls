Option Explicit

' KPIConsolidator.cls - Optimized and Revamped
' Dependencies: Microsoft Scripting Runtime
' Assumes k_config and wsRef set externally

Private Const MODULE_NAME As String = "KPIConsolidator"
Private Const DICTIONARY_COMPARE_MODE As Integer = vbTextCompare
Private Const MAX_BRANCH_CODE_LENGTH As Integer = 6
Private Const MIN_BRANCH_CODE_LENGTH As Integer = 6

' Enhanced types for better structure
Private Type THeaderInfo
    Row As Long
    BranchCodeCol As Long
    LastRow As Long
    LastCol As Long
End Type

Private Type TFileInfo
    Pattern As String
    LabelMap As String
    Processed As Boolean
End Type

Private Type TProcessingStats
    FilesProcessed As Long
    RecordsProcessed As Long
    ErrorCount As Long
    StartTime As Date
    EndTime As Date
End Type

' Internal state with improved organization
Private m_Headers As Scripting.Dictionary
Private m_Data As Scripting.Dictionary
Private m_CommonHeaders As Scripting.Dictionary
Private m_LabelMaps As Scripting.Dictionary
Private m_FilePatterns() As TFileInfo
Private m_ProcessedFiles As Scripting.Dictionary
Private m_ProcessingStats As TProcessingStats
Private m_ErrorLog As String

' Performance optimization flags
Private m_ScreenUpdating As Boolean
Private m_Calculation As XlCalculation
Private m_EnableEvents As Boolean

' --- Class Lifecycle ---

Private Sub Class_Initialize()
    InitializeDictionaries
    InitializeFilePatterns
    InitializePerformanceSettings
    ResetProcessingStats
End Sub

Private Sub Class_Terminate()
    CleanupResources
    RestorePerformanceSettings
End Sub

' --- Public Properties ---

Public Property Get ErrorLog() As String
    ErrorLog = m_ErrorLog
End Property

Public Property Get DataDict() As Object
    Set DataDict = m_Data
End Property

Public Property Get HasFiles() As Boolean
    HasFiles = m_ProcessingStats.FilesProcessed > 0
End Property

Public Property Get ProcessingStats() As String
    With m_ProcessingStats
        ProcessingStats = "Files: " & .FilesProcessed & _
                         ", Records: " & .RecordsProcessed & _
                         ", Errors: " & .ErrorCount & _
                         ", Duration: " & Format(.EndTime - .StartTime, "hh:nn:ss")
    End With
End Property

' --- Main Public Methods ---

Public Sub AddData(ByVal Key As String, ByVal Value As Variant)
    If ValidateKey(Key) Then
        m_Data.Add Key, Value
    End If
End Sub

Public Sub LoadHeaders(ByVal LabelMap As String)
    On Error GoTo ErrHandler
    
    ValidateParameter LabelMap, "LabelMap"
    
    ' Load common headers if not already loaded
    If m_CommonHeaders.Count = 0 Then
        LoadCommonHeaders
    End If
    
    ' Load specific label mappings
    LoadLabelMappings LabelMap
    
    Exit Sub
ErrHandler:
    LogError "LoadHeaders", Err.Number, Err.Description
End Sub

Public Sub ProcessFiles(ByVal FilePath As String, ByVal DateVals As Variant)
    On Error GoTo ErrHandler
    
    m_ProcessingStats.StartTime = Now
    
    ' Validate inputs
    ValidateProcessFilesInputs FilePath, DateVals
    
    ' Optimize Excel performance
    OptimizeExcelPerformance True
    
    ' Initialize processing
    InitializeProcessing
    
    ' Process files for each date
    ProcessDateRange DateVals, FilePath
    
    ' Finalize processing
    FinalizeProcessing
    
    Exit Sub
ErrHandler:
    LogError "ProcessFiles", Err.Number, Err.Description
    m_ProcessingStats.ErrorCount = m_ProcessingStats.ErrorCount + 1
Finally:
    OptimizeExcelPerformance False
    m_ProcessingStats.EndTime = Now
End Sub

Public Sub ClearData()
    ClearDictionary m_Data
    ClearDictionary m_Headers
    Set m_ProcessedFiles = NewDict()
    ResetProcessingStats
End Sub

' --- Private Initialization Methods ---

Private Sub InitializeDictionaries()
    Set m_Headers = NewDict()
    Set m_Data = NewDict()
    Set m_CommonHeaders = NewDict()
    Set m_LabelMaps = NewDict()
    Set m_ProcessedFiles = NewDict()
End Sub

Private Sub InitializeFilePatterns()
    Dim patterns As Variant
    Dim i As Long
    
    patterns = Split(k_config.FILE_PATTERNS, ",")
    ReDim m_FilePatterns(LBound(patterns) To UBound(patterns))
    
    For i = LBound(patterns) To UBound(patterns)
        With m_FilePatterns(i)
            .Pattern = "*" & Trim$(patterns(i)) & "*.xlsx"
            .LabelMap = ""
            .Processed = False
        End With
    Next i
End Sub

Private Sub InitializePerformanceSettings()
    With Application
        m_ScreenUpdating = .ScreenUpdating
        m_Calculation = .Calculation
        m_EnableEvents = .EnableEvents
    End With
End Sub

Private Sub ResetProcessingStats()
    With m_ProcessingStats
        .FilesProcessed = 0
        .RecordsProcessed = 0
        .ErrorCount = 0
        .StartTime = 0
        .EndTime = 0
    End With
End Sub

' --- Private Processing Methods ---

Private Sub ProcessDateRange(ByVal DateVals As Variant, ByVal FilePath As String)
    Dim processedDates As Scripting.Dictionary
    Set processedDates = NewDict()
    
    Dim dateVal As Variant
    For Each dateVal In DateVals
        If IsValidDate(dateVal) Then
            Dim dateKey As String
            dateKey = CStr(dateVal)
            
            If Not processedDates.Exists(dateKey) Then
                processedDates.Add dateKey, True
                ProcessSingleDate dateKey, FilePath
            End If
        End If
    Next dateVal
End Sub

Private Sub ProcessSingleDate(ByVal DateKey As String, ByVal FilePath As String)
    ' Try KPI file first
    If ProcessKPIFile(DateKey) Then
        Exit Sub
    End If
    
    ' Process pattern-based files
    ProcessPatternFiles DateKey, FilePath
    
    ' Calculate totals for this date
    CalculateTotals DateKey
End Sub

Private Function ProcessKPIFile(ByVal DateKey As String) As Boolean
    Dim kpiFile As String
    kpiFile = k_config.FILE_KPI_PATH & DateKey & "eKPI.xlsx"
    
    If FileExists(kpiFile) And Not m_ProcessedFiles.Exists(kpiFile) Then
        m_ProcessedFiles.Add kpiFile, True
        LoadKPIFileFromWB kpiFile
        m_ProcessingStats.FilesProcessed = m_ProcessingStats.FilesProcessed + 1
        ProcessKPIFile = True
    End If
End Function

Private Sub ProcessPatternFiles(ByVal DateKey As String, ByVal FilePath As String)
    Dim i As Long
    For i = LBound(m_FilePatterns) To UBound(m_FilePatterns)
        ProcessPatternFile DateKey, FilePath, i
    Next i
End Sub

Private Sub ProcessPatternFile(ByVal DateKey As String, ByVal FilePath As String, ByVal PatternIndex As Long)
    On Error GoTo ErrHandler
    
    With m_FilePatterns(PatternIndex)
        ' Load headers for this pattern
        Dim LabelMap As String
        LabelMap = UCase$(k_config.FILE_INFO.Cells(PatternIndex + 1, 2).Value2)
        LoadHeaders LabelMap
        
        ' Find and process matching files
        Dim filePattern As String
        filePattern = DateKey & " " & .Pattern
        
        Dim fileName As String
        fileName = Dir(FilePath & filePattern)
        
        Do While Len(fileName) > 0
            Dim fileKey As String
            fileKey = DateKey & "_" & .Pattern & "_" & fileName
            
            If Not m_ProcessedFiles.Exists(fileKey) Then
                m_ProcessedFiles.Add fileKey, True
                ProcessWorkbookFromFile FilePath & fileName, DateKey, Mid(.Pattern, 2, 9)
                m_ProcessingStats.FilesProcessed = m_ProcessingStats.FilesProcessed + 1
            End If
            
            fileName = Dir
        Loop
    End With
    
    Exit Sub
ErrHandler:
    LogError "ProcessPatternFile", Err.Number, Err.Description
End Sub

Private Sub LoadKPIFileFromWB(ByVal KpiFile As String)
    On Error GoTo ErrHandler
    
    Dim wb As Workbook
    Set wb = Workbooks.Open(KpiFile, ReadOnly:=True, UpdateLinks:=False)
    
    If WorksheetExists(wb, "KPI") Then
        LoadKPIFile wb.Sheets("KPI").UsedRange.Value2, wb.Name
    End If
    
    wb.Close SaveChanges:=False
    
    Exit Sub
ErrHandler:
    LogError "LoadKPIFileFromWB", Err.Number, Err.Description
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

Private Sub ProcessWorkbookFromFile(ByVal FullPath As String, ByVal DateKey As String, ByVal FileIde As String)
    On Error GoTo ErrHandler
    
    Dim wb As Workbook
    Set wb = Workbooks.Open(FullPath, ReadOnly:=True, UpdateLinks:=False)
    
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If IsValidWorksheet(ws) Then
            ProcessWorksheetData ws, DateKey
        End If
    Next ws
    
    wb.Close SaveChanges:=False
    
    Exit Sub
ErrHandler:
    LogError "ProcessWorkbookFromFile", Err.Number, Err.Description
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

Private Sub ProcessWorksheetData(ByVal ws As Worksheet, ByVal DateKey As String)
    On Error GoTo ErrHandler
    
    Dim headerInfo As THeaderInfo
    If Not GetHeaderInfo(ws, headerInfo) Then Exit Sub
    
    Dim dataArray As Variant
    dataArray = GetWorksheetData(ws, headerInfo)
    
    If Not IsEmpty(dataArray) Then
        ProcessDataArray dataArray, DateKey, ws.Parent.Name
        m_ProcessingStats.RecordsProcessed = m_ProcessingStats.RecordsProcessed + UBound(dataArray, 1) - 1
    End If
    
    Exit Sub
ErrHandler:
    LogError "ProcessWorksheetData", Err.Number, Err.Description
End Sub

Private Function GetWorksheetData(ByVal ws As Worksheet, ByVal headerInfo As THeaderInfo) As Variant
    With ws
        If headerInfo.LastRow > headerInfo.Row Then
            GetWorksheetData = .Range(.Cells(headerInfo.Row, headerInfo.BranchCodeCol), _
                                    .Cells(headerInfo.LastRow, headerInfo.LastCol)).Value2
        End If
    End With
End Function

Private Sub ProcessDataArray(ByRef DataArray As Variant, ByVal DateVal As String, ByVal WbName As String)
    On Error GoTo ErrHandler
    
    Dim divisor As Double
    divisor = GetDivisor(WbName)
    
    Dim headerIndices As Scripting.Dictionary
    Set headerIndices = BuildKPILookup(DataArray)
    
    Dim branchDict As Scripting.Dictionary
    Set branchDict = GetOrCreateBranchDict(DateVal)
    
    ProcessBranchData DataArray, branchDict, headerIndices, divisor
    
    Exit Sub
ErrHandler:
    LogError "ProcessDataArray", Err.Number, Err.Description
End Sub

Private Sub ProcessBranchData(ByRef DataArray As Variant, ByVal BranchDict As Scripting.Dictionary, _
                             ByVal HeaderIndices As Scripting.Dictionary, ByVal Divisor As Double)
    Dim i As Long, j As Long
    Dim branchCode As String
    Dim branchData As Scripting.Dictionary
    
    For i = 2 To UBound(DataArray, 1)
        branchCode = GetNormalizedBranchCode(DataArray(i, 1))
        
        If IsValidBranchCode(branchCode) Then
            Set branchData = GetOrCreateBranchData(BranchDict, branchCode)
            ProcessBranchKPIs DataArray, i, branchData, HeaderIndices, Divisor
        End If
    Next i
End Sub

Private Sub ProcessBranchKPIs(ByRef DataArray As Variant, ByVal RowIndex As Long, _
                             ByVal BranchData As Scripting.Dictionary, _
                             ByVal HeaderIndices As Scripting.Dictionary, ByVal Divisor As Double)
    Dim j As Long
    Dim key As String
    Dim value As Double
    
    For j = 1 To UBound(DataArray, 2)
        If j <> 1 Then ' Skip branch code column
            key = NormalizeHeader(DataArray(1, j))
            
            If IsValidKPIKey(key) Then
                If m_LabelMaps.Exists(key) Then key = m_LabelMaps(key)
                
                value = GetNumericValue(Nz(DataArray(RowIndex, j), 0)) / Divisor
                
                If value <> 0 Then
                    BranchData(key) = value
                    If Not m_Headers.Exists(key) Then m_Headers.Add key, True
                End If
            End If
        End If
    Next j
End Sub

' --- Private Helper Methods ---

Private Function NewDict() As Scripting.Dictionary
    Set NewDict = New Scripting.Dictionary
    NewDict.CompareMode = DICTIONARY_COMPARE_MODE
End Function

Private Function IsValidDate(ByVal Value As Variant) As Boolean
    IsValidDate = Not IsEmpty(Value) And LenB(CStr(Value)) > 0
End Function

Private Function IsValidBranchCode(ByVal BranchCode As String) As Boolean
    IsValidBranchCode = Len(BranchCode) >= MIN_BRANCH_CODE_LENGTH
End Function

Private Function IsValidKPIKey(ByVal Key As String) As Boolean
    IsValidKPIKey = Len(Key) > 0 And Not m_CommonHeaders.Exists(Key)
End Function

Private Function IsValidWorksheet(ByVal ws As Worksheet) As Boolean
    IsValidWorksheet = ws.Visible = xlSheetVisible And ws.UsedRange.Rows.Count > 1
End Function

Private Function GetNormalizedBranchCode(ByVal Value As Variant) As String
    Dim normalized As String
    normalized = NormalizeHeader(Value)
    If Len(normalized) >= MAX_BRANCH_CODE_LENGTH Then
        GetNormalizedBranchCode = Left$(normalized, MAX_BRANCH_CODE_LENGTH)
    Else
        GetNormalizedBranchCode = normalized
    End If
End Function

Private Function GetOrCreateBranchDict(ByVal DateVal As String) As Scripting.Dictionary
    If Not m_Data.Exists(DateVal) Then
        Set GetOrCreateBranchDict = NewDict()
        Set m_Data(DateVal) = GetOrCreateBranchDict
    Else
        Set GetOrCreateBranchDict = m_Data(DateVal)
    End If
End Function

Private Function GetOrCreateBranchData(ByVal BranchDict As Scripting.Dictionary, _
                                      ByVal BranchCode As String) As Scripting.Dictionary
    If Not BranchDict.Exists(BranchCode) Then
        Set GetOrCreateBranchData = NewDict()
        BranchDict.Add BranchCode, GetOrCreateBranchData
    Else
        Set GetOrCreateBranchData = BranchDict(BranchCode)
    End If
End Function

Private Function GetHeaderInfo(ByVal ws As Worksheet, ByRef headerInfo As THeaderInfo) As Boolean
    On Error GoTo ErrHandler
    
    With ws
        ' Find header row and branch code column
        headerInfo.Row = FindHeaderRow(ws)
        headerInfo.BranchCodeCol = FindBranchCodeColumn(ws, headerInfo.Row)
        
        If headerInfo.Row > 0 And headerInfo.BranchCodeCol > 0 Then
            headerInfo.LastRow = .Cells(.Rows.Count, headerInfo.BranchCodeCol).End(xlUp).Row
            headerInfo.LastCol = .Cells(headerInfo.Row, .Columns.Count).End(xlToLeft).Column
            GetHeaderInfo = headerInfo.LastRow > headerInfo.Row
        End If
    End With
    
    Exit Function
ErrHandler:
    GetHeaderInfo = False
End Function

Private Function FindHeaderRow(ByVal ws As Worksheet) As Long
    ' Implementation depends on your specific header detection logic
    ' This is a simplified version
    FindHeaderRow = 1
End Function

Private Function FindBranchCodeColumn(ByVal ws As Worksheet, ByVal HeaderRow As Long) As Long
    ' Implementation depends on your specific branch code detection logic
    ' This is a simplified version
    FindBranchCodeColumn = 1
End Function

Private Function HasValidDataStructure(ByVal ws As Worksheet, ByRef headerInfo As THeaderInfo) As Boolean
    HasValidDataStructure = GetHeaderInfo(ws, headerInfo)
End Function

Private Function WorksheetExists(ByVal wb As Workbook, ByVal SheetName As String) As Boolean
    On Error GoTo ErrHandler
    WorksheetExists = Not wb.Sheets(SheetName) Is Nothing
    Exit Function
ErrHandler:
    WorksheetExists = False
End Function

Private Function FileExists(ByVal FilePath As String) As Boolean
    FileExists = Len(Dir(FilePath)) > 0
End Function

Private Sub LoadCommonHeaders()
    On Error GoTo ErrHandler
    
    Dim arr As Variant
    arr = wsRef.Range("CommonHeaders").Value2
    
    Dim i As Long
    For i = 1 To UBound(arr, 1)
        AddIfValid m_CommonHeaders, arr(i, 1), arr(i, 2)
    Next i
    
    Exit Sub
ErrHandler:
    LogError "LoadCommonHeaders", Err.Number, Err.Description
End Sub

Private Sub LoadLabelMappings(ByVal LabelMap As String)
    On Error GoTo ErrHandler
    
    m_LabelMaps.RemoveAll
    
    Dim arr As Variant
    arr = wsRef.Range(LabelMap).Value2
    
    Dim i As Long
    For i = 1 To UBound(arr, 1)
        AddIfValid m_LabelMaps, arr(i, 1), arr(i, 2)
    Next i
    
    Exit Sub
ErrHandler:
    LogError "LoadLabelMappings", Err.Number, Err.Description
End Sub

Private Sub CalculateTotals(ByVal DateVal As String)
    On Error GoTo ErrHandler
    
    If Not m_Data.Exists(DateVal) Then Exit Sub
    
    Dim branchDict As Scripting.Dictionary
    Set branchDict = m_Data(DateVal)
    
    Dim circleDict As Scripting.Dictionary
    Set circleDict = NewDict()
    
    CalculateBranchTotals branchDict, circleDict
    
    If circleDict.Count > 0 Then
        Set m_Data(DateVal)("CIRCLE") = circleDict
    End If
    
    Exit Sub
ErrHandler:
    LogError "CalculateTotals", Err.Number, Err.Description
End Sub

Private Sub CalculateBranchTotals(ByVal BranchDict As Scripting.Dictionary, _
                                 ByVal CircleDict As Scripting.Dictionary)
    Dim branch As Variant
    Dim key As Variant
    
    For Each branch In BranchDict.Keys
        If branch <> "CIRCLE" Then
            Dim kpiDict As Scripting.Dictionary
            Set kpiDict = BranchDict(branch)
            
            For Each key In kpiDict.Keys
                If IsNumeric(kpiDict(key)) And key <> "BRANCH CODE" Then
                    If Not CircleDict.Exists(key) Then CircleDict(key) = 0
                    CircleDict(key) = CircleDict(key) + CDbl(kpiDict(key))
                End If
            Next key
        End If
    Next branch
End Sub

Private Sub LoadKPIFile(ByVal DataArray As Variant, ByVal WbName As String)
    ' Implementation for loading KPI file data
    ' This would be similar to ProcessDataArray but specific to KPI files
    ProcessDataArray DataArray, ExtractDateFromFileName(WbName), WbName
End Sub

Private Function ExtractDateFromFileName(ByVal FileName As String) As String
    ' Extract date from filename - implementation depends on your naming convention
    ' This is a placeholder implementation
    ExtractDateFromFileName = Format(Date, "yyyymmdd")
End Function

' --- Utility Functions ---

Private Function Nz(ByVal Value As Variant, Optional ByVal DefaultValue As Variant = "") As Variant
    If IsNull(Value) Or IsEmpty(Value) Then
        Nz = DefaultValue
    Else
        Nz = Value
    End If
End Function

Private Function NormalizeHeader(ByVal Value As Variant) As String
    If IsEmpty(Value) Or IsNull(Value) Then
        NormalizeHeader = ""
    Else
        NormalizeHeader = UCase$(Trim$(CStr(Value)))
    End If
End Function

Private Function GetNumericValue(ByVal Value As Variant) As Double
    If IsNumeric(Value) Then
        GetNumericValue = CDbl(Value)
    Else
        Dim cleanValue As String
        cleanValue = Replace(Trim$(Nz(Value, "")), ",", "")
        GetNumericValue = Val(cleanValue)
    End If
End Function

Private Function BuildKPILookup(ByRef DataArray As Variant) As Scripting.Dictionary
    Set BuildKPILookup = NewDict()
    
    Dim j As Long
    For j = 1 To UBound(DataArray, 2)
        Dim key As String
        key = UCase$(Nz(DataArray(1, j), ""))
        
        If key <> "" And Not BuildKPILookup.Exists(key) Then
            BuildKPILookup.Add key, j
        End If
    Next j
End Function

Private Function GetDivisor(ByVal WbName As String) As Double
    If InStr(WbName, "PSFID-011") > 0 Or InStr(WbName, "SMEAD-014") > 0 Then
        GetDivisor = 100000
    Else
        GetDivisor = 1
    End If
End Function

' --- Validation Methods ---

Private Sub ValidateParameter(ByVal Value As Variant, ByVal ParameterName As String)
    If IsEmpty(Value) Or IsNull(Value) Or Len(Trim$(CStr(Value))) = 0 Then
        Err.Raise 5, MODULE_NAME, "Invalid parameter: " & ParameterName
    End If
End Sub

Private Sub ValidateProcessFilesInputs(ByVal FilePath As String, ByVal DateVals As Variant)
    If Not IsArray(DateVals) Or IsEmpty(DateVals) Then
        Err.Raise 1001, MODULE_NAME, "DateVals must be a valid array"
    End If
    
    If Len(Dir(FilePath, vbDirectory)) = 0 Then
        Err.Raise 1002, MODULE_NAME, "Invalid file path: " & FilePath
    End If
End Sub

Private Function ValidateKey(ByVal Key As String) As Boolean
    ValidateKey = Len(Trim$(Key)) > 0 And Not m_Data.Exists(Key)
End Function

' --- Performance Optimization ---

Private Sub OptimizeExcelPerformance(ByVal Enable As Boolean)
    If Enable Then
        With Application
            .ScreenUpdating = False
            .Calculation = xlCalculationManual
            .EnableEvents = False
        End With
    Else
        RestorePerformanceSettings
    End If
End Sub

Private Sub RestorePerformanceSettings()
    With Application
        .ScreenUpdating = m_ScreenUpdating
        .Calculation = m_Calculation
        .EnableEvents = m_EnableEvents
    End With
End Sub

' --- Cleanup Methods ---

Private Sub CleanupResources()
    ClearDictionary m_Data
    ClearDictionary m_Headers
    ClearDictionary m_CommonHeaders
    ClearDictionary m_LabelMaps
    ClearDictionary m_ProcessedFiles
    
    Set m_Data = Nothing
    Set m_Headers = Nothing
    Set m_CommonHeaders = Nothing
    Set m_LabelMaps = Nothing
    Set m_ProcessedFiles = Nothing
End Sub

Private Sub ClearDictionary(ByVal Dict As Scripting.Dictionary)
    If Not Dict Is Nothing Then
        Dict.RemoveAll
    End If
End Sub

Private Sub AddIfValid(ByRef Dict As Scripting.Dictionary, ByVal Key As Variant, ByVal Value As Variant)
    Dim normalizedKey As String
    Dim normalizedValue As String
    
    normalizedKey = UCase$(Trim$(Nz(Key, "")))
    normalizedValue = Trim$(Nz(Value, ""))
    
    If LenB(normalizedKey) > 0 And LenB(normalizedValue) > 0 And Not Dict.Exists(normalizedKey) Then
        Dict.Add normalizedKey, normalizedValue
    End If
End Sub

Private Sub InitializeProcessing()
    ClearDictionary m_Data
    ClearDictionary m_Headers
    Set m_ProcessedFiles = NewDict()
    ResetProcessingStats
End Sub

Private Sub FinalizeProcessing()
    ' Any final processing steps
    m_ProcessingStats.EndTime = Now
End Sub

' --- Error Handling ---

Private Sub LogError(ByVal ProcName As String, ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    Dim errorEntry As String
    errorEntry = Format(Now, "yyyy-mm-dd hh:nn:ss") & " [" & MODULE_NAME & "." & ProcName & "] " & _
                 ErrorNumber & ": " & ErrorDescription
    
    m_ErrorLog = m_ErrorLog & vbCrLf & errorEntry
    m_ProcessingStats.ErrorCount = m_ProcessingStats.ErrorCount + 1
    
    ' Optional: Write to debug window for development
    Debug.Print errorEntry
End Sub
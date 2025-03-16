VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KPIConsolidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' KPIConsolidator.cls
Option Explicit
Private Type THeaderInfo
    row As Long
    branchCodeCol As Long
End Type

Private Type TKPIData
    Headers As Scripting.Dictionary
    data As Scripting.Dictionary
    CommonHeaders As Scripting.Dictionary
    LabelMaps As Scripting.Dictionary
    DateProcessor As DateProcessor
End Type

Private Type TFilePatterns
    Patterns() As String
End Type

Private this As TKPIData
Private FilePatterns As TFilePatterns

Private Sub Class_Initialize()
    ' Initialize dictionaries
    Set this.Headers = New Scripting.Dictionary
    Set this.data = New Scripting.Dictionary
    Set this.CommonHeaders = New Scripting.Dictionary
    Set this.LabelMaps = New Scripting.Dictionary
    
    ' Set file patterns to match
    ReDim FilePatterns.Patterns(3)
    FilePatterns.Patterns(0) = "*SMEAD-002*.xlsx"
    FilePatterns.Patterns(1) = "*SMEAD-009*.xlsx"
    FilePatterns.Patterns(2) = "*SMEAD-001*.xlsx"
    FilePatterns.Patterns(3) = "*PSFID-011*.xlsx"
End Sub
Public Sub LoadCommonHeaders()
    Dim headerRange As Range
    Set headerRange = wsRefer.Range("CommonHeaders")
    
    Dim header As Range
    For Each header In headerRange.Columns(1).Cells
        If Not this.CommonHeaders.Exists(header.value) Then
            this.CommonHeaders.Add header.value, header.offset(0, 1).value
        End If
    Next header
End Sub
Public Sub LoadLabelMaps()
    Dim headerRange As Range
    Set headerRange = wsRefer.Range("LabelMap")
    
    Dim header As Range
    For Each header In headerRange.Columns(1).Cells
        If Not this.LabelMaps.Exists(header.value) Then
            this.LabelMaps.Add header.value, header.offset(0, 1).value
        End If
    Next header
End Sub

Public Sub ProcessFiles(ByVal FilePath As String, ByVal dateVal As String)
    ' Complete reset of dictionaries
    Set this.Headers = New Scripting.Dictionary
    Set this.data = New Scripting.Dictionary
    ' Don't recreate CommonHeaders as it's loaded separately
    
    Dim fName As String
    Dim filesProcessed As Long
    
    fName = Dir(FilePath & dateVal & " *.xlsx")
    
    Do While fName <> ""
        If MatchesPattern(fName) Then
            On Error Resume Next
            ProcessWorkbook FilePath & fName
            If Err.Number <> 0 Then
                Debug.Print "Error processing file " & fName & ": " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
            filesProcessed = filesProcessed + 1
        End If
        fName = Dir
    Loop
    
    If filesProcessed = 0 Then
        Err.Raise 1000, "ProcessFiles", "No matching files found"
    End If
End Sub
Private Function MatchesPattern(ByVal fileName As String) As Boolean
    Dim i As Long
    For i = 0 To UBound(FilePatterns.Patterns)
        If fileName Like FilePatterns.Patterns(i) Then
            MatchesPattern = True
            Exit Function
        End If
    Next i
End Function

Private Sub ProcessWorkbook(ByVal FilePath As String)
    Dim wb As Workbook
    Set wb = Workbooks.Open(FilePath, ReadOnly:=True)
    
    On Error GoTo CleanUp
    ProcessWorksheet wb.Sheets(1)
    
CleanUp:
    wb.Close False
    If Err.Number <> 0 Then Err.Raise Err.Number
End Sub
Private Function FindHeaderInfo(ByVal ws As Worksheet, ByVal searchTerms As Variant) As THeaderInfo
    Dim rng As Range
    Dim term As Variant
    
    ' Initialize the return value
    FindHeaderInfo.row = 0
    FindHeaderInfo.branchCodeCol = 0
    
    ' Loop through each search term
    For Each term In searchTerms.Keys
        Set rng = ws.UsedRange.Find(term, LookIn:=xlValues, Lookat:=xlWhole)
        
        If Not rng Is Nothing Then
            FindHeaderInfo.row = rng.row
            FindHeaderInfo.branchCodeCol = rng.Column
            Exit For ' Exit the loop once a match is found
        End If
    Next term
End Function
Private Sub ProcessData(ByRef dataArray As Variant, ByVal PKeycol As Long)
    Dim i As Long, j As Long
    Dim BranchCode As String, HeaderName As String, newHeader As String
    Dim currentFilename As String
    Dim divisor As Double
    
    ' Get current filename to check if it's PSFID-011
    currentFilename = ActiveWorkbook.name
    
    ' Set the divisor based on the filename
    If InStr(currentFilename, "PSFID-011") > 0 Then
        divisor = 100000 ' Convert Rs. to Lakhs
    Else
        divisor = 1 ' No conversion needed
    End If
    
    ' Calculate relative position in array for branch code column
    For i = 2 To UBound(dataArray, 1)
        BranchCode = Left(dataArray(i, PKeycol), 6)
        
        If Not this.data.Exists(BranchCode) Then
            Set this.data(BranchCode) = New Scripting.Dictionary
        End If
        
        For j = 1 To UBound(dataArray, 2)
            HeaderName = Trim(UCase(dataArray(1, j)))
            
            If this.LabelMaps.Exists(HeaderName) Then
                newHeader = this.LabelMaps(HeaderName)
'                newHeader = this.CommonHeaders(HeaderName)
            Else
                newHeader = HeaderName
            End If
            
            If Not this.Headers.Exists(newHeader) Then
                this.Headers(newHeader) = this.Headers.Count + 1
            End If
            
            ' Apply conversion if needed, skipping the branch code and identifier columns
            If IsNumeric(dataArray(i, j)) And j <> PKeycol Then  'And Not (HeaderName Like "*CODE*" Or HeaderName Like "*ID*" Or HeaderName = "BRANCH CODE")
                this.data(BranchCode)(newHeader) = dataArray(i, j) / divisor
            Else
                this.data(BranchCode)(newHeader) = dataArray(i, j)
            End If
        Next j
    Next i
End Sub
Public Sub GenerateReport(ByVal FilePath As String, ByVal dateVal As String)
    Dim wb As Workbook
    Set wb = Workbooks.Add
    
    With wb.Sheets(1)
        .name = "KPI"
        WriteData .Cells, dateVal
        FormatWorksheet wb.Sheets(1)
    End With
    
        wb.SaveAs FilePath & dateVal & "KPI.xlsx"
        wb.Close False
End Sub
Private Sub WriteData(ByVal startCell As Range, ByVal dateVal As String)
    ' Write headers
    startCell.Cells(1, 1).value = "REPORT DATE"
'    startCell.Cells(1, 2).value = "BRANCH CODE"
    
    Dim Col As Long
    Col = 2
    
    ' Keep track of headers and their column positions
    Dim writtenHeaders As New Scripting.Dictionary
    
    ' Write common headers first
    Dim key As Variant
    For Each key In this.CommonHeaders.Keys
        Dim headervalue As String
        headervalue = this.CommonHeaders(key)
        
        If this.Headers.Exists(headervalue) And Not writtenHeaders.Exists(headervalue) Then
            startCell.Cells(1, Col).value = headervalue
            writtenHeaders.Add headervalue, Col
            Col = Col + 1
        End If
    Next key
    
    ' Write remaining headers
    For Each key In this.Headers.Keys
        If Not writtenHeaders.Exists(key) Then
            startCell.Cells(1, Col).value = key
            writtenHeaders.Add key, Col
            Col = Col + 1
        End If
    Next key
    
    ' Write data using the same column mapping from writtenHeaders
    Dim row As Long
    row = 2
    Dim branch As Variant
    
    For Each branch In this.data.Keys
        ' Write date and branch
        startCell.Cells(row, 1).value = dateVal
        startCell.Cells(row, 2).NumberFormat = "@"
'        startCell.Cells(row, 2).value = branch
        
        ' Write all data using column positions from writtenHeaders
        Dim dataKey As Variant
        For Each dataKey In this.data(branch).Keys
            If writtenHeaders.Exists(dataKey) Then
                startCell.Cells(row, writtenHeaders(dataKey)).value = this.data(branch)(dataKey)
            End If
        Next dataKey
        
        row = row + 1
    Next branch
End Sub
Private Function IsCommonHeader(ByVal HeaderName As String) As Boolean
    Dim key As Variant
    For Each key In this.CommonHeaders.Keys
        If this.CommonHeaders(key) = HeaderName Then
            IsCommonHeader = True
            Exit Function
        End If
    Next key
    IsCommonHeader = False
End Function

Private Sub FormatWorksheet(ByVal ws As Worksheet)
    With ws
        .Range(.Cells(1, 1), .Cells(1, .Columns.Count)).Font.Bold = True
        .UsedRange.Borders.LineStyle = xlContinuous
        .Columns.AutoFit
        .Range("A2").Select
        ActiveWindow.FreezePanes = True
    End With
End Sub

Private Sub Class_Terminate()
    Set this.Headers = Nothing
    Set this.data = Nothing
    Set this.CommonHeaders = Nothing
End Sub



Private Sub ProcessWorksheet(ByVal ws As Worksheet)
    Dim headerInfo As THeaderInfo
    Dim currentFilename As String
        
        ' Regular processing for other file types
        headerInfo = FindHeaderInfo(ws, this.CommonHeaders)
        
        If headerInfo.row = 0 Then
            Err.Raise 1001, "ProcessWorksheet", "Header row not found"
            Exit Sub
        End If
        
        With ws
            Dim lastCol As Long, firstCol As Long, lastRow As Long
            lastCol = .Cells(headerInfo.row, .Columns.Count).End(xlToLeft).Column
            firstCol = .Cells(headerInfo.row, headerInfo.branchCodeCol).End(xlToLeft).Column
            lastRow = .Cells(.Rows.Count, firstCol).End(xlUp).row
            
            Dim dataArray As Variant
            dataArray = .Range(.Cells(headerInfo.row, firstCol), .Cells(lastRow, lastCol)).Value2
            headerInfo.branchCodeCol = headerInfo.branchCodeCol - firstCol + 1
        End With
        
    currentFilename = ActiveWorkbook.name
    ' Check if this is SMEAD-001 file that needs row-to-column transformation
    If InStr(currentFilename, "SMEAD-001") > 0 Then
        ' Process using the row-to-column transformation
        ProcessSMEAD1 dataArray, 1 ' Using option 1 as default
        ' Find branch code column in transformed data
    Dim i As Long, branchCodeCol As Long
        For i = 1 To UBound(dataArray, 2)
            If UCase(Trim(dataArray(1, i))) = "BRANCH CODE" Then
                headerInfo.branchCodeCol = i
                Exit For
            End If
        Next i
    End If

'After Array Conversion reset the Branch Code Column Position
ProcessData dataArray, headerInfo.branchCodeCol

End Sub
Sub ProcessSMEAD1(inputData As Variant, Optional npaOption As Integer = 1) 'DSK
    ' Constants for column names and default values
    Const DEFAULT_CATEGORY As String = "TOTAL_ADVANCE"
    Const BRANCH_CODE_FORMAT As String = "000000"
    Const OUTPUT_SHEET_NAME As String = "Business Figure "
    
    Dim outputData As Variant, headerRow As Variant
    Dim categoryDict As Object, branchDict As Object, colMap As Object
    Dim i As Long, j As Long, k As Long, rowCounter As Long, valCol As Long, colIndex As Long
    Dim categoryArray As Variant, branchKeys As Variant, branchKey As String, categoryKey As Variant
    Dim totalCols As Long, categoryCount As Long, valueColsCount As Long
    Dim specificCategory As String, BranchCode As String
    Dim branchCodeCol As Long, totNosCol As Long, totalBalCol As Long, categoryCol As Long
    Dim zoneCol As Long, circleCol As Long, branchNameCol As Long
    Dim requiredCols As Variant, colName As Variant
    Dim wsNew As Worksheet
    
    On Error GoTo ErrorHandler
    
    npaOption = InputBox("How to handle NPA data?" & vbCrLf & vbCrLf & _
                        "1 = Take from TOTAL_ADVANCE only (default)" & vbCrLf & _
                        "2 = Take from each Category (Creates Many Columns)" & vbCrLf & _
                        "3 = Take from a Specific Category of Your Choice", _
                        "NPA Option", 1)
    
    If npaOption < 1 Or npaOption > 3 Then npaOption = 1
    
    ' Initialize required columns
    requiredCols = Array("ZONE NAME", "CIRCLE NAME", "BRANCH NAME", "BRANCH CODE", "CATEGORY", "TOT NOS", "TOTAL BAL")
    
    ' Setup
    Set categoryDict = CreateObject("Scripting.Dictionary")
    Set branchDict = CreateObject("Scripting.Dictionary")
    Set colMap = CreateObject("Scripting.Dictionary")
    headerRow = Application.Index(inputData, 1, 0)
    
    ' Map columns
    i = 1
    For Each colName In headerRow
        If this.LabelMaps.Exists(UCase(headerRow(i))) Then colName = this.LabelMaps(UCase(headerRow(i)))
        colMap(UCase(Trim(colName))) = i
        i = i + 1
    Next colName
    
    ' Find required columns
    For Each colName In requiredCols
        If Not colMap.Exists(colName) Then
            MsgBox "Could not find required column: " & colName, vbCritical
            Exit Sub
        End If
    Next colName

    ' Assign column indices
    zoneCol = colMap("ZONE NAME")
    circleCol = colMap("CIRCLE NAME")
    branchNameCol = colMap("BRANCH NAME")
    branchCodeCol = colMap("BRANCH CODE")
    categoryCol = colMap("CATEGORY")
    totNosCol = colMap("TOT NOS")
    totalBalCol = colMap("TOTAL BAL")
    
    ' Count value columns (everything after total balance column)
    valueColsCount = UBound(headerRow) - totalBalCol
    
    ' Get unique branches and categories
    For i = 2 To UBound(inputData, 1)
        BranchCode = inputData(i, branchCodeCol)
        branchKey = Trim(BranchCode)
        
        If Len(branchKey) > 0 Then
            If Not branchDict.Exists(branchKey) Then
                branchDict.Add branchKey, Array( _
                    inputData(i, zoneCol), _
                    inputData(i, circleCol), _
                    inputData(i, branchNameCol), _
                    branchKey)
            End If
            
            If Not categoryDict.Exists(UCase(Trim(inputData(i, categoryCol)))) Then
                categoryDict.Add UCase(Trim(inputData(i, categoryCol))), True
            End If
        End If
    Next i
    
    categoryCount = categoryDict.Count
    categoryArray = categoryDict.Keys
    
    ' Handle option 3 (specific category)
    If npaOption = 3 Then
        Dim catList As String, selectedIndex As Integer
        catList = ""
        For i = 0 To categoryCount - 1
            catList = catList & i + 1 & ". " & categoryArray(i) & vbCrLf
        Next i
        
        selectedIndex = InputBox("Select the category number to use for NPA data:" & vbCrLf & vbCrLf & _
                                catList, "Select Category", 1)
                                
        If selectedIndex < 1 Or selectedIndex > categoryCount Then
            specificCategory = DEFAULT_CATEGORY
        Else
            specificCategory = categoryArray(selectedIndex - 1)
        End If
    End If
    
    ' Calculate output size
    Select Case npaOption
        Case 1, 3: totalCols = 4 + (categoryCount * 2) + valueColsCount
        Case 2: totalCols = 4 + (categoryCount * (2 + valueColsCount))
    End Select
    
    ' Create output array
    ReDim outputData(1 To branchDict.Count + 1, 1 To totalCols)
    
    ' Set basic headers
    outputData(1, 1) = "ZONE NAME"
    outputData(1, 2) = "CIRCLE NAME"
    outputData(1, 3) = "BRANCH NAME"
    outputData(1, 4) = "BRANCH CODE"
    
    ' Set category headers
    colIndex = 5
    Select Case npaOption
        Case 1: ' Default category only
            For i = 0 To categoryCount - 1
                outputData(1, colIndex) = categoryArray(i) & " Nos"
                outputData(1, colIndex + 1) = categoryArray(i)
                colIndex = colIndex + 2
            Next i
            
            For i = totalBalCol + 1 To UBound(headerRow)
                outputData(1, colIndex) = headerRow(i)
                colIndex = colIndex + 1
            Next i
            
        Case 2: ' Each category with its own NPA
            For i = 0 To categoryCount - 1
                outputData(1, colIndex) = categoryArray(i) & " Nos"
                outputData(1, colIndex + 1) = categoryArray(i)
                colIndex = colIndex + 2
                
                For j = totalBalCol + 1 To UBound(headerRow)
                    outputData(1, colIndex) = categoryArray(i) & " " & headerRow(j)
                    colIndex = colIndex + 1
                Next j
            Next i
            
        Case 3: ' Specific category
            For i = 0 To categoryCount - 1
                outputData(1, colIndex) = categoryArray(i) & " Nos"
                outputData(1, colIndex + 1) = categoryArray(i)
                colIndex = colIndex + 2
            Next i
            
            For i = totalBalCol + 1 To UBound(headerRow)
                outputData(1, colIndex) = headerRow(i) & " (from " & specificCategory & ")"
                colIndex = colIndex + 1
            Next i
    End Select
    
    ' Fill data
    rowCounter = 2
    branchKeys = branchDict.Keys
    
    For i = 0 To branchDict.Count - 1
        branchKey = branchKeys(i)
        Dim branchInfo As Variant
        branchInfo = branchDict(branchKey)
        
        ' Branch info
        outputData(rowCounter, 1) = branchInfo(0) ' Zone Name
        outputData(rowCounter, 2) = branchInfo(1) ' Circle Name
        outputData(rowCounter, 3) = branchInfo(2) ' Branch Name
        outputData(rowCounter, 4) = CStr(branchInfo(3)) ' Branch Code
        
        colIndex = 5
        
        Select Case npaOption
            Case 1: ' Default category only
                For j = 0 To categoryCount - 1
                    categoryKey = categoryArray(j)
                    For k = 2 To UBound(inputData, 1)
                        If Format(inputData(k, branchCodeCol), BRANCH_CODE_FORMAT) = branchKey And _
                           UCase(Trim(inputData(k, categoryCol))) = categoryKey Then
                            
                            outputData(rowCounter, colIndex) = inputData(k, totNosCol)
                            outputData(rowCounter, colIndex + 1) = inputData(k, totalBalCol)
                            Exit For
                        End If
                    Next k
                    colIndex = colIndex + 2
                Next j
                
                For k = 2 To UBound(inputData, 1)
                    If Format(inputData(k, branchCodeCol), BRANCH_CODE_FORMAT) = branchKey And _
                       UCase(Trim(inputData(k, categoryCol))) = UCase(DEFAULT_CATEGORY) Then
                        
                        For j = 1 To valueColsCount
                            outputData(rowCounter, colIndex + j - 1) = inputData(k, totalBalCol + j)
                        Next j
                        Exit For
                    End If
                Next k
                
            Case 2: ' Each category with NPA
                For j = 0 To categoryCount - 1
                    categoryKey = categoryArray(j)
                    For k = 2 To UBound(inputData, 1)
                        If Format(inputData(k, branchCodeCol), BRANCH_CODE_FORMAT) = branchKey And _
                           UCase(Trim(inputData(k, categoryCol))) = categoryKey Then
                            
                            outputData(rowCounter, colIndex) = inputData(k, totNosCol)
                            outputData(rowCounter, colIndex + 1) = inputData(k, totalBalCol)
                            
                            For valCol = 1 To valueColsCount
                                outputData(rowCounter, colIndex + 1 + valCol) = _
                                    inputData(k, totalBalCol + valCol)
                            Next valCol
                            
                            Exit For
                        End If
                    Next k
                    colIndex = colIndex + 2 + valueColsCount
                Next j
                
            Case 3: ' Specific category for NPA
                For j = 0 To categoryCount - 1
                    categoryKey = categoryArray(j)
                    For k = 2 To UBound(inputData, 1)
                        If Format(inputData(k, branchCodeCol), BRANCH_CODE_FORMAT) = branchKey And _
                           UCase(Trim(inputData(k, categoryCol))) = categoryKey Then
                            
                            outputData(rowCounter, colIndex) = inputData(k, totNosCol)
                            outputData(rowCounter, colIndex + 1) = inputData(k, totalBalCol)
                            Exit For
                        End If
                    Next k
                    colIndex = colIndex + 2
                Next j
                
                For k = 2 To UBound(inputData, 1)
                    If Format(inputData(k, branchCodeCol), BRANCH_CODE_FORMAT) = branchKey And _
                       UCase(Trim(inputData(k, categoryCol))) = UCase(specificCategory) Then
                        
                        For j = 1 To valueColsCount
                            outputData(rowCounter, colIndex + j - 1) = inputData(k, totalBalCol + j)
                        Next j
                        Exit For
                    End If
                Next k
        End Select
        
        rowCounter = rowCounter + 1
    Next i
    Erase inputData
    inputData = outputData

    Dim msgText As String
    Select Case npaOption
        Case 1
            msgText = "Data transformation complete! NPA data taken from " & DEFAULT_CATEGORY & " category."
        Case 2
            msgText = "Data transformation complete! NPA data included for each category."
        Case 3
            msgText = "Data transformation complete! NPA data taken from " & specificCategory & " category."
    End Select
    
    MsgBox msgText, vbInformation


    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Application.ScreenUpdating = True
End Sub


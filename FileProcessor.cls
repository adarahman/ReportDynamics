VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' --- FileProcessor.cls ---
Option Explicit

Private Type TFileProcessor
    folderPath As String
    filePattern As String
    ProcessedCount As Long
    DateProcessor As DateProcessor
End Type

Private this As TFileProcessor

Private Const BRANCH_NAME As Long = 6
Private Const ACCOUNT_NUMBER As Long = 7
Private Const ACCOUNT_NAME As Long = 8
Private Const SCHEME_TYPE As Long = 12
Private Const VARIATION As Long = 15

Public Event ProcessingProgress(ByVal fileName As String, ByVal currentCount As Long)

Public Property Get ProcessedCount() As Long
    ProcessedCount = this.ProcessedCount
End Property

Public Function Initialize(ByVal folderPath As String, ByVal startDate As String, ByVal endDate As String) As Boolean
    ' Initialize date processor
    Set this.DateProcessor = New DateProcessor
    If Not this.DateProcessor.Initialize(startDate, endDate) Then
        Initialize = False
        Exit Function
    End If
    
    ' Initialize file processor
    this.folderPath = folderPath
    If Right(this.folderPath, 1) <> "\" Then this.folderPath = this.folderPath & "\"
    
    ' Validate folder path
    If Dir(this.folderPath, vbDirectory) = "" Then
        MsgBox "Folder path does not exist: " & this.folderPath, vbCritical
        Initialize = False
        Exit Function
    End If
    
    this.ProcessedCount = 0
    Initialize = True
End Function

Public Function ProcessFiles(ByRef DataAccumulator As DataAccumulator) As Boolean
    Dim currentDate As Date
    Dim wb As Workbook
    Dim filePattern As String
    Dim fileName As String
    Dim FilePath As String
    
    On Error GoTo ErrorHandler
    
    ' Process each date in range
    currentDate = this.DateProcessor.GetNextDate
    Do While currentDate <> 0
        ' Create file pattern for current date
        filePattern = this.DateProcessor.FormatFileDate(currentDate) & " SMEAD-008*.xlsx"
        fileName = Dir(this.folderPath & filePattern)
        
        ' Process all files for current date
        Do While fileName <> ""
            this.ProcessedCount = this.ProcessedCount + 1
            RaiseEvent ProcessingProgress(fileName, this.ProcessedCount)
            
            FilePath = this.folderPath & fileName
            Set wb = Workbooks.Open(FilePath, ReadOnly:=True)
            ProcessWorkbook wb, DataAccumulator
            wb.Close SaveChanges:=False
            
            fileName = Dir
        Loop
        
        currentDate = this.DateProcessor.GetNextDate
    Loop
    
    If this.ProcessedCount = 0 Then
        MsgBox "No files found in the specified date range.", vbExclamation
        ProcessFiles = False
        Exit Function
    End If
    
    ProcessFiles = True
    Exit Function
    
ErrorHandler:
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    MsgBox "Error processing files: " & Err.Description, vbCritical
    ProcessFiles = False
End Function

Private Sub ProcessWorkbook(ByVal wb As Workbook, ByRef DataAccumulator As DataAccumulator)
    Dim ws As Worksheet
    Dim dataArray As Variant
    Dim lastRow As Long
    Dim i As Long
    Dim value As Double
    
    Set ws = wb.Sheets(1)
    lastRow = ws.Cells(ws.Rows.Count, BRANCH_NAME).End(xlUp).row
    dataArray = ws.Range(ws.Cells(4, 1), ws.Cells(lastRow, VARIATION)).Value2
    
    For i = 2 To UBound(dataArray, 1)
        If IsNumeric(dataArray(i, VARIATION)) Then
            value = IIf(dataArray(i, VARIATION) <> 0, _
                       Format(dataArray(i, VARIATION) / 10000000, "0.00"), _
                       0)
                       
            DataAccumulator.AddValue _
                CStr(dataArray(i, BRANCH_NAME)), _
                CStr(dataArray(i, ACCOUNT_NUMBER)), _
                CStr(dataArray(i, ACCOUNT_NAME)), _
                CStr(dataArray(i, SCHEME_TYPE)), _
                value
        End If
    Next i
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataAccumulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' --- DataAccumulator.cls ---
Option Explicit

Private Type TDataAccumulator
    Dict As Object ' Stores the actual values
    TextDict As Object ' Stores the text representations
End Type

Private this As TDataAccumulator

Private Sub Class_Initialize()
    Set this.Dict = CreateObject("Scripting.Dictionary")
    Set this.TextDict = CreateObject("Scripting.Dictionary")
    this.Dict.CompareMode = TextCompare
    this.TextDict.CompareMode = TextCompare
End Sub

Private Sub Class_Terminate()
    Set this.Dict = Nothing
    Set this.TextDict = Nothing
End Sub

Public Sub AddValue(ByVal branchName As String, ByVal accountNumber As String, _
                   ByVal accountName As String, ByVal schemeType As String, _
                   ByVal value As Double)
    Dim key As String
    key = JoinKey(branchName, accountNumber, accountName, schemeType)
    
    ' Store the formatted text values
    If Not this.TextDict.Exists(key) Then
        Dim textValues As Collection
        Set textValues = New Collection
        textValues.Add CStr(branchName)    ' Branch Name
        textValues.Add CStr(accountNumber) ' Account Number
        textValues.Add CStr(accountName)   ' Account Name
        textValues.Add CStr(schemeType)    ' Scheme Type
        this.TextDict.Add key, textValues
    End If
    
    ' Store the numeric value
    If this.Dict.Exists(key) Then
        this.Dict(key) = this.Dict(key) + value
    Else
        this.Dict.Add key, value
    End If
End Sub

Private Function JoinKey(ParamArray values() As Variant) As String
    JoinKey = Join(values, "|")
End Function

Public Function GetCount() As Long
    GetCount = this.Dict.Count
End Function

Public Sub OutputToWorksheet(ByVal ws As Worksheet)
    Const OUT_SRNO As Long = 1
    Const OUT_BRANCH As Long = 2
    Const OUT_ACCNO As Long = 3
    Const OUT_ACCNAME As Long = 4
    Const OUT_SCHEME As Long = 5
    Const OUT_VAR As Long = 6
    
    Dim outputArray() As Variant
    Dim i As Long
    Dim key As Variant
    Dim textValues As Collection
    
    ' Prepare output array
    ReDim outputArray(1 To this.Dict.Count + 1, 1 To 6)
    
    ' Set headers
    outputArray(1, OUT_SRNO) = "SR NO"
    outputArray(1, OUT_BRANCH) = "BRANCH NAME"
    outputArray(1, OUT_ACCNO) = "ACCOUNT NUMBER"
    outputArray(1, OUT_ACCNAME) = "ACCOUNT NAME"
    outputArray(1, OUT_SCHEME) = "SCHEME TYPE"
    outputArray(1, OUT_VAR) = "VARIATION"
    
    ' Fill data
    i = 2
    For Each key In this.Dict.Keys
        Set textValues = this.TextDict(key)
        
        outputArray(i, OUT_SRNO) = i - 1
        outputArray(i, OUT_BRANCH) = textValues(1)  ' Branch Name
        outputArray(i, OUT_ACCNO) = textValues(2)   ' Account Number
        outputArray(i, OUT_ACCNAME) = textValues(3) ' Account Name
        outputArray(i, OUT_SCHEME) = textValues(4)  ' Scheme Type
        outputArray(i, OUT_VAR) = this.Dict(key)    ' Value
        
        i = i + 1
    Next key
    
    ' Output and format the worksheet
    With ws
        .Cells.Clear
        
        With .Columns("C")
            .NumberFormat = "@"            ' Branch name as text
            .HorizontalAlignment = xlLeft
        End With
        
        ' Output the array
        .Range("A1").Resize(UBound(outputArray, 1), UBound(outputArray, 2)).value = outputArray
        SortRange ws.Range("B1:F" & UBound(outputArray, 1)), 4
        FormatHeader ws
                        
    End With
End Sub
Private Sub SortRange(ByRef dataRange As Range, Col As Long)
    ' Check if the range is valid
    If dataRange Is Nothing Then
        MsgBox "The data range is not valid."
        Exit Sub
    End If
    
    ' Sort data
    With dataRange
        .Sort _
            Key1:=.Columns(Col), _
            Order1:=xlAscending, _
            header:=xlYes, _
            MatchCase:=False, _
            Orientation:=xlTopToBottom
            
        ' Apply formatting
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' --- DateProcessor.cls ---
Option Explicit

Private Type TDateProcessor
    startDate As Date
    endDate As Date
    monthDate As Date
    qtrDate As Date
    yoyDate As Date
    fyDate As Date
    pfyDate As Date
    currentDate As Date
End Type

Private this As TDateProcessor

Public Property Get startDate() As Date
    startDate = this.startDate
End Property

Public Property Get endDate() As Date
    endDate = this.endDate
End Property

Public Property Get currentDate() As Date
    currentDate = this.currentDate
End Property

Public Function Initialize(ByVal startDateStr As String, ByVal endDateStr As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Convert dd.mm.yy to dates
    this.startDate = CDate(Replace(startDateStr, ".", "/"))
    this.endDate = CDate(Replace(endDateStr, ".", "/"))
    this.currentDate = this.startDate
    
    ' Validate date range
    If this.startDate > this.endDate Then
        MsgBox "Start date cannot be later than end date!", vbExclamation
        Exit Function
    End If
    
    Initialize = True
    Exit Function
    
ErrorHandler:
    MsgBox "Invalid date format. Please use dd.mm.yy", vbExclamation
    Initialize = False
End Function

Public Function GetNextDate() As Date
    If this.currentDate <= this.endDate Then
        GetNextDate = this.currentDate
        this.currentDate = DateAdd("d", 1, this.currentDate)
    Else
        GetNextDate = 0
    End If
End Function

Public Function FormatFileDate(ByVal inputDate As Date) As String
    FormatFileDate = Format(inputDate, "dd.mm.yy")
End Function

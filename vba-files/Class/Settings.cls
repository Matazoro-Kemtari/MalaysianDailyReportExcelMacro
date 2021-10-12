VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private canInitialize As Boolean

Private fDailyReportDirectory As String
Private fSummaryDirectory As String
Private fDailyReportFileName As String
Private fSummaryFileName As String

Private Sub Class_Initialize ()
    canInitialize = True
End Sub

' コンストラクタに引数を持たす代用
Public Sub BuildOnce(ByVal DailyReportDirectory As String, _
                     ByVal SummaryDirectory As String, _
                     ByVal DailyReportFileName As String, _
                     ByVal SummaryFileName As String)
    if canInitialize Then
        fDailyReportDirectory = DailyReportDirectory
        fSummaryDirectory = SummaryDirectory
        fDailyReportFileName = DailyReportFileName
        fSummaryFileName = SummaryFileName
        canInitialize = False
    End If
End Sub

Public Property Get DailyReportDirectory() As String
    DailyReportDirectory = fDailyReportDirectory
End Property

Public Property Get SummaryDirectory() As String
    SummaryDirectory = fSummaryDirectory
End Property

Public Property Get DailyReportFileName() As String
    DailyReportFileName = fDailyReportFileName
End Property

Public Property Get SummaryFileName() As String
    SummaryFileName = fSummaryFileName
End Property

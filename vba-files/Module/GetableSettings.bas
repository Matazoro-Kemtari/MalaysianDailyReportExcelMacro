Attribute VB_Name = "GetableSettings"
Option Explicit

Function GetSettings() As Settings
    Dim MySettings As New Settings
    Call MySettings.BuildOnce( _
        Sheet1.Range("DailyReportDirectory"), _
        Sheet1.Range("SummaryDirectory"), _
        Sheet1.Range("DailyReportFileName"), _
        Sheet1.Range("SummaryFileName") _
    )
    Set GetSettings = MySettings
End Function

Function GetTargets() As Variant
    ' ‘ÎÛÒ‚ğæ“¾
    Dim buf As Variant
    buf = Sheet1.Range("ûW‘ÎÛ").Value
    Dim Target() As Variant
    If IsArray(buf) Then
        Target = buf
    Else
        ReDim Target(1 To 1, 1 To 1)
        Target(1, 1) = buf
    End If
    Erase buf
    GetTargets = Target
End Function

Function GetTargetYear() As Long
    GetTargetYear = Sheet1.Range("ProcessYear")
End Function

Function GetTargetMonth() As Long
    GetTargetMonth = Sheet1.Range("ProcessMonth")
End Function

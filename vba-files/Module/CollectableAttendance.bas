Attribute VB_Name = "CollectableAttendance"
Option Explicit

' 勤怠情報を収集する
Sub CollectAttendance(ByRef MySettings As Settings, _
                      ByVal ProcessYear As Long, _
                      ByVal ProcessMonth As Long, _
                      ByRef Targets() As Variant)
    If MySettings Is Nothing Then
        Err.Raise ARGUMENT_NULL_EXCEPTION, "CollectAttendance", "引数の値がNulです MySettings"
    End If

    If ProcessYear < 2021 Then
        Err.Raise ARGUMENT_OUT_OF_RANGE_EXCEPTION, "CollectAttendance", "引数の値が範囲外です ProcessYear"
    End If

    If ProcessMonth < 1 Or ProcessMonth > 12 Then
        Err.Raise ARGUMENT_OUT_OF_RANGE_EXCEPTION, "CollectAttendance", "引数の値が範囲外です ProcessMonth"
    End If

    If LBound(Targets) <> 1 Then
        Err.Raise ARGUMENT_OUT_OF_RANGE_EXCEPTION, "CollectAttendance", "引数の値が範囲外です Target"
    End If

    Dim tar
    For Each tar In Targets
        ' 日報を開く
        Dim ReportBook As Workbook
        Set ReportBook = OpenDailyReport(MySettings, tar)
        If Not ReportBook Is Nothing Then
            ' フィルタ
            If FilterReport(ProcessYear, ProcessMonth, ReportBook) > 0 Then
                ' コピー
                ' 削除
                ' フィルタ解除
                Call ClearFilter(ReportBook)
                ' 保存
                ReportBook.Save
            End If
            ReportBook.Close SaveChanges:=False
            Set ReportBook = Nothing
        End If
    Next tar
End Sub

Private Function OpenDailyReport(ByRef MySettings As Settings, _
                                 ByVal Target As String) As Workbook
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' ファイル名を設定のフォルダとファイル名のサフィックスで作成
    Dim FileName As String
    FileName = fso.BuildPath(MySettings.DailyReportDirectory, _
                             Target & MySettings.DailyReportSuffix)
    If Not fso.FileExists(FileName) Then
        Exit Function
    End If

    On Error GoTo DUPLICATE_FILE_NAMES
    Set OpenDailyReport = Workbooks.Open(FileName:=FileName)
    Exit Function

DUPLICATE_FILE_NAMES:
    If Err = 1004 Then
        Err.Raise DUPLICATE_WORKSHEET_NAMES_EXCEPTION, Err.Description
    Else
        Err.Raise Err
    End If
End Function

private Function FilterReport(ByVal ProcessYear As Long, _
                              ByVal ProcessMonth As Long, _
                              ByRef ReportBook As Workbook) As Long
    Call ReportBook.Sheets(1).ListObjects("入力テーブル").Range.AutoFilter(Field:=1, Operator:= _
        xlFilterValues, Criteria2:=Array(1, CStr(ProcessMonth) & "/1/" & CStr(ProcessYear)))
    FilterReport = ReportBook.Sheets(1).ListObjects("入力テーブル").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
End Function

Private Sub ClearFilter(ByRef ReportBook As Workbook)
    ReportBook.Sheets(1).ShowAllData
End Sub

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

    ' 集計クラス
    Dim Collector As AttendanceCollector
    Set Collector = New AttendanceCollector

    Dim tar
    For Each tar In Targets
        ' 日報を開く
        Dim ReportBook As Workbook
        Set ReportBook = OpenDailyReport(MySettings, tar)
        If Not ReportBook Is Nothing Then
            ' フィルタ
            If FilterReport(ProcessYear, ProcessMonth, ReportBook) > 0 Then
                ' コピー
                Call CopyAttendance(ReportBook, Collector)
                ' 削除
                Call RemoveAttendance(ReportBook)
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
        ' Err1004はファイルがない場合・すでにファイル名が同じブックを開いている場合
        Err.Raise DUPLICATE_WORKSHEET_NAMES_EXCEPTION, Err.Description
    Else
        Err.Raise Err
    End If
End Function

private Function FilterReport(ByVal ProcessYear As Long, _
                              ByVal ProcessMonth As Long, _
                              ByRef ReportBook As Workbook) As Long
    ' オートフィルタで、月単位でフィルタをかける Arrayの1が月単位 日は指定したが無視される
    Call ReportBook.Sheets(1).ListObjects("入力テーブル").Range.AutoFilter(Field:=1, Operator:= _
        xlFilterValues, Criteria2:=Array(1, CStr(ProcessMonth) & "/1/" & CStr(ProcessYear)))
    FilterReport = ReportBook.Sheets(1).ListObjects("入力テーブル").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
End Function

Private Sub ClearFilter(ByRef ReportBook As Workbook)
    ReportBook.Sheets(1).ListObjects("入力テーブル").Range.AutoFilter
End Sub

Private Function CopyAttendance(ByRef ReportBook As Workbook, ByRef Collector As AttendanceCollector)
    ' ここでは、フィルタ等で表示セルのみコピーされる特性を
    ' 活かして、対象のみ集計テーブルにコピーする
    ' http://officetanaka.net/excel/vba/tips/tips155c.htm

    Dim repoTbl As ListObject
    Set repoTbl = ReportBook.Sheets(1).ListObjects("入力テーブル")
    Dim cnt As Long
    cnt = repoTbl.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1

    repoTbl.ListColumns("date").DataBodyRange.Copy
    Collector.NewRange("日付").PasteSpecial Paste:=xlPasteValues
    Collector.NewRange("日付").Resize(RowSize:=cnt).NumberFormatLocal = "yyyy/m/d"
    ReportBook.Sheets(1).Range("EmployeeNumber").Copy
    Collector.NewRange("社員番号").Resize(RowSize:=cnt).PasteSpecial Paste:=xlPasteValues
    ReportBook.Sheets(1).Range("EmployeeName").Copy
    Collector.NewRange("氏名").Resize(RowSize:=cnt).PasteSpecial Paste:=xlPasteValues
    repoTbl.ListColumns("Work type").DataBodyRange.Copy
    Collector.NewRange("残業区分").PasteSpecial Paste:=xlPasteValues
    repoTbl.ListColumns("Time").DataBodyRange.Copy
    Collector.NewRange("実働時間").PasteSpecial Paste:=xlPasteValues
    repoTbl.ListColumns("work number").DataBodyRange.Copy
    Collector.NewRange("作業番号").PasteSpecial Paste:=xlPasteValues
    repoTbl.ListColumns("code").DataBodyRange.Copy
    Collector.NewRange("コード").PasteSpecial Paste:=xlPasteValues
    repoTbl.ListColumns("Notes").DataBodyRange.Copy
    Collector.NewRange("特記事項").PasteSpecial Paste:=xlPasteValues
    repoTbl.ListColumns("class 1").DataBodyRange.Copy
    Collector.NewRange("大分類").PasteSpecial Paste:=xlPasteValues
    repoTbl.ListColumns("class 2").DataBodyRange.Copy
    Collector.NewRange("中分類").PasteSpecial Paste:=xlPasteValues

    Collector.Flash
End Function

Private Function RemoveAttendance(ByRef ReportBook As Workbook)
    Dim repoTbl As ListObject
    Set repoTbl = ReportBook.Sheets(1).ListObjects("入力テーブル")
    Dim cnt As Long
    cnt = repoTbl.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1

    repoTbl.ListColumns("date").DataBodyRange.EntireRow.Delete
    Dim i As Long
    For i = 1 To cnt
        repoTbl.ListRows.Add
    Next i
End Function

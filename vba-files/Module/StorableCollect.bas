Attribute VB_Name = "StorableCollect"
Option Explicit

Public Sub CollectBookExists(ByRef MySettings As Settings, _
                                  ByVal ProcessYear As Long, _
                                  ByVal ProcessMonth As Long)
    Dim FileName As String
    FileName = MakeCollectFileName(MySettings.SummaryDirectory, _
                                   ProcessYear, _
                                   ProcessMonth, _
                                   MySettings.SummaryFileName)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(FileName) Then
        Err.Raise COLLECT_BOOK_EXISTS_EXCEPTION, "CollectAttendance", "集計ファイルが存在しています" & vbNewLine & FileName
    End If
End Sub

Public Sub SaveCollect(ByRef MySettings As Settings, _
                       ByVal ProcessYear As Long, _
                       ByVal ProcessMonth As Long)
    Call CollectBookExists(MySettings, ProcessYear, ProcessMonth)

    Dim NewCollectBook As Workbook
    Set NewCollectBook = Workbooks.Add
    Sheet2.Copy Before:=NewCollectBook.Sheets(1)
    NewCollectBook.Sheets(2).Delete
    Dim FileName As String
    FileName = MakeCollectFileName(MySettings.SummaryDirectory, _
                                   ProcessYear, _
                                   ProcessMonth, _
                                   MySettings.SummaryFileName)
    NewCollectBook.SaveAs FileName:=FileName _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Sheet1.Activate
End Sub

Private  Function MakeCollectFileName(ByVal Path As String, _
                                      ByVal ProcessYear As Long, _
                                      ByVal ProcessMonth As Long, _
                                      ByVal CollectSufix) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' 集計ファイルを設定から作成する
    Dim FileName As String
    FileName = fso.BuildPath(Path, _
            CStr(ProcessYear) & Format("00", ProcessMonth) & CollectSufix)
    MakeCollectFileName = FileName
End Function

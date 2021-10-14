VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AttendanceCollector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 集計に使用するシート
Private fCollectionSheet As Worksheet
' 新規レコードのセルコレクション
Private fNewRange As Collection

Private Sub Class_Initialize ()
    Set fCollectionSheet = Sheet2
    If fCollectionSheet.AutoFilterMode Then
        fCollectionSheet.ShowAllData
    End If
    If fCollectionSheet.ListObjects("集計テーブル").Range.Columns(1).Count > 0 Then
        ' 内容を消す 消すとIndexがなくなるので、下から消す
        Dim i As Long
        For i = fCollectionSheet.ListObjects("集計テーブル").ListRows.Count To 1 Step -1
            Call fCollectionSheet.ListObjects("集計テーブル").ListRows(i).Delete
        Next i
    End If

    Call Flash
End Sub

Public Function Flash()
    Set fNewRange = New Collection
    Dim tbl As ListObject
    Set tbl = fCollectionSheet.ListObjects("集計テーブル")
    If tbl.ListColumns("日付").DataBodyRange Is Nothing Then
        fNewRange.Add Key:="日付", Item:=tbl.ListColumns("日付").Range.Cells(2)
        fNewRange.Add Key:="社員番号", Item:=tbl.ListColumns("社員番号").Range.Cells(2)
        fNewRange.Add Key:="氏名", Item:=tbl.ListColumns("氏名").Range.Cells(2)
        fNewRange.Add Key:="残業区分", Item:=tbl.ListColumns("残業区分").Range.Cells(2)
        fNewRange.Add Key:="実働時間", Item:=tbl.ListColumns("実働時間").Range.Cells(2)
        fNewRange.Add Key:="作業番号", Item:=tbl.ListColumns("作業番号").Range.Cells(2)
        fNewRange.Add Key:="コード", Item:=tbl.ListColumns("コード").Range.Cells(2)
        fNewRange.Add Key:="特記事項", Item:=tbl.ListColumns("特記事項").Range.Cells(2)
        fNewRange.Add Key:="大分類", Item:=tbl.ListColumns("大分類").Range.Cells(2)
        fNewRange.Add Key:="中分類", Item:=tbl.ListColumns("中分類").Range.Cells(2)
    Else
        fNewRange.Add Key:="日付", Item:=tbl.ListColumns("日付").DataBodyRange(tbl.ListColumns("日付").DataBodyRange.Count).Offset(RowOffset:=1)
        fNewRange.Add Key:="社員番号", Item:=tbl.ListColumns("社員番号").DataBodyRange(tbl.ListColumns("社員番号").DataBodyRange.Count).Offset(RowOffset:=1)
        fNewRange.Add Key:="氏名", Item:=tbl.ListColumns("氏名").DataBodyRange(tbl.ListColumns("氏名").DataBodyRange.Count).Offset(RowOffset:=1)
        fNewRange.Add Key:="残業区分", Item:=tbl.ListColumns("残業区分").DataBodyRange(tbl.ListColumns("残業区分").DataBodyRange.Count).Offset(RowOffset:=1)
        fNewRange.Add Key:="実働時間", Item:=tbl.ListColumns("実働時間").DataBodyRange(tbl.ListColumns("実働時間").DataBodyRange.Count).Offset(RowOffset:=1)
        fNewRange.Add Key:="作業番号", Item:=tbl.ListColumns("作業番号").DataBodyRange(tbl.ListColumns("作業番号").DataBodyRange.Count).Offset(RowOffset:=1)
        fNewRange.Add Key:="コード", Item:=tbl.ListColumns("コード").DataBodyRange(tbl.ListColumns("コード").DataBodyRange.Count).Offset(RowOffset:=1)
        fNewRange.Add Key:="特記事項", Item:=tbl.ListColumns("特記事項").DataBodyRange(tbl.ListColumns("特記事項").DataBodyRange.Count).Offset(RowOffset:=1)
        fNewRange.Add Key:="大分類", Item:=tbl.ListColumns("大分類").DataBodyRange(tbl.ListColumns("大分類").DataBodyRange.Count).Offset(RowOffset:=1)
        fNewRange.Add Key:="中分類", Item:=tbl.ListColumns("中分類").DataBodyRange(tbl.ListColumns("中分類").DataBodyRange.Count).Offset(RowOffset:=1)
    End If
End Function

Public Property Get NewRange() As Collection
    Set NewRange = fNewRange
End Property

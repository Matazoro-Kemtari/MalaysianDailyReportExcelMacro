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

' �W�v�Ɏg�p����V�[�g
Private fCollectionSheet As Worksheet
' �V�K���R�[�h�̃Z���R���N�V����
Private fNewRange As Collection

Private Sub Class_Initialize ()
    Set fCollectionSheet = Sheet2
    If fCollectionSheet.AutoFilterMode Then
        fCollectionSheet.ShowAllData
    End If
    If fCollectionSheet.ListObjects("�W�v�e�[�u��").Range.Columns(1).Count > 0 Then
        ' ���e������ ������Index���Ȃ��Ȃ�̂ŁA���������
        Dim i As Long
        For i = fCollectionSheet.ListObjects("�W�v�e�[�u��").ListRows.Count To 1 Step -1
            Call fCollectionSheet.ListObjects("�W�v�e�[�u��").ListRows(i).Delete
        Next i
    End If

    Call Flash
End Sub

Public Function Flash()
    Set fNewRange = New Collection
    Dim tbl As ListObject
    Set tbl = fCollectionSheet.ListObjects("�W�v�e�[�u��")
    If tbl.ListColumns("���t").DataBodyRange Is Nothing Then
        fNewRange.Add Key:="���t", Item:=tbl.ListColumns("���t").Range.Cells(2)
        fNewRange.Add Key:="�Ј��ԍ�", Item:=tbl.ListColumns("�Ј��ԍ�").Range.Cells(2)
        fNewRange.Add Key:="����", Item:=tbl.ListColumns("����").Range.Cells(2)
        fNewRange.Add Key:="�c�Ƌ敪", Item:=tbl.ListColumns("�c�Ƌ敪").Range.Cells(2)
        fNewRange.Add Key:="��������", Item:=tbl.ListColumns("��������").Range.Cells(2)
        fNewRange.Add Key:="��Ɣԍ�", Item:=tbl.ListColumns("��Ɣԍ�").Range.Cells(2)
        fNewRange.Add Key:="�R�[�h", Item:=tbl.ListColumns("�R�[�h").Range.Cells(2)
        fNewRange.Add Key:="���L����", Item:=tbl.ListColumns("���L����").Range.Cells(2)
        fNewRange.Add Key:="�啪��", Item:=tbl.ListColumns("�啪��").Range.Cells(2)
        fNewRange.Add Key:="������", Item:=tbl.ListColumns("������").Range.Cells(2)
    Else
        fNewRange.Add Key:="���t", Item:=tbl.ListColumns("���t").DataBodyRange(tbl.ListColumns("���t").DataBodyRange.Count).Offset(RowOffset:=1)
        fNewRange.Add Key:="�Ј��ԍ�", Item:=tbl.ListColumns("�Ј��ԍ�").DataBodyRange(tbl.ListColumns("�Ј��ԍ�").DataBodyRange.Count).Offset(RowOffset:=1)
        fNewRange.Add Key:="����", Item:=tbl.ListColumns("����").DataBodyRange(tbl.ListColumns("����").DataBodyRange.Count).Offset(RowOffset:=1)
        fNewRange.Add Key:="�c�Ƌ敪", Item:=tbl.ListColumns("�c�Ƌ敪").DataBodyRange(tbl.ListColumns("�c�Ƌ敪").DataBodyRange.Count).Offset(RowOffset:=1)
        fNewRange.Add Key:="��������", Item:=tbl.ListColumns("��������").DataBodyRange(tbl.ListColumns("��������").DataBodyRange.Count).Offset(RowOffset:=1)
        fNewRange.Add Key:="��Ɣԍ�", Item:=tbl.ListColumns("��Ɣԍ�").DataBodyRange(tbl.ListColumns("��Ɣԍ�").DataBodyRange.Count).Offset(RowOffset:=1)
        fNewRange.Add Key:="�R�[�h", Item:=tbl.ListColumns("�R�[�h").DataBodyRange(tbl.ListColumns("�R�[�h").DataBodyRange.Count).Offset(RowOffset:=1)
        fNewRange.Add Key:="���L����", Item:=tbl.ListColumns("���L����").DataBodyRange(tbl.ListColumns("���L����").DataBodyRange.Count).Offset(RowOffset:=1)
        fNewRange.Add Key:="�啪��", Item:=tbl.ListColumns("�啪��").DataBodyRange(tbl.ListColumns("�啪��").DataBodyRange.Count).Offset(RowOffset:=1)
        fNewRange.Add Key:="������", Item:=tbl.ListColumns("������").DataBodyRange(tbl.ListColumns("������").DataBodyRange.Count).Offset(RowOffset:=1)
    End If
End Function

Public Property Get NewRange() As Collection
    Set NewRange = fNewRange
End Property
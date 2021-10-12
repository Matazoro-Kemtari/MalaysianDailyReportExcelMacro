Attribute VB_Name = "CollectableAttendance"
Option Explicit

' �Αӏ������W����
Sub CollectAttendance(ByRef MySettings As Settings, _
                      ByVal ProcessYear As Long, _
                      ByVal ProcessMonth As Long, _
                      ByRef Targets() As Variant)
    If MySettings Is Nothing Then
        Err.Raise ARGUMENT_NULL_EXCEPTION, "CollectAttendance", "�����̒l��Nul�ł� MySettings"
    End If

    If ProcessYear < 2021 Then
        Err.Raise ARGUMENT_OUT_OF_RANGE_EXCEPTION, "CollectAttendance", "�����̒l���͈͊O�ł� ProcessYear"
    End If

    If ProcessMonth < 1 Or ProcessMonth > 12 Then
        Err.Raise ARGUMENT_OUT_OF_RANGE_EXCEPTION, "CollectAttendance", "�����̒l���͈͊O�ł� ProcessMonth"
    End If

    If LBound(Targets) <> 1 Then
        Err.Raise ARGUMENT_OUT_OF_RANGE_EXCEPTION, "CollectAttendance", "�����̒l���͈͊O�ł� Target"
    End If

    Dim tar
    For Each tar In Targets
        ' ������J��
        Dim ReportBook As Workbook
        Set ReportBook = OpenDailyReport(MySettings, tar)
        If Not ReportBook Is Nothing Then
            ' �t�B���^
            If FilterReport(ProcessYear, ProcessMonth, ReportBook) > 0 Then
                ' �R�s�[
                ' �폜
                ' �t�B���^����
                Call ClearFilter(ReportBook)
                ' �ۑ�
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
    ' �t�@�C������ݒ�̃t�H���_�ƃt�@�C�����̃T�t�B�b�N�X�ō쐬
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
    Call ReportBook.Sheets(1).ListObjects("���̓e�[�u��").Range.AutoFilter(Field:=1, Operator:= _
        xlFilterValues, Criteria2:=Array(1, CStr(ProcessMonth) & "/1/" & CStr(ProcessYear)))
    FilterReport = ReportBook.Sheets(1).ListObjects("���̓e�[�u��").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
End Function

Private Sub ClearFilter(ByRef ReportBook As Workbook)
    ReportBook.Sheets(1).ShowAllData
End Sub

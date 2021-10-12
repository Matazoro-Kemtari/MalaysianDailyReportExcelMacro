Attribute VB_Name = "Program"
Option Explicit

' ���[�U�[�G���[�ԍ�
Public Const FILE_NOT_FOUND_EXCEPTION As Long = vbObjectError + 514
Public Const ADD_IN_INSTALLED_EXCEPTION As Long = vbObjectError + 515
Public Const DUPULICATE_ADD_IN_EXCEPTION As Long = vbObjectError + 516
Public Const UNDEFINED_ADD_IN_EXCEPTION As Long = vbObjectError + 517
Public Const DIFFERENCE_ADD_IN_HASH As Long = vbObjectError + 518
Public Const ARGUMENT_OUT_OF_RANGE_EXCEPTION As Long = vbObjectError + 519
Public Const ARGUMENT_NULL_EXCEPTION As Long = vbObjectError + 520
Public Const DUPLICATE_WORKSHEET_NAMES_EXCEPTION As Long = vbObjectError + 521

' ������W�{�^���C�x���g
Sub CollectAttendanceButton_Click()
    On Error Goto CATCH_EXCEPTION

    ' �ݒ���擾
    Dim MySettings As Settings
    Set MySettings = GetSettings()
    Dim TargetYear As Long
    TargetYear = GetTargetYear()
    Dim TargetMonth As Long
    TargetMonth = GetTargetMonth()
    Dim TargetWorkers() As Variant
    TargetWorkers = GetTargets()

    ' ������W
    Call CollectAttendance(MySettings, _
                           TargetYear, _
                           TargetMonth, _
                           TargetWorkers)

    ' ���W�ۑ�

    Exit Sub

CATCH_EXCEPTION:
    Select Case Err.Number
    Case ARGUMENT_OUT_OF_RANGE_EXCEPTION, ARGUMENT_NULL_EXCEPTION, DUPLICATE_WORKSHEET_NAMES_EXCEPTION
        MsgBox Err.Description, vbCritical
    Case Else
        Err.Raise Err
    End Select
End Sub

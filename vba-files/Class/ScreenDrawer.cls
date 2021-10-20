VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScreenDrawer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mDrowEnabled As Boolean

Public Property Get RefreshDrow() As Boolean
    RefreshDrow = mDrowEnabled
End Property

Public Property Let RefreshDrow(ByVal vNewValue As Boolean)
    Select Case vNewValue
        Case True
            Call DrowStart
        Case False
            Call DroStop
    End Select
End Property

Private Sub class_initialize()
    '描画処理ストップ
    Call DroStop
End Sub

Private Sub Class_Terminate()
    '画面描画再開
    If Not mDrowEnabled Then
        Call DrowStart
    End If
End Sub

Private Sub DrowStart()
    '画面描画再開
    mDrowEnabled = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    '計算処理再開
    Application.Calculation = xlCalculationAutomatic
    '再計算を行う
    Application.Calculate
End Sub

Private Sub DroStop()
    '描画処理ストップ
    mDrowEnabled = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    '計算処理ストップ
    Application.Calculation = xlCalculationManual
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APITimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "APITimer"


Public Event Refresh()

Private mlngTimerID      As Long

Friend Property Let ID(ByVal plngValue As Long)
    mlngTimerID = plngValue
End Property

Friend Property Get ID() As Long
    ID = mlngTimerID
End Property

Public Sub StartTimer(ByVal Interval As Long)
    If mlngTimerID = 0 Then
        AddTimer Me, Interval
    End If
End Sub

Public Sub StopTimer()
    If mlngTimerID > 0 Then
        RemoveTimer Me
    End If
End Sub

Private Sub Class_Terminate()
    StopTimer
End Sub

Friend Sub RaiseTimerEvent()
    RaiseEvent Refresh
End Sub

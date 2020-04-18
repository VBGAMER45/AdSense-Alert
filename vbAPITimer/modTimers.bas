Attribute VB_Name = "modTimers"
Option Explicit

Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private mcolItems   As Collection

Public Sub AddTimer(ByRef pobjTimer As APITimer, ByVal plngInterval As Long)
    If mcolItems Is Nothing Then
        Set mcolItems = New Collection
    End If
    pobjTimer.ID = SetTimer(0, 0, plngInterval, AddressOf Timer_CBK)
    mcolItems.Add ObjPtr(pobjTimer), pobjTimer.ID & "K"
End Sub

Public Sub RemoveTimer(ByRef pobjTimer As APITimer)
On Error GoTo ErrHandler
    mcolItems.Remove pobjTimer.ID & "K"
    KillTimer 0, pobjTimer.ID
    pobjTimer.ID = 0
    If mcolItems.Count = 0 Then
        Set mcolItems = Nothing
    End If
    Exit Sub
ErrHandler:
    
End Sub

Public Sub Timer_CBK(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal SysTime As Long)
Dim lngPointer  As Long
Dim objTimer    As APITimer
On Error GoTo ErrHandler
    lngPointer = mcolItems.Item(idEvent & "K")
    Set objTimer = PtrObj(lngPointer)
    objTimer.RaiseTimerEvent
    Set objTimer = Nothing
    Exit Sub
ErrHandler:

End Sub

Private Function PtrObj(ByVal Pointer As Long) As Object
Dim objObject   As Object
    CopyMemory objObject, Pointer, 4&
    Set PtrObj = objObject
    CopyMemory objObject, 0&, 4&
End Function

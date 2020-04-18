VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PopUpMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Event PopupMessageClicked(ByRef Item As PopUpMessage)

Public Enum vbPopUpDirection
    vbPopUp = 0
    vbPopDown
    vbPopLeft
    vbPopRight
End Enum

Private mcolItems           As Collection

Private menmDirection       As vbPopUpDirection

Private mlngShowDelay       As Long
Private mlngScrollDelay     As Long
Private mlngMovementIndex   As Long
Private mlngMessageWidth    As Long
Private mlngMessageHeight   As Long
Private mlngXPos            As Long
Private mlngYPos            As Long
Private mlngParentHandle    As Long
Private mblnAllowFading     As Boolean

Private WithEvents mobjTimer       As APITimer
Attribute mobjTimer.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set mcolItems = New Collection
    menmDirection = vbPopUp
    mlngShowDelay = 3000
    mlngScrollDelay = 30
    mlngMovementIndex = 5
    mlngMessageWidth = 2500
    mlngMessageHeight = 1900
    mlngXPos = GetDesktopWidth - mlngMessageWidth
    mlngYPos = GetDesktopHeight
End Sub

Private Sub Class_Terminate()
    CloseDownGraceFully
End Sub

Public Property Get Item(ByVal Index As Variant) As PopUpMessage
    Set Item = mcolItems.Item(Index)
End Property

Public Property Get Count() As Long
    Count = mcolItems.Count
End Property

Public Function Show(ByRef PopUp As PopUpMessage)
    If Not PopUp.Visible Then
        PopUp.Initialise Me
        mcolItems.Add PopUp, CStr(PopUp.hWnd) & "K"
        StartTimer
    Else
        PopUp.ClosePopUp
    End If
End Function

Friend Sub Remove(ByRef pobjPopup As PopUpMessage)
    mcolItems.Remove CStr(pobjPopup.hWnd) & "K"
    If mcolItems.Count = 0 Then
        StopTimer
    End If
End Sub

Public Property Let PopUpDirection(ByVal Value As vbPopUpDirection)
    menmDirection = Value
End Property

Public Property Get PopUpDirection() As vbPopUpDirection
    PopUpDirection = menmDirection
End Property

Public Property Let ShowDelay(ByVal Value As Long)
    mlngShowDelay = Value
End Property

Public Property Get ShowDelay() As Long
    ShowDelay = mlngShowDelay
End Property

Public Property Let ScrollDelay(ByVal Value As Long)
    If Not (mlngScrollDelay = Value) Then
        mlngScrollDelay = Value
        If Not (mobjTimer Is Nothing) Then
            StopTimer
            StartTimer
        End If
    End If
End Property

Public Property Get ScrollDelay() As Long
    ScrollDelay = mlngScrollDelay
End Property

Public Property Let MovementIndex(ByVal Value As Long)
    mlngMovementIndex = Value
End Property

Public Property Get MovementIndex() As Long
    MovementIndex = mlngMovementIndex
End Property

Public Property Let MessageWidth(ByVal Value As Long)
    mlngMessageWidth = Value
End Property

Public Property Get MessageWidth() As Long
    MessageWidth = mlngMessageWidth
End Property

Public Property Let MessageHeight(ByVal Value As Long)
    mlngMessageHeight = Value
End Property

Public Property Get MessageHeight() As Long
    MessageHeight = mlngMessageHeight
End Property

Public Property Let XPos(ByVal Value As Long)
    mlngXPos = Value
End Property

Public Property Get XPos() As Long
    XPos = mlngXPos
End Property

Public Property Let YPos(ByVal Value As Long)
    mlngYPos = Value
End Property

Public Property Get YPos() As Long
    YPos = mlngYPos
End Property

Public Property Let ParentHandle(ByVal Value As Long)
    mlngParentHandle = Value
End Property

Public Property Get ParentHandle() As Long
    ParentHandle = mlngParentHandle
End Property

Public Property Let AllowFading(ByVal Value As Boolean)
    mblnAllowFading = Value
End Property

Public Property Get AllowFading() As Boolean
    AllowFading = mblnAllowFading
End Property

Private Sub StartTimer()
    If mobjTimer Is Nothing Then
        Set mobjTimer = New APITimer
        mobjTimer.StartTimer mlngScrollDelay
    End If
End Sub

Private Sub StopTimer()
    mobjTimer.StopTimer
    Set mobjTimer = Nothing
End Sub

Private Sub mobjTimer_Refresh()
    MoveMessages
End Sub

Private Sub MoveMessages()
Dim objPopUp        As PopUpMessage
Dim lngOffset       As Long
Dim lngCount        As Long
Dim lngIndex        As Long
Dim lngSystemTime   As Long
Dim lngParentWidth  As Long
Dim lngParentHeight As Long
Dim lngLeft         As Long
Dim lngTop          As Long
Dim lngHeight       As Long
Dim lngWidth        As Long
    lngSystemTime = GetTickCount
    Select Case menmDirection
        Case vbPopUpDirection.vbPopUp
            lngCount = (mlngYPos - (mlngYPos Mod mlngMessageHeight)) / mlngMessageHeight
        Case vbPopUpDirection.vbPopDown
            lngParentHeight = GetParentHeight(mlngParentHandle)
            lngCount = ((lngParentHeight - mlngYPos) - ((lngParentHeight - mlngYPos) Mod mlngMessageHeight)) / mlngMessageHeight
        Case vbPopUpDirection.vbPopLeft
            lngCount = (mlngXPos - (mlngXPos Mod mlngMessageWidth)) / mlngMessageWidth
        Case vbPopUpDirection.vbPopRight
            lngParentWidth = GetParentWidth(mlngParentHandle)
            lngCount = ((lngParentWidth - mlngXPos) - ((lngParentWidth - mlngXPos) Mod mlngMessageWidth)) / mlngMessageWidth
    End Select
    
    If lngCount > mcolItems.Count Then
        lngCount = mcolItems.Count
    End If
    If menmDirection = vbPopUp Or menmDirection = vbPopDown Then
        lngOffset = mlngYPos - (mlngYPos Mod Screen.TwipsPerPixelY)
    Else
        lngOffset = mlngXPos - (mlngXPos Mod Screen.TwipsPerPixelX)
    End If
    For Each objPopUp In mcolItems
        lngIndex = lngIndex + 1
        With objPopUp
            If .PercentShown = 100 And Not .Closing Then
                If CursorInWindow(.hWnd) Then
                    If lngSystemTime - .ShowTime > (mlngShowDelay / 2) Then
                        .ShowTime = lngSystemTime - (mlngShowDelay / 2)
                    End If
                Else
                    If (lngSystemTime - .ShowTime > mlngShowDelay) And Not (.ProgressBar And .Value < 100) And Not .Sticky Then
                        .PercentShown = .PercentShown - mlngMovementIndex
                    End If
                End If
            Else
                If Not .FullyShown And Not .Closing Then
                    If .PercentShown = 0 Then
                        PlayWavFile .WavFile
                    End If
                    .PercentShown = .PercentShown + mlngMovementIndex
                Else
                    If .PercentShown > 0 Then
                        .PercentShown = .PercentShown - mlngMovementIndex
                    End If
                End If
            End If
            If Not .FullyShown And Not .Closing Then
                If .PercentShown >= 100 Then
                    .PercentShown = 100
                    .FullyShown = True
                    .ShowTime = lngSystemTime
                End If
            Else
                If .PercentShown < 0 Then
                    .PercentShown = 0
                End If
            End If
            Select Case menmDirection
                Case vbPopUpDirection.vbPopUp
                    lngHeight = CLng(mlngMessageHeight * (.PercentShown / 100))
                    lngHeight = lngHeight - (lngHeight Mod Screen.TwipsPerPixelY)
                    lngWidth = mlngMessageWidth
                    lngLeft = mlngXPos
                    lngTop = lngOffset - lngHeight
                Case vbPopUpDirection.vbPopDown
                    lngHeight = CLng(mlngMessageHeight * (.PercentShown / 100))
                    lngHeight = lngHeight - (lngHeight Mod Screen.TwipsPerPixelY)
                    lngWidth = mlngMessageWidth
                    lngLeft = mlngXPos
                    lngTop = lngOffset
                Case vbPopUpDirection.vbPopLeft
                    lngHeight = mlngMessageHeight
                    lngWidth = CLng(mlngMessageWidth * (.PercentShown / 100))
                    lngWidth = lngWidth - (lngWidth Mod Screen.TwipsPerPixelX)
                    lngLeft = lngOffset - lngWidth
                    lngTop = mlngYPos
                Case vbPopUpDirection.vbPopRight
                    lngHeight = mlngMessageHeight
                    lngWidth = CLng(mlngMessageWidth * (.PercentShown / 100))
                    lngWidth = lngWidth - (lngWidth Mod Screen.TwipsPerPixelX)
                    lngLeft = lngOffset
                    lngTop = mlngYPos
            End Select
            .MoveMessage lngLeft, lngTop, lngWidth, lngHeight
            Select Case menmDirection
                Case vbPopUpDirection.vbPopUp
                    lngOffset = lngOffset - lngHeight
                Case vbPopUpDirection.vbPopDown
                    lngOffset = lngOffset + lngHeight
                Case vbPopUpDirection.vbPopLeft
                    lngOffset = lngOffset - lngWidth
                Case vbPopUpDirection.vbPopRight
                    lngOffset = lngOffset + lngWidth
            End Select
            If .Finished Then
                .DestroyPopup
            End If
        End With
        If lngIndex = lngCount Then
            Set objPopUp = Nothing
            Exit For
        End If
    Next objPopUp
    If mblnAllowFading Then
        For Each objPopUp In mcolItems
            objPopUp.SetNewTranslucency
        Next objPopUp
    End If
End Sub

Private Sub CloseDownGraceFully()
Dim objPopUp    As PopUpMessage
    For Each objPopUp In mcolItems
        objPopUp.DestroyPopup
    Next objPopUp
End Sub

Friend Sub PopupClicked(ByRef pobjPopup As PopUpMessage)
    RaiseEvent PopupMessageClicked(pobjPopup)
End Sub

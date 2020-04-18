Attribute VB_Name = "modFunctions"
Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA As Long = 48

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type



Public Function PtrObj(ByVal Pointer As Long) As Object
Dim objObject   As Object
    CopyMemory objObject, Pointer, 4&
    Set PtrObj = objObject
    CopyMemory objObject, 0&, 4&
End Function

Public Function GetDesktopWidth() As Long
Dim udtRect     As RECT
    SystemParametersInfo SPI_GETWORKAREA, 0, udtRect, 0
    GetDesktopWidth = udtRect.Right * Screen.TwipsPerPixelX
End Function

Public Function GetDesktopHeight() As Long
Dim udtRect     As RECT
    SystemParametersInfo SPI_GETWORKAREA, 0, udtRect, 0
    GetDesktopHeight = udtRect.Bottom * Screen.TwipsPerPixelY
End Function

Public Sub SetWindowToTop(ByVal plnghWnd As Long)
    SetWindowPos plnghWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Public Function GetParentHeight(ByVal plnghWnd As Long) As Long
Dim lngHeight   As Long
Dim udtRect     As RECT
    If plnghWnd = 0 Then
        lngHeight = GetDesktopHeight
    Else
        GetWindowRect plnghWnd, udtRect
        lngHeight = udtRect.Bottom - udtRect.Top
        lngHeight = lngHeight * Screen.TwipsPerPixelY
    End If
    GetParentHeight = lngHeight
End Function

Public Function GetParentWidth(ByVal plnghWnd As Long) As Long
Dim lngWidth    As Long
Dim udtRect     As RECT
    If plnghWnd = 0 Then
        lngWidth = GetDesktopWidth
    Else
        GetWindowRect plnghWnd, udtRect
        lngWidth = udtRect.Right - udtRect.Left
        lngWidth = lngWidth * Screen.TwipsPerPixelY
    End If
    GetParentWidth = lngWidth
End Function

Public Function CursorInWindow(ByVal plnghWnd As Long) As Boolean
Dim udtPt       As POINTAPI
Dim udtRect     As RECT
    GetCursorPos udtPt
    GetWindowRect plnghWnd, udtRect
    If udtPt.X >= udtRect.Left And udtPt.X <= udtRect.Right Then
        If udtPt.Y >= udtRect.Top And udtPt.Y <= udtRect.Bottom Then
            CursorInWindow = True
        End If
    End If
End Function

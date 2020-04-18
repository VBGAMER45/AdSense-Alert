Attribute VB_Name = "modTranslucency"
Option Explicit

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1&
Private Const LWA_ALPHA = &H2&

Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

Public Function SetTranslucency(ByVal plnghWnd As Long, ByVal pbytAlpha As Byte) As Boolean
Dim lngStyle    As Long
Dim lnghWnd     As Long
    If OSSupportsLayers Then
        lnghWnd = GetTopLevel(plnghWnd)
        lngStyle = GetWindowLong(lnghWnd, GWL_EXSTYLE)
        If lngStyle <> (lngStyle Or WS_EX_LAYERED) Then
            lngStyle = (lngStyle Or WS_EX_LAYERED)
            SetWindowLong lnghWnd, GWL_EXSTYLE, lngStyle
        End If
        SetTranslucency = CBool(SetLayeredWindowAttributes(lnghWnd, 0, CLng(pbytAlpha), LWA_ALPHA))
    End If
End Function

Private Function ClearTranslucency(ByVal plnghWnd As Long) As Boolean
Dim lngStyle    As Long
Dim lnghWnd     As Long
    If OSSupportsLayers Then
        lnghWnd = GetTopLevel(plnghWnd)
        Call SetLayeredWindowAttributes(lnghWnd, 0, 255&, LWA_ALPHA)
        lngStyle = GetWindowLong(lnghWnd, GWL_EXSTYLE) And Not WS_EX_LAYERED
        ClearTranslucency = CBool(SetWindowLong(lnghWnd, GWL_EXSTYLE, lngStyle))
    End If
End Function

Private Function OSSupportsLayers() As Boolean
Dim udtOS       As OSVERSIONINFO
    With udtOS
        .dwOSVersionInfoSize = Len(udtOS)
        Call GetVersionEx(udtOS)
        If .dwPlatformId = VER_PLATFORM_WIN32_NT Then
            OSSupportsLayers = (.dwMajorVersion > 4)
        End If
    End With
End Function

Private Function GetTopLevel(ByVal plngChild As Long) As Long
Dim lnghWnd As Long
    lnghWnd = plngChild
    Do While IsWindowVisible(GetParent(lnghWnd))
        lnghWnd = GetParent(plngChild)
        plngChild = lnghWnd
    Loop
    GetTopLevel = lnghWnd
End Function

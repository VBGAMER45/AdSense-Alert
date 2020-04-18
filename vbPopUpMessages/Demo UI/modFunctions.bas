Attribute VB_Name = "modFunctions"
Option Explicit

Private Declare Function SystemParametersInfo Lib "User32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Const SPI_GETWORKAREA As Long = 48

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Function GetDesktopWidth() As Long
Dim udtRECT     As RECT
    SystemParametersInfo SPI_GETWORKAREA, 0, udtRECT, 0
    GetDesktopWidth = udtRECT.Right * Screen.TwipsPerPixelX
End Function

Public Function GetDesktopHeight() As Long
Dim udtRECT     As RECT
    SystemParametersInfo SPI_GETWORKAREA, 0, udtRECT, 0
    GetDesktopHeight = udtRECT.Bottom * Screen.TwipsPerPixelY
End Function

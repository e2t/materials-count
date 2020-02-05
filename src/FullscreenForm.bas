Attribute VB_Name = "FullscreenForm"
Option Explicit

Private Declare PtrSafe Function GetSystemMetrics Lib "user32.dll" (ByVal index As Long) As Long
Private Const SM_CXFULLSCREEN = 16
Private Const SM_CYFULLSCREEN = 17
Private Const SM_CYCAPTION = 4

Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Function DPI() As Long
    Const LOGPIXELSX = 88
    Dim hWndDesktop As Long
    Dim hDCDesktop As Long
    
    hWndDesktop = GetDesktopWindow()
    hDCDesktop = GetDC(hWndDesktop)
    DPI = GetDeviceCaps(hDCDesktop, LOGPIXELSX)
    ReleaseDC hWndDesktop, hDCDesktop
End Function

Function PixelsToPoints() As Double
    PixelsToPoints = 72 / DPI  '96 pixels -> 72 points/inch = 0.75
End Function

Function MaximizedWidth() As Long
    MaximizedWidth = GetSystemMetrics(SM_CXFULLSCREEN) * PixelsToPoints
End Function

Function MaximizedHeight() As Long
    MaximizedHeight = (GetSystemMetrics(SM_CYFULLSCREEN) + GetSystemMetrics(SM_CYCAPTION)) * PixelsToPoints
End Function

Attribute VB_Name = "Module1"
Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal HWnd, ByVal hdc)
Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Declare Function GetDC Lib "user32" (ByVal HWnd)
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Global Const SRCPAINT = &HEE0086

Global Const SRCAND = &H8800C6 ' (DWORD) dest = source AND dest
Global Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Global Const SRCINVERT = &H660046 ' (DWORD) dest = source XOR dest
Global Const TRANSCOLOR = &H0&
Global Const TRANSCOLOR2 = &HFFFFFF


Function TransparentBlt(hDestDC As Long, nDestX, nDestY, nWidth, nHeight, hSourceDC As Long, nSourceX, nSourceY, TRANSCOLOR As Long)
    Dim lOldColor As Long
    Dim hMaskDC As Long
    Dim hMaskBmp As Long
    Dim hOldMaskBmp As Long
    Dim hTempBmp As Long
    Dim hTempDC As Long
    Dim hOldTempBmp As Long
    Dim hDummy As Long
    lOldColor = SetBkColor&(hSourceDC, TRANSCOLOR)
    lOldColor = SetBkColor&(hDestDC, TRANSCOLOR)
    hMaskDC = CreateCompatibleDC(hDestDC)
    hMaskBmp = CreateCompatibleBitmap(hDestDC, nWidth, nHeight)
    hOldMaskBmp = SelectObject(hMaskDC, hMaskBmp)
    hTempBmp = CreateBitmap(nWidth, nHeight, 1, 1, 0&)
    hTempDC = CreateCompatibleDC(hDestDC)
    hOldTempBmp = SelectObject(hTempDC, hTempBmp)
    If BitBlt(hTempDC, 0, 0, nWidth, nHeight, hSourceDC, nSourceX, nSourceY, SRCCOPY) Then
        hDummy = BitBlt(hMaskDC, 0, 0, nWidth, nHeight, hTempDC, 0, 0, SRCCOPY)
    End If
    hTempBmp = SelectObject(hTempDC, hOldTempBmp)
    hDummy = DeleteObject(hTempBmp)
    hDummy = DeleteDC(hTempDC)
    If BitBlt(hDestDC, nDestX, nDestY, nWidth, nHeight, hSourceDC, nSourceX, nSourceY, SRCINVERT) Then
      If BitBlt(hDestDC, nDestX, nDestY, nWidth, nHeight, hMaskDC, 0, 0, SRCAND) Then
        If BitBlt(hDestDC, nDestX, nDestY, nWidth, nHeight, hSourceDC, nSourceX, nSourceY, SRCINVERT) Then
           TransparentBlt = True
        End If
      End If
    End If
    hMaskBmp = SelectObject(hMaskDC, hOldMaskBmp)
    hDummy = DeleteObject(hMaskBmp)
    hDummy = DeleteDC(hMaskDC)
  End Function


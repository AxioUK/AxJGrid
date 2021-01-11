Attribute VB_Name = "mdlRender"
'// Render Strech
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal Hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32.dll" (ByVal Hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal Hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal Hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal Hdc As Long) As Long
Private Declare Function GdiTransparentBlt Lib "gdi32.dll" (ByVal Hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal Hdc As Long, ByVal hObject As Long) As Long







'Function Render Stretch
Public Function RenderStretchFromPicture(ByVal destDC As Long, ByVal destX As Long, ByVal destY As Long, ByVal DestW As Long, ByVal DestH As Long, ByVal SrcPicture As StdPicture, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Size As Long, Optional MaskColor As Long = -1)
    On Error Resume Next
    Dim DC          As Long
    Dim hOldBmp    As Long
    DC = CreateCompatibleDC(0)
    hOldBmp = SelectObject(DC, SrcPicture.Handle)
    RenderStretchFromDC destDC, destX, destY, DestW, DestH, DC, X, Y, Width, Height, Size, MaskColor
    hOldBmp = SelectObject(DC, hOldBmp)
    DeleteDC DC
End Function
Private Function RenderStretchFromDC(ByVal destDC As Long, ByVal destX As Long, ByVal destY As Long, ByVal DestW As Long, ByVal DestH As Long, ByVal SrcDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Size As Long, Optional MaskColor As Long = -1)
Dim Sx2 As Long
        Sx2 = Size * 2
        If MaskColor <> -1 Then
            Dim mDC         As Long
            Dim mX          As Long
            Dim mY          As Long
            Dim DC          As Long
            Dim hBmp        As Long
            Dim hOldBmp     As Long
        
            mDC = destDC
            DC = GetDC(0)
            destDC = CreateCompatibleDC(0)
            hBmp = CreateCompatibleBitmap(DC, DestW, DestH)
            hOldBmp = SelectObject(destDC, hBmp) ' save the original BMP for later reselection
            mX = destX: mY = destY
            destX = 0: destY = 0
        End If

        SetStretchBltMode destDC, vbPaletteModeNone
        
        BitBlt destDC, destX, destY, Size, Size, SrcDC, X, Y, vbSrcCopy  'TOP_LEFT
        StretchBlt destDC, destX + Size, destY, DestW - Sx2, Size, SrcDC, X + Size, Y, Width - Sx2, Size, vbSrcCopy 'TOP_CENTER
        BitBlt destDC, destX + DestW - Size, destY, Size, Size, SrcDC, X + Width - Size, Y, vbSrcCopy 'TOP_RIGHT
        StretchBlt destDC, destX, destY + Size, Size, DestH - Sx2, SrcDC, X, Y + Size, Size, Height - Sx2, vbSrcCopy 'MID_LEFT
        StretchBlt destDC, destX + Size, destY + Size, DestW - Sx2, DestH - Sx2, SrcDC, X + Size, Y + Size, Width - Sx2, Height - Sx2, vbSrcCopy 'MID_CENTER
        StretchBlt destDC, destX + DestW - Size, destY + Size, Size, DestH - Sx2, SrcDC, X + Width - Size, Y + Size, Size, Height - Sx2, vbSrcCopy 'MID_RIGHT
        BitBlt destDC, destX, destY + DestH - Size, Size, Size, SrcDC, X, Y + Height - Size, vbSrcCopy 'BOTTOM_LEFT
        StretchBlt destDC, destX + Size, destY + DestH - Size, DestW - Sx2, Size, SrcDC, X + Size, Y + Height - Size, Width - Sx2, Size, vbSrcCopy   'BOTTOM_CENTER
        BitBlt destDC, destX + DestW - Size, destY + DestH - Size, Size, Size, SrcDC, X + Width - Size, Y + Height - Size, vbSrcCopy 'BOTTOM_RIGHT

    If MaskColor <> -1 Then
        GdiTransparentBlt mDC, mX, mY, DestW, DestH, destDC, 0, 0, DestW, DestH, MaskColor
        SelectObject destDC, hOldBmp
        DeleteObject hBmp
        DeleteDC DC
        DeleteDC destDC
    End If

End Function

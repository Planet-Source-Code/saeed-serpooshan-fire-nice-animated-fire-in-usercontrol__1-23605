Attribute VB_Name = "UcFireModuleAPI"
' this module contains some declaring sub for use some API functions.
' this is needed by UcFire.Ctl UserControl

Public Type POINTAPI
        x As Long
        y As Long
End Type

'as you can see you can declare API s As Function Or Sub!
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Sub SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long)

'------------- BitBlt: -----------------------------------------------------------------------
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const SRCAND = &H8800C6          ' (DWORD) dest = source AND dest
Public Const SRCCOPY = &HCC0020         ' (DWORD) dest = source
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest

Public Declare Sub BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long)
' Note : usually you can use paintpicture function instead of bitblt.
'
'1 Copy old background back over sprite's old position
'2 Copy the background where the sprite will go
'3 Draw the mask
'4 Draw the sprite
'    temp = BitBlt(hdc, x, y, PicWidth, PicHeight, picCopy.hdc, 0, 0, SRCCOPY)
'    temp = BitBlt(picCopy.hdc, 0, 0, PicWidth, PicHeight, hdc, NewX, NewY, SRCCOPY)
'    temp = BitBlt(hdc, NewX, NewY, PicWidth, PicHeight, picMask.hdc, 0, 0, SRCAND)
'    temp = BitBlt(hdc, NewX, NewY, PicWidth, PicHeight, picSprite.hdc, 0, 0, SRCINVERT)
'--------------------------------------------------------------------------------------------

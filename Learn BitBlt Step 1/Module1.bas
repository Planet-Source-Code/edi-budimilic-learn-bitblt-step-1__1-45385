Attribute VB_Name = "Module1"
'These are the API functions that makes it all possible (to use BitBlt and other functions)

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCERASE = &H4400328
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

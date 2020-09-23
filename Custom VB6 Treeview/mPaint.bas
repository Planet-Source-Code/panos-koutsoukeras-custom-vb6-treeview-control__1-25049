Attribute VB_Name = "mPaint"
Option Explicit

' Gradient drawing
Public Const GRADIENT_FILL_RECT_H As Long = &H0
Public Const GRADIENT_FILL_RECT_V  As Long = &H1
Public Const GRADIENT_FILL_TRIANGLE As Long = &H2

Public Const DSna = &H220326

' General window definitions
Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' GetBitmapIntoDC
Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

' Don't use the VB API declare of PAINTSTRUCT-
' it misses out the full length of reserved data
' bytes, causing a GPF under NT
Public Type PAINTSTRUCT
   hdc As Long ' Handle to the display DC to be used for painting
   fErase As Long ' Specifies whether the background must be erased
   rcPaint As RECT ' Specifies a RECT structure that specifies the upper left and lower right corners of the rectangle in which the painting is requested.
   fRestore As Long ' Reserved; used internally by the system.
   fIncUpdate As Long ' Reserved; used internally by the system.
   rgbReserved(0 To 31) As Byte ' Reserved; used internally by the system.
End Type

' The Memory DC structure
Public Type MemoryDC
   MemHDC As Long
   MemBmp As Long
   MemBmpOld As Long
   MemWidth As Long
   MemHeight As Long
End Type

Public Type TRIVERTEX
    x As Long
    y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type
    
Public Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type
    
Public Type GRADIENT_TRIANGLE
    Vertex1 As Long
    Vertex2 As Long
    Vertex3 As Long
End Type

Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

' Called from GetBitmapIntoDC (UserControl)
Public Declare Function GetDesktopWindow Lib "user32" () As Long

' Called from DoTvNotify(mTreeview), Resize(UserControl)
' and TvWMPaint(UserControl)
Public Declare Function InvalidateRect Lib "user32" _
                            (ByVal hWnd As Long, _
                             ByVal lpRect As Long, _
                             ByVal bErase As Long) As Long

' GDI object functions:

' Used widely in UserControl
Public Declare Function SelectObject Lib "gdi32" _
                            (ByVal hdc As Long, _
                             ByVal hObject As Long) As Long
                             
' Used widely in UserControl
Public Declare Function DeleteObject Lib "gdi32" _
                            (ByVal hObject As Long) As Long
                            
' Called from GetBitmapIntoDC (UserControl)
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" _
                            (ByVal hObject As Long, _
                             ByVal nCount As Long, _
                             lpObject As Any) As Long
                            
' Called from GetBitmapIntoDC (UserControl)
Public Declare Function ReleaseDC Lib "user32" _
                            (ByVal hWnd As Long, _
                             ByVal hdc As Long) As Long
                             
' Called from GetBitmapIntoDC (UserControl)
Public Declare Function GetDC Lib "user32" _
                            (ByVal hWnd As Long) As Long
                            
' Used widely in UserControl
Public Declare Function DeleteDC Lib "gdi32" _
                            (ByVal hdc As Long) As Long
                             
' Called from MakeMemDC(UserControl) and GetBitmapIntoDC(UserControl)
Public Declare Function CreateCompatibleDC Lib "gdi32" _
                            (ByVal hdc As Long) As Long
                            
' Called from MakeMemDC(UserControl) and GetBitmapIntoDC(UserControl)
Public Declare Function CreateCompatibleBitmap Lib "gdi32" _
                            (ByVal hdc As Long, _
                             ByVal nWidth As Long, _
                             ByVal nHeight As Long) As Long
                             
' Called from MakeMemDC(UserControl) and SetBackColor(UserControl)
Public Declare Function CreateSolidBrush Lib "gdi32" _
                            (ByVal crColor As Long) As Long
                            
' Called from TvWMPaint(UserControl)
Public Declare Function SetTextColor Lib "gdi32" _
                            (ByVal hdc As Long, _
                             ByVal crColor As Long) As Long
                             
' Called from TvWMPaint(UserControl)
Public Declare Function SetBkColor Lib "gdi32" _
                            (ByVal hdc As Long, _
                             ByVal crColor As Long) As Long
                             
' Used widely in UserControl
Public Declare Function BitBlt Lib "gdi32" _
                            (ByVal hDestDC As Long, _
                             ByVal x As Long, _
                             ByVal y As Long, _
                             ByVal nWidth As Long, _
                             ByVal nHeight As Long, _
                             ByVal hSrcDC As Long, _
                             ByVal xSrc As Long, _
                             ByVal ySrc As Long, _
                             ByVal dwRop As Long) As Long
                             
' Called from MakeMemDC(UserControl) and SetBackColor(UserControl)
Public Declare Function FillRect Lib "user32" _
                            (ByVal hdc As Long, _
                             lpRect As RECT, _
                             ByVal hBrush As Long) As Long
                         
' Called from TvWMPaint(UserControl)
Public Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" _
                            (ByVal hdc As Long, _
                             pVertex As TRIVERTEX, _
                             ByVal dwNumVertex As Long, _
                             pMesh As GRADIENT_RECT, _
                             ByVal dwNumMesh As Long, _
                             ByVal dwMode As Long) As Long

' Called from TvWMPaint(UserControl)
Public Declare Function GradientFillTri Lib "msimg32" Alias "GradientFill" _
                            (ByVal hdc As Long, _
                             pVertex As TRIVERTEX, _
                             ByVal dwNumVertex As Long, _
                             pMesh As GRADIENT_TRIANGLE, _
                             ByVal dwNumMesh As Long, _
                             ByVal dwMode As Long) As Long

' Called from TranslateColor
Private Declare Function OleTranslateColor Lib "oleaut32.dll" _
                            (ByVal lOleColor As Long, _
                             ByVal lHPalette As Long, _
                             lColorRef As Long) As Long

' Called from ForeColor and Enabled properties(UserControl)
Public Declare Function InvalidateRectAsNull Lib "user32" Alias "InvalidateRect" _
                            (ByVal hWnd As Long, _
                             ByVal lpRect As Long, _
                             ByVal bErase As Long) As Long
                         
' Called from ForeColor and Enabled properties(UserControl)
Public Declare Function UpdateWindow Lib "user32" _
                            (ByVal hWnd As Long) As Long

' Called from TvWMPaint(UserControl)
Public Declare Function BeginPaint Lib "user32" _
                            (ByVal hWnd As Long, _
                             lpPaint As Any) As Long

' Called from TvWMPaint(UserControl)
Public Declare Function EndPaint Lib "user32" _
                            (ByVal hWnd As Long, _
                             lpPaint As Any) As Long

' Use LockWindowUpdate with care! - If you call it and there
' is an attempt to resize the control or draw something that
' was previously hidden then there is screen flicker.
' Called from ShowShellContextMenu(mIShellFolder) ' and DoTvNotify(mTreeview)
Public Declare Function LockWindowUpdate Lib "user32" _
                            (ByVal hwndLock As Long) As Long
                         
' Used widely in UserControl
Public Declare Function GetClientRect Lib "user32" _
                            (ByVal hWnd As Long, _
                             lpRect As RECT) As Long

' Used widely
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                            (ByVal hWnd As Long, _
                             ByVal wMsg As Long, _
                             ByVal wParam As Long, _
                             lParam As Any) As Long

' Used widely in UserControl
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
                            (ByVal hWnd As Long, _
                             ByVal wMsg As Long, _
                             ByVal wParam As Long, _
                             ByVal lParam As Long) As Long
                         
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
                            (ByVal hWnd As Long, _
                             ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
                            (ByVal hWnd As Long, _
                             ByVal nIndex As Long, _
                             ByVal dwNewLong As Long) As Long

' Used widely
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
                            (pDest As Any, _
                             pSource As Any, _
                             ByVal dwLength As Long)
                         
' Used widely
Public Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" _
                            (pDest As Any, _
                             ByVal dwLength As Long, _
                             ByVal bFill As Byte)

' Converts a OLE_COLOR to Long
' Used widely in UserControl
Public Function TranslateColor(lColor As Long, Optional ByVal hPal As Long = 0) As Long

  Dim lR As Long

    OleTranslateColor lColor, hPal, lR
    TranslateColor = lR

End Function






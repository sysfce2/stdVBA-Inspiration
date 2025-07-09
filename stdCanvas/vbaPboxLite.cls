VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vbaPboxLite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

 '   ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
 '   ||||||||||||       VBAPICTUREBOX LITE (v0.1)         |||||||||||||
 '   ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
 '
 '   AUTHOR:   Kallun Willock
 '   URL:      https://github.com/KallunWillock/??/
 '   PURPOSE:  The ImageMsgBox is a Unicode-enabled MessageBox that can display a custom icon (ICO file).                                                                                     ' _
 '             It also features a timeout feature resulting in it closing down after a designated period of time.
 '   LICENSE:  MIT
 '   VERSION:  0.1        20/05/2025         Initial commit.
 '
 '   TODO:     [ ] Add AutoRedraw functionality          [ ] GDI + GDI+ Shapes
 '             [ ] Text Routines                         [ ] Left / Right Justification of content
 '             [ ] Better Pen/Brush handling             [ ] Drawing/Pen style
 '
 '   NOTES:    -
 
#If False Then
  
  ' Setup
  
  Private WithEvents PBox As vbaPbox
Attribute PBox.VB_VarHelpID = -1
  
  Private Sub UserForm_Initialize()
    Set CV = New vbaPbox
    CV.InitializeCanvas Me.Frame1
  End Sub

#End If

' API Declarations

#If VBA7 Then
  Private Declare PtrSafe Function SetPixel Lib "gdi32" (ByVal hDC As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
  Private Declare PtrSafe Function GetPixel_ Lib "gdi32" Alias "GetPixel" (ByVal hDC As LongPtr, ByVal nXPos As Long, ByVal nYPos As Long) As Long
  Private Declare PtrSafe Function GdipGetImageWidth Lib "gdiplus.dll" (ByVal Image As LongPtr, Width As Long) As Long
  Private Declare PtrSafe Function GdipGetImageHeight Lib "gdiplus.dll" (ByVal Image As LongPtr, height As Long) As Long
  Private Declare PtrSafe Function CreateHatchBrush Lib "gdi32.dll" (ByVal nIndex As Long, ByVal crColor As Long) As LongPtr
  Private Declare PtrSafe Function Polygon Lib "gdi32.dll" (ByVal hDC As LongPtr, lpPoint As POINTAPI, ByVal nCount As Long) As Long
  Private Declare PtrSafe Function MoveToEx Lib "gdi32.dll" (ByVal hDC As LongPtr, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
  Private Declare PtrSafe Function LineTo Lib "gdi32.dll" (ByVal hDC As LongPtr, ByVal X As Long, ByVal Y As Long) As Long
  Private Declare PtrSafe Function SetBkMode Lib "gdi32.dll" (ByVal hDC As LongPtr, ByVal nBkMode As Long) As Long
  Private Declare PtrSafe Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As LongPtr
  Private Declare PtrSafe Function GetDC Lib "user32.dll" (ByVal hWnd As LongPtr) As LongPtr
  Private Declare PtrSafe Function ReleaseDC Lib "user32.dll" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
  Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As LongPtr) As LongPtr
  Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
  Private Declare PtrSafe Function SelectObject Lib "gdi32.dll" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
  Private Declare PtrSafe Function DeleteObject Lib "gdi32.dll" (ByVal hObject As LongPtr) As Long
  Private Declare PtrSafe Function DeleteDC Lib "gdi32.dll" (ByVal hDC As LongPtr) As Long
  Private Declare PtrSafe Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
  Private Declare PtrSafe Function StretchBlt Lib "gdi32.dll" (ByVal hDC As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
  Private Declare PtrSafe Function TransparentBlt Lib "msimg32.dll" (ByVal hdcDest As LongPtr, ByVal xoriginDest As Long, ByVal yoriginDest As Long, ByVal wDest As Long, ByVal hDest As Long, ByVal hdcSrc As LongPtr, ByVal xoriginSrc As Long, ByVal yoriginSrc As Long, ByVal wSrc As Long, ByVal hSrc As Long, ByVal crTransparent As Long) As Boolean
  Private Declare PtrSafe Function FillRect Lib "user32.dll" (ByVal hDC As LongPtr, lpRect As RECT, ByVal hBrush As LongPtr) As Long
  Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As LongPtr
  Private Declare PtrSafe Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As LongPtr
  Private Declare PtrSafe Function Rectangle Lib "gdi32.dll" (ByVal hDC As LongPtr, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
  Private Declare PtrSafe Function Ellipse Lib "gdi32.dll" (ByVal hDC As LongPtr, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
  Private Declare PtrSafe Function OleTranslateColor Lib "oleaut32.dll" (ByVal clr As OLE_COLOR, ByVal palet As LongPtr, col As Long) As Long
  Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As LongPtr)
  Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hWnd As LongPtr) As Long
  Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32.dll" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
  Private Declare PtrSafe Function GdipLoadImageFromFile Lib "gdiplus.dll" (ByVal Filename As LongPtr, Image As LongPtr) As Long
  Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "gdiplus.dll" (ByVal bitmap As LongPtr, hbmReturn As LongPtr, ByVal background As Long) As Long
  Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus.dll" (ByVal Image As LongPtr) As Long
  Private Declare PtrSafe Function GdipSaveImageToFile Lib "gdiplus.dll" (ByVal Image As LongPtr, ByVal Filename As LongPtr, clsidEncoder As GUID, encoderParams As Any) As Long
  Private Declare PtrSafe Function CLSIDFromString Lib "ole32.dll" (ByVal lpsz As LongPtr, pclsid As GUID) As Long
  Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
  Private Declare PtrSafe Function GetDpiForWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
#Else
  Public Enum LongPtr
  [_]
  End Enum
  Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
  Private Declare Function GetPixel_ Lib "gdi32" Alias "GetPixel" (ByVal hDC As LongPtr, ByVal nXPos As Long, ByVal nYPos As Long) As Long
  Private Declare Function GdipGetImageWidth Lib "gdiplus.dll" (ByVal Image As LongPtr, Width As Long) As Long
  Private Declare Function GdipGetImageHeight Lib "gdiplus.dll" (ByVal Image As LongPtr, Height As Long) As Long
  Private Declare Function CreateHatchBrush Lib "gdi32.dll" (ByVal nIndex As Long, ByVal crColor As Long) As LongPtr
  Private Declare Function Polygon Lib "gdi32.dll" (ByVal hDC As LongPtr, lpPoint As POINTAPI, ByVal nCount As Long) As Long
  Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hDC As LongPtr, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
  Private Declare Function LineTo Lib "gdi32.dll" (ByVal hDC As LongPtr, ByVal X As Long, ByVal Y As Long) As Long
  Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As LongPtr, ByVal nBkMode As Long) As Long
  Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As LongPtr
  Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As LongPtr) As LongPtr
  Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
  Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As LongPtr) As LongPtr
  Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
  Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As LongPtr) As Long
  Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As LongPtr) As Long
  Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
  Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hDC As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
  Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdcDest As LongPtr, ByVal xoriginDest As Long, ByVal yoriginDest As Long, ByVal wDest As Long, ByVal hDest As Long, ByVal hdcSrc As LongPtr, ByVal xoriginSrc As Long, ByVal yoriginSrc As Long, ByVal wSrc As Long, ByVal hSrc As Long, ByVal crTransparent As Long) As Boolean
  Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As LongPtr, lpRect As Rect, ByVal hBrush As LongPtr) As Long
  Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As LongPtr
  Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As LongPtr
  Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hDC As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
  Private Declare Function Ellipse Lib "gdi32.dll" (ByVal hDC As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
  Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal clr As OLE_COLOR, ByVal palet As LongPtr, col As Long) As Long
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
  Private Declare Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hWnd As LongPtr) As Long
  Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
  Private Declare Function GdiplusStartup Lib "gdiplus.dll" (token As LongPtr, inputbuf As gdiplusStartupInput, Optional ByVal outputbuf As LongPtr = 0) As Long
  Private Declare Sub GdiplusShutdown Lib "gdiplus.dll" (ByVal token As LongPtr)
  Private Declare Function GdipLoadImageFromFile Lib "gdiplus.dll" (ByVal Filename As LongPtr, Image As LongPtr) As Long
  Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus.dll" (ByVal bitmap As LongPtr, hbmReturn As LongPtr, ByVal background As Long) As Long
  Private Declare Function GdipDisposeImage Lib "gdiplus.dll" (ByVal Image As LongPtr) As Long
  Private Declare Function GdipSaveImageToFile Lib "gdiplus.dll" (ByVal Image As LongPtr, ByVal Filename As LongPtr, clsidEncoder As Guid, encoderParams As Any) As Long
  Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpsz As LongPtr, pclsid As Guid) As Long
  Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As PICTDESC, RefIID As Guid, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
#End If

' Constants / Types / Enums

Private Const SRCCOPY As Long = &HCC0020
Private Const PS_SOLID As Long = 0
Private Const NULL_BRUSH As Long = 7 '5
Private Const NULL_PEN As Long = 8
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const GM_ADVANCED As Long = 2
Private Const GP_STAT_OK As Long = 0

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(0 To 7) As Byte
End Type

Private Type EncoderParameter
  GUID As GUID
  NumberOfValues As Long
  Type As Long
  Value As LongPtr
End Type
  
Private Type EncoderParameters
  Count As Long
  Parameter(0 To 0) As EncoderParameter
End Type

Private Type PICTDESC
  size As Long
  Type As Long
  hPic As LongPtr
  hPal As LongPtr
End Type

Public Enum PicFileTypeEnum
  pictypeBMP = 1
  pictypeGIF = 2
  pictypePNG = 3
  pictypeJPG = 4
End Enum

Public ForeColor As OLE_COLOR
Public BackColor As OLE_COLOR
Public FillColor As OLE_COLOR
Public PenWidth As Long

Private Const BKMODE_TRANSPARENT As Long = 1
Private Const BKMODE_OPAQUE As Long = 2

Private Const HS_HORIZONTAL As Long = 0
Private Const HS_VERTICAL As Long = 1
Private Const HS_FDIAGONAL As Long = 2
Private Const HS_BDIAGONAL As Long = 3
Private Const HS_CROSS As Long = 4
Private Const HS_DIAGCROSS As Long = 5

Public Enum FillStyleConstants
  vbFSSolid = 0
  vbFSTransparent = 1
  vbHorizontalLine = 2
  vbVerticalLine = 3
  vbUpwardDiagonal = 4
  vbDownwardDiagonal = 5
  vbCross = 6
  vbDiagonalCross = 7
End Enum

Private Type PictureType
  FilePath As String
  Width As Long
  height As Long
End Type

Private Type SettingsType
  DPI   As Long
  PixelsPerPoint As Single
  PointsPerPixel As Single
End Type

Private Type ClassVariablesType
  hDC As LongPtr
  hWnd As LongPtr
  MemoryDC    As LongPtr
  Image As PictureType
  Settings As SettingsType
  CanvasWidth As Long
  CanvasHeight As Long
  MemoryBitmap As LongPtr
  OldMemoryBitmap As LongPtr
End Type

Private This As ClassVariablesType
Private WithEvents HostFrame As MSForms.frame

Public Event KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Public Event KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Public Event KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

' Properties

Public Property Get hDC() As LongPtr: hDC = This.MemoryDC: End Property
Public Property Get hWnd() As LongPtr:  hWnd = This.hWnd: End Property
Public Property Get Enabled() As Boolean: Enabled = HostFrame.Enabled: End Property
Public Property Let Enabled(ByVal RHS As Boolean): HostFrame.Enabled = RHS: End Property
Public Property Get Left() As Long: Left = HostFrame.Left: End Property
Public Property Let Left(ByVal RHS As Long): HostFrame.Left = RHS: End Property
Public Property Get Top() As Long: Top = HostFrame.Top: End Property
Public Property Let Top(ByVal RHS As Long): HostFrame.Top = RHS: End Property
Public Property Get Width() As Long: Width = HostFrame.Width: End Property
Public Property Let Width(ByVal RHS As Long): HostFrame.Width = RHS: End Property
Public Property Get height() As Long: height = HostFrame.height: End Property
Public Property Let height(ByVal RHS As Long): HostFrame.height = RHS: End Property
Public Property Get CanvasWidth() As Long: CanvasWidth = This.CanvasWidth: End Property
Public Property Get CanvasHeight() As Long: CanvasHeight = This.CanvasHeight: End Property

Public Sub InitializeCanvas(ByRef frmHost As MSForms.frame)
  
  Dim Result As Long
  
  If frmHost Is Nothing Then Exit Sub
  Set HostFrame = frmHost
  'Set m_Parent = frmHost.Parent
  frmHost.BackColor = frmHost.Parent.BackColor
  
  ' Strip the styling unless the frame control unless non-default styles have been applied.
  If frmHost.BorderStyle = fmBorderStyleNone Or frmHost.BorderColor = &H80000012 Or frmHost.SpecialEffect = fmSpecialEffectEtched Then
    frmHost.SpecialEffect = 0
    frmHost.BorderStyle = fmBorderStyleNone
    frmHost.Caption = vbNullString
  End If
    
  Call IUnknown_GetWindow(frmHost, VarPtr(This.hWnd))
  
  This.hDC = GetDC(This.hWnd)
  
  If This.hDC = 0 Then Exit Sub
  
  'This.CanvasWidth = PointsToPixelsX(frmHost.Width)
  'This.CanvasHeight = PointsToPixelsY(frmHost.Height)
  This.CanvasWidth = frmHost.Width / This.Settings.PointsPerPixel
  This.CanvasHeight = frmHost.height / This.Settings.PointsPerPixel

  
  If This.CanvasWidth <= 0 Or This.CanvasHeight <= 0 Then
    This.CanvasWidth = IIf(This.CanvasWidth <= 0, 100, This.CanvasWidth)
    This.CanvasHeight = IIf(This.CanvasHeight <= 0, 100, This.CanvasHeight)
  End If
  
  This.MemoryDC = CreateCompatibleDC(This.hDC)
  This.MemoryBitmap = CreateCompatibleBitmap(This.hDC, This.CanvasWidth, This.CanvasHeight)
  This.OldMemoryBitmap = SelectObject(This.MemoryDC, This.MemoryBitmap)
  
  Me.ForeColor = RGB(0, 0, 0)
  Me.BackColor = RGB(255, 255, 255)
  Me.FillColor = BackColor
  Me.PenWidth = 1
 
  Me.CLS
  Redraw
End Sub

' All drawing operations are first done to a memory DC, which is then bitblt'd to the acutal hDC.
' This will need to be adjusted to provide the option of drawing direct to hDC. The benefit to
' drawing to memory first though, is that it's quicker and it avoids flickering.

Public Sub Redraw()
  If This.MemoryDC <> 0 And This.hDC <> 0 And This.CanvasWidth > 0 And This.CanvasHeight > 0 Then
    BitBlt This.hDC, 0, 0, This.CanvasWidth, This.CanvasHeight, This.MemoryDC, 0, 0, SRCCOPY
  End If
End Sub

Public Sub CLS(Optional ByVal ClearColor As OLE_COLOR = -1)
  Dim rc As RECT, hBrush As LongPtr, TargetColor As OLE_COLOR, lngColor As Long
  
  If This.MemoryDC = 0 Then Exit Sub
  
  TargetColor = IIf(ClearColor = -1, Me.BackColor, ClearColor)
  OleTranslateColor TargetColor, 0, lngColor
  
  rc.Left = 0: rc.Top = 0
  rc.Right = This.CanvasWidth
  rc.Bottom = This.CanvasHeight
  
  hBrush = CreateSolidBrush(lngColor)
  If hBrush <> 0 Then
    FillRect This.MemoryDC, rc, hBrush
    DeleteObject hBrush
  End If
End Sub

Public Sub DrawPixel(ByVal TargetX As Long, ByVal TargetY As Long, Optional ByVal pixelColor As OLE_COLOR = -1)
  If This.MemoryDC = 0 Then Exit Sub
  
  Dim lngPixelColor As Long, Result As Long
  OleTranslateColor IIf(pixelColor = -1, Me.ForeColor, pixelColor), 0, lngPixelColor
  
  Result = SetPixel(This.MemoryDC, TargetX, TargetY, lngPixelColor)

End Sub

Public Function Point(ByVal TargetX As Long, ByVal TargetY As Long) As OLE_COLOR
  ' POINT is a method of the VB6 PictureBox which performns the same function as GetPixel.
  Point = GetPixel(This.MemoryDC, TargetX, TargetY)
End Function

Public Function GetPixel(ByVal TargetX As Long, ByVal TargetY As Long) As OLE_COLOR
  GetPixel = GetPixel_(This.MemoryDC, TargetX, TargetY)
End Function

Public Sub DrawRectangle(ByVal x1 As Long, ByVal y1 As Long, ByVal Width As Long, ByVal height As Long, Optional ByVal BorderColor As OLE_COLOR = -1, Optional ByVal CurrentFillColor As OLE_COLOR = -1, Optional ByVal CurrentPenWidth As Long = -1, Optional ByVal CurrentFillStyle As FillStyleConstants = vbFSSolid)
  
  If This.MemoryDC = 0 Then Exit Sub
  
  Dim hPen As LongPtr, hOldPen As LongPtr, hBrush As LongPtr, hOldBrush As LongPtr
  Dim lngBorderCOlor As Long, lngActualFillColor As Long
  Dim effPenWidth As Long
  Dim scaledX1 As Long, scaledY1 As Long, scaledWidth As Long, scaledHeight As Long
  Dim oldBkMode As Long
  Dim brushCreatedForHatch As Boolean
  
  effPenWidth = IIf(CurrentPenWidth = -1, Me.PenWidth, CurrentPenWidth)
  If effPenWidth < 1 Then effPenWidth = 1
  
    OleTranslateColor IIf(BorderColor = -1, Me.ForeColor, BorderColor), 0, lngBorderCOlor
    
    hPen = CreatePen(PS_SOLID, effPenWidth, lngBorderCOlor)
    If hPen = 0 Then Exit Sub
    hOldPen = SelectObject(This.MemoryDC, hPen)
  If CurrentFillColor = -1 Then lngActualFillColor = Me.FillColor Else OleTranslateColor CurrentFillColor, 0, lngActualFillColor
  
  brushCreatedForHatch = False
  If CurrentFillStyle = vbFSTransparent Or CurrentFillColor = -1 And CurrentFillStyle <> vbFSSolid Then ' If explicitly transparent or no fill color given for non-solid
    hBrush = GetStockObject(NULL_BRUSH)
  ElseIf CurrentFillStyle = vbFSSolid Then
    hBrush = CreateSolidBrush(lngActualFillColor)
    If hBrush = 0 Then hBrush = GetStockObject(NULL_BRUSH)
  Else
    Dim hatchIndex As Long: hatchIndex = CurrentFillStyle - 2
    hBrush = CreateHatchBrush(hatchIndex, lngActualFillColor)
    If hBrush = 0 Then hBrush = GetStockObject(NULL_BRUSH) Else brushCreatedForHatch = True
  End If
  hOldBrush = SelectObject(This.MemoryDC, hBrush)
  
  If brushCreatedForHatch Then oldBkMode = SetBkMode(This.MemoryDC, BKMODE_TRANSPARENT)
  
  Rectangle This.MemoryDC, x1, y1, x1 + Width, y1 + height
  If brushCreatedForHatch Then SetBkMode This.MemoryDC, oldBkMode
  SelectObject This.MemoryDC, hOldBrush
  SelectObject This.MemoryDC, hOldPen
  DeleteObject hPen
  If hBrush <> GetStockObject(NULL_BRUSH) Then DeleteObject hBrush
End Sub

Public Sub DrawCircle(ByVal centerX As Long, _
                      ByVal centerY As Long, _
                      ByVal radius As Long, _
                      Optional ByVal BorderColor As OLE_COLOR = -1, _
                      Optional ByVal CurrentFillColor As OLE_COLOR = -1, _
                      Optional ByVal CurrentPenWidth As Long = -1, _
                      Optional ByVal CurrentFillStyle As FillStyleConstants = vbFSSolid)

  Dim hPen As LongPtr, hOldPen As LongPtr, hBrush As LongPtr, hOldBrush As LongPtr, lngBorderCOlor As Long, lngActualFillColor As Long
  Dim effPenWidth As Long, oldBkMode As Long
  
  effPenWidth = IIf(CurrentPenWidth = -1, Me.PenWidth, CurrentPenWidth)
  If effPenWidth < 1 Then effPenWidth = 1
  
  OleTranslateColor IIf(BorderColor = -1, Me.ForeColor, BorderColor), 0, lngBorderCOlor
  
  hPen = CreatePen(PS_SOLID, effPenWidth, lngBorderCOlor)
  If hPen = 0 Then Exit Sub
  
  hOldPen = SelectObject(This.MemoryDC, hPen)
  If CurrentFillColor = -1 Then lngActualFillColor = Me.FillColor Else OleTranslateColor CurrentFillColor, 0, lngActualFillColor
  
  
  If CurrentFillStyle = vbFSTransparent Or CurrentFillColor = -1 And CurrentFillStyle <> vbFSSolid Then
    hBrush = GetStockObject(NULL_BRUSH)
  ElseIf CurrentFillStyle = vbFSSolid Then
    hBrush = CreateSolidBrush(lngActualFillColor)
    If hBrush = 0 Then hBrush = GetStockObject(NULL_BRUSH)
  End If
  hOldBrush = SelectObject(This.MemoryDC, hBrush)
  
  Ellipse This.MemoryDC, centerX - radius, centerY - radius, centerX + radius, centerY + radius
  
  SelectObject This.MemoryDC, hOldBrush
  SelectObject This.MemoryDC, hOldPen
  DeleteObject hPen
  If hBrush <> GetStockObject(NULL_BRUSH) Then DeleteObject hBrush
End Sub

Public Sub DrawLine(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, Optional ByVal LineColor As OLE_COLOR = -1, Optional ByVal CurrentPenWidth As Long = -1)
  If This.MemoryDC = 0 Then Exit Sub
  Dim hPen As LongPtr, hOldPen As LongPtr
  Dim lngLineColor As Long
  Dim effPenWidth As Long
  Dim pt As POINTAPI
  Dim lngBorderCOlor As Long
  
  effPenWidth = IIf(CurrentPenWidth = -1, Me.PenWidth, CurrentPenWidth)
  If effPenWidth < 1 Then effPenWidth = 1
  
  OleTranslateColor IIf(LineColor = -1, Me.ForeColor, LineColor), 0, lngLineColor
  
  hPen = CreatePen(PS_SOLID, effPenWidth, lngLineColor)
  If hPen = 0 Then Exit Sub
  hOldPen = SelectObject(This.MemoryDC, hPen)
  
  MoveToEx This.MemoryDC, x1, y1, pt
  LineTo This.MemoryDC, x2, y2
  
  SelectObject This.MemoryDC, hOldPen
  DeleteObject hPen
End Sub

Function ARGB(ByVal Alpha As Long, ByVal RGBValue As Long) As Long
  ' Source: https://www.vbforums.com/showthread.php?623200-RGB-ARGB-Colours
  If Alpha > 127 Then ' handle high bit and prevent overflow
    ARGB = RGBValue Or ((Alpha And Not &H80) * &H1000000) Or &H80000000
  Else
    ARGB = RGBValue Or (Alpha * &H1000000)
  End If
End Function

Public Function LoadPictureEx(ByVal FilePath As String) As StdPicture
  Dim hGdipImage As LongPtr, hBitmap As LongPtr, WidthPX As Long, HeightPx As Long
  Dim Pic As StdPicture
  Dim Result As Long
  
  Result = GdipLoadImageFromFile(StrPtr(FilePath), hGdipImage)
  If Result <> GP_STAT_OK Or hGdipImage = 0 Then Exit Function
  
  Result = GdipGetImageWidth(hGdipImage, WidthPX)
  Result = GdipGetImageHeight(hGdipImage, HeightPx)
  
  Result = GdipCreateHBITMAPFromBitmap(hGdipImage, hBitmap, 0)
  GdipDisposeImage hGdipImage
  If Result <> GP_STAT_OK Or hBitmap = 0 Then Exit Function
  
  Set Pic = BitmapToPicture(hBitmap)
  Set LoadPictureEx = Pic
End Function

Public Function SavePictureEx(ByVal FilePath As String, _
                              Optional ByVal FileType As PicFileTypeEnum = pictypePNG, _
                              Optional ByVal Quality As Long = 90) As Boolean
                              
 ' To be added
 
End Function

Public Sub RenderPicture(ByRef Pic As StdPicture, _
                         ByVal X As Long, _
                         ByVal Y As Long, _
                         Optional ByVal DestWidth As Long = -1, _
                         Optional ByVal DestHeight As Long = -1, _
                         Optional ByVal TransparentKeyColor As OLE_COLOR = -1)
  
  If This.MemoryDC = 0 Or Pic Is Nothing Then Exit Sub
  
  Dim hSrcDC As LongPtr
  Dim hOldBitmap As LongPtr
  Dim srcWidthPx As Long, srcHeightPx As Long
  Dim finalDestWidth As Long, finalDestHeight As Long
  Dim lngTransColor As Long
  
  hSrcDC = CreateCompatibleDC(This.MemoryDC)
  If hSrcDC = 0 Then Exit Sub
  hOldBitmap = SelectObject(hSrcDC, Pic.Handle)
  
  srcWidthPx = PointsToPixelsX(Pic.Width)
  srcHeightPx = PointsToPixelsY(Pic.height)
  
  finalDestWidth = IIf(DestWidth = -1, srcWidthPx, DestWidth)
  finalDestHeight = IIf(DestHeight = -1, srcHeightPx, DestHeight)
  
  If TransparentKeyColor = -1 Then
    StretchBlt This.MemoryDC, X, Y, finalDestWidth, finalDestHeight, hSrcDC, 0, 0, srcWidthPx, srcHeightPx, SRCCOPY
  Else
    OleTranslateColor TransparentKeyColor, 0, lngTransColor
    TransparentBlt This.MemoryDC, X, Y, finalDestWidth, finalDestHeight, hSrcDC, 0, 0, srcWidthPx, srcHeightPx, lngTransColor
  End If
  
  SelectObject hSrcDC, hOldBitmap
  DeleteDC hSrcDC
End Sub

Public Function GetImageFileDimensions(ByVal FilePath As String, ByRef WidthPX As Long, ByRef HeightPx As Long) As Boolean
  Dim hGdipImage As LongPtr, Result As Long, Success As Boolean
  
  If Len(Dir(FilePath)) = 0 Then Debug.Print "File not found. Exiting":   Exit Function
  ' Use GetObject instead?
  Result = GdipLoadImageFromFile(StrPtr(FilePath), hGdipImage)
  
  If hGdipImage <> 0 Then
    Result = GdipGetImageWidth(hGdipImage, WidthPX)
    Result = GdipGetImageHeight(hGdipImage, HeightPx)
    GetImageFileDimensions = True
  End If
  
  If WidthPX = 0 Or HeightPx = 0 Then Debug.Print "GdipLoadImageFromFile failed for '" & FilePath & ""
  
  If hGdipImage <> 0 Then
    GdipDisposeImage hGdipImage
    hGdipImage = 0
  End If
  
End Function

' --- Helper Functions ---
Public Function PointsToPixelsX(Points As Single) As Long
  PointsToPixelsX = CLng(This.Settings.PointsPerPixel * Points)
End Function

Private Function PointsToPixelsY(Points As Single) As Long
  PointsToPixelsY = CLng(This.Settings.PointsPerPixel * Points)
End Function

Private Function BitmapToPicture(ByVal hBitmap As LongPtr) As StdPicture
  Dim Pic As StdPicture, PicDesc As PICTDESC, IID_IPicture As GUID
  If hBitmap = 0 Then Exit Function
  With IID_IPicture
    .Data1 = &H7BF80980: .Data2 = &HBF32: .Data3 = &H101A
    .Data4(0) = &H8B: .Data4(1) = &HBB: .Data4(2) = &H0: .Data4(3) = &HAA
    .Data4(4) = &H0: .Data4(5) = &H30: .Data4(6) = &HC: .Data4(7) = &HAB
  End With
  With PicDesc
    .size = LenB(PicDesc): .Type = 1: .hPic = hBitmap: .hPal = 0
  End With
  OleCreatePictureIndirect PicDesc, IID_IPicture, True, Pic
  Set BitmapToPicture = Pic
End Function

Private Function ClampLng(Value As Long, MinVal As Long, MaxVal As Long) As Long
  If Value < MinVal Then
    ClampLng = MinVal
  ElseIf Value > MaxVal Then
    ClampLng = MaxVal
  Else
    ClampLng = Value
  End If
End Function

Private Sub CleanUpResources()
  If This.MemoryDC <> 0 Then
    If This.OldMemoryBitmap <> 0 Then SelectObject This.MemoryDC, This.OldMemoryBitmap: This.OldMemoryBitmap = 0
  End If
  If This.MemoryBitmap <> 0 Then DeleteObject This.MemoryBitmap: This.MemoryBitmap = 0
  If This.MemoryDC <> 0 Then DeleteDC This.MemoryDC: This.MemoryDC = 0
  If This.hDC <> 0 And This.hWnd <> 0 Then ReleaseDC This.hWnd, This.hDC: This.hDC = 0
  This.hWnd = 0
  Set HostFrame = Nothing
End Sub

Private Sub Class_Initialize()
  This.Settings.DPI = GetDpiForWindow(Application.hWnd)
  This.Settings.PixelsPerPoint = This.Settings.DPI / 72
  This.Settings.PointsPerPixel = 72 / This.Settings.DPI
End Sub

Private Sub Class_Terminate()
  CleanUpResources
End Sub

Private Sub HostFrame_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub HostFrame_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub HostFrame_Keyup(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

' The host frames layout event is co-opted as a pseudo Resize event, which is not ordinarily
' available through event sinks. This method of updated the internal variable by referencing the
' property is because the property calculates the repsonse basewd on the HostFrame dimensions
' as at that time. Arguably, better to do this there rather than here.
Private Sub HostFrame_Layout()
  '
End Sub

' The vbaPictureBox recalculates all mouse movements to pixels. This code can be improved.
Private Sub HostFrame_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AdjustMouseCoordinates X, Y
  RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub HostFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AdjustMouseCoordinates X, Y
  RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub HostFrame_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AdjustMouseCoordinates X, Y
  RaiseEvent MouseUp(Button, Shift, X, Y)
  
  If Button = vbKeyLButton Then

  End If
  
  If Shift = 1 Then           ' Shift Button = 1
    '
  ElseIf Shift = 2 Then       ' Control Button = 2
    '
  ElseIf Shift = 4 Then       ' Alt Button = 4
    '
  End If
End Sub

Private Sub AdjustMouseCoordinates(ByRef X As Single, ByRef Y As Single)
  X = X / This.Settings.PointsPerPixel
  Y = Y / This.Settings.PointsPerPixel
End Sub

Sub Pause(Optional ByVal Period As Single = 1)
  Period = Period + Timer
  Do
    DoEvents
  Loop While Period > Timer
End Sub
Attribute VB_Name = "modImageSearch"

'Author: https://github.com/KallunWillock

Public Type SearchResult
  FoundTarget As Boolean
  FoundX As Long
  FoundY As Long
  SearchAreaDescription As String
  TemplateSizeDescription As String
  PixelsInSearchArea As Long
  PixelComparisonsMade As Double
  Message As String
End Type

Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function BitBlt Lib "gdi32" (ByVal hDestDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function GetDIBits Lib "gdi32" (ByVal hdc As LongPtr, ByVal hBitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Const SRCCOPY As Long = &HCC0020
Private Const DIB_RGB_COLORS As Long = 0

Private Type RGBQUAD
  rgbBlue As Byte
  rgbGreen As Byte
  rgbRed As Byte
  rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type

Private Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors As RGBQUAD
End Type

Public Function ImageSearch(ByVal SearchLeft As Long, ByVal SearchTop As Long, ByVal SearchWidth As Long, ByVal SearchHeight As Long, ByVal TemplatePath As String) As SearchResult
  Dim Report As SearchResult
  Dim ScreenDC As LongPtr, MemoryDC As LongPtr
  Dim hBitmap As LongPtr, hOldBitmap As LongPtr, TemplateW As Long, TemplateH As Long
  Dim TemplatePic As stdole.IPicture, ScreenPixels() As Byte, templatePixels() As Byte
  Dim ScreenWidth As Long, ScreenHeight As Long, TemplateWidth As Long, TemplateHeight As Long
  
  Dim x As Long, y As Long, i As Long, j As Long
  Dim MatchFound As Boolean, PixelComparisons As Double
  
  Report.FoundTarget = False
  Report.FoundX = -1
  Report.FoundY = -1
  Report.SearchAreaDescription = "Rect(" & SearchLeft & ", " & SearchTop & ", " & SearchWidth & ", " & SearchHeight & ")"
  Report.PixelsInSearchArea = SearchWidth * SearchHeight
  PixelComparisons = 0
  
  On Error Resume Next
  Set TemplatePic = LoadPictureEx(TemplatePath)
  On Error GoTo 0
  If TemplatePic Is Nothing Then 
    Report.Message = "No image found: " & TemplatePath
    ImageSearch = Report
    Exit Function
  End If
  
  TemplateW = TemplatePic.Width \ 26.6667
  TemplateH = TemplatePic.Height \ 26.6667
  Report.TemplateSizeDescription = TemplateW & "x" & TemplateH & " pixels"
  
  ScreenDC = GetDC(0)
  MemoryDC = CreateCompatibleDC(ScreenDC)
  hBitmap = CreateCompatibleBitmap(ScreenDC, SearchWidth, SearchHeight)
  hOldBitmap = SelectObject(MemoryDC, hBitmap)
  BitBlt MemoryDC, 0, 0, SearchWidth, SearchHeight, ScreenDC, SearchLeft, SearchTop, SRCCOPY
  
  Call GetBitmapPixels(hBitmap, ScreenPixels, ScreenWidth, ScreenHeight)
  Call GetBitmapPixels(TemplatePic.handle, templatePixels, TemplateW, TemplateH)
    
  For y = 0 To ScreenHeight - TemplateH
    For x = 0 To ScreenWidth - TemplateW
      MatchFound = True
      For j = 0 To TemplateH - 1
        For i = 0 To TemplateW - 1
          PixelComparisons = PixelComparisons + 1

          If templatePixels(0, i, j) <> ScreenPixels(0, x + i, y + j) Or templatePixels(1, i, j) <> ScreenPixels(1, x + i, y + j) Or templatePixels(2, i, j) <> ScreenPixels(2, x + i, y + j) Then
            MatchFound = False
            Exit For
          End If
        Next i
        If Not MatchFound Then Exit For
      Next j
      
      If MatchFound Then
        Report.FoundTarget = True
        Report.FoundX = SearchLeft + x
        Report.FoundY = SearchTop + y
        GoTo FinishReport
      End If
    Next x
  Next y
  
FinishReport:
  Report.PixelComparisonsMade = PixelComparisons
  If Report.FoundTarget Then
    Report.Message = "Image found." & vbCrLf & "Location (X, Y): (" & Report.FoundX & ", " & Report.FoundY & ")" & vbCrLf & "Searched Area: " & Report.SearchAreaDescription & vbCrLf & "Template Size: " & Report.TemplateSizeDescription & vbCrLf & "Pixel Comparisons: " & format(PixelComparisons, "#,##0")
  Else
    Report.Message = "Image not found." & vbCrLf & "Searched Area: " & Report.SearchAreaDescription & vbCrLf & "Template Size: " & Report.TemplateSizeDescription & vbCrLf & "Total Pixel Comparisons: " & format(PixelComparisons, "#,##0")
  End If
  
Cleanup:
  If hOldBitmap <> 0 Then SelectObject MemoryDC, hOldBitmap
  If hBitmap <> 0 Then DeleteObject hBitmap
  If MemoryDC <> 0 Then DeleteDC MemoryDC
  If ScreenDC <> 0 Then ReleaseDC 0, ScreenDC
  Set TemplatePic = Nothing
  ImageSearch = Report
End Function

Private Function GetBitmapPixels(ByVal hBitmap As LongPtr, ByRef OutPixels() As Byte, ByRef Width As Long, ByRef Height As Long) As Boolean
  On Error GoTo ErrorHandler
  Dim bmi As BITMAPINFO
  Dim hdc As LongPtr
  hdc = CreateCompatibleDC(0)
  If hdc = 0 Then Exit Function
  bmi.bmiHeader.biSize = Len(bmi.bmiHeader)
  If GetDIBits(hdc, hBitmap, 0, 0, ByVal 0&, bmi, DIB_RGB_COLORS) = 0 Then GoTo ErrorHandler
  With bmi.bmiHeader
    Width = .biWidth
    Height = Abs(.biHeight)
    .biHeight = -Height
    .biBitCount = 32
    .biCompression = 0
  End With
  ReDim OutPixels(0 To 3, 0 To Width - 1, 0 To Height - 1)
  If GetDIBits(hdc, hBitmap, 0, Height, OutPixels(0, 0, 0), bmi, DIB_RGB_COLORS) > 0 Then
    GetBitmapPixels = True
  End If
ErrorHandler:
  If hdc <> 0 Then DeleteDC hdc
End Function

Public Sub TestThis()
  Dim Report As SearchResult
  Dim imagePath As String
  imagePath = "C:\CODE\icons\template.png"
  Report = ImageSearch(0, 0, 1920, 1280, imagePath)
  Debug.Print Report.Message, IIf(Report.FoundTarget, vbInformation, vbExclamation), "Image Search Report"
  If Report.FoundTarget Then
    Debug.Print "Found at: " & Report.FoundX, Report.FoundY
  Else
    Debug.Print "Search failed after " & Report.PixelComparisonsMade & " comparisons."
  End If
End Sub

Private Function LoadPictureEx(Optional ByVal Filename As String, Optional ByVal LoadToPBox As Boolean = False) As StdPicture
  With CreateObject("WIA.Imagefile")
    .LoadFile Filename
    Set LoadPictureEx = .FileData.Picture
  End With
End Function





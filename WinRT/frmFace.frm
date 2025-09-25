VERSION 5.00
Begin VB.Form frmFace 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FaceDetector"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   11295
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pbImage 
      Height          =   7935
      Left            =   0
      ScaleHeight     =   7875
      ScaleWidth      =   11235
      TabIndex        =   0
      Top             =   0
      Width           =   11295
   End
   Begin VB.Menu mnu_OpenImage 
      Caption         =   "&Open Image"
   End
End
Attribute VB_Name = "frmFace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' Namespace Windows
Private Windows As New Windows

' Namespace Windows.Media.FaceAnalysis.FaceDetector
Private FaceDetector As FaceDetector

Private Sub Form_Load()
    pbImage.ScaleMode = vbPixels
    pbImage.AutoRedraw = True
    pbImage.ForeColor = vbRed
    pbImage.DrawWidth = 2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SetAllNothing
End Sub

Private Sub SetAllNothing()
    If IsNotNothing(FaceDetector) Then Set FaceDetector = Nothing
End Sub

Private Sub mnu_OpenImage_Click()
    Dim FileOpenPicker As FileOpenPicker
    Set FileOpenPicker = Windows.Storage.Pickers.FileOpenPicker
    If IsNotNothing(FileOpenPicker) Then
        FileOpenPicker.ParentHwnd = Me.hwnd
        Dim FileTypeFilter As List_String
        Set FileTypeFilter = FileOpenPicker.FileTypeFilter
        If IsNotNothing(FileTypeFilter) Then
            Call FileTypeFilter.Append(".bmp")
            Call FileTypeFilter.Append(".gif")
            Call FileTypeFilter.Append(".jpg")
            Call FileTypeFilter.Append(".png")
            Dim StorageFile As StorageFile
            Set StorageFile = FileOpenPicker.PickSingleFileAsync
            If IsNotNothing(StorageFile) Then
                Call SetAllNothing
                Call OpenImage(StorageFile)
            End If
        End If
    End If
End Sub

Public Sub OpenImage(ByVal imageFile As StorageFile)
    If IsNotNothing(imageFile) Then
        Dim RandomAccessStream As RandomAccessStream
        Set RandomAccessStream = imageFile.OpenAsync(FileAccessMode_Read)
        If IsNotNothing(RandomAccessStream) Then
            pbImage.Picture = GetPictureFromIStream(RandomAccessStream.ToIStream)
            Dim SoftwareBitmap As SoftwareBitmap
            Set SoftwareBitmap = Windows.Graphics.Imaging.BitmapDecoder.CreateAsync(RandomAccessStream).GetSoftwareBitmapAsync
            If IsNotNothing(SoftwareBitmap) Then
                Call FaceDetection(SoftwareBitmap)
            End If
        End If
    End If
End Sub

Public Sub FaceDetection(ByVal bitmap As SoftwareBitmap)
    If IsNotNothing(bitmap) Then
        Set FaceDetector = Windows.Media.FaceAnalysis.FaceDetector.CreateAsync
        If IsNotNothing(FaceDetector) Then
            If FaceDetector.IsSupported Then
                Debug.Print "FaceDetector.IsSupported = " & FaceDetector.IsSupported
                Debug.Print "FaceDetector.MinDetectableFaceSize = " & FaceDetector.MinDetectableFaceSize.ToString
                Debug.Print "FaceDetector.MaxDetectableFaceSize = " & FaceDetector.MaxDetectableFaceSize.ToString
                Dim SupportedPixelFormats As ReadOnlyList_1 'ReadOnlyList_BitmapPixelFormat
                Set SupportedPixelFormats = FaceDetector.GetSupportedBitmapPixelFormats
                If IsNotNothing(SupportedPixelFormats) Then
                    Dim PixelFormatCount As Long
                    PixelFormatCount = SupportedPixelFormats.Size
                    If PixelFormatCount > 0& Then
                        Dim PixelFormatItem As Long
                        For PixelFormatItem = 0 To PixelFormatCount - 1
                            Debug.Print "SupportedPixelFormat = " & SupportedPixelFormats.GetAt(PixelFormatItem)
                        Next
                    End If
                    Dim convBitmap As SoftwareBitmap
                    Set convBitmap = Windows.Graphics.Imaging.SoftwareBitmap.Convert(bitmap, SupportedPixelFormats.GetAt(0))
                    If IsNotNothing(convBitmap) Then
                        Dim DetectedFaces As ReadOnlyList_1 'ReadOnlyList_DetectedFace
                        Set DetectedFaces = FaceDetector.DetectFacesAsync(convBitmap)
                        If IsNotNothing(DetectedFaces) Then
                            Dim DetectedFaceCount As Long
                            DetectedFaceCount = DetectedFaces.Size
                            Debug.Print "DetectedFaces = " & DetectedFaceCount
                            If DetectedFaceCount > 0& Then
                                Dim DetectedFaceItem As Long
                                For DetectedFaceItem = 0 To DetectedFaceCount - 1
                                    Dim BitmapBounds As BitmapBounds
                                    Set BitmapBounds = DetectedFaces.GetAt(DetectedFaceItem).FaceBox
                                    If IsNotNothing(BitmapBounds) Then
                                        Debug.Print "DetectedFace.FaceBox.BitmapBounds = " & BitmapBounds.ToString
                                        pbImage.Line (BitmapBounds.X, BitmapBounds.Y)-(BitmapBounds.X + _
                                                                                       BitmapBounds.Width, _
                                                                                       BitmapBounds.Y + _
                                                                                       BitmapBounds.Height), _
                                                                                       pbImage.ForeColor, B
                                    
                                        Set BitmapBounds = Nothing
                                    End If
                                Next
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

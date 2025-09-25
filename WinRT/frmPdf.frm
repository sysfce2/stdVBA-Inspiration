VERSION 5.00
Begin VB.Form frmPdf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PdfViewer"
   ClientHeight    =   7950
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   11295
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pbPdf 
      Height          =   7935
      Left            =   0
      ScaleHeight     =   7875
      ScaleWidth      =   10965
      TabIndex        =   1
      Top             =   0
      Width           =   11025
   End
   Begin VB.VScrollBar sbPdf 
      Height          =   7935
      Left            =   11040
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Menu mnu_OpenPdf 
      Caption         =   "&Open Pdf"
   End
End
Attribute VB_Name = "frmPdf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' Namespace Windows
Private Windows As New Windows

' Namespace Windows.Data.Pdf
Private PdfDocument As PdfDocument

Private Sub Form_Load()
    pbPdf.ScaleMode = vbPixels
    pbPdf.AutoRedraw = True
    sbPdf.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SetAllNothing
End Sub

Private Sub SetAllNothing()
    If IsNotNothing(PdfDocument) Then Set PdfDocument = Nothing
End Sub

Private Sub mnu_OpenPdf_Click()
    Dim FileOpenPicker As FileOpenPicker
    Set FileOpenPicker = Windows.Storage.Pickers.FileOpenPicker
    If IsNotNothing(FileOpenPicker) Then
        FileOpenPicker.ParentHwnd = Me.hWnd
        Dim FileTypeFilter As List_String
        Set FileTypeFilter = FileOpenPicker.FileTypeFilter
        If IsNotNothing(FileTypeFilter) Then
            Call FileTypeFilter.Append(".pdf")
            Dim StorageFile As StorageFile
            Set StorageFile = FileOpenPicker.PickSingleFileAsync
            If IsNotNothing(StorageFile) Then
                Call SetAllNothing
                Call OpenPdf(StorageFile)
            End If
        End If
    End If
End Sub

Private Sub OpenPdf(ByVal pdfFile As StorageFile)
    If IsNotNothing(pdfFile) Then
        Debug.Print "StorageFile.Name = " & pdfFile.Name
        Set PdfDocument = Windows.Data.Pdf.PdfDocument.LoadFromFileAsync(pdfFile)
        If IsNothing(PdfDocument) Then
            'passwort protected
            'Set PdfDocument = Windows.Data.Pdf.PdfDocument.LoadFromFileWithPasswordAsync(pdfFile, "password")
        End If
        If IsNotNothing(PdfDocument) Then
            Debug.Print "PdfDocument.IsPasswordProtected = " & CStr(PdfDocument.IsPasswordProtected)
            Debug.Print "PdfDocument.PageCount = " & CStr(PdfDocument.PageCount)
            With sbPdf
                .Min = 0
                .value = 0
                .Max = PdfDocument.PageCount - 1
                .Enabled = True
            End With
            Call GetPdfPage(PdfDocument, 0)
        End If
    End If
End Sub

Private Sub GetPdfPage(ByVal pdfDoc As PdfDocument, ByVal pageIndex As Long)
    pbPdf.SetFocus
    If IsNotNothing(pdfDoc) Then
        If pdfDoc.PageCount > 0 And _
           pageIndex <= pdfDoc.PageCount Then
            Dim PdfPage As PdfPage
            Set PdfPage = pdfDoc.GetPage(pageIndex)
            If IsNotNothing(PdfPage) Then
                Me.Caption = "PdfViewer: Seite " & CStr(pageIndex + 1) & " von " & CStr(pdfDoc.PageCount)
                Debug.Print "PdfPage.Index = " & CStr(PdfPage.index)
                Debug.Print "PdfPage.Size = " & PdfPage.Size.ToString
                Debug.Print "PdfPage.Rotation = " & CStr(PdfPage.Rotation)
                Debug.Print "PdfPage.PreferredZoom = " & CStr(PdfPage.PreferredZoom)
                Debug.Print "PdfPage.Dimensions.ArtBox = " & PdfPage.Dimensions.ArtBox.ToString
                Debug.Print "PdfPage.Dimensions.BleedBox = " & PdfPage.Dimensions.BleedBox.ToString
                Debug.Print "PdfPage.Dimensions.CropBox = " & PdfPage.Dimensions.CropBox.ToString
                Debug.Print "PdfPage.Dimensions.MediaBox = " & PdfPage.Dimensions.MediaBox.ToString
                Debug.Print "PdfPage.Dimensions.TrimBox = " & PdfPage.Dimensions.TrimBox.ToString
                If PdfPage.PreparePageAsync Then
                    Dim pIStream As Long
                    pIStream = SHCreateMemStream(0&, 0&)
                    If pIStream <> 0& Then
                        Dim RandomAccessStream As RandomAccessStream
                        Set RandomAccessStream = Windows.Storage.Streams.RandomAccessStream.FromIStream(pIStream, BSOS_DEFAULT)
                        If IsNotNothing(RandomAccessStream) Then
                            Dim Size As Size
                            Set Size = PdfPage.Size
                            Dim RatioX As Single
                            Dim RatioY As Single
                            Dim PageRatio As Single
                            RatioX = CSng((pbPdf.ScaleWidth - 20) / Size.Width)
                            RatioY = CSng((pbPdf.ScaleHeight - 20) / Size.Height)
                            If RatioX > RatioY Then PageRatio = RatioY Else PageRatio = RatioX
                            Dim PdfPageRenderOptions As PdfPageRenderOptions
                            Set PdfPageRenderOptions = Windows.Data.Pdf.PdfPageRenderOptions
                            If IsNotNothing(PdfPageRenderOptions) Then
                                PdfPageRenderOptions.DestinationWidth = Size.Width * PageRatio
                                PdfPageRenderOptions.DestinationHeight = Size.Height * PageRatio
                                PdfPageRenderOptions.BackgroundColor = Windows.UI.Colors.White
                                If PdfPage.RenderWithOptionsToStreamAsync(RandomAccessStream, PdfPageRenderOptions) Then
                                    Dim PdfPicture As StdPicture
                                    Set PdfPicture = GetPictureFromIStream(pIStream)
                                    Dim DrawTop As Long
                                    Dim DrawLeft As Long
                                    Dim DrawRight As Long
                                    Dim DrawBottom As Long
                                    DrawRight = CLng(pbPdf.ScaleX(PdfPicture.Width, vbHimetric, vbPixels))
                                    DrawBottom = CLng(pbPdf.ScaleY(PdfPicture.Height, vbHimetric, vbPixels))
                                    DrawTop = CLng((pbPdf.ScaleHeight - DrawBottom) \ 2)
                                    DrawLeft = CLng((pbPdf.ScaleWidth - DrawRight) \ 2)
                                    pbPdf.Cls
                                    Call pbPdf.PaintPicture(PdfPicture, DrawLeft, DrawTop)
                                    pbPdf.Line (DrawLeft - 1&, DrawTop - 1&)- _
                                               (DrawLeft + DrawRight, _
                                                DrawTop + DrawBottom), &H0&, B
                                    Set PdfPicture = Nothing
                                End If
                            End If
                        End If
                        Call ReleaseIfc(pIStream)
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub sbPdf_Change()
    Call GetPdfPage(PdfDocument, sbPdf.value)
End Sub

Private Sub sbPdf_Scroll()
    Call GetPdfPage(PdfDocument, sbPdf.value)
End Sub

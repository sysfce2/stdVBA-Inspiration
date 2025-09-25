VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitmapDecoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Graphics_Imaging_BitmapDecoder As String = "Windows.Graphics.Imaging.BitmapDecoder"
Private Const IID_IBitmapFrame As String = "{72a49a1c-8081-438d-91bc-94ecfc8185c6}"
Private Const IID_IBitmapDecoder As String = "{acef22ba-1d74-4c91-9dfc-9620745233e6}"
Private Const IID_IBitmapDecoderStatics As String = "{438ccb26-bcef-4e95-bad6-23a822e58d01}"
Private Const IID_IBitmapDecoderStatics2 As String = "{50ba68ea-99a1-40c4-80d9-aef0dafa6c3f}"
Private Const IID_IBitmapFrameWithSoftwareBitmap As String = "{fe287c9a-420c-4963-87ad-691436e08383}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IBitmapDecoderStatics
    IBitmapDecoderStatics_GetBmpDecoderId = 6
    IBitmapDecoderStatics_GetJpegDecoderId = 7
    IBitmapDecoderStatics_GetPngDecoderId = 8
    IBitmapDecoderStatics_GetTiffDecoderId = 9
    IBitmapDecoderStatics_GetGifDecoderId = 10
    IBitmapDecoderStatics_GetJpegXRDecoderId = 11
    IBitmapDecoderStatics_GetIcoDecoderId = 12
    IBitmapDecoderStatics_GetDecoderInformationEnumerator = 13
    IBitmapDecoderStatics_CreateAsync = 14
    IBitmapDecoderStatics_CreateWithIdAsync = 15
                        
    ' IBitmapDecoderStatics2
    IBitmapDecoderStatics2_GetHeifDecoderId = 6
    IBitmapDecoderStatics2_GetWebpDecoderId = 7

    ' IBitmapDecoder
    IBitmapDecoder_GetBitmapContainerProperties = 6
    IBitmapDecoder_GetDecoderInformation = 7
    IBitmapDecoder_GetFrameCount = 8
    IBitmapDecoder_GetPreviewAsync = 9
    IBitmapDecoder_GetFrameAsync = 10

    ' IBitmapFrame
    IBitmapFrame_GetThumbnailAsync = 6
    IBitmapFrame_GetBitmapProperties = 7
    IBitmapFrame_GetBitmapPixelFormat = 8
    IBitmapFrame_GetBitmapAlphaMode = 9
    IBitmapFrame_GetDpiX = 10
    IBitmapFrame_GetDpiY = 11
    IBitmapFrame_GetPixelWidth = 12
    IBitmapFrame_GetPixelHeight = 13
    IBitmapFrame_GetOrientedPixelWidth = 14
    IBitmapFrame_GetOrientedPixelHeight = 15
    IBitmapFrame_GetPixelDataAsync = 16
    IBitmapFrame_GetPixelDataTransformedAsync = 17
    
    ' IBitmapFrameWithSoftwareBitmap
    IBitmapFrameWithSoftwareBitmap_GetSoftwareBitmapAsync = 6
    IBitmapFrameWithSoftwareBitmap_GetSoftwareBitmapConvertedAsync = 7
    IBitmapFrameWithSoftwareBitmap_GetSoftwareBitmapTransformedAsync = 8
                        
End Enum

' ----==== Variablen ====----
Private m_pIBitmapFrame As Long
Private m_pIBitmapDecoder As Long
Private m_pIBitmapDecoderStatics As Long
Private m_pIBitmapDecoderStatics2 As Long
Private m_pIBitmapFrameWithSoftwareBitmap As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Graphics_Imaging_BitmapDecoder, _
                            IID_IBitmapDecoderStatics, _
                            m_pIBitmapDecoderStatics) Then
        Call QueryIfc(m_pIBitmapDecoderStatics, _
                      IID_IBitmapDecoderStatics2, _
                      m_pIBitmapDecoderStatics2)
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIBitmapFrameWithSoftwareBitmap)
    Call ReleaseIfc(m_pIBitmapFrame)
    Call ReleaseIfc(m_pIBitmapDecoder)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIBitmapDecoder
End Property

Public Property Let Ifc(ByVal pIBitmapDecoder As Long)
    Call ReleaseIfc(m_pIBitmapFrameWithSoftwareBitmap)
    Call ReleaseIfc(m_pIBitmapFrame)
    Call ReleaseIfc(m_pIBitmapDecoder)
    m_pIBitmapDecoder = pIBitmapDecoder
    Call QueryIfc(m_pIBitmapDecoder, IID_IBitmapFrame, m_pIBitmapFrame)
    Call QueryIfc(m_pIBitmapDecoder, IID_IBitmapFrameWithSoftwareBitmap, m_pIBitmapFrameWithSoftwareBitmap)
End Property

' IBitmapDecoderStatics
Friend Property Get BmpDecoderId() As GUID
    Dim Ret As GUID
    If m_pIBitmapDecoderStatics <> 0& Then
        Dim value As GUID
        If Invoke(m_pIBitmapDecoderStatics, _
                  IBitmapDecoderStatics_GetBmpDecoderId, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    BmpDecoderId = Ret
End Property

Friend Property Get JpegDecoderId() As GUID
    Dim Ret As GUID
    If m_pIBitmapDecoderStatics <> 0& Then
        Dim value As GUID
        If Invoke(m_pIBitmapDecoderStatics, _
                  IBitmapDecoderStatics_GetJpegDecoderId, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    JpegDecoderId = Ret
End Property

Friend Property Get PngDecoderId() As GUID
    Dim Ret As GUID
    If m_pIBitmapDecoderStatics <> 0& Then
        Dim value As GUID
        If Invoke(m_pIBitmapDecoderStatics, _
                  IBitmapDecoderStatics_GetPngDecoderId, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    PngDecoderId = Ret
End Property

Friend Property Get TiffDecoderId() As GUID
    Dim Ret As GUID
    If m_pIBitmapDecoderStatics <> 0& Then
        Dim value As GUID
        If Invoke(m_pIBitmapDecoderStatics, _
                  IBitmapDecoderStatics_GetTiffDecoderId, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    TiffDecoderId = Ret
End Property

Friend Property Get GifDecoderId() As GUID
    Dim Ret As GUID
    If m_pIBitmapDecoderStatics <> 0& Then
        Dim value As GUID
        If Invoke(m_pIBitmapDecoderStatics, _
                  IBitmapDecoderStatics_GetGifDecoderId, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    GifDecoderId = Ret
End Property

Friend Property Get JpegXRDecoderId() As GUID
    Dim Ret As GUID
    If m_pIBitmapDecoderStatics <> 0& Then
        Dim value As GUID
        If Invoke(m_pIBitmapDecoderStatics, _
                  IBitmapDecoderStatics_GetJpegXRDecoderId, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    JpegXRDecoderId = Ret
End Property

Friend Property Get IcoDecoderId() As GUID
    Dim Ret As GUID
    If m_pIBitmapDecoderStatics <> 0& Then
        Dim value As GUID
        If Invoke(m_pIBitmapDecoderStatics, _
                  IBitmapDecoderStatics_GetIcoDecoderId, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    IcoDecoderId = Ret
End Property

' IBitmapDecoderStatics2
Friend Property Get HeifDecoderId() As GUID
    Dim Ret As GUID
    If m_pIBitmapDecoderStatics2 <> 0& Then
        Dim value As GUID
        If Invoke(m_pIBitmapDecoderStatics2, _
                  IBitmapDecoderStatics2_GetHeifDecoderId, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    HeifDecoderId = Ret
End Property

Friend Property Get WebpDecoderId() As GUID
    Dim Ret As GUID
    If m_pIBitmapDecoderStatics2 <> 0& Then
        Dim value As GUID
        If Invoke(m_pIBitmapDecoderStatics2, _
                  IBitmapDecoderStatics2_GetWebpDecoderId, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    WebpDecoderId = Ret
End Property

' IBitmapDecoder
Public Property Get BitmapContainerProperties() As BitmapPropertiesView
    Dim Ret As BitmapPropertiesView
    If m_pIBitmapDecoder <> 0& Then
        Dim pIBitmapPropertiesView As Long
        If Invoke(m_pIBitmapDecoder, _
                  IBitmapDecoder_GetBitmapContainerProperties, _
                  VarPtr(pIBitmapPropertiesView)) = S_OK Then
            If pIBitmapPropertiesView <> 0& Then
                Set Ret = New BitmapPropertiesView
                Ret.Ifc = pIBitmapPropertiesView
            End If
        End If
    End If
    Set BitmapProperties = Ret
End Property

Public Property Get DecoderInformation() As BitmapCodecInformation
    Dim Ret As BitmapCodecInformation
    If m_pIBitmapDecoder <> 0& Then
        Dim pIBitmapCodecInformation As Long
        If Invoke(m_pIBitmapDecoder, _
                  IBitmapDecoder_GetDecoderInformation, _
                  VarPtr(pIBitmapCodecInformation)) = S_OK Then
            If pIBitmapCodecInformation <> 0& Then
                Set Ret = New BitmapCodecInformation
                Ret.Ifc = pIBitmapCodecInformation
            End If
        End If
    End If
    Set DecoderInformation = Ret
End Property

Public Property Get FrameCount() As Long
    Dim Ret As Long
    If m_pIBitmapDecoder <> 0& Then
        Dim value As Long
        If Invoke(m_pIBitmapDecoder, _
                  IBitmapDecoder_GetFrameCount, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    FrameCount = Ret
End Property

' IBitmapFrame
Public Property Get BitmapProperties() As BitmapPropertiesView
    Dim Ret As BitmapPropertiesView
    If m_pIBitmapFrame <> 0& Then
        Dim pBitmapPropertiesView As Long
        If Invoke(m_pIBitmapFrame, _
                  IBitmapFrame_GetBitmapProperties, _
                  VarPtr(pBitmapPropertiesView)) = S_OK Then
            If pBitmapPropertiesView <> 0& Then
                Set Ret = New BitmapPropertiesView
                Ret.Ifc = pBitmapPropertiesView
            End If
        End If
    End If
    Set BitmapProperties = Ret
End Property

Public Property Get BitmapPixelFormat() As BitmapPixelFormat
    Dim Ret As BitmapPixelFormat
    If m_pIBitmapFrame <> 0& Then
        Dim value As Long
        If Invoke(m_pIBitmapFrame, _
                  IBitmapFrame_GetBitmapPixelFormat, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    BitmapPixelFormat = Ret
End Property

Public Property Get BitmapAlphaMode() As BitmapAlphaMode
    Dim Ret As BitmapAlphaMode
    If m_pIBitmapFrame <> 0& Then
        Dim value As Long
        If Invoke(m_pIBitmapFrame, _
                  IBitmapFrame_GetBitmapAlphaMode, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    BitmapAlphaMode = Ret
End Property

Public Property Get dpiX() As Double
    Dim Ret As Double
    If m_pIBitmapFrame <> 0& Then
        Dim value As Double
        If Invoke(m_pIBitmapFrame, _
                  IBitmapFrame_GetDpiX, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    dpiX = Ret
End Property

Public Property Get dpiY() As Double
    Dim Ret As Double
    If m_pIBitmapFrame <> 0& Then
        Dim value As Double
        If Invoke(m_pIBitmapFrame, _
                  IBitmapFrame_GetDpiY, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    dpiY = Ret
End Property

Public Property Get PixelWidth() As Long
    Dim Ret As Long
    If m_pIBitmapFrame <> 0& Then
        Dim value As Long
        If Invoke(m_pIBitmapFrame, _
                  IBitmapFrame_GetPixelWidth, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    PixelWidth = Ret
End Property

Public Property Get PixelHeight() As Long
    Dim Ret As Long
    If m_pIBitmapFrame <> 0& Then
        Dim value As Long
        If Invoke(m_pIBitmapFrame, _
                  IBitmapFrame_GetPixelHeight, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    PixelHeight = Ret
End Property

Public Property Get OrientedPixelWidth() As Long
    Dim Ret As Long
    If m_pIBitmapFrame <> 0& Then
        Dim value As Long
        If Invoke(m_pIBitmapFrame, _
                  IBitmapFrame_GetOrientedPixelWidth, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    OrientedPixelWidth = Ret
End Property

Public Property Get OrientedPixelHeight() As Long
    Dim Ret As Long
    If m_pIBitmapFrame <> 0& Then
        Dim value As Long
        If Invoke(m_pIBitmapFrame, _
                  IBitmapFrame_GetOrientedPixelHeight, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    OrientedPixelHeight = Ret
End Property

' ----==== Functions ====----
' IBitmapDecoderStatics
Public Function GetDecoderInformationEnumerator() As ReadOnlyList_1 'ReadOnlyList_BitmapCodecInformation
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_BitmapCodecInformation
    If m_pIBitmapDecoderStatics <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIBitmapDecoderStatics, _
                  IBitmapDecoderStatics_GetDecoderInformationEnumerator, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_BitmapCodecInformation, pIVectorView) 'ReadOnlyList_BitmapCodecInformation
            End If
        End If
    End If
    Set GetDecoderInformationEnumerator = Ret
End Function

Public Function CreateAsync(ByVal stream As RandomAccessStream) As BitmapDecoder
    Dim Ret As BitmapDecoder
    If m_pIBitmapDecoderStatics <> 0& And _
       IsNotNothing(stream) Then
        Dim pIBitmapDecoder As Long
        If Invoke(m_pIBitmapDecoderStatics, _
                  IBitmapDecoderStatics_CreateAsync, _
                  stream.Ifc, _
                  VarPtr(pIBitmapDecoder)) = S_OK Then
            If pIBitmapDecoder <> 0& Then
                If Await(pIBitmapDecoder) Then
                    Set Ret = New BitmapDecoder
                    Ret.Ifc = pIBitmapDecoder
                End If
            End If
        End If
    End If
    Set CreateAsync = Ret
End Function

Friend Function CreateWithIdAsync(ByRef decoderId As GUID, _
                                  ByVal stream As RandomAccessStream) As BitmapDecoder
    Dim Ret As BitmapDecoder
    If m_pIBitmapDecoderStatics <> 0& And _
       IsNotNothing(stream) Then
        Dim pIBitmapDecoder As Long
        If Invoke(m_pIBitmapDecoderStatics, _
                  IBitmapDecoderStatics_CreateWithIdAsync, _
                  decoderId.data1, _
                  decoderId.data2, _
                  decoderId.data3, _
                  decoderId.data4, _
                  stream.Ifc, _
                  VarPtr(pIBitmapDecoder)) = S_OK Then
            If pIBitmapDecoder <> 0& Then
                If Await(pIBitmapDecoder) Then
                    Set Ret = New BitmapDecoder
                    Ret.Ifc = pIBitmapDecoder
                End If
            End If
        End If
    End If
    Set CreateWithIdAsync = Ret
End Function

' IBitmapDecoder
Public Function GetPreviewAsync() As ImageStream
    Dim Ret As ImageStream
    If m_pIBitmapDecoder <> 0& Then
        Dim pIImageStream As Long
        If Invoke(m_pIBitmapDecoder, _
                  IBitmapDecoder_GetPreviewAsync, _
                  VarPtr(pIImageStream)) = S_OK Then
            If pIImageStream <> 0& Then
                If Await(pIImageStream) Then
                    Set Ret = New ImageStream
                    Ret.Ifc = pIImageStream
                End If
            End If
        End If
    End If
    Set GetPreviewAsync = Ret
End Function

Public Function GetFrameAsync() As BitmapFrame
    Dim Ret As BitmapFrame
    If m_pIBitmapDecoder <> 0& Then
        Dim pIBitmapFrame As Long
        If Invoke(m_pIBitmapDecoder, _
                  IBitmapDecoder_GetFrameAsync, _
                  VarPtr(pIBitmapFrame)) = S_OK Then
            If pIBitmapFrame <> 0& Then
                If Await(pIBitmapFrame) Then
                    Set Ret = New BitmapFrame
                    Ret.Ifc = pIBitmapFrame
                End If
            End If
        End If
    End If
    Set GetFrameAsync = Ret
End Function

' IBitmapFrame
Public Function GetThumbnailAsync() As ImageStream
    Dim Ret As ImageStream
    If m_pIBitmapFrame <> 0& Then
        Dim pIImageStream As Long
        If Invoke(m_pIBitmapFrame, _
                  IBitmapFrame_GetThumbnailAsync, _
                  VarPtr(pIImageStream)) = S_OK Then
            If pIImageStream <> 0& Then
                If Await(pIImageStream) Then
                    Set Ret = New ImageStream
                    Ret.Ifc = pIImageStream
                End If
            End If
        End If
    End If
    Set GetThumbnailAsync = Ret
End Function

Public Function GetPixelDataAsync() As PixelDataProvider
    Dim Ret As PixelDataProvider
    If m_pIBitmapFrame <> 0& Then
        Dim pIPixelDataProvider As Long
        If Invoke(m_pIBitmapFrame, _
                  IBitmapFrame_GetPixelDataAsync, _
                  VarPtr(pIPixelDataProvider)) = S_OK Then
            If pIPixelDataProvider <> 0& Then
                If Await(pIPixelDataProvider) Then
                    Set Ret = New PixelDataProvider
                    Ret.Ifc = pIPixelDataProvider
                End If
            End If
        End If
    End If
    Set GetPixelDataAsync = Ret
End Function

Public Function GetPixelDataTransformedAsync(ByVal PixelFormat As BitmapPixelFormat, _
                                             ByVal alphaMode As BitmapAlphaMode, _
                                             ByVal transform As BitmapTransform, _
                                             ByVal exOrientationMode As ExifOrientationMode, _
                                             ByVal colManagementMode As ColorManagementMode) As PixelDataProvider
    Dim Ret As PixelDataProvider
    If m_pIBitmapFrame <> 0& And _
       IsNotNothing(transform) Then
        Dim pIPixelDataProvider As Long
        If Invoke(m_pIBitmapFrame, _
                  IBitmapFrame_GetPixelDataTransformedAsync, _
                  PixelFormat, _
                  alphaMode, _
                  transform.Ifc, _
                  exOrientationMode, _
                  colManagementMode, _
                  VarPtr(pIPixelDataProvider)) = S_OK Then
            If pIPixelDataProvider <> 0& Then
                If Await(pIPixelDataProvider) Then
                    Set Ret = New PixelDataProvider
                    Ret.Ifc = pIPixelDataProvider
                End If
            End If
        End If
    End If
    Set GetPixelDataTransformedAsync = Ret
End Function

' IBitmapFrameWithSoftwareBitmap
Public Function GetSoftwareBitmapAsync() As SoftwareBitmap
    Dim Ret As SoftwareBitmap
    If m_pIBitmapFrameWithSoftwareBitmap <> 0& Then
        Dim pISoftwareBitmap As Long
        If Invoke(m_pIBitmapFrameWithSoftwareBitmap, _
                  IBitmapFrameWithSoftwareBitmap_GetSoftwareBitmapAsync, _
                  VarPtr(pISoftwareBitmap)) = S_OK Then
            If pISoftwareBitmap <> 0& Then
                If Await(pISoftwareBitmap) Then
                    Set Ret = New SoftwareBitmap
                    Ret.Ifc = pISoftwareBitmap
                End If
            End If
        End If
    End If
    Set GetSoftwareBitmapAsync = Ret
End Function

Public Function GetSoftwareBitmapConvertedAsync(ByVal PixelFormat As BitmapPixelFormat, _
                                                ByVal alphaMode As BitmapAlphaMode) As SoftwareBitmap
    Dim Ret As SoftwareBitmap
    If m_pIBitmapFrameWithSoftwareBitmap <> 0& Then
        Dim pISoftwareBitmap As Long
        If Invoke(m_pIBitmapFrameWithSoftwareBitmap, _
                  IBitmapFrameWithSoftwareBitmap_GetSoftwareBitmapConvertedAsync, _
                  PixelFormat, _
                  alphaMode, _
                  VarPtr(pISoftwareBitmap)) = S_OK Then
            If pISoftwareBitmap <> 0& Then
                If Await(pISoftwareBitmap) Then
                    Set Ret = New SoftwareBitmap
                    Ret.Ifc = pISoftwareBitmap
                End If
            End If
        End If
    End If
    Set GetSoftwareBitmapConvertedAsync = Ret
End Function

Public Function GetSoftwareBitmapTransformedAsync(ByVal PixelFormat As BitmapPixelFormat, _
                                                  ByVal alphaMode As BitmapAlphaMode, _
                                                  ByVal transform As BitmapTransform, _
                                                  ByVal exOrientationMode As ExifOrientationMode, _
                                                  ByVal colManagementMode As ColorManagementMode) As SoftwareBitmap
    Dim Ret As SoftwareBitmap
    If m_pIBitmapFrameWithSoftwareBitmap <> 0& And _
       IsNotNothing(transform) Then
        Dim pISoftwareBitmap As Long
        If Invoke(m_pIBitmapFrameWithSoftwareBitmap, _
                  IBitmapFrameWithSoftwareBitmap_GetSoftwareBitmapTransformedAsync, _
                  PixelFormat, _
                  alphaMode, _
                  transform.Ifc, _
                  exOrientationMode, _
                  colManagementMode, _
                  VarPtr(pISoftwareBitmap)) = S_OK Then
            If pISoftwareBitmap <> 0& Then
                If Await(pISoftwareBitmap) Then
                    Set Ret = New SoftwareBitmap
                    Ret.Ifc = pISoftwareBitmap
                End If
            End If
        End If
    End If
    Set GetSoftwareBitmapTransformedAsync = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


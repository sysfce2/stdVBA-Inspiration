VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitmapEncoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Graphics_Imaging_BitmapEncoder As String = "Windows.Graphics.Imaging.BitmapEncoder"
Private Const IID_IBitmapEncoder As String = "{2bc468e3-e1f8-4b54-95e8-32919551ce62}"
Private Const IID_IBitmapEncoderStatics As String = "{a74356a7-a4e4-4eb9-8e40-564de7e1ccb2}"
Private Const IID_IBitmapEncoderStatics2 As String = "{33cbc259-fe31-41b1-b812-086d21e87e16}"
Private Const IID_IBitmapEncoderWithSoftwareBitmap As String = "{686cd241-4330-4c77-ace4-0334968b1768}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IBitmapEncoderStatics
    IBitmapEncoderStatics_GetBmpEncoderId = 6
    IBitmapEncoderStatics_GetJpegEncoderId = 7
    IBitmapEncoderStatics_GetPngEncoderId = 8
    IBitmapEncoderStatics_GetTiffEncoderId = 9
    IBitmapEncoderStatics_GetGifEncoderId = 10
    IBitmapEncoderStatics_GetJpegXREncoderId = 11
    IBitmapEncoderStatics_GetEncoderInformationEnumerator = 12
    IBitmapEncoderStatics_CreateAsync = 13
'    IBitmapEncoderStatics_CreateWithEncodingOptionsAsync = 14
    IBitmapEncoderStatics_CreateForTranscodingAsync = 15
    IBitmapEncoderStatics_CreateForInPlacePropertyEncodingAsync = 16
                        
    ' IBitmapEncoderStatics2
    IBitmapEncoderStatics2_GetHeifEncoderId = 6
                        
    ' IBitmapEncoder
    IBitmapEncoder_GetEncoderInformation = 6
    IBitmapEncoder_GetBitmapProperties = 7
    IBitmapEncoder_GetBitmapContainerProperties = 8
    IBitmapEncoder_GetIsThumbnailGenerated = 9
    IBitmapEncoder_PutIsThumbnailGenerated = 10
    IBitmapEncoder_GetGeneratedThumbnailWidth = 11
    IBitmapEncoder_PutGeneratedThumbnailWidth = 12
    IBitmapEncoder_GetGeneratedThumbnailHeight = 13
    IBitmapEncoder_PutGeneratedThumbnailHeight = 14
    IBitmapEncoder_GetBitmapTransform = 15
    IBitmapEncoder_SetPixelData = 16
    IBitmapEncoder_GoToNextFrameAsync = 17
'    IBitmapEncoder_GoToNextFrameWithEncodingOptionsAsync = 18
    IBitmapEncoder_FlushAsync = 19
                        
    ' IBitmapEncoderWithSoftwareBitmap
    IBitmapEncoderWithSoftwareBitmap_SetSoftwareBitmap = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIBitmapEncoder As Long
Private m_pIBitmapEncoderStatics As Long
Private m_pIBitmapEncoderStatics2 As Long
Private m_pIBitmapEncoderWithSoftwareBitmap As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Graphics_Imaging_BitmapEncoder, _
                            IID_IBitmapEncoderStatics, _
                            m_pIBitmapEncoderStatics) Then
        Call QueryIfc(m_pIBitmapEncoderStatics, _
                      IID_IBitmapEncoderStatics2, _
                      m_pIBitmapEncoderStatics2)
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIBitmapEncoderWithSoftwareBitmap)
    Call ReleaseIfc(m_pIBitmapEncoder)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIBitmapEncoder
End Property

Public Property Let Ifc(ByVal pIBitmapEncoder As Long)
    Call ReleaseIfc(m_pIBitmapEncoderWithSoftwareBitmap)
    Call ReleaseIfc(m_pIBitmapEncoder)
    m_pIBitmapEncoder = pIBitmapEncoder
    Call QueryIfc(m_pIBitmapEncoder, IID_IBitmapEncoderWithSoftwareBitmap, m_pIBitmapEncoderWithSoftwareBitmap)
End Property

' IBitmapEncoderStatics
Friend Property Get BmpEncoderId() As GUID
    Dim Ret As GUID
    If m_pIBitmapEncoderStatics <> 0& Then
        Dim value As GUID
        If Invoke(m_pIBitmapEncoderStatics, _
                  IBitmapEncoderStatics_GetBmpEncoderId, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    BmpEncoderId = Ret
End Property

Friend Property Get JpegEncoderId() As GUID
    Dim Ret As GUID
    If m_pIBitmapEncoderStatics <> 0& Then
        Dim value As GUID
        If Invoke(m_pIBitmapEncoderStatics, _
                  IBitmapEncoderStatics_GetJpegEncoderId, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    JpegEncoderId = Ret
End Property

Friend Property Get PngEncoderId() As GUID
    Dim Ret As GUID
    If m_pIBitmapEncoderStatics <> 0& Then
        Dim value As GUID
        If Invoke(m_pIBitmapEncoderStatics, _
                  IBitmapEncoderStatics_GetPngEncoderId, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    PngEncoderId = Ret
End Property

Friend Property Get TiffEncoderId() As GUID
    Dim Ret As GUID
    If m_pIBitmapEncoderStatics <> 0& Then
        Dim value As GUID
        If Invoke(m_pIBitmapEncoderStatics, _
                  IBitmapEncoderStatics_GetTiffEncoderId, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    TiffEncoderId = Ret
End Property

Friend Property Get GifEncoderId() As GUID
    Dim Ret As GUID
    If m_pIBitmapEncoderStatics <> 0& Then
        Dim value As GUID
        If Invoke(m_pIBitmapEncoderStatics, _
                  IBitmapEncoderStatics_GetGifEncoderId, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    GifEncoderId = Ret
End Property

Friend Property Get JpegXREncoderId() As GUID
    Dim Ret As GUID
    If m_pIBitmapEncoderStatics <> 0& Then
        Dim value As GUID
        If Invoke(m_pIBitmapEncoderStatics, _
                  IBitmapEncoderStatics_GetJpegXREncoderId, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    JpegXREncoderId = Ret
End Property

' IBitmapEncoderStatics2
Friend Property Get HeifEncoderId() As GUID
    Dim Ret As GUID
    If m_pIBitmapEncoderStatics2 <> 0& Then
        Dim value As GUID
        If Invoke(m_pIBitmapEncoderStatics2, _
                  IBitmapEncoderStatics2_GetHeifEncoderId, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    HeifEncoderId = Ret
End Property

' IBitmapEncoder
Public Property Get EncoderInformation() As BitmapCodecInformation
    Dim Ret As BitmapCodecInformation
    If m_pIBitmapEncoder <> 0& Then
        Dim pIBitmapCodecInformation As Long
        If Invoke(m_pIBitmapEncoder, _
                  IBitmapEncoder_GetEncoderInformation, _
                  VarPtr(pIBitmapCodecInformation)) = S_OK Then
            If pIBitmapCodecInformation <> 0& Then
                Set Ret = New BitmapCodecInformation
                Ret.Ifc = pIBitmapCodecInformation
            End If
        End If
    End If
    Set EncoderInformation = Ret
End Property

Public Property Get BitmapProperties() As BitmapProperties
    Dim Ret As BitmapProperties
    If m_pIBitmapEncoder <> 0& Then
        Dim pIBitmapProperties As Long
        If Invoke(m_pIBitmapEncoder, _
                  IBitmapEncoder_GetBitmapProperties, _
                  VarPtr(pIBitmapProperties)) = S_OK Then
            If pIBitmapProperties <> 0& Then
                Set Ret = New BitmapProperties
                Ret.Ifc = pIBitmapProperties
            End If
        End If
    End If
    Set BitmapProperties = Ret
End Property

Public Property Get BitmapContainerProperties() As BitmapProperties
    Dim Ret As BitmapProperties
    If m_pIBitmapEncoder <> 0& Then
        Dim pIBitmapProperties As Long
        If Invoke(m_pIBitmapEncoder, _
                  IBitmapEncoder_GetBitmapContainerProperties, _
                  VarPtr(pIBitmapProperties)) = S_OK Then
            If pIBitmapProperties <> 0& Then
                Set Ret = New BitmapProperties
                Ret.Ifc = pIBitmapProperties
            End If
        End If
    End If
    Set BitmapContainerProperties = Ret
End Property

Public Property Get IsThumbnailGenerated() As Boolean
    Dim Ret As Boolean
    If m_pIBitmapEncoder <> 0& Then
        Dim value As Long
        If Invoke(m_pIBitmapEncoder, _
                  IBitmapEncoder_GetIsThumbnailGenerated, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsThumbnailGenerated = Ret
End Property

Public Property Let IsThumbnailGenerated(ByVal value As Boolean)
    If m_pIBitmapEncoder <> 0& Then
        If Invoke(m_pIBitmapEncoder, _
                  IBitmapEncoder_PutIsThumbnailGenerated, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get GeneratedThumbnailWidth() As Long
    Dim Ret As Long
    If m_pIBitmapEncoder <> 0& Then
        Dim value As Long
        If Invoke(m_pIBitmapEncoder, _
                  IBitmapEncoder_GetGeneratedThumbnailWidth, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    GeneratedThumbnailWidth = Ret
End Property

Public Property Let GeneratedThumbnailWidth(ByVal value As Long)
    If m_pIBitmapEncoder <> 0& Then
        If Invoke(m_pIBitmapEncoder, _
                  IBitmapEncoder_PutGeneratedThumbnailWidth, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get GeneratedThumbnailHeight() As Long
    Dim Ret As Long
    If m_pIBitmapEncoder <> 0& Then
        Dim value As Long
        If Invoke(m_pIBitmapEncoder, _
                  IBitmapEncoder_GetGeneratedThumbnailHeight, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    GeneratedThumbnailHeight = Ret
End Property

Public Property Let GeneratedThumbnailHeight(ByVal value As Long)
    If m_pIBitmapEncoder <> 0& Then
        If Invoke(m_pIBitmapEncoder, _
                  IBitmapEncoder_PutGeneratedThumbnailHeight, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get BitmapTransform() As BitmapTransform
    Dim Ret As BitmapTransform
    If m_pIBitmapEncoder <> 0& Then
        Dim pIBitmapTransform As Long
        If Invoke(m_pIBitmapEncoder, _
                  IBitmapEncoder_GetBitmapTransform, _
                  VarPtr(pIBitmapTransform)) = S_OK Then
            If pIBitmapTransform <> 0& Then
                Set Ret = New BitmapTransform
                Ret.Ifc = pIBitmapTransform
            End If
        End If
    End If
    Set BitmapTransform = Ret
End Property

' ----==== Functions ====----
' IBitmapEncoderStatics
Public Function GetEncoderInformationEnumerator() As ReadOnlyList_1 'ReadOnlyList_BitmapCodecInformation
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_BitmapCodecInformation
    If m_pIBitmapEncoderStatics <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIBitmapEncoderStatics, _
                  IBitmapEncoderStatics_GetEncoderInformationEnumerator, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_BitmapCodecInformation, pIVectorView) 'ReadOnlyList_BitmapCodecInformation
            End If
        End If
    End If
    Set GetEncoderInformationEnumerator = Ret
End Function

Friend Function CreateAsync(ByRef encoderId As GUID, _
                            ByVal stream As RandomAccessStream) As BitmapEncoder
    Dim Ret As BitmapEncoder
    If m_pIBitmapEncoderStatics <> 0& And _
       IsNotNothing(stream) Then
        Dim pIBitmapEncoder As Long
        If Invoke(m_pIBitmapEncoderStatics, _
                  IBitmapEncoderStatics_CreateAsync, _
                  encoderId.data1, _
                  encoderId.data2, _
                  encoderId.data3, _
                  encoderId.data4, _
                  stream.Ifc, _
                  VarPtr(pIBitmapEncoder)) = S_OK Then
            If pIBitmapEncoder <> 0& Then
                If Await(pIBitmapEncoder) Then
                    Set Ret = New BitmapEncoder
                    Ret.Ifc = pIBitmapEncoder
                End If
            End If
        End If
    End If
    Set CreateAsync = Ret
End Function

'    IBitmapEncoderStatics_CreateWithEncodingOptionsAsync = 14

Public Function CreateForTranscodingAsync(ByVal stream As RandomAccessStream, _
                                          ByVal decoder As BitmapDecoder) As BitmapEncoder
    Dim Ret As BitmapEncoder
    If m_pIBitmapEncoderStatics <> 0& And _
       IsNotNothing(stream) And _
       IsNotNothing(decoder) Then
        Dim pIBitmapEncoder As Long
        If Invoke(m_pIBitmapEncoderStatics, _
                  IBitmapEncoderStatics_CreateForTranscodingAsync, _
                  stream.Ifc, _
                  decoder.Ifc, _
                  VarPtr(pIBitmapEncoder)) = S_OK Then
            If pIBitmapEncoder <> 0& Then
                If Await(pIBitmapEncoder) Then
                    Set Ret = New BitmapEncoder
                    Ret.Ifc = pIBitmapEncoder
                End If
            End If
        End If
    End If
    Set CreateForTranscodingAsync = Ret
End Function

Public Function CreateForInPlacePropertyEncodingAsync(ByVal decoder As BitmapDecoder) As BitmapEncoder
    Dim Ret As BitmapEncoder
    If m_pIBitmapEncoderStatics <> 0& And _
       IsNotNothing(decoder) Then
        Dim pIBitmapEncoder As Long
        If Invoke(m_pIBitmapEncoderStatics, _
                  IBitmapEncoderStatics_CreateForInPlacePropertyEncodingAsync, _
                  decoder.Ifc, _
                  VarPtr(pIBitmapEncoder)) = S_OK Then
            If pIBitmapEncoder <> 0& Then
                If Await(pIBitmapEncoder) Then
                    Set Ret = New BitmapEncoder
                    Ret.Ifc = pIBitmapEncoder
                End If
            End If
        End If
    End If
    Set CreateForInPlacePropertyEncodingAsync = Ret
End Function

' IBitmapEncoder
Public Function SetPixelData(ByVal PixelFormat As BitmapPixelFormat, _
                             ByVal alphaMode As BitmapAlphaMode, _
                             ByVal Width As Long, _
                             ByVal Height As Long, _
                             ByVal dpiX As Double, _
                             ByVal dpiY As Double, _
                             ByVal pixelsLength As Long, _
                             ByRef pixels() As Byte) As Boolean
    Dim Ret As Boolean
    If m_pIBitmapEncoder <> 0& Then
        If Invoke(m_pIBitmapEncoder, _
                  IBitmapEncoder_SetPixelData, _
                  PixelFormat, _
                  alphaMode, _
                  Width, _
                  Height, _
                  dpiX, _
                  dpiY, _
                  pixelsLength, _
                  VarPtr(pixels(0))) = S_OK Then
            Ret = True
        End If
    End If
    SetPixelData = Ret
End Function

'    IBitmapEncoder_GoToNextFrameAsync = 17
Public Function GoToNextFrameAsync() As Boolean
    Dim Ret As Boolean
    If m_pIBitmapEncoder <> 0& Then
        Dim pIAsyncAction As Long
        If Invoke(m_pIBitmapEncoder, _
                  IBitmapEncoder_GoToNextFrameAsync, _
                  VarPtr(pIAsyncAction)) = S_OK Then
            If pIAsyncAction <> 0& Then
                If Await(pIAsyncAction, False) Then
                    Ret = True
                End If
            End If
        End If
    End If
    GoToNextFrameAsync = Ret
End Function

'    IBitmapEncoder_GoToNextFrameWithEncodingOptionsAsync = 18

Public Function FlushAsync() As Boolean
    Dim Ret As Boolean
    If m_pIBitmapEncoder <> 0& Then
        Dim pIAsyncAction As Long
        If Invoke(m_pIBitmapEncoder, _
                  IBitmapEncoder_FlushAsync, _
                  VarPtr(pIAsyncAction)) = S_OK Then
            If pIAsyncAction <> 0& Then
                If Await(pIAsyncAction, False) Then
                    Ret = True
                End If
            End If
        End If
    End If
    FlushAsync = Ret
End Function

' IBitmapEncoderWithSoftwareBitmap
Public Function SetSoftwareBitmap(ByVal bitmap As SoftwareBitmap) As Boolean
    Dim Ret As Boolean
    If m_pIBitmapEncoderWithSoftwareBitmap <> 0& And _
       IsNotNothing(bitmap) Then
        If Invoke(m_pIBitmapEncoderWithSoftwareBitmap, _
                  IBitmapEncoderWithSoftwareBitmap_SetSoftwareBitmap, _
                  bitmap.Ifc) = S_OK Then
            Ret = True
        End If
    End If
    SetSoftwareBitmap = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



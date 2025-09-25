VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Capture_MediaCapture As String = "Windows.Media.Capture.MediaCapture"
Private Const IID_IMediaCapture As String = "{c61afbb4-fb10-4a34-ac18-ca80d9c8e7ee}"
Private Const IID_IMediaCapture2 As String = "{9cc68260-7da1-4043-b652-21b8878daff9}"
Private Const IID_IMediaCapture3 As String = "{d4136f30-1564-466e-bc0a-af94e02ab016}"
Private Const IID_IMediaCapture4 As String = "{bacd6fd6-fb08-4947-aea2-ce14eff0ce13}"
Private Const IID_IMediaCapture5 As String = "{da787c22-3a9b-4720-a71e-97900a316e5a}"
Private Const IID_IMediaCapture6 As String = "{228948bd-4b20-4bb1-9fd6-a583212a1012}"
Private Const IID_IMediaCapture7 As String = "{9169f102-8888-541a-95bc-24e4d462542a}"
Private Const IID_IMediaCaptureStatics As String = "{acef81ff-99ed-4645-965e-1925cfc63834}"
Private Const IID_IMediaCaptureVideoPreview As String = "{27727073-549e-447f-a20a-4f03c479d8c0}"

' * Class implements the following interface:
' *    Windows.Foundation.IClosable
 
' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IMediaCaptureStatics
    IMediaCaptureStatics_IsVideoProfileSupported = 6
    IMediaCaptureStatics_FindAllVideoProfiles = 7
    IMediaCaptureStatics_FindConcurrentProfiles = 8
    IMediaCaptureStatics_FindKnownVideoProfiles = 9
    
    ' IMediaCapture
    IMediaCapture_InitializeAsync = 6
    IMediaCapture_InitializeWithSettingsAsync = 7
    IMediaCapture_StartRecordToStorageFileAsync = 8
    IMediaCapture_StartRecordToStreamAsync = 9
'    IMediaCapture_StartRecordToCustomSinkAsync = 10
'    IMediaCapture_StartRecordToCustomSinkIdAsync = 11
    IMediaCapture_StopRecordAsync = 12
    IMediaCapture_CapturePhotoToStorageFileAsync = 13
    IMediaCapture_CapturePhotoToStreamAsync = 14
    ' DEPRECATED ("AddEffectAsync might not be available in future versions of Windows.
    ' Starting with Windows Threshold, use AddAudioEffectAsync and AddVideoEffectAsync instead")
'    IMediaCapture_AddEffectAsync = 15
    IMediaCapture_ClearEffectsAsync = 16
    IMediaCapture_SetEncoderProperty = 17
    IMediaCapture_GetEncoderProperty = 18
    IMediaCapture_AddFailed = 19
    IMediaCapture_RemoveFailed = 20
    IMediaCapture_AddRecordLimitationExceeded = 21
    IMediaCapture_RemoveRecordLimitationExceeded = 22
    IMediaCapture_GetMediaCaptureSettings = 23
'    IMediaCapture_GetAudioDeviceController = 24
'    IMediaCapture_GetVideoDeviceController = 25
    IMediaCapture_SetPreviewMirroring = 26
    IMediaCapture_GetPreviewMirroring = 27
    IMediaCapture_SetPreviewRotation = 27
    IMediaCapture_GetPreviewRotation = 29
    IMediaCapture_SetRecordRotation = 30
    IMediaCapture_GetRecordRotation = 31
    
    ' IMediaCaptureVideoPreview
    IMediaCaptureVideoPreview_StartPreviewAsync = 6
'    IMediaCaptureVideoPreview_StartPreviewToCustomSinkAsync = 7
'    IMediaCaptureVideoPreview_StartPreviewToCustomSinkIdAsync = 8
    IMediaCaptureVideoPreview_StopPreviewAsync = 9

End Enum

' ----==== Variablen ====----
Private m_pIMediaCapture As Long
Private m_pIMediaCapture2 As Long
Private m_pIMediaCapture3 As Long
Private m_pIMediaCapture4 As Long
Private m_pIMediaCapture5 As Long
Private m_pIMediaCapture6 As Long
Private m_pIMediaCapture7 As Long
Private m_pIMediaCaptureStatics As Long
Private m_pIMediaCaptureVideoPreview As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Media_Capture_MediaCapture, _
                            IID_IMediaCaptureStatics, _
                            m_pIMediaCaptureStatics) Then
    End If
    Dim pIMediaCapture As Long
    If GetActivateInstance(Windows_Media_Capture_MediaCapture, _
                           IID_IMediaCapture, _
                           pIMediaCapture) Then
        Me.Ifc = pIMediaCapture
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMediaCaptureVideoPreview)
    Call ReleaseIfc(m_pIMediaCapture7)
    Call ReleaseIfc(m_pIMediaCapture6)
    Call ReleaseIfc(m_pIMediaCapture5)
    Call ReleaseIfc(m_pIMediaCapture4)
    Call ReleaseIfc(m_pIMediaCapture3)
    Call ReleaseIfc(m_pIMediaCapture2)
    Call DisposeIfc(m_pIMediaCapture) ' <- IClosable
    Call ReleaseIfc(m_pIMediaCaptureStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMediaCapture
End Property

Public Property Let Ifc(ByVal pIMediaCapture As Long)
    Call ReleaseIfc(m_pIMediaCaptureVideoPreview)
    Call ReleaseIfc(m_pIMediaCapture7)
    Call ReleaseIfc(m_pIMediaCapture6)
    Call ReleaseIfc(m_pIMediaCapture5)
    Call ReleaseIfc(m_pIMediaCapture4)
    Call ReleaseIfc(m_pIMediaCapture3)
    Call ReleaseIfc(m_pIMediaCapture2)
    Call DisposeIfc(m_pIMediaCapture) ' <- IClosable
    m_pIMediaCapture = pIMediaCapture
    Call QueryIfc(m_pIMediaCapture, IID_IMediaCapture2, m_pIMediaCapture2)
    Call QueryIfc(m_pIMediaCapture, IID_IMediaCapture3, m_pIMediaCapture3)
    Call QueryIfc(m_pIMediaCapture, IID_IMediaCapture4, m_pIMediaCapture4)
    Call QueryIfc(m_pIMediaCapture, IID_IMediaCapture5, m_pIMediaCapture5)
    Call QueryIfc(m_pIMediaCapture, IID_IMediaCapture6, m_pIMediaCapture6)
    Call QueryIfc(m_pIMediaCapture, IID_IMediaCapture7, m_pIMediaCapture7)
    Call QueryIfc(m_pIMediaCapture, IID_IMediaCaptureVideoPreview, m_pIMediaCaptureVideoPreview)
End Property

' IMediaCapture
Public Property Get MediaCaptureSettings() As MediaCaptureSettings
    Dim Ret As MediaCaptureSettings
    If m_pIMediaCapture <> 0& Then
        Dim pIMediaCaptureSettings As Long
        If Invoke(m_pIMediaCapture, _
                  IMediaCapture_GetMediaCaptureSettings, _
                  VarPtr(pIMediaCaptureSettings)) = S_OK Then
            If pIMediaCaptureSettings <> 0& Then
                Set Ret = New MediaCaptureSettings
                Ret.Ifc = pIMediaCaptureSettings
            End If
        End If
    End If
    Set MediaCaptureSettings = Ret
End Property

'Public Property Get AudioDeviceController() As AudioDeviceController
'    Dim Ret As AudioDeviceController
'    If m_pIMediaCapture <> 0& Then
'        Dim pIAudioDeviceController As Long
'        If Invoke(m_pIMediaCapture, _
'                  IMediaCapture_GetAudioDeviceController, _
'                  VarPtr(pIAudioDeviceController)) = S_OK Then
'            If pIAudioDeviceController <> 0& Then
'                Set Ret = New AudioDeviceController
'                Ret.Ifc = pIAudioDeviceController
'            End If
'        End If
'    End If
'    Set AudioDeviceController = Ret
'End Property

'Public Property Get VideoDeviceController() As VideoDeviceController
'    Dim Ret As VideoDeviceController
'    If m_pIMediaCapture <> 0& Then
'        Dim pIVideoDeviceController As Long
'        If Invoke(m_pIMediaCapture, _
'                  IMediaCapture_GetVideoDeviceController, _
'                  VarPtr(pIVideoDeviceController)) = S_OK Then
'            If pIVideoDeviceController <> 0& Then
'                Set Ret = New VideoDeviceController
'                Ret.Ifc = pIVideoDeviceController
'            End If
'        End If
'    End If
'    Set VideoDeviceController = Ret
'End Property

' ----==== Functions ====----
' IMediaCaptureStatics
Public Function IsVideoProfileSupported(ByVal VideoDeviceId As String) As Boolean
    Dim Ret As Boolean
    If m_pIMediaCaptureStatics <> 0& Then
        Dim value As Long
        Dim hString As Long
        hString = CreateWindowsString(VideoDeviceId)
        If Invoke(m_pIMediaCaptureStatics, _
                  IMediaCaptureStatics_IsVideoProfileSupported, _
                  hString, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
        Call DeleteWindowsString(hString)
    End If
    IsVideoProfileSupported = Ret
End Function

Public Function FindAllVideoProfiles(ByVal VideoDeviceId As String) As ReadOnlyList_1 'ReadOnlyList_CaptureVideoProfile
    Dim Ret As ReadOnlyList_1 ' ReadOnlyList_CaptureVideoProfile
    If m_pIMediaCaptureStatics <> 0& Then
        Dim pIVectorView As Long
        Dim hString As Long
        hString = CreateWindowsString(VideoDeviceId)
        If Invoke(m_pIMediaCaptureStatics, _
                  IMediaCaptureStatics_FindAllVideoProfiles, _
                  hString, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_MediaCaptureVideoProfile, pIVectorView) 'ReadOnlyList_CaptureVideoProfile
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set FindAllVideoProfiles = Ret
End Function

Public Function FindConcurrentProfiles(ByVal VideoDeviceId As String) As ReadOnlyList_1 'ReadOnlyList_CaptureVideoProfile
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_CaptureVideoProfile
    If m_pIMediaCaptureStatics <> 0& Then
        Dim pIVectorView As Long
        Dim hString As Long
        hString = CreateWindowsString(VideoDeviceId)
        If Invoke(m_pIMediaCaptureStatics, _
                  IMediaCaptureStatics_FindConcurrentProfiles, _
                  hString, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_MediaCaptureVideoProfile, pIVectorView) 'ReadOnlyList_CaptureVideoProfile
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set FindConcurrentProfiles = Ret
End Function

Public Function FindKnownVideoProfiles(ByVal VideoDeviceId As String, _
                                       ByVal Name As KnownVideoProfile) As ReadOnlyList_1 'ReadOnlyList_CaptureVideoProfile
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_CaptureVideoProfile
    If m_pIMediaCaptureStatics <> 0& Then
        Dim pIVectorView As Long
        Dim hString As Long
        hString = CreateWindowsString(VideoDeviceId)
        If Invoke(m_pIMediaCaptureStatics, _
                  IMediaCaptureStatics_FindKnownVideoProfiles, _
                  hString, _
                  Name, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_MediaCaptureVideoProfile, pIVectorView) 'ReadOnlyList_CaptureVideoProfile
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set FindKnownVideoProfiles = Ret
End Function

' IMediaCapture
Public Function InitializeAsync() As Boolean
    Dim Ret As Boolean
    If m_pIMediaCapture <> 0& Then
        Dim pIAsyncAction As Long
        If Invoke(m_pIMediaCapture, _
                  IMediaCapture_InitializeAsync, _
                  VarPtr(pIAsyncAction)) = S_OK Then
            If Await(pIAsyncAction, False) Then
                Ret = True
            End If
        End If
    End If
    InitializeAsync = Ret
End Function

Public Function InitializeWithSettingsAsync(ByVal InitializationSettings As MediaCaptureInitializationSettings) As Boolean
    Dim Ret As Boolean
    If m_pIMediaCapture <> 0& And _
       IsNotNothing(InitializationSettings) Then
        Dim pIAsyncAction As Long
        If Invoke(m_pIMediaCapture, _
                  IMediaCapture_InitializeWithSettingsAsync, _
                  InitializationSettings.Ifc, _
                  VarPtr(pIAsyncAction)) = S_OK Then
            If Await(pIAsyncAction, False) Then
                Ret = True
            End If
        End If
    End If
    InitializeWithSettingsAsync = Ret
End Function

Public Function StartRecordToStorageFileAsync(ByVal EncodingProfile As MediaEncodingProfile, _
                                              ByVal file As StorageFile) As Boolean
    Dim Ret As Boolean
    If m_pIMediaCapture <> 0& And _
       IsNotNothing(EncodingProfile) And _
       IsNotNothing(file) Then
        Dim pIAsyncAction As Long
        If Invoke(m_pIMediaCapture, _
                  IMediaCapture_StartRecordToStorageFileAsync, _
                  EncodingProfile.Ifc, _
                  file.Ifc, _
                  VarPtr(pIAsyncAction)) = S_OK Then
            If Await(pIAsyncAction, False) Then
                Ret = True
            End If
        End If
    End If
    StartRecordToStorageFileAsync = Ret
End Function

Public Function StartRecordToStreamAsync(ByVal EncodingProfile As MediaEncodingProfile, _
                                         ByVal stream As RandomAccessStream) As Boolean
    Dim Ret As Boolean
    If m_pIMediaCapture <> 0& And _
       IsNotNothing(EncodingProfile) And _
       IsNotNothing(stream) Then
        Dim pIAsyncAction As Long
        If Invoke(m_pIMediaCapture, _
                  IMediaCapture_StartRecordToStreamAsync, _
                  EncodingProfile.Ifc, _
                  stream.Ifc, _
                  VarPtr(pIAsyncAction)) = S_OK Then
            If Await(pIAsyncAction, False) Then
                Ret = True
            End If
        End If
    End If
    StartRecordToStreamAsync = Ret
End Function

'    IMediaCapture_StartRecordToCustomSinkAsync = 10
'    IMediaCapture_StartRecordToCustomSinkIdAsync = 11

Public Function StopRecordAsync() As Boolean
    Dim Ret As Boolean
    If m_pIMediaCapture <> 0& Then
        Dim pIAsyncAction As Long
        If Invoke(m_pIMediaCapture, _
                  IMediaCapture_StopRecordAsync, _
                  VarPtr(pIAsyncAction)) = S_OK Then
            If Await(pIAsyncAction, False) Then
                Ret = True
            End If
        End If
    End If
    StopRecordAsync = Ret
End Function

Public Function CapturePhotoToStorageFileAsync(ByVal ImageEncoding As ImageEncodingProperties, _
                                               ByVal file As StorageFile) As Boolean
    Dim Ret As Boolean
    If m_pIMediaCapture <> 0& And _
       IsNotNothing(ImageEncoding) And _
       IsNotNothing(file) Then
        Dim pIAsyncAction As Long
        If Invoke(m_pIMediaCapture, _
                  IMediaCapture_CapturePhotoToStorageFileAsync, _
                  ImageEncoding.Ifc, _
                  file.Ifc, _
                  VarPtr(pIAsyncAction)) = S_OK Then
            If Await(pIAsyncAction, False) Then
                Ret = True
            End If
        End If
    End If
    CapturePhotoToStorageFileAsync = Ret
End Function

Public Function CapturePhotoToStreamAsync(ByVal ImageEncoding As ImageEncodingProperties, _
                                          ByVal stream As RandomAccessStream) As Boolean
    Dim Ret As Boolean
    If m_pIMediaCapture <> 0& And _
       IsNotNothing(ImageEncoding) And _
       IsNotNothing(stream) Then
        Dim pIAsyncAction As Long
        If Invoke(m_pIMediaCapture, _
                  IMediaCapture_CapturePhotoToStreamAsync, _
                  ImageEncoding.Ifc, _
                  stream.Ifc, _
                  VarPtr(pIAsyncAction)) = S_OK Then
            If Await(pIAsyncAction, False) Then
                Ret = True
            End If
        End If
    End If
    CapturePhotoToStreamAsync = Ret
End Function

'    IMediaCapture_ClearEffectsAsync = 16
Public Function ClearEffectsAsync(ByVal StreamType As MediaStreamType) As Boolean
    Dim Ret As Boolean
    If m_pIMediaCapture <> 0& Then
        Dim pIAsyncAction As Long
        If Invoke(m_pIMediaCapture, _
                  IMediaCapture_ClearEffectsAsync, _
                  StreamType, _
                  VarPtr(pIAsyncAction)) = S_OK Then
            If Await(pIAsyncAction, False) Then
                Ret = True
            End If
        End If
    End If
    ClearEffectsAsync = Ret
End Function

Friend Function SetEncoderProperty(ByVal StreamType As MediaStreamType, _
                                   ByRef PropertyId As GUID, _
                                   ByVal PropertyValue As Inspectable) As Boolean
    Dim Ret As Boolean
    If m_pIMediaCapture <> 0& And _
       IsNotNothing(PropertyValue) Then
        If Invoke(m_pIMediaCapture, _
                  IMediaCapture_SetEncoderProperty, _
                  StreamType, _
                  VarPtr(PropertyId), _
                  PropertyValue.Ifc) = S_OK Then
            Ret = True
        End If
    End If
    SetEncoderProperty = Ret
End Function

Friend Function GetEncoderProperty(ByVal StreamType As MediaStreamType, _
                                   ByRef PropertyId As GUID) As Inspectable
    Dim Ret As Inspectable
    If m_pIMediaCapture <> 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIMediaCapture, _
                  IMediaCapture_GetEncoderProperty, _
                  StreamType, _
                  VarPtr(PropertyId), _
                  VarPtr(pIInspectable)) = S_OK Then
            If pIInspectable <> 0& Then
                Set Ret = New Inspectable
                Ret.Ifc = pIInspectable
            End If
        End If
    End If
    Set GetEncoderProperty = Ret
End Function

Public Function AddFailed(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaCapture <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaCapture, _
                  IMediaCapture_AddFailed, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddFailed = Ret
End Function

Public Function RemoveFailed(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaCapture <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaCapture, _
                  IMediaCapture_RemoveFailed, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveFailed = Ret
End Function

Public Function AddRecordLimitationExceeded(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaCapture <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaCapture, _
                  IMediaCapture_AddRecordLimitationExceeded, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddRecordLimitationExceeded = Ret
End Function

Public Function RemoveRecordLimitationExceeded(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaCapture <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaCapture, _
                  IMediaCapture_RemoveRecordLimitationExceeded, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveRecordLimitationExceeded = Ret
End Function

Public Function SetPreviewMirroring(ByVal value As Boolean) As Boolean
    Dim Ret As Boolean
    If m_pIMediaCapture <> 0& Then
        If Invoke(m_pIMediaCapture, _
                  IMediaCapture_SetPreviewMirroring, _
                  value) = S_OK Then
            Ret = True
        End If
    End If
    SetPreviewMirroring = Ret
End Function

Public Function GetPreviewMirroring() As Boolean
    Dim Ret As Boolean
    If m_pIMediaCapture <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaCapture, _
                  IMediaCapture_GetPreviewMirroring, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    GetPreviewMirroring = Ret
End Function

Public Function SetPreviewRotation(ByVal value As VideoRotation) As Boolean
    Dim Ret As Boolean
    If m_pIMediaCapture <> 0& Then
        If Invoke(m_pIMediaCapture, _
                  IMediaCapture_SetPreviewRotation, _
                  value) = S_OK Then
            Ret = True
        End If
    End If
    SetPreviewRotation = Ret
End Function

Public Function GetPreviewRotation() As VideoRotation
    Dim Ret As VideoRotation
    If m_pIMediaCapture <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaCapture, _
                  IMediaCapture_GetPreviewRotation, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    GetPreviewRotation = Ret
End Function

Public Function SetRecordRotation(ByVal value As VideoRotation) As Boolean
    Dim Ret As Boolean
    If m_pIMediaCapture <> 0& Then
        If Invoke(m_pIMediaCapture, _
                  IMediaCapture_SetRecordRotation, _
                  value) = S_OK Then
            Ret = True
        End If
    End If
    SetRecordRotation = Ret
End Function

Public Function GetRecordRotation() As VideoRotation
    Dim Ret As VideoRotation
    If m_pIMediaCapture <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaCapture, _
                  IMediaCapture_GetRecordRotation, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    GetRecordRotation = Ret
End Function

' IMediaCaptureVideoPreview
Public Function StartPreviewAsync() As Boolean
    Dim Ret As Boolean
    If m_pIMediaCaptureVideoPreview <> 0& Then
        Dim pIAsyncAction As Long
        If Invoke(m_pIMediaCaptureVideoPreview, _
                  IMediaCaptureVideoPreview_StartPreviewAsync, _
                  VarPtr(pIAsyncAction)) = S_OK Then
            If Await(pIAsyncAction, False) Then
                Ret = True
            End If
        End If
    End If
    StartPreviewAsync = Ret
End Function

'    IMediaCaptureVideoPreview_StartPreviewToCustomSinkAsync = 7
'    IMediaCaptureVideoPreview_StartPreviewToCustomSinkIdAsync = 8

Public Function StopPreviewAsync() As Boolean
    Dim Ret As Boolean
    If m_pIMediaCaptureVideoPreview <> 0& Then
        Dim pIAsyncAction As Long
        If Invoke(m_pIMediaCaptureVideoPreview, _
                  IMediaCaptureVideoPreview_StopPreviewAsync, _
                  VarPtr(pIAsyncAction)) = S_OK Then
            If Await(pIAsyncAction, False) Then
                Ret = True
            End If
        End If
    End If
    StopPreviewAsync = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


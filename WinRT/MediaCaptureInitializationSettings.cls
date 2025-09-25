VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaCaptureInitializationSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Capture_MediaCaptureInitializationSettings As String = "Windows.Media.Capture.MediaCaptureInitializationSettings"
Private Const IID_IMediaCaptureInitializationSettings As String = "{9782ba70-ea65-4900-9356-8ca887726884}"
Private Const IID_IMediaCaptureInitializationSettings2 As String = "{404e0626-c9dc-43e9-aee4-e6bf1b57b44c}"
Private Const IID_IMediaCaptureInitializationSettings3 As String = "{4160519d-be48-4730-8104-0cf6e9e97948}"
Private Const IID_IMediaCaptureInitializationSettings4 As String = "{f502a537-4cb7-4d28-95ed-4f9f012e0518}"
Private Const IID_IMediaCaptureInitializationSettings5 As String = "{d5a2e3b8-2626-4e94-b7b3-5308a0f64b1a}"
Private Const IID_IMediaCaptureInitializationSettings6 As String = "{b2e26b47-3db1-4d33-ab63-0ffa09056585}"
Private Const IID_IMediaCaptureInitializationSettings7 As String = "{41546967-f58a-5d82-9ef4-ed572fb5e34e}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IMediaCaptureInitializationSettings
    IMediaCaptureInitializationSettings_PutAudioDeviceId = 6
    IMediaCaptureInitializationSettings_GetAudioDeviceId = 7
    IMediaCaptureInitializationSettings_PutVideoDeviceId = 8
    IMediaCaptureInitializationSettings_GetVideoDeviceId = 9
    IMediaCaptureInitializationSettings_PutStreamingCaptureMode = 10
    IMediaCaptureInitializationSettings_GetStreamingCaptureMode = 11
    IMediaCaptureInitializationSettings_PutPhotoCaptureSource = 12
    IMediaCaptureInitializationSettings_GetPhotoCaptureSource = 13
                        
    ' IMediaCaptureInitializationSettings2
    IMediaCaptureInitializationSettings2_PutMediaCategory = 6
    IMediaCaptureInitializationSettings2_GetMediaCategory = 7
    IMediaCaptureInitializationSettings2_PutAudioProcessing = 8
    IMediaCaptureInitializationSettings2_GetAudioProcessing = 9
                        
    ' IMediaCaptureInitializationSettings3
'    IMediaCaptureInitializationSettings3_PutAudioSource = 6
'    IMediaCaptureInitializationSettings3_GetAudioSource = 7
'    IMediaCaptureInitializationSettings3_PutVideoSource = 8
'    IMediaCaptureInitializationSettings3_GetVideoSource = 9
                        
    ' IMediaCaptureInitializationSettings4
'    IMediaCaptureInitializationSettings4_GetVideoProfile = 6
'    IMediaCaptureInitializationSettings4_PutVideoProfile = 7
'    IMediaCaptureInitializationSettings4_GetPreviewMediaDescription = 8
'    IMediaCaptureInitializationSettings4_PutPreviewMediaDescription = 9
'    IMediaCaptureInitializationSettings4_GetRecordMediaDescription = 10
'    IMediaCaptureInitializationSettings4_PutRecordMediaDescription = 11
'    IMediaCaptureInitializationSettings4_GetPhotoMediaDescription = 12
'    IMediaCaptureInitializationSettings4_PutPhotoMediaDescription = 13
                        
    ' IMediaCaptureInitializationSettings5
'    IMediaCaptureInitializationSettings5_GetSourceGroup = 6
'    IMediaCaptureInitializationSettings5_PutSourceGroup = 7
'    IMediaCaptureInitializationSettings5_GetSharingMode = 8
'    IMediaCaptureInitializationSettings5_PutSharingMode = 9
'    IMediaCaptureInitializationSettings5_GetMemoryPreference = 10
'    IMediaCaptureInitializationSettings5_PutMemoryPreference = 11
                        
    ' IMediaCaptureInitializationSettings6
'    IMediaCaptureInitializationSettings6_GetAlwaysPlaySystemShutterSound = 6
'    IMediaCaptureInitializationSettings6_PutAlwaysPlaySystemShutterSound = 7
    
    ' IMediaCaptureInitializationSettings7
'    IMediaCaptureInitializationSettings7_GetDeviceUriPasswordCredential = 6
'    IMediaCaptureInitializationSettings7_PutDeviceUriPasswordCredential = 7
'    IMediaCaptureInitializationSettings7_GetDeviceUri = 8
'    IMediaCaptureInitializationSettings7_PutDeviceUri = 9
                        
End Enum

' ----==== Variablen ====----
Private m_pIMediaCaptureInitializationSettings As Long
Private m_pIMediaCaptureInitializationSettings2 As Long
Private m_pIMediaCaptureInitializationSettings3 As Long
Private m_pIMediaCaptureInitializationSettings4 As Long
Private m_pIMediaCaptureInitializationSettings5 As Long
Private m_pIMediaCaptureInitializationSettings6 As Long
Private m_pIMediaCaptureInitializationSettings7 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    Dim pIMediaCaptureInitializationSettings As Long
    If GetActivateInstance(Windows_Media_Capture_MediaCaptureInitializationSettings, _
                           IID_IMediaCaptureInitializationSettings, _
                           pIMediaCaptureInitializationSettings) Then
        Me.Ifc = pIMediaCaptureInitializationSettings
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMediaCaptureInitializationSettings7)
    Call ReleaseIfc(m_pIMediaCaptureInitializationSettings6)
    Call ReleaseIfc(m_pIMediaCaptureInitializationSettings5)
    Call ReleaseIfc(m_pIMediaCaptureInitializationSettings4)
    Call ReleaseIfc(m_pIMediaCaptureInitializationSettings3)
    Call ReleaseIfc(m_pIMediaCaptureInitializationSettings2)
    Call ReleaseIfc(m_pIMediaCaptureInitializationSettings)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMediaCaptureInitializationSettings
End Property

Public Property Let Ifc(ByVal pIMediaCaptureInitializationSettings As Long)
    Call ReleaseIfc(m_pIMediaCaptureInitializationSettings7)
    Call ReleaseIfc(m_pIMediaCaptureInitializationSettings6)
    Call ReleaseIfc(m_pIMediaCaptureInitializationSettings5)
    Call ReleaseIfc(m_pIMediaCaptureInitializationSettings4)
    Call ReleaseIfc(m_pIMediaCaptureInitializationSettings3)
    Call ReleaseIfc(m_pIMediaCaptureInitializationSettings2)
    Call ReleaseIfc(m_pIMediaCaptureInitializationSettings)
    m_pIMediaCaptureInitializationSettings = pIMediaCaptureInitializationSettings
    Call QueryIfc(m_pIMediaCaptureInitializationSettings, IID_IMediaCaptureInitializationSettings2, m_pIMediaCaptureInitializationSettings2)
    Call QueryIfc(m_pIMediaCaptureInitializationSettings, IID_IMediaCaptureInitializationSettings3, m_pIMediaCaptureInitializationSettings3)
    Call QueryIfc(m_pIMediaCaptureInitializationSettings, IID_IMediaCaptureInitializationSettings4, m_pIMediaCaptureInitializationSettings4)
    Call QueryIfc(m_pIMediaCaptureInitializationSettings, IID_IMediaCaptureInitializationSettings5, m_pIMediaCaptureInitializationSettings5)
    Call QueryIfc(m_pIMediaCaptureInitializationSettings, IID_IMediaCaptureInitializationSettings6, m_pIMediaCaptureInitializationSettings6)
    Call QueryIfc(m_pIMediaCaptureInitializationSettings, IID_IMediaCaptureInitializationSettings7, m_pIMediaCaptureInitializationSettings7)
End Property

' IMediaCaptureInitializationSettings
Public Property Let AudioDeviceId(ByVal Value As String)
    If m_pIMediaCaptureInitializationSettings <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Value)
        If Invoke(m_pIMediaCaptureInitializationSettings, _
                  IMediaCaptureInitializationSettings_PutAudioDeviceId, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get AudioDeviceId() As String
    Dim Ret As String
    If m_pIMediaCaptureInitializationSettings <> 0& Then
        Dim hString As Long
        If Invoke(m_pIMediaCaptureInitializationSettings, _
                  IMediaCaptureInitializationSettings_GetAudioDeviceId, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    AudioDeviceId = Ret
End Property

Public Property Let VideoDeviceId(ByVal Value As String)
    If m_pIMediaCaptureInitializationSettings <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Value)
        If Invoke(m_pIMediaCaptureInitializationSettings, _
                  IMediaCaptureInitializationSettings_PutVideoDeviceId, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get VideoDeviceId() As String
    Dim Ret As String
    If m_pIMediaCaptureInitializationSettings <> 0& Then
        Dim hString As Long
        If Invoke(m_pIMediaCaptureInitializationSettings, _
                  IMediaCaptureInitializationSettings_GetVideoDeviceId, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    VideoDeviceId = Ret
End Property

Public Property Let StreamingCaptureMode(ByVal Value As StreamingCaptureMode)
    If m_pIMediaCaptureInitializationSettings <> 0& Then
        If Invoke(m_pIMediaCaptureInitializationSettings, _
                  IMediaCaptureInitializationSettings_PutStreamingCaptureMode, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get StreamingCaptureMode() As StreamingCaptureMode
    Dim Ret As StreamingCaptureMode
    If m_pIMediaCaptureInitializationSettings <> 0& Then
        Dim Value As Long
        If Invoke(m_pIMediaCaptureInitializationSettings, _
                  IMediaCaptureInitializationSettings_GetStreamingCaptureMode, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    StreamingCaptureMode = Ret
End Property

Public Property Let PhotoCaptureSource(ByVal Value As PhotoCaptureSource)
    If m_pIMediaCaptureInitializationSettings <> 0& Then
        If Invoke(m_pIMediaCaptureInitializationSettings, _
                  IMediaCaptureInitializationSettings_PutPhotoCaptureSource, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get PhotoCaptureSource() As PhotoCaptureSource
    Dim Ret As PhotoCaptureSource
    If m_pIMediaCaptureInitializationSettings <> 0& Then
        Dim Value As Long
        If Invoke(m_pIMediaCaptureInitializationSettings, _
                  IMediaCaptureInitializationSettings_GetPhotoCaptureSource, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    PhotoCaptureSource = Ret
End Property

' IMediaCaptureInitializationSettings2
Public Property Let MediaCategory(ByVal Value As MediaCategory)
    If m_pIMediaCaptureInitializationSettings2 <> 0& Then
        If Invoke(m_pIMediaCaptureInitializationSettings2, _
                  IMediaCaptureInitializationSettings2_PutMediaCategory, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get MediaCategory() As MediaCategory
    Dim Ret As MediaCategory
    If m_pIMediaCaptureInitializationSettings2 <> 0& Then
        Dim Value As Long
        If Invoke(m_pIMediaCaptureInitializationSettings2, _
                  IMediaCaptureInitializationSettings2_GetMediaCategory, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    MediaCategory = Ret
End Property

Public Property Let AudioProcessing(ByVal Value As AudioProcessing)
    If m_pIMediaCaptureInitializationSettings2 <> 0& Then
        If Invoke(m_pIMediaCaptureInitializationSettings2, _
                  IMediaCaptureInitializationSettings2_PutAudioProcessing, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get AudioProcessing() As AudioProcessing
    Dim Ret As AudioProcessing
    If m_pIMediaCaptureInitializationSettings2 <> 0& Then
        Dim Value As Long
        If Invoke(m_pIMediaCaptureInitializationSettings2, _
                  IMediaCaptureInitializationSettings2_GetAudioProcessing, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    AudioProcessing = Ret
End Property
                        
' IMediaCaptureInitializationSettings3
'    IMediaCaptureInitializationSettings3_PutAudioSource = 6
'    IMediaCaptureInitializationSettings3_GetAudioSource = 7
'    IMediaCaptureInitializationSettings3_PutVideoSource = 8
'    IMediaCaptureInitializationSettings3_GetVideoSource = 9
                        
' IMediaCaptureInitializationSettings4
'    IMediaCaptureInitializationSettings4_GetVideoProfile = 6
'    IMediaCaptureInitializationSettings4_PutVideoProfile = 7
'    IMediaCaptureInitializationSettings4_GetPreviewMediaDescription = 8
'    IMediaCaptureInitializationSettings4_PutPreviewMediaDescription = 9
'    IMediaCaptureInitializationSettings4_GetRecordMediaDescription = 10
'    IMediaCaptureInitializationSettings4_PutRecordMediaDescription = 11
'    IMediaCaptureInitializationSettings4_GetPhotoMediaDescription = 12
'    IMediaCaptureInitializationSettings4_PutPhotoMediaDescription = 13
                        
' IMediaCaptureInitializationSettings5
'    IMediaCaptureInitializationSettings5_GetSourceGroup = 6
'    IMediaCaptureInitializationSettings5_PutSourceGroup = 7
'    IMediaCaptureInitializationSettings5_GetSharingMode = 8
'    IMediaCaptureInitializationSettings5_PutSharingMode = 9
'    IMediaCaptureInitializationSettings5_GetMemoryPreference = 10
'    IMediaCaptureInitializationSettings5_PutMemoryPreference = 11
                        
' IMediaCaptureInitializationSettings6
'    IMediaCaptureInitializationSettings6_GetAlwaysPlaySystemShutterSound = 6
'    IMediaCaptureInitializationSettings6_PutAlwaysPlaySystemShutterSound = 7
    
' IMediaCaptureInitializationSettings7
'    IMediaCaptureInitializationSettings7_GetDeviceUriPasswordCredential = 6
'    IMediaCaptureInitializationSettings7_PutDeviceUriPasswordCredential = 7
'    IMediaCaptureInitializationSettings7_GetDeviceUri = 8
'    IMediaCaptureInitializationSettings7_PutDeviceUri = 9

' ----==== Functions ====----

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



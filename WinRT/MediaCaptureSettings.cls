VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaCaptureSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 11/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Capture_MediaCaptureSettings As String = "Windows.Media.Capture.MediaCaptureSettings"
Private Const IID_IMediaCaptureSettings As String = "{1d83aafe-6d45-4477-8dc4-ac5bc01c4091}"
Private Const IID_IMediaCaptureSettings2 As String = "{6f9e7cfb-fa9f-4b13-9cbe-5ab94f1f3493}"
Private Const IID_IMediaCaptureSettings3 As String = "{303c67c2-8058-4b1b-b877-8c2ef3528440}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IMediaCaptureSettings
    IMediaCaptureSettings_GetAudioDeviceId = 6
    IMediaCaptureSettings_GetVideoDeviceId = 7
    IMediaCaptureSettings_GetStreamingCaptureMode = 8
    IMediaCaptureSettings_GetPhotoCaptureSource = 9
    IMediaCaptureSettings_GetVideoDeviceCharacteristic = 10
                        
    ' IMediaCaptureSettings2
    IMediaCaptureSettings2_GetConcurrentRecordAndPhotoSupported = 6
    IMediaCaptureSettings2_GetConcurrentRecordAndPhotoSequenceSupported = 7
    IMediaCaptureSettings2_GetCameraSoundRequiredForRegion = 8
    IMediaCaptureSettings2_GetHorizontal35mmEquivalentFocalLength = 9
    IMediaCaptureSettings2_GetPitchOffsetDegrees = 10
    IMediaCaptureSettings2_GetVertical35mmEquivalentFocalLength = 11
    IMediaCaptureSettings2_GetMediaCategory = 12
    IMediaCaptureSettings2_GetAudioProcessing = 13
                        
    ' IMediaCaptureSettings3
    IMediaCaptureSettings3_GetDirect3D11Device = 6

    ' IReference
    IReference_GetValue = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIMediaCaptureSettings As Long
Private m_pIMediaCaptureSettings2 As Long
Private m_pIMediaCaptureSettings3 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMediaCaptureSettings3)
    Call ReleaseIfc(m_pIMediaCaptureSettings2)
    Call ReleaseIfc(m_pIMediaCaptureSettings)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMediaCaptureSettings
End Property

Public Property Let Ifc(ByVal pIMediaCaptureSettings As Long)
    Call ReleaseIfc(m_pIMediaCaptureSettings3)
    Call ReleaseIfc(m_pIMediaCaptureSettings2)
    Call ReleaseIfc(m_pIMediaCaptureSettings)
    m_pIMediaCaptureSettings = pIMediaCaptureSettings
    Call QueryIfc(m_pIMediaCaptureSettings, IID_IMediaCaptureSettings2, m_pIMediaCaptureSettings2)
    Call QueryIfc(m_pIMediaCaptureSettings, IID_IMediaCaptureSettings3, m_pIMediaCaptureSettings3)
End Property

' IMediaCaptureSettings
Public Property Get AudioDeviceId() As String
    Dim Ret As String
    If m_pIMediaCaptureSettings <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaCaptureSettings, _
                  IMediaCaptureSettings_GetAudioDeviceId, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    AudioDeviceId = Ret
End Property

Public Property Get VideoDeviceId() As String
    Dim Ret As String
    If m_pIMediaCaptureSettings <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaCaptureSettings, _
                  IMediaCaptureSettings_GetVideoDeviceId, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    VideoDeviceId = Ret
End Property

Public Property Get StreamingCaptureMode() As StreamingCaptureMode
    Dim Ret As StreamingCaptureMode
    If m_pIMediaCaptureSettings <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaCaptureSettings, _
                  IMediaCaptureSettings_GetStreamingCaptureMode, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    StreamingCaptureMode = Ret
End Property

Public Property Get PhotoCaptureSource() As PhotoCaptureSource
    Dim Ret As PhotoCaptureSource
    If m_pIMediaCaptureSettings <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaCaptureSettings, _
                  IMediaCaptureSettings_GetPhotoCaptureSource, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    PhotoCaptureSource = Ret
End Property

Public Property Get VideoDeviceCharacteristic() As VideoDeviceCharacteristic
    Dim Ret As VideoDeviceCharacteristic
    If m_pIMediaCaptureSettings <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaCaptureSettings, _
                  IMediaCaptureSettings_GetVideoDeviceCharacteristic, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    VideoDeviceCharacteristic = Ret
End Property
                        
' IMediaCaptureSettings2
Public Property Get ConcurrentRecordAndPhotoSupported() As Boolean
    Dim Ret As Boolean
    If m_pIMediaCaptureSettings2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaCaptureSettings2, _
                  IMediaCaptureSettings2_GetConcurrentRecordAndPhotoSupported, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    ConcurrentRecordAndPhotoSupported = Ret
End Property

Public Property Get ConcurrentRecordAndPhotoSequenceSupported() As Boolean
    Dim Ret As Boolean
    If m_pIMediaCaptureSettings2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaCaptureSettings2, _
                  IMediaCaptureSettings2_GetConcurrentRecordAndPhotoSequenceSupported, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    ConcurrentRecordAndPhotoSequenceSupported = Ret
End Property

Public Property Get CameraSoundRequiredForRegion() As Boolean
    Dim Ret As Boolean
    If m_pIMediaCaptureSettings2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaCaptureSettings2, _
                  IMediaCaptureSettings2_GetCameraSoundRequiredForRegion, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    CameraSoundRequiredForRegion = Ret
End Property

Public Property Get Horizontal35mmEquivalentFocalLength() As Long
    Dim Ret As Long
    If m_pIMediaCaptureSettings2 <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIMediaCaptureSettings2, _
                  IMediaCaptureSettings2_GetHorizontal35mmEquivalentFocalLength, _
                  VarPtr(pIReference)) = S_OK Then
            If pIReference <> 0& Then
                Dim value As Long
                If Invoke(pIReference, IReference_GetValue, VarPtr(value)) = S_OK Then
                    Ret = value
                End If
                Call ReleaseIfc(pIReference)
            End If
        End If
    End If
    Horizontal35mmEquivalentFocalLength = Ret
End Property

Public Property Get PitchOffsetDegrees() As Long
    Dim Ret As Long
    If m_pIMediaCaptureSettings2 <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIMediaCaptureSettings2, _
                  IMediaCaptureSettings2_GetPitchOffsetDegrees, _
                  VarPtr(pIReference)) = S_OK Then
            If pIReference <> 0& Then
                Dim value As Long
                If Invoke(pIReference, IReference_GetValue, VarPtr(value)) = S_OK Then
                    Ret = value
                End If
                Call ReleaseIfc(pIReference)
            End If
        End If
    End If
    PitchOffsetDegrees = Ret
End Property

Public Property Get Vertical35mmEquivalentFocalLength() As Long
    Dim Ret As Long
    If m_pIMediaCaptureSettings2 <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIMediaCaptureSettings2, _
                  IMediaCaptureSettings2_GetVertical35mmEquivalentFocalLength, _
                  VarPtr(pIReference)) = S_OK Then
            If pIReference <> 0& Then
                Dim value As Long
                If Invoke(pIReference, IReference_GetValue, VarPtr(value)) = S_OK Then
                    Ret = value
                End If
                Call ReleaseIfc(pIReference)
            End If
        End If
    End If
    Vertical35mmEquivalentFocalLength = Ret
End Property

Public Property Get MediaCategory() As MediaCategory
    Dim Ret As MediaCategory
    If m_pIMediaCaptureSettings2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaCaptureSettings2, _
                  IMediaCaptureSettings2_GetMediaCategory, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    MediaCategory = Ret
End Property

Public Property Get AudioProcessing() As AudioProcessing
    Dim Ret As AudioProcessing
    If m_pIMediaCaptureSettings2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaCaptureSettings2, _
                  IMediaCaptureSettings2_GetAudioProcessing, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    AudioProcessing = Ret
End Property
                        
' IMediaCaptureSettings3
Public Property Get Direct3D11Device() As Direct3DDevice
    Dim Ret As Direct3DDevice
    If m_pIMediaCaptureSettings3 <> 0& Then
        Dim pIDirect3DDevice As Long
        If Invoke(m_pIMediaCaptureSettings3, _
                  IMediaCaptureSettings3_GetDirect3D11Device, _
                  VarPtr(pIDirect3DDevice)) = S_OK Then
            If pIDirect3DDevice <> 0& Then
                Set Ret = New Direct3DDevice
                Ret.Ifc = pIDirect3DDevice
            End If
        End If
    End If
    Set Direct3D11Device = Ret
End Property

' ----==== Functions ====----

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



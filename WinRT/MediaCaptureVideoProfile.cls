VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaCaptureVideoProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Capture_MediaCaptureVideoProfile As String = "Windows.Media.Capture.MediaCaptureVideoProfile"
Private Const IID_IMediaCaptureVideoProfile As String = "{21a073bf-a3ee-4ecf-9ef6-50b0bc4e1305}"
Private Const IID_IMediaCaptureVideoProfile2 As String = "{97ddc95f-94ce-468f-9316-fc5bc2638f6b}"
 
' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IMediaCaptureVideoProfile
    IMediaCaptureVideoProfile_GetId = 6
    IMediaCaptureVideoProfile_GetVideoDeviceId = 7
'    IMediaCaptureVideoProfile_GetSupportedPreviewMediaDescription = 8
'    IMediaCaptureVideoProfile_GetSupportedRecordMediaDescription = 9
'    IMediaCaptureVideoProfile_GetSupportedPhotoMediaDescription = 10
'    IMediaCaptureVideoProfile_GetConcurrency = 11
                        
    ' IMediaCaptureVideoProfile2
'    IMediaCaptureVideoProfile_GetFrameSourceInfos = 6
'    IMediaCaptureVideoProfile_GetProperties = 7
                        
End Enum

' ----==== Variablen ====----
Private m_pIMediaCaptureVideoProfile As Long
Private m_pIMediaCaptureVideoProfile2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMediaCaptureVideoProfile2)
    Call ReleaseIfc(m_pIMediaCaptureVideoProfile)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMediaCaptureVideoProfile
End Property

Public Property Let Ifc(ByVal pIMediaCaptureVideoProfile As Long)
    Call ReleaseIfc(m_pIMediaCaptureVideoProfile2)
    Call ReleaseIfc(m_pIMediaCaptureVideoProfile)
    m_pIMediaCaptureVideoProfile = pIMediaCaptureVideoProfile
    Call QueryIfc(m_pIMediaCaptureVideoProfile, IID_IMediaCaptureVideoProfile2, m_pIMediaCaptureVideoProfile2)
End Property

    ' IMediaCaptureVideoProfile
Public Property Get Id() As String
    Dim Ret As String
    If m_pIMediaCaptureVideoProfile <> 0& Then
        Dim Value As Long
        If Invoke(m_pIMediaCaptureVideoProfile, _
                  IMediaCaptureVideoProfile_GetId, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    Id = Ret
End Property

Public Property Get VideoDeviceId() As String
    Dim Ret As String
    If m_pIMediaCaptureVideoProfile <> 0& Then
        Dim Value As Long
        If Invoke(m_pIMediaCaptureVideoProfile, _
                  IMediaCaptureVideoProfile_GetVideoDeviceId, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    VideoDeviceId = Ret
End Property

'    MediaCaptureVideoProfile_GetSupportedPreviewMediaDescription = 8
'    MediaCaptureVideoProfile_GetSupportedRecordMediaDescription = 9
'    MediaCaptureVideoProfile_GetSupportedPhotoMediaDescription = 10
                        
    ' IMediaCaptureVideoProfile2
'    MediaCaptureVideoProfile_GetFrameSourceInfos = 6
'    MediaCaptureVideoProfile_GetProperties = 7

' ----==== Functions ====----
    ' IMediaCaptureVideoProfile
'    MediaCaptureVideoProfile_GetConcurrency = 11

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



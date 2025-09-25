VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaEncodingProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_MediaProperties_MediaEncodingProfile As String = "Windows.Media.MediaProperties.MediaEncodingProfile"
Private Const IID_IMediaEncodingProfile As String = "{e7dbf5a8-1db9-4783-876b-3dfe12acfdb3}"
Private Const IID_IMediaEncodingProfile2 As String = "{349b3e0a-4035-488e-9877-85632865ed10}"
Private Const IID_IMediaEncodingProfile3 As String = "{ba6ebe88-7570-4e69-accf-5611ad015f88}"
Private Const IID_IMediaEncodingProfileStatics As String = "{197f352c-2ede-4a45-a896-817a4854f8fe}"
Private Const IID_IMediaEncodingProfileStatics2 As String = "{ce8de74f-6af4-4288-8fe2-79adf1f79a43}"
Private Const IID_IMediaEncodingProfileStatics3 As String = "{90dac5aa-cf76-4294-a9ed-1a1420f51f6b}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IMediaEncodingProfileStatics
    IMediaEncodingProfileStatics_CreateM4a = 6
    IMediaEncodingProfileStatics_CreateMp3 = 7
    IMediaEncodingProfileStatics_CreateWma = 8
    IMediaEncodingProfileStatics_CreateMp4 = 9
    IMediaEncodingProfileStatics_CreateWmv = 10
    IMediaEncodingProfileStatics_CreateFromFileAsync = 11
'    IMediaEncodingProfileStatics_CreateFromStreamAsync = 12
                        
    ' IMediaEncodingProfileStatics2
    IMediaEncodingProfileStatics2_CreateWav = 6
    IMediaEncodingProfileStatics2_CreateAvi = 7
    
    ' IMediaEncodingProfileStatics3
    IMediaEncodingProfileStatics3_CreateAlac = 6
    IMediaEncodingProfileStatics3_CreateFlac = 7
    IMediaEncodingProfileStatics3_CreateHevc = 8
    
    ' IMediaEncodingProfile
    IMediaEncodingProfile_PutAudio = 6
    IMediaEncodingProfile_GetAudio = 7
    IMediaEncodingProfile_PutVideo = 8
    IMediaEncodingProfile_GetVideo = 9
'    IMediaEncodingProfile_PutContainer = 10
'    IMediaEncodingProfile_GetContainer = 11
                        
    ' IMediaEncodingProfile2
'    IMediaEncodingProfile2_SetAudioTracks = 6
'    IMediaEncodingProfile2_GetAudioTracks = 7
'    IMediaEncodingProfile2_SetVideoTracks = 8
'    IMediaEncodingProfile2_GetVideoTracks = 9
                        
    ' IMediaEncodingProfile3
'    IMediaEncodingProfile3_SetTimedMetadataTracks = 6
'    IMediaEncodingProfile3_GetTimedMetadataTracks = 7
    
End Enum

' ----==== Variablen ====----
Private m_pIMediaEncodingProfile As Long
Private m_pIMediaEncodingProfile2 As Long
Private m_pIMediaEncodingProfile3 As Long
Private m_pIMediaEncodingProfileStatics As Long
Private m_pIMediaEncodingProfileStatics2 As Long
Private m_pIMediaEncodingProfileStatics3 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Media_MediaProperties_MediaEncodingProfile, _
                            IID_IMediaEncodingProfileStatics, _
                            m_pIMediaEncodingProfileStatics) Then
        Call QueryIfc(m_pIMediaEncodingProfileStatics, IID_IMediaEncodingProfileStatics2, m_pIMediaEncodingProfileStatics2)
        Call QueryIfc(m_pIMediaEncodingProfileStatics, IID_IMediaEncodingProfileStatics3, m_pIMediaEncodingProfileStatics3)
    End If
    Dim pIMediaEncodingProfile As Long
    If GetActivateInstance(Windows_Media_MediaProperties_MediaEncodingProfile, _
                           IID_IMediaEncodingProfile, _
                           pIMediaEncodingProfile) Then
        Me.Ifc = pIMediaEncodingProfile
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMediaEncodingProfile3)
    Call ReleaseIfc(m_pIMediaEncodingProfile2)
    Call ReleaseIfc(m_pIMediaEncodingProfile)
    Call ReleaseIfc(m_pIMediaEncodingProfileStatics3)
    Call ReleaseIfc(m_pIMediaEncodingProfileStatics2)
    Call ReleaseIfc(m_pIMediaEncodingProfileStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMediaEncodingProfile
End Property

Public Property Let Ifc(ByVal pIMediaEncodingProfile As Long)
    Call ReleaseIfc(m_pIMediaEncodingProfile3)
    Call ReleaseIfc(m_pIMediaEncodingProfile2)
    Call ReleaseIfc(m_pIMediaEncodingProfile)
    m_pIMediaEncodingProfile = pIMediaEncodingProfile
    Call QueryIfc(m_pIMediaEncodingProfile, IID_IMediaEncodingProfile2, m_pIMediaEncodingProfile2)
    Call QueryIfc(m_pIMediaEncodingProfile, IID_IMediaEncodingProfile3, m_pIMediaEncodingProfile3)
End Property

' IMediaEncodingProfile
Public Property Let Audio(ByVal value As AudioEncodingProperties)
    If m_pIMediaEncodingProfile <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pIMediaEncodingProfile, _
                  IMediaEncodingProfile_PutAudio, _
                  value.Ifc) = S_OK Then
        End If
    End If
End Property

Public Property Get Audio() As AudioEncodingProperties
    Dim Ret As AudioEncodingProperties
    If m_pIMediaEncodingProfile <> 0& Then
        Dim pIAudioEncodingProperties As Long
        If Invoke(m_pIMediaEncodingProfile, _
                  IMediaEncodingProfile_GetAudio, _
                  VarPtr(pIAudioEncodingProperties)) = S_OK Then
            If pIAudioEncodingProperties <> 0& Then
                Set Ret = New AudioEncodingProperties
                Ret.Ifc = pIAudioEncodingProperties
            End If
        End If
    End If
    Set Audio = Ret
End Property

Public Property Let Video(ByVal value As VideoEncodingProperties)
    If m_pIMediaEncodingProfile <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pIMediaEncodingProfile, _
                  IMediaEncodingProfile_PutVideo, _
                  value.Ifc) = S_OK Then
        End If
    End If
End Property

Public Property Get Video() As VideoEncodingProperties
    Dim Ret As VideoEncodingProperties
    If m_pIMediaEncodingProfile <> 0& Then
        Dim pIVideoEncodingProperties As Long
        If Invoke(m_pIMediaEncodingProfile, _
                  IMediaEncodingProfile_GetVideo, _
                  VarPtr(pIVideoEncodingProperties)) = S_OK Then
            If pIVideoEncodingProperties <> 0& Then
                Set Ret = New VideoEncodingProperties
                Ret.Ifc = pIVideoEncodingProperties
            End If
        End If
    End If
    Set Video = Ret
End Property

'    IMediaEncodingProfile_PutContainer = 10
'    IMediaEncodingProfile_GetContainer = 11

' ----==== Functions ====----
' IMediaEncodingProfileStatics
Public Function CreateM4a(ByVal quality As AudioEncodingQuality) As MediaEncodingProfile
    Dim Ret As MediaEncodingProfile
    If m_pIMediaEncodingProfileStatics <> 0& Then
        Dim pIMediaEncodingProfile As Long
        If Invoke(m_pIMediaEncodingProfileStatics, _
                  IMediaEncodingProfileStatics_CreateM4a, _
                  quality, _
                  VarPtr(pIMediaEncodingProfile)) = S_OK Then
            If pIMediaEncodingProfile <> 0& Then
                Set Ret = New MediaEncodingProfile
                Ret.Ifc = pIMediaEncodingProfile
            End If
        End If
    End If
    Set CreateM4a = Ret
End Function

Public Function CreateMp3(ByVal quality As AudioEncodingQuality) As MediaEncodingProfile
    Dim Ret As MediaEncodingProfile
    If m_pIMediaEncodingProfileStatics <> 0& Then
        Dim pIMediaEncodingProfile As Long
        If Invoke(m_pIMediaEncodingProfileStatics, _
                  IMediaEncodingProfileStatics_CreateMp3, _
                  quality, _
                  VarPtr(pIMediaEncodingProfile)) = S_OK Then
            If pIMediaEncodingProfile <> 0& Then
                Set Ret = New MediaEncodingProfile
                Ret.Ifc = pIMediaEncodingProfile
            End If
        End If
    End If
    Set CreateMp3 = Ret
End Function

Public Function CreateWma(ByVal quality As AudioEncodingQuality) As MediaEncodingProfile
    Dim Ret As MediaEncodingProfile
    If m_pIMediaEncodingProfileStatics <> 0& Then
        Dim pIMediaEncodingProfile As Long
        If Invoke(m_pIMediaEncodingProfileStatics, _
                  IMediaEncodingProfileStatics_CreateWma, _
                  quality, _
                  VarPtr(pIMediaEncodingProfile)) = S_OK Then
            If pIMediaEncodingProfile <> 0& Then
                Set Ret = New MediaEncodingProfile
                Ret.Ifc = pIMediaEncodingProfile
            End If
        End If
    End If
    Set CreateWma = Ret
End Function

Public Function CreateMp4(ByVal quality As VideoEncodingQuality) As MediaEncodingProfile
    Dim Ret As MediaEncodingProfile
    If m_pIMediaEncodingProfileStatics <> 0& Then
        Dim pIMediaEncodingProfile As Long
        If Invoke(m_pIMediaEncodingProfileStatics, _
                  IMediaEncodingProfileStatics_CreateMp4, _
                  quality, _
                  VarPtr(pIMediaEncodingProfile)) = S_OK Then
            If pIMediaEncodingProfile <> 0& Then
                Set Ret = New MediaEncodingProfile
                Ret.Ifc = pIMediaEncodingProfile
            End If
        End If
    End If
    Set CreateMp4 = Ret
End Function

Public Function CreateWmv(ByVal quality As VideoEncodingQuality) As MediaEncodingProfile
    Dim Ret As MediaEncodingProfile
    If m_pIMediaEncodingProfileStatics <> 0& Then
        Dim pIMediaEncodingProfile As Long
        If Invoke(m_pIMediaEncodingProfileStatics, _
                  IMediaEncodingProfileStatics_CreateWmv, _
                  quality, _
                  VarPtr(pIMediaEncodingProfile)) = S_OK Then
            If pIMediaEncodingProfile <> 0& Then
                Set Ret = New MediaEncodingProfile
                Ret.Ifc = pIMediaEncodingProfile
            End If
        End If
    End If
    Set CreateWmv = Ret
End Function

Public Function CreateFromFileAsync(ByVal File As StorageFile) As MediaEncodingProfile
    Dim Ret As MediaEncodingProfile
    If m_pIMediaEncodingProfileStatics <> 0& And _
       IsNotNothing(File) Then
        Dim pIMediaEncodingProfile As Long
        If Invoke(m_pIMediaEncodingProfileStatics, _
                  IMediaEncodingProfileStatics_CreateFromFileAsync, _
                  File.Ifc, _
                  VarPtr(pIMediaEncodingProfile)) = S_OK Then
            If pIMediaEncodingProfile <> 0& Then
                If Await(pIMediaEncodingProfile) Then
                    Ret.Ifc = pIMediaEncodingProfile
                Else
                    Set Ret = Nothing
                End If
            Else
                Set Ret = Nothing
            End If
        Else
            Set Ret = Nothing
        End If
    End If
    Set CreateFromFileAsync = Ret
End Function

'    IMediaEncodingProfileStatics_CreateFromStreamAsync = 12
                        
' IMediaEncodingProfileStatics2
Public Function CreateWav(ByVal quality As AudioEncodingQuality) As MediaEncodingProfile
    Dim Ret As MediaEncodingProfile
    If m_pIMediaEncodingProfileStatics2 <> 0& Then
        Dim pIMediaEncodingProfile As Long
        If Invoke(m_pIMediaEncodingProfileStatics2, _
                  IMediaEncodingProfileStatics2_CreateWav, _
                  quality, _
                  VarPtr(pIMediaEncodingProfile)) = S_OK Then
            If pIMediaEncodingProfile <> 0& Then
                Set Ret = New MediaEncodingProfile
                Ret.Ifc = pIMediaEncodingProfile
            End If
        End If
    End If
    Set CreateWav = Ret
End Function

Public Function CreateAvi(ByVal quality As VideoEncodingQuality) As MediaEncodingProfile
    Dim Ret As MediaEncodingProfile
    If m_pIMediaEncodingProfileStatics2 <> 0& Then
        Dim pIMediaEncodingProfile As Long
        If Invoke(m_pIMediaEncodingProfileStatics2, _
                  IMediaEncodingProfileStatics2_CreateAvi, _
                  quality, _
                  VarPtr(pIMediaEncodingProfile)) = S_OK Then
            If pIMediaEncodingProfile <> 0& Then
                Set Ret = New MediaEncodingProfile
                Ret.Ifc = pIMediaEncodingProfile
            End If
        End If
    End If
    Set CreateAvi = Ret
End Function
    
' IMediaEncodingProfileStatics3
Public Function CreateAlac(ByVal quality As AudioEncodingQuality) As MediaEncodingProfile
    Dim Ret As MediaEncodingProfile
    If m_pIMediaEncodingProfileStatics3 <> 0& Then
        Dim pIMediaEncodingProfile As Long
        If Invoke(m_pIMediaEncodingProfileStatics3, _
                  IMediaEncodingProfileStatics3_CreateAlac, _
                  quality, _
                  VarPtr(pIMediaEncodingProfile)) = S_OK Then
            If pIMediaEncodingProfile <> 0& Then
                Set Ret = New MediaEncodingProfile
                Ret.Ifc = pIMediaEncodingProfile
            End If
        End If
    End If
    Set CreateAlac = Ret
End Function

Public Function CreateFlac(ByVal quality As AudioEncodingQuality) As MediaEncodingProfile
    Dim Ret As MediaEncodingProfile
    If m_pIMediaEncodingProfileStatics3 <> 0& Then
        Dim pIMediaEncodingProfile As Long
        If Invoke(m_pIMediaEncodingProfileStatics3, _
                  IMediaEncodingProfileStatics3_CreateFlac, _
                  quality, _
                  VarPtr(pIMediaEncodingProfile)) = S_OK Then
            If pIMediaEncodingProfile <> 0& Then
                Set Ret = New MediaEncodingProfile
                Ret.Ifc = pIMediaEncodingProfile
            End If
        End If
    End If
    Set CreateFlac = Ret
End Function

Public Function CreateHevc(ByVal quality As VideoEncodingQuality) As MediaEncodingProfile
    Dim Ret As MediaEncodingProfile
    If m_pIMediaEncodingProfileStatics3 <> 0& Then
        Dim pIMediaEncodingProfile As Long
        If Invoke(m_pIMediaEncodingProfileStatics3, _
                  IMediaEncodingProfileStatics3_CreateHevc, _
                  quality, _
                  VarPtr(pIMediaEncodingProfile)) = S_OK Then
            If pIMediaEncodingProfile <> 0& Then
                Set Ret = New MediaEncodingProfile
                Ret.Ifc = pIMediaEncodingProfile
            End If
        End If
    End If
    Set CreateHevc = Ret
End Function

' IMediaEncodingProfile2
'    IMediaEncodingProfile2_SetAudioTracks = 6
'    IMediaEncodingProfile2_GetAudioTracks = 7
'    IMediaEncodingProfile2_SetVideoTracks = 8
'    IMediaEncodingProfile2_GetVideoTracks = 9
                        
' IMediaEncodingProfile3
'    IMediaEncodingProfile3_SetTimedMetadataTracks = 6
'    IMediaEncodingProfile3_GetTimedMetadataTracks = 7

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



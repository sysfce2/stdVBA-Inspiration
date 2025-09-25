VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VideoEncodingProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_MediaProperties_VideoEncodingProperties As String = "Windows.Media.MediaProperties.VideoEncodingProperties"
Private Const IID_IMediaEncodingProperties As String = "{b4002af6-acd4-4e5a-a24b-5d7498a8b8c4}"
Private Const IID_IVideoEncodingProperties As String = "{76ee6c9a-37c2-4f2a-880a-1282bbb4373d}"
Private Const IID_IVideoEncodingProperties2 As String = "{f743a1ef-d465-4290-a94b-ef0f1528f8e3}"
Private Const IID_IVideoEncodingProperties3 As String = "{386bcdc4-873a-479f-b3eb-56c1fcbec6d7}"
Private Const IID_IVideoEncodingProperties4 As String = "{724ef014-c10c-40f2-9d72-3ee13b45fa8e}"
Private Const IID_IVideoEncodingProperties5 As String = "{4959080f-272f-4ece-a4df-c0ccdb33d840}"
Private Const IID_IVideoEncodingPropertiesStatics As String = "{3ce14d44-1dc5-43db-9f38-ebebf90152cb}"
Private Const IID_IVideoEncodingPropertiesStatics2 As String = "{cf1ebd5d-49fe-4d00-b59a-cfa4dfc51944}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IVideoEncodingPropertiesStatics
    IVideoEncodingPropertiesStatics_CreateH264 = 6
    IVideoEncodingPropertiesStatics_CreateMpeg2 = 7
    IVideoEncodingPropertiesStatics_CreateUncompressed = 8
    
    ' IVideoEncodingPropertiesStatics2
    IVideoEncodingPropertiesStatics2_CreateHevc = 6
        
    ' IMediaEncodingProperties
'    IMediaEncodingProperties_GetProperties = 6
    IMediaEncodingProperties_GetType = 7
    IMediaEncodingProperties_PutSubtype = 8
    IMediaEncodingProperties_GetSubtype = 9
                    
    ' IVideoEncodingProperties
    IVideoEncodingProperties_PutBitrate = 6
    IVideoEncodingProperties_GetBitrate = 7
    IVideoEncodingProperties_PutWidth = 8
    IVideoEncodingProperties_GetWidth = 9
    IVideoEncodingProperties_PutHeight = 10
    IVideoEncodingProperties_GetHeight = 11
'    IVideoEncodingProperties_GetFrameRate = 12
'    IVideoEncodingProperties_GetPixelAspectRatio = 13
    
    ' IVideoEncodingProperties2
'    IVideoEncodingProperties2_SetFormatUserData = 6
'    IVideoEncodingProperties2_GetFormatUserData = 7
    IVideoEncodingProperties2_PutProfileId = 8
    IVideoEncodingProperties2_GetProfileId = 9
    
    ' IVideoEncodingProperties3
    IVideoEncodingProperties3_GetStereoscopicVideoPackingMode = 6
    
    ' IVideoEncodingProperties4
    IVideoEncodingProperties4_GetSphericalVideoFrameFormat = 6
    
    ' IVideoEncodingProperties5
    IVideoEncodingProperties5_Copy = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIMediaEncodingProperties As Long
Private m_pIVideoEncodingProperties As Long
Private m_pIVideoEncodingProperties2 As Long
Private m_pIVideoEncodingProperties3 As Long
Private m_pIVideoEncodingProperties4 As Long
Private m_pIVideoEncodingProperties5 As Long
Private m_pIVideoEncodingPropertiesStatics As Long
Private m_pIVideoEncodingPropertiesStatics2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Media_MediaProperties_VideoEncodingProperties, _
                            IID_IVideoEncodingPropertiesStatics, _
                            m_pIVideoEncodingPropertiesStatics) Then
        Call QueryIfc(m_pIVideoEncodingPropertiesStatics, IID_IVideoEncodingPropertiesStatics2, m_pIVideoEncodingPropertiesStatics2)
    End If
    Dim pIVideoEncodingProperties As Long
    If GetActivateInstance(Windows_Media_MediaProperties_VideoEncodingProperties, _
                           IID_IVideoEncodingProperties, _
                           pIVideoEncodingProperties) Then
        Me.Ifc = pIVideoEncodingProperties
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMediaEncodingProperties)
    Call ReleaseIfc(m_pIVideoEncodingProperties5)
    Call ReleaseIfc(m_pIVideoEncodingProperties4)
    Call ReleaseIfc(m_pIVideoEncodingProperties3)
    Call ReleaseIfc(m_pIVideoEncodingProperties2)
    Call ReleaseIfc(m_pIVideoEncodingProperties)
    Call ReleaseIfc(m_pIVideoEncodingPropertiesStatics2)
    Call ReleaseIfc(m_pIVideoEncodingPropertiesStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIVideoEncodingProperties
End Property

Public Property Let Ifc(ByVal pIVideoEncodingProperties As Long)
    Call ReleaseIfc(m_pIMediaEncodingProperties)
    Call ReleaseIfc(m_pIVideoEncodingProperties5)
    Call ReleaseIfc(m_pIVideoEncodingProperties4)
    Call ReleaseIfc(m_pIVideoEncodingProperties3)
    Call ReleaseIfc(m_pIVideoEncodingProperties2)
    Call ReleaseIfc(m_pIVideoEncodingProperties)
    m_pIVideoEncodingProperties = pIVideoEncodingProperties
    Call QueryIfc(m_pIVideoEncodingProperties, IID_IMediaEncodingProperties, m_pIMediaEncodingProperties)
    Call QueryIfc(m_pIVideoEncodingProperties, IID_IVideoEncodingProperties, m_pIVideoEncodingProperties)
    Call QueryIfc(m_pIVideoEncodingProperties, IID_IVideoEncodingProperties2, m_pIVideoEncodingProperties2)
    Call QueryIfc(m_pIVideoEncodingProperties, IID_IVideoEncodingProperties3, m_pIVideoEncodingProperties3)
    Call QueryIfc(m_pIVideoEncodingProperties, IID_IVideoEncodingProperties4, m_pIVideoEncodingProperties4)
    Call QueryIfc(m_pIVideoEncodingProperties, IID_IVideoEncodingProperties5, m_pIVideoEncodingProperties5)
End Property

' IMediaEncodingProperties
'    IMediaEncodingProperties_GetProperties = 6

Public Property Get MediaType() As String
    Dim Ret As String
    If m_pIMediaEncodingProperties <> 0& Then
        Dim hString As Long
        If Invoke(m_pIMediaEncodingProperties, _
                  IMediaEncodingProperties_GetType, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    MediaType = Ret
End Property

Public Property Let MediaSubtype(ByVal Value As String)
    If m_pIMediaEncodingProperties <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Value)
        If Invoke(m_pIMediaEncodingProperties, _
                  IMediaEncodingProperties_PutSubtype, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get MediaSubtype() As String
    Dim Ret As String
    If m_pIMediaEncodingProperties <> 0& Then
        Dim hString As Long
        If Invoke(m_pIMediaEncodingProperties, _
                  IMediaEncodingProperties_GetSubtype, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    MediaSubtype = Ret
End Property

' IVideoEncodingProperties
Public Property Let Bitrate(ByVal Value As Long)
    If m_pIVideoEncodingProperties <> 0& Then
        If Invoke(m_pIVideoEncodingProperties, _
                  IVideoEncodingProperties_PutBitrate, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get Bitrate() As Long
    Dim Ret As Long
    If m_pIVideoEncodingProperties <> 0& Then
        Dim Value As Long
        If Invoke(m_pIVideoEncodingProperties, _
                  IVideoEncodingProperties_GetBitrate, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Bitrate = Ret
End Property

Public Property Let Width(ByVal Value As Long)
    If m_pIVideoEncodingProperties <> 0& Then
        If Invoke(m_pIVideoEncodingProperties, _
                  IVideoEncodingProperties_PutWidth, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get Width() As Long
    Dim Ret As Long
    If m_pIVideoEncodingProperties <> 0& Then
        Dim Value As Long
        If Invoke(m_pIVideoEncodingProperties, _
                  IVideoEncodingProperties_GetWidth, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Width = Ret
End Property

Public Property Let Height(ByVal Value As Long)
    If m_pIVideoEncodingProperties <> 0& Then
        If Invoke(m_pIVideoEncodingProperties, _
                  IVideoEncodingProperties_PutHeight, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get Height() As Long
    Dim Ret As Long
    If m_pIVideoEncodingProperties <> 0& Then
        Dim Value As Long
        If Invoke(m_pIVideoEncodingProperties, _
                  IVideoEncodingProperties_GetHeight, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Height = Ret
End Property

'    IVideoEncodingProperties_GetFrameRate = 12
'    IVideoEncodingProperties_GetPixelAspectRatio = 13

' IVideoEncodingProperties2
Public Property Let ProfileId(ByVal Value As Long)
    If m_pIVideoEncodingProperties2 <> 0& Then
        If Invoke(m_pIVideoEncodingProperties2, _
                  IVideoEncodingProperties2_PutProfileId, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get ProfileId() As Long
    Dim Ret As Long
    If m_pIVideoEncodingProperties2 <> 0& Then
        Dim Value As Long
        If Invoke(m_pIVideoEncodingProperties2, _
                  IVideoEncodingProperties2_GetProfileId, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    ProfileId = Ret
End Property

' IVideoEncodingProperties3
Public Property Get StereoscopicVideoPackingMode() As StereoscopicVideoPackingMode
    Dim Ret As StereoscopicVideoPackingMode
    If m_pIVideoEncodingProperties3 <> 0& Then
        Dim Value As Long
        If Invoke(m_pIVideoEncodingProperties3, _
                  IVideoEncodingProperties3_GetStereoscopicVideoPackingMode, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    StereoscopicVideoPackingMode = Ret
End Property
    
' IVideoEncodingProperties4
Public Property Get SphericalVideoFrameFormat() As SphericalVideoFrameFormat
    Dim Ret As SphericalVideoFrameFormat
    If m_pIVideoEncodingProperties4 <> 0& Then
        Dim Value As Long
        If Invoke(m_pIVideoEncodingProperties4, _
                  IVideoEncodingProperties4_GetSphericalVideoFrameFormat, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    SphericalVideoFrameFormat = Ret
End Property

' ----==== Functions ====----
' IVideoEncodingPropertiesStatics
Public Function CreateH264() As VideoEncodingProperties
    Dim Ret As VideoEncodingProperties
    If m_pIVideoEncodingPropertiesStatics <> 0& Then
        Dim pIVideoEncodingProperties As Long
        If Invoke(m_pIVideoEncodingPropertiesStatics, _
                  IVideoEncodingPropertiesStatics_CreateH264, _
                  VarPtr(pIVideoEncodingProperties)) = S_OK Then
            If pIVideoEncodingProperties <> 0& Then
                Set Ret = New VideoEncodingProperties
                Ret.Ifc = pIVideoEncodingProperties
            End If
        End If
    End If
    Set CreateH264 = Ret
End Function

Public Function CreateMpeg2() As VideoEncodingProperties
    Dim Ret As VideoEncodingProperties
    If m_pIVideoEncodingPropertiesStatics <> 0& Then
        Dim pIVideoEncodingProperties As Long
        If Invoke(m_pIVideoEncodingPropertiesStatics, _
                  IVideoEncodingPropertiesStatics_CreateMpeg2, _
                  VarPtr(pIVideoEncodingProperties)) = S_OK Then
            If pIVideoEncodingProperties <> 0& Then
                Set Ret = New VideoEncodingProperties
                Ret.Ifc = pIVideoEncodingProperties
            End If
        End If
    End If
    Set CreateMpeg2 = Ret
End Function

Public Function CreateUncompressed(ByVal subtype As String, _
                                   ByVal Width As Long, _
                                   ByVal Height As Long) As VideoEncodingProperties
    Dim Ret As VideoEncodingProperties
    If m_pIVideoEncodingPropertiesStatics <> 0& Then
        Dim pIVideoEncodingProperties As Long
        Dim hString As Long
        hString = CreateWindowsString(subtype)
        If Invoke(m_pIVideoEncodingPropertiesStatics2, _
                  IVideoEncodingPropertiesStatics_CreateUncompressed, _
                  hString, _
                  Width, _
                  Height, _
                  VarPtr(pIVideoEncodingProperties)) = S_OK Then
            If pIVideoEncodingProperties <> 0& Then
                Set Ret = New VideoEncodingProperties
                Ret.Ifc = pIVideoEncodingProperties
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set CreateUncompressed = Ret
End Function

' IVideoEncodingPropertiesStatics2
Public Function CreateHevc() As VideoEncodingProperties
    Dim Ret As VideoEncodingProperties
    If m_pIVideoEncodingPropertiesStatics2 <> 0& Then
        Dim pIVideoEncodingProperties As Long
        If Invoke(m_pIVideoEncodingPropertiesStatics2, _
                  IVideoEncodingPropertiesStatics2_CreateHevc, _
                  VarPtr(pIVideoEncodingProperties)) = S_OK Then
            If pIVideoEncodingProperties <> 0& Then
                Set Ret = New VideoEncodingProperties
                Ret.Ifc = pIVideoEncodingProperties
            End If
        End If
    End If
    Set CreateHevc = Ret
End Function
    
' IVideoEncodingProperties2
'    IVideoEncodingProperties2_SetFormatUserData = 6
'    IVideoEncodingProperties2_GetFormatUserData = 7
    
' IVideoEncodingProperties5
Public Function Copy() As VideoEncodingProperties
    Dim Ret As VideoEncodingProperties
    If m_pIVideoEncodingProperties5 <> 0& Then
        Dim pIVideoEncodingProperties As Long
        If Invoke(m_pIVideoEncodingProperties5, _
                  IVideoEncodingProperties5_Copy, _
                  VarPtr(pIVideoEncodingProperties)) = S_OK Then
            If pIVideoEncodingProperties <> 0& Then
                Set Ret = New VideoEncodingProperties
                Ret.Ifc = pIVideoEncodingProperties
            End If
        End If
    End If
    Set Copy = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



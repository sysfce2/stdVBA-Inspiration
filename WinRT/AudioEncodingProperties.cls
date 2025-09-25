VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AudioEncodingProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 04/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_MediaProperties_AudioEncodingProperties As String = "Windows.Media.MediaProperties.AudioEncodingProperties"
Private Const IID_IMediaEncodingProperties As String = "{b4002af6-acd4-4e5a-a24b-5d7498a8b8c4}"
Private Const IID_IAudioEncodingProperties As String = "{62bc7a16-005c-4b3b-8a0b-0a090e9687f3}"
Private Const IID_IAudioEncodingProperties2 As String = "{c45d54da-80bd-4c23-80d5-72d4a181e894}"
Private Const IID_IAudioEncodingProperties3 As String = "{87600341-748c-4f8d-b0fd-10caf08ff087}"
Private Const IID_IAudioEncodingPropertiesStatics As String = "{0cad332c-ebe9-4527-b36d-e42a13cf38db}"
Private Const IID_IAudioEncodingPropertiesStatics2 As String = "{7489316f-77a0-433d-8ed5-4040280e8665}"
Private Const IID_IAudioEncodingPropertiesWithFormatUserData As String = "{98f10d79-13ea-49ff-be70-2673db69702c}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IAudioEncodingPropertiesStatics
    IAudioEncodingPropertiesStatics_CreateAac = 6
    IAudioEncodingPropertiesStatics_CreateAacAdts = 7
    IAudioEncodingPropertiesStatics_CreateMp3 = 8
    IAudioEncodingPropertiesStatics_CreatePcm = 9
    IAudioEncodingPropertiesStatics_CreateWma = 10
                        
    ' IAudioEncodingPropertiesStatics2
    IAudioEncodingPropertiesStatics2_CreateAlac = 6
    IAudioEncodingPropertiesStatics2_CreateFlac = 7
                        
    ' IMediaEncodingProperties
'    IMediaEncodingProperties_GetProperties = 6
    IMediaEncodingProperties_GetType = 7
    IMediaEncodingProperties_PutSubtype = 8
    IMediaEncodingProperties_GetSubtype = 9
    
    ' IAudioEncodingProperties
    IAudioEncodingProperties_PutBitrate = 6
    IAudioEncodingProperties_GetBitrate = 7
    IAudioEncodingProperties_PutChannelCount = 8
    IAudioEncodingProperties_GetChannelCount = 9
    IAudioEncodingProperties_PutSampleRate = 10
    IAudioEncodingProperties_GetSampleRate = 11
    IAudioEncodingProperties_PutBitsPerSample = 12
    IAudioEncodingProperties_GetBitsPerSample = 13
                        
    ' IAudioEncodingProperties2
    IAudioEncodingProperties2_GetIsSpatial = 6

    ' IAudioEncodingProperties3
    IAudioEncodingProperties3_Copy = 6

    ' IAudioEncodingPropertiesWithFormatUserData
    IAudioEncodingPropertiesWithFormatUserData_SetFormatUserData = 6
    IAudioEncodingPropertiesWithFormatUserData_GetFormatUserData = 7
                        
End Enum

' ----==== Variablen ====----
Private m_pIMediaEncodingProperties As Long
Private m_pIAudioEncodingProperties As Long
Private m_pIAudioEncodingProperties2 As Long
Private m_pIAudioEncodingProperties3 As Long
Private m_pIAudioEncodingPropertiesStatics As Long
Private m_pIAudioEncodingPropertiesStatics2 As Long
Private m_pIAudioEncodingPropertiesWithFormatUserData As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Media_MediaProperties_AudioEncodingProperties, _
                            IID_IAudioEncodingPropertiesStatics, _
                            m_pIAudioEncodingPropertiesStatics) Then
        Call QueryIfc(m_pIAudioEncodingPropertiesStatics, IID_IAudioEncodingPropertiesStatics2, m_pIAudioEncodingPropertiesStatics2)
    End If
    Dim pIAudioEncodingProperties As Long
    If GetActivateInstance(Windows_Media_MediaProperties_AudioEncodingProperties, _
                           IID_IAudioEncodingProperties, _
                           pIAudioEncodingProperties) Then
        Me.Ifc = pIAudioEncodingProperties
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIAudioEncodingPropertiesWithFormatUserData)
    Call ReleaseIfc(m_pIMediaEncodingProperties)
    Call ReleaseIfc(m_pIAudioEncodingProperties3)
    Call ReleaseIfc(m_pIAudioEncodingProperties2)
    Call ReleaseIfc(m_pIAudioEncodingProperties)
    Call ReleaseIfc(m_pIAudioEncodingPropertiesStatics2)
    Call ReleaseIfc(m_pIAudioEncodingPropertiesStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIAudioEncodingProperties
End Property

Public Property Let Ifc(ByVal pIAudioEncodingProperties As Long)
    Call ReleaseIfc(m_pIAudioEncodingPropertiesWithFormatUserData)
    Call ReleaseIfc(m_pIMediaEncodingProperties)
    Call ReleaseIfc(m_pIAudioEncodingProperties3)
    Call ReleaseIfc(m_pIAudioEncodingProperties2)
    Call ReleaseIfc(m_pIAudioEncodingProperties)
    m_pIAudioEncodingProperties = pIAudioEncodingProperties
    Call QueryIfc(m_pIAudioEncodingProperties, IID_IMediaEncodingProperties, m_pIMediaEncodingProperties)
    Call QueryIfc(m_pIAudioEncodingProperties, IID_IAudioEncodingProperties, m_pIAudioEncodingProperties)
    Call QueryIfc(m_pIAudioEncodingProperties, IID_IAudioEncodingProperties2, m_pIAudioEncodingProperties2)
    Call QueryIfc(m_pIAudioEncodingProperties, IID_IAudioEncodingProperties3, m_pIAudioEncodingProperties3)
    Call QueryIfc(m_pIAudioEncodingProperties, IID_IAudioEncodingPropertiesWithFormatUserData, m_pIAudioEncodingPropertiesWithFormatUserData)
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

Public Property Let MediaSubtype(ByVal value As String)
    If m_pIMediaEncodingProperties <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
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

' IAudioEncodingProperties
Public Property Let Bitrate(ByVal value As Long)
    If m_pIAudioEncodingProperties <> 0& Then
        If Invoke(m_pIAudioEncodingProperties, _
                  IAudioEncodingProperties_PutBitrate, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get Bitrate() As Long
    Dim Ret As Long
    If m_pIAudioEncodingProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIAudioEncodingProperties, _
                  IAudioEncodingProperties_GetBitrate, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Bitrate = Ret
End Property

Public Property Let ChannelCount(ByVal value As Long)
    If m_pIAudioEncodingProperties <> 0& Then
        If Invoke(m_pIAudioEncodingProperties, _
                  IAudioEncodingProperties_PutChannelCount, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get ChannelCount() As Long
    Dim Ret As Long
    If m_pIAudioEncodingProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIAudioEncodingProperties, _
                  IAudioEncodingProperties_GetChannelCount, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    ChannelCount = Ret
End Property

Public Property Let SampleRate(ByVal value As Long)
    If m_pIAudioEncodingProperties <> 0& Then
        If Invoke(m_pIAudioEncodingProperties, _
                  IAudioEncodingProperties_PutSampleRate, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get SampleRate() As Long
    Dim Ret As Long
    If m_pIAudioEncodingProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIAudioEncodingProperties, _
                  IAudioEncodingProperties_GetSampleRate, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    SampleRate = Ret
End Property

Public Property Let BitsPerSample(ByVal value As Long)
    If m_pIAudioEncodingProperties <> 0& Then
        If Invoke(m_pIAudioEncodingProperties, _
                  IAudioEncodingProperties_PutBitsPerSample, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get BitsPerSample() As Long
    Dim Ret As Long
    If m_pIAudioEncodingProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIAudioEncodingProperties, _
                  IAudioEncodingProperties_GetBitsPerSample, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    BitsPerSample = Ret
End Property
                        
' IAudioEncodingProperties2
Public Property Get IsSpatial() As Boolean
    Dim Ret As Boolean
    If m_pIAudioEncodingProperties2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIAudioEncodingProperties2, _
                  IAudioEncodingProperties2_GetIsSpatial, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsSpatial = Ret
End Property
 
' ----==== Functions ====----
' IAudioEncodingPropertiesStatics
Public Function CreateAac(ByVal SampleRate As Long, _
                          ByVal ChannelCount As Long, _
                          ByVal BitsPerSample As Long) As AudioEncodingProperties
    Dim Ret As AudioEncodingProperties
    If m_pIAudioEncodingPropertiesStatics <> 0& Then
        Dim pIAudioEncodingProperties As Long
        If Invoke(m_pIAudioEncodingPropertiesStatics, _
                  IAudioEncodingPropertiesStatics_CreateAac, _
                  SampleRate, _
                  ChannelCount, _
                  BitsPerSample, _
                  VarPtr(pIAudioEncodingProperties)) = S_OK Then
            If pIAudioEncodingProperties <> 0& Then
                Set Ret = New AudioEncodingProperties
                Ret.Ifc = pIAudioEncodingProperties
            End If
        End If
    End If
    Set CreateAac = Ret
End Function

Public Function CreateAacAdts(ByVal SampleRate As Long, _
                              ByVal ChannelCount As Long, _
                              ByVal BitsPerSample As Long) As AudioEncodingProperties
    Dim Ret As AudioEncodingProperties
    If m_pIAudioEncodingPropertiesStatics <> 0& Then
        Dim pIAudioEncodingProperties As Long
        If Invoke(m_pIAudioEncodingPropertiesStatics, _
                  IAudioEncodingPropertiesStatics_CreateAacAdts, _
                  SampleRate, _
                  ChannelCount, _
                  BitsPerSample, _
                  VarPtr(pIAudioEncodingProperties)) = S_OK Then
            If pIAudioEncodingProperties <> 0& Then
                Set Ret = New AudioEncodingProperties
                Ret.Ifc = pIAudioEncodingProperties
            End If
        End If
    End If
    Set CreateAacAdts = Ret
End Function

Public Function CreateMp3(ByVal SampleRate As Long, _
                          ByVal ChannelCount As Long, _
                          ByVal BitsPerSample As Long) As AudioEncodingProperties
    Dim Ret As AudioEncodingProperties
    If m_pIAudioEncodingPropertiesStatics <> 0& Then
        Dim pIAudioEncodingProperties As Long
        If Invoke(m_pIAudioEncodingPropertiesStatics, _
                  IAudioEncodingPropertiesStatics_CreateMp3, _
                  SampleRate, _
                  ChannelCount, _
                  BitsPerSample, _
                  VarPtr(pIAudioEncodingProperties)) = S_OK Then
            If pIAudioEncodingProperties <> 0& Then
                Set Ret = New AudioEncodingProperties
                Ret.Ifc = pIAudioEncodingProperties
            End If
        End If
    End If
    Set CreateMp3 = Ret
End Function

Public Function CreatePcm(ByVal SampleRate As Long, _
                          ByVal ChannelCount As Long, _
                          ByVal BitsPerSample As Long) As AudioEncodingProperties
    Dim Ret As AudioEncodingProperties
    If m_pIAudioEncodingPropertiesStatics <> 0& Then
        Dim pIAudioEncodingProperties As Long
        If Invoke(m_pIAudioEncodingPropertiesStatics, _
                  IAudioEncodingPropertiesStatics_CreatePcm, _
                  SampleRate, _
                  ChannelCount, _
                  BitsPerSample, _
                  VarPtr(pIAudioEncodingProperties)) = S_OK Then
            If pIAudioEncodingProperties <> 0& Then
                Set Ret = New AudioEncodingProperties
                Ret.Ifc = pIAudioEncodingProperties
            End If
        End If
    End If
    Set CreatePcm = Ret
End Function

Public Function CreateWma(ByVal SampleRate As Long, _
                          ByVal ChannelCount As Long, _
                          ByVal BitsPerSample As Long) As AudioEncodingProperties
    Dim Ret As AudioEncodingProperties
    If m_pIAudioEncodingPropertiesStatics <> 0& Then
        Dim pIAudioEncodingProperties As Long
        If Invoke(m_pIAudioEncodingPropertiesStatics, _
                  IAudioEncodingPropertiesStatics_CreateWma, _
                  SampleRate, _
                  ChannelCount, _
                  BitsPerSample, _
                  VarPtr(pIAudioEncodingProperties)) = S_OK Then
            If pIAudioEncodingProperties <> 0& Then
                Set Ret = New AudioEncodingProperties
                Ret.Ifc = pIAudioEncodingProperties
            End If
        End If
    End If
    Set CreateWma = Ret
End Function
                        
' IAudioEncodingPropertiesStatics2
Public Function CreateAlac(ByVal SampleRate As Long, _
                           ByVal ChannelCount As Long, _
                           ByVal BitsPerSample As Long) As AudioEncodingProperties
    Dim Ret As AudioEncodingProperties
    If m_pIAudioEncodingPropertiesStatics2 <> 0& Then
        Dim pIAudioEncodingProperties As Long
        If Invoke(m_pIAudioEncodingPropertiesStatics2, _
                  IAudioEncodingPropertiesStatics2_CreateAlac, _
                  SampleRate, _
                  ChannelCount, _
                  BitsPerSample, _
                  VarPtr(pIAudioEncodingProperties)) = S_OK Then
            If pIAudioEncodingProperties <> 0& Then
                Set Ret = New AudioEncodingProperties
                Ret.Ifc = pIAudioEncodingProperties
            End If
        End If
    End If
    Set CreateAlac = Ret
End Function

Public Function CreateFlac(ByVal SampleRate As Long, _
                           ByVal ChannelCount As Long, _
                           ByVal BitsPerSample As Long) As AudioEncodingProperties
    Dim Ret As AudioEncodingProperties
    If m_pIAudioEncodingPropertiesStatics2 <> 0& Then
        Dim pIAudioEncodingProperties As Long
        If Invoke(m_pIAudioEncodingPropertiesStatics2, _
                  IAudioEncodingPropertiesStatics2_CreateFlac, _
                  SampleRate, _
                  ChannelCount, _
                  BitsPerSample, _
                  VarPtr(pIAudioEncodingProperties)) = S_OK Then
            If pIAudioEncodingProperties <> 0& Then
                Set Ret = New AudioEncodingProperties
                Ret.Ifc = pIAudioEncodingProperties
            End If
        End If
    End If
    Set CreateFlac = Ret
End Function

' IAudioEncodingProperties3
Public Function Copy() As AudioEncodingProperties
    Dim Ret As AudioEncodingProperties
    If m_pIAudioEncodingProperties3 <> 0& Then
        Dim pIAudioEncodingProperties As Long
        If Invoke(m_pIAudioEncodingProperties3, _
                  IAudioEncodingProperties3_Copy, _
                  VarPtr(pIAudioEncodingProperties)) = S_OK Then
            If pIAudioEncodingProperties <> 0& Then
                Set Ret = New AudioEncodingProperties
                Ret.Ifc = pIAudioEncodingProperties
            End If
        End If
    End If
    Set Copy = Ret
End Function

' IAudioEncodingPropertiesWithFormatUserData
Public Function SetFormatUserData(ByVal valueLength As Long, _
                                  ByRef value() As Byte) As Boolean
    Dim Ret As Boolean
    If m_pIAudioEncodingPropertiesWithFormatUserData <> 0& Then
        If Invoke(m_pIAudioEncodingPropertiesWithFormatUserData, _
                  IAudioEncodingPropertiesWithFormatUserData_SetFormatUserData, _
                  valueLength, _
                  VarPtr(value(0))) = S_OK Then
            Ret = True
        End If
    End If
    SetFormatUserData = Ret
End Function

Public Function GetFormatUserData(ByRef valueLength As Long, _
                                  ByRef value() As Byte) As Boolean
    Dim Ret As Boolean
    If m_pIAudioEncodingPropertiesWithFormatUserData <> 0& Then
        If Invoke(m_pIAudioEncodingPropertiesWithFormatUserData, _
                  IAudioEncodingPropertiesWithFormatUserData_GetFormatUserData, _
                  VarPtr(valueLength), _
                  VarPtr(value(0))) = S_OK Then
            Ret = True
        End If
    End If
    GetFormatUserData = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



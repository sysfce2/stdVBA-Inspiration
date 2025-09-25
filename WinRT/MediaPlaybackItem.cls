VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaPlaybackItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Playback_MediaPlaybackItem As String = "Windows.Media.Playback.MediaPlaybackItem"
Private Const IID_IMediaPlaybackItemFactory As String = "{7133fce1-1769-4ff9-a7c1-38d2c4d42360}"
Private Const IID_IMediaPlaybackItemFactory2 As String = "{d77cdf3a-b947-4972-b35d-adfb931a71e6}"
Private Const IID_IMediaPlaybackItemStatics As String = "{4b1be7f4-4345-403c-8a67-f5de91df4c86}"
Private Const IID_IMediaPlaybackItem As String = "{047097d2-e4af-48ab-b283-6929e674ece2}"
Private Const IID_IMediaPlaybackItem2 As String = "{d859d171-d7ef-4b81-ac1f-f40493cbb091}"
Private Const IID_IMediaPlaybackItem3 As String = "{0d328220-b80a-4d09-9ff8-f87094a1c831}"
Private Const IID_IMediaPlaybackSource As String = "{ef9dc2bc-9317-4696-b051-2bad643177b5}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    'IMediaPlaybackItemFactory
    IMediaPlaybackItemFactory_Create = 6
    
    'IMediaPlaybackItemFactory2
    IMediaPlaybackItemFactory2_CreateWithStartTime = 6
    IMediaPlaybackItemFactory2_CreateWithStartTimeAndDurationLimit = 7
    
    'IMediaPlaybackItemStatics
    IMediaPlaybackItemStatics_FindFromMediaSource = 6

    'IMediaPlaybackItem
    IMediaPlaybackItem_AddAudioTracksChanged = 6
    IMediaPlaybackItem_RemoveAudioTracksChanged = 7
    IMediaPlaybackItem_AddVideoTracksChanged = 8
    IMediaPlaybackItem_RemoveVideoTracksChanged = 9
    IMediaPlaybackItem_AddTimedMetadataTracksChanged = 10
    IMediaPlaybackItem_RemoveTimedMetadataTracksChanged = 11
    IMediaPlaybackItem_GetSource = 12
    IMediaPlaybackItem_GetAudioTracks = 13
    IMediaPlaybackItem_GetVideoTracks = 14
'    IMediaPlaybackItem_GetTimedMetadataTracks = 15
    
    'IMediaPlaybackItem2
    IMediaPlaybackItem2_GetBreakSchedule = 6
    IMediaPlaybackItem2_GetStartTime = 7
    IMediaPlaybackItem2_GetDurationLimit = 8
    IMediaPlaybackItem2_GetCanSkip = 9
    IMediaPlaybackItem2_PutCanSkip = 10
    IMediaPlaybackItem2_GetDisplayProperties = 11
    IMediaPlaybackItem2_ApplyDisplayProperties = 12
    
    'IMediaPlaybackItem3
    IMediaPlaybackItem3_GetIsDisabledInPlaybackList = 6
    IMediaPlaybackItem3_PutIsDisabledInPlaybackList = 7
    IMediaPlaybackItem3_GetTotalDownloadProgress = 8
    IMediaPlaybackItem3_GetAutoLoadedDisplayProperties = 9
    IMediaPlaybackItem3_PutAutoLoadedDisplayProperties = 10
                        
    'IMediaPlaybackSource
    ' has no member
    
    ' IReference
    IReference_GetValue = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIMediaPlaybackItemFactory As Long
Private m_pIMediaPlaybackItemFactory2 As Long
Private m_pIMediaPlaybackItemStatics As Long
Private m_pIMediaPlaybackItem As Long
Private m_pIMediaPlaybackItem2 As Long
Private m_pIMediaPlaybackItem3 As Long
Private m_pIMediaPlaybackSource As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Media_Playback_MediaPlaybackItem, _
                            IID_IMediaPlaybackItemFactory, _
                            m_pIMediaPlaybackItemFactory) Then
        Call QueryIfc(m_pIMediaPlaybackItemFactory, IID_IMediaPlaybackItemFactory2, m_pIMediaPlaybackItemFactory2)
    End If
    If GetActivationFactory(Windows_Media_Playback_MediaPlaybackItem, _
                            IID_IMediaPlaybackItemStatics, _
                            m_pIMediaPlaybackItemStatics) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMediaPlaybackSource)
    Call ReleaseIfc(m_pIMediaPlaybackItem3)
    Call ReleaseIfc(m_pIMediaPlaybackItem2)
    Call ReleaseIfc(m_pIMediaPlaybackItem)
    Call ReleaseIfc(m_pIMediaPlaybackItemStatics)
    Call ReleaseIfc(m_pIMediaPlaybackItemFactory2)
    Call ReleaseIfc(m_pIMediaPlaybackItemFactory)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMediaPlaybackItem
End Property

Public Property Let Ifc(ByVal pIMediaPlaybackItem As Long)
    Call ReleaseIfc(m_pIMediaPlaybackSource)
    Call ReleaseIfc(m_pIMediaPlaybackItem3)
    Call ReleaseIfc(m_pIMediaPlaybackItem2)
    Call ReleaseIfc(m_pIMediaPlaybackItem)
    m_pIMediaPlaybackItem = pIMediaPlaybackItem
    Call QueryIfc(m_pIMediaPlaybackItem, IID_IMediaPlaybackSource, m_pIMediaPlaybackSource)
    Call QueryIfc(m_pIMediaPlaybackItem, IID_IMediaPlaybackItem2, m_pIMediaPlaybackItem2)
    Call QueryIfc(m_pIMediaPlaybackItem, IID_IMediaPlaybackItem3, m_pIMediaPlaybackItem3)
End Property

'IMediaPlaybackItem
Public Property Get Source() As MediaSource
    Dim Ret As MediaSource
    If m_pIMediaPlaybackItem <> 0& Then
        Dim pIMediaSource As Long
        If Invoke(m_pIMediaPlaybackItem, _
                  IMediaPlaybackItem_GetSource, _
                  VarPtr(pIMediaSource)) = S_OK Then
            If pIMediaSource <> 0& Then
                Set Ret = New MediaSource
                Ret.Ifc = pIMediaSource
            End If
        End If
    End If
    Set Source = Ret
End Property

Public Property Get AudioTracks() As ReadOnlyList_1 'ReadOnlyList_AudioTrack
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_AudioTrack
    If m_pIMediaPlaybackItem <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIMediaPlaybackItem, _
                  IMediaPlaybackItem_GetAudioTracks, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_AudioTrack, pIVectorView) 'ReadOnlyList_AudioTrack
            End If
        End If
    End If
    Set AudioTracks = Ret
End Property

Public Property Get VideoTracks() As ReadOnlyList_1 'ReadOnlyList_VideoTrack
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_VideoTrack
    If m_pIMediaPlaybackItem <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIMediaPlaybackItem, _
                  IMediaPlaybackItem_GetVideoTracks, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_VideoTrack, pIVectorView) 'ReadOnlyList_VideoTrack
            End If
        End If
    End If
    Set VideoTracks = Ret
End Property

'    IMediaPlaybackItem_GetTimedMetadataTracks
    
'IMediaPlaybackItem2
Public Property Get BreakSchedule() As MediaBreakSchedule
    Dim Ret As MediaBreakSchedule
    If m_pIMediaPlaybackItem2 <> 0& Then
        Dim pIMediaBreakSchedule As Long
        If Invoke(m_pIMediaPlaybackItem2, _
                  IMediaPlaybackItem2_GetBreakSchedule, _
                  VarPtr(pIMediaBreakSchedule)) = S_OK Then
            If pIMediaBreakSchedule <> 0& Then
                Set Ret = New MediaBreakSchedule
                Ret.Ifc = pIMediaBreakSchedule
            End If
        End If
    End If
    Set BreakSchedule = Ret
End Property

Public Property Get StartTime() As TimeSpan
    Dim Ret As TimeSpan
    If m_pIMediaPlaybackItem2 <> 0& Then
        Dim value As Currency
        If Invoke(m_pIMediaPlaybackItem2, _
                  IMediaPlaybackItem2_GetStartTime, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New TimeSpan
            Ret.TimeSpanValue = value
        End If
    End If
    Set StartTime = Ret
End Property

Public Property Get DurationLimit() As TimeSpan
    Dim Ret As TimeSpan
    If m_pIMediaPlaybackItem2 <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIMediaPlaybackItem2, _
                  IMediaPlaybackItem2_GetDurationLimit, _
                  VarPtr(pIReference)) = S_OK Then
            If pIReference <> 0& Then
                Dim value As Currency
                If Invoke(pIReference, _
                          IReference_GetValue, _
                          VarPtr(value)) = S_OK Then
                    Set Ret = New TimeSpan
                    Ret.TimeSpanValue = value
                End If
                Call ReleaseIfc(pIReference)
            End If
        End If
    End If
    Set DurationLimit = Ret
End Property

Public Property Get CanSkip() As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackItem2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlaybackItem2, _
                  IMediaPlaybackItem2_GetCanSkip, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    CanSkip = Ret
End Property
    
Public Property Let CanSkip(ByVal value As Boolean)
    If m_pIMediaPlaybackItem2 <> 0& Then
        If Invoke(m_pIMediaPlaybackItem2, _
                  IMediaPlaybackItem2_PutCanSkip, _
                  value) = S_OK Then
        End If
    End If
End Property
    
'IMediaPlaybackItem3
Public Property Get IsDisabledInPlaybackList() As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackItem3 <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlaybackItem3, _
                  IMediaPlaybackItem3_GetIsDisabledInPlaybackList, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsDisabledInPlaybackList = Ret
End Property
    
Public Property Let IsDisabledInPlaybackList(ByVal value As Boolean)
    If m_pIMediaPlaybackItem3 <> 0& Then
        If Invoke(m_pIMediaPlaybackItem3, _
                  IMediaPlaybackItem3_PutIsDisabledInPlaybackList, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get TotalDownloadProgress() As Double
    Dim Ret As Double
    If m_pIMediaPlaybackItem3 <> 0& Then
        Dim value As Double
        If Invoke(m_pIMediaPlaybackItem3, _
                  IMediaPlaybackItem3_GetTotalDownloadProgress, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    TotalDownloadProgress = Ret
End Property

Public Property Get AutoLoadedDisplayProperties() As AutoLoadedDisplayPropertyKind
    Dim Ret As AutoLoadedDisplayPropertyKind
    If m_pIMediaPlaybackItem3 <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlaybackItem3, _
                  IMediaPlaybackItem3_GetAutoLoadedDisplayProperties, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    AutoLoadedDisplayProperties = Ret
End Property
    
Public Property Let AutoLoadedDisplayProperties(ByVal value As AutoLoadedDisplayPropertyKind)
    If m_pIMediaPlaybackItem3 <> 0& Then
        If Invoke(m_pIMediaPlaybackItem3, _
                  IMediaPlaybackItem3_PutAutoLoadedDisplayProperties, _
                  value) = S_OK Then
        End If
    End If
End Property

' ----==== Functions ====----
'IMediaPlaybackItemFactory
Public Function Create(ByVal Source As MediaSource) As MediaPlaybackItem
    Dim Ret As MediaPlaybackItem
    If m_pIMediaPlaybackItemFactory <> 0& And _
       IsNotNothing(Source) Then
        Dim pIMediaPlaybackItem As Long
        If Invoke(m_pIMediaPlaybackItemFactory, _
                  IMediaPlaybackItemFactory_Create, _
                  Source.Ifc, _
                  VarPtr(pIMediaPlaybackItem)) = S_OK Then
            If pIMediaPlaybackItem <> 0& Then
                Set Ret = New MediaPlaybackItem
                Ret.Ifc = pIMediaPlaybackItem
            End If
        End If
    End If
    Set Create = Ret
End Function
    
'IMediaPlaybackItemFactory2
Public Function CreateWithStartTime(ByVal Source As MediaSource, _
                                    ByVal StartTime As TimeSpan) As MediaPlaybackItem
    Dim Ret As MediaPlaybackItem
    If m_pIMediaPlaybackItemFactory2 <> 0& And _
       IsNotNothing(Source) And _
       IsNotNothing(StartTime) Then
        Dim pIMediaPlaybackItem As Long
        If Invoke(m_pIMediaPlaybackItemFactory2, _
                  IMediaPlaybackItemFactory2_CreateWithStartTime, _
                  Source.Ifc, _
                  StartTime.TimeSpanValue, _
                  VarPtr(pIMediaPlaybackItem)) = S_OK Then
            If pIMediaPlaybackItem <> 0& Then
                Set Ret = New MediaPlaybackItem
                Ret.Ifc = pIMediaPlaybackItem
            End If
        End If
    End If
    Set CreateWithStartTime = Ret
End Function

Public Function CreateWithStartTimeAndDurationLimit(ByVal Source As MediaSource, _
                                                    ByVal StartTime As TimeSpan, _
                                                    ByVal DurationLimit As TimeSpan) As MediaPlaybackItem
    Dim Ret As MediaPlaybackItem
    If m_pIMediaPlaybackItemFactory2 <> 0& And _
       IsNotNothing(Source) And _
       IsNotNothing(StartTime) And _
       IsNotNothing(DurationLimit) Then
        Dim pIMediaPlaybackItem As Long
        If Invoke(m_pIMediaPlaybackItemFactory2, _
                  IMediaPlaybackItemFactory2_CreateWithStartTimeAndDurationLimit, _
                  Source.Ifc, _
                  StartTime.TimeSpanValue, _
                  DurationLimit.TimeSpanValue, _
                  VarPtr(pIMediaPlaybackItem)) = S_OK Then
            If pIMediaPlaybackItem <> 0& Then
                Set Ret = New MediaPlaybackItem
                Ret.Ifc = pIMediaPlaybackItem
            End If
        End If
    End If
    Set CreateWithStartTimeAndDurationLimit = Ret
End Function
    
'IMediaPlaybackItemStatics
Public Function FindFromMediaSource(ByVal Source As MediaSource) As MediaPlaybackItem
    Dim Ret As MediaPlaybackItem
    If m_pIMediaPlaybackItemStatics <> 0& And _
       IsNotNothing(Source) Then
        Dim pIMediaPlaybackItem As Long
        If Invoke(m_pIMediaPlaybackItemStatics, _
                  IMediaPlaybackItemStatics_FindFromMediaSource, _
                  Source.Ifc, _
                  VarPtr(pIMediaPlaybackItem)) = S_OK Then
            If pIMediaPlaybackItem <> 0& Then
                Set Ret = New MediaPlaybackItem
                Ret.Ifc = pIMediaPlaybackItem
            End If
        End If
    End If
    Set FindFromMediaSource = Ret
End Function

'IMediaPlaybackItem
Public Function AddAudioTracksChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackItem <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackItem, _
                  IMediaPlaybackItem_AddAudioTracksChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddAudioTracksChanged = Ret
End Function

Public Function RemoveAudioTracksChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackItem <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackItem, _
                  IMediaPlaybackItem_RemoveAudioTracksChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveAudioTracksChanged = Ret
End Function

Public Function AddVideoTracksChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackItem <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackItem, _
                  IMediaPlaybackItem_AddVideoTracksChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddVideoTracksChanged = Ret
End Function

Public Function RemoveVideoTracksChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackItem <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackItem, _
                  IMediaPlaybackItem_RemoveVideoTracksChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveVideoTracksChanged = Ret
End Function

Public Function AddTimedMetadataTracksChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackItem <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackItem, _
                  IMediaPlaybackItem_AddTimedMetadataTracksChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddTimedMetadataTracksChanged = Ret
End Function

Public Function RemoveTimedMetadataTracksChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackItem <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackItem, _
                  IMediaPlaybackItem_RemoveTimedMetadataTracksChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveTimedMetadataTracksChanged = Ret
End Function
    
'IMediaPlaybackItem2
Public Function GetDisplayProperties() As MediaItemDisplayProperties
    Dim Ret As MediaItemDisplayProperties
    If m_pIMediaPlaybackItem2 <> 0& Then
        Dim pIMediaItemDisplayProperties As Long
        If Invoke(m_pIMediaPlaybackItem2, _
                  IMediaPlaybackItem2_GetDisplayProperties, _
                  VarPtr(pIMediaItemDisplayProperties)) = S_OK Then
            If pIMediaItemDisplayProperties <> 0& Then
                Set Ret = New MediaItemDisplayProperties
                Ret.Ifc = pIMediaItemDisplayProperties
            End If
        End If
    End If
    Set GetDisplayProperties = Ret
End Function

Public Function ApplyDisplayProperties(ByVal value As MediaItemDisplayProperties) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackItem2 <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pIMediaPlaybackItem2, _
                  IMediaPlaybackItem2_ApplyDisplayProperties, _
                  value.Ifc) = S_OK Then
            Ret = True
        End If
    End If
    ApplyDisplayProperties = Ret
End Function
    
' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaBreakSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 11/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Playback_MediaBreakSchedule As String = "Windows.Media.Playback.MediaBreakSchedule"
Private Const IID_IMediaBreakSchedule As String = "{a19a5813-98b6-41d8-83da-f971d22b7bba}"
 
' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IMediaBreakSchedule
    IMediaBreakSchedule_AddScheduleChanged = 6
    IMediaBreakSchedule_RemoveScheduleChanged = 7
    IMediaBreakSchedule_InsertMidrollBreak = 8
    IMediaBreakSchedule_RemoveMidrollBreak = 9
    IMediaBreakSchedule_GetMidrollBreaks = 10
    IMediaBreakSchedule_PutPrerollBreak = 11
    IMediaBreakSchedule_GetPrerollBreak = 12
    IMediaBreakSchedule_PutPostrollBreak = 13
    IMediaBreakSchedule_GetPostrollBreak = 14
    IMediaBreakSchedule_GetPlaybackItem = 15
                        
End Enum

' ----==== Variablen ====----
Private m_pIMediaBreakSchedule As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMediaBreakSchedule)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMediaBreakSchedule
End Property

Public Property Let Ifc(ByVal pIMediaBreakSchedule As Long)
    Call ReleaseIfc(m_pIMediaBreakSchedule)
    m_pIMediaBreakSchedule = pIMediaBreakSchedule
End Property
          
' IMediaBreakSchedule
Public Property Get MidrollBreaks() As ReadOnlyList_1 'ReadOnlyList_MediaBreak
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_MediaBreak
    If m_pIMediaBreakSchedule <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIMediaBreakSchedule, _
                  IMediaBreakSchedule_GetMidrollBreaks, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_MediaBreak, pIVectorView) 'ReadOnlyList_MediaBreak
            End If
        End If
    End If
    Set MidrollBreaks = Ret
End Property

Public Property Let PrerollBreak(ByVal value As MediaBreak)
    If m_pIMediaBreakSchedule <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pIMediaBreakSchedule, _
                  IMediaBreakSchedule_PutPrerollBreak, _
                  value.Ifc) = S_OK Then
        End If
    End If
End Property

Public Property Get PrerollBreak() As MediaBreak
    Dim Ret As MediaBreak
    If m_pIMediaBreakSchedule <> 0& Then
        Dim pIMediaBreak As Long
        If Invoke(m_pIMediaBreakSchedule, _
                  IMediaBreakSchedule_GetPrerollBreak, _
                  VarPtr(pIMediaBreak)) = S_OK Then
            If pIMediaBreak <> 0& Then
                Set Ret = New MediaBreak
                Ret.Ifc = pIMediaBreak
            End If
        End If
    End If
    Set PrerollBreak = Ret
End Property

Public Property Let PostrollBreak(ByVal value As MediaBreak)
    If m_pIMediaBreakSchedule <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pIMediaBreakSchedule, _
                  IMediaBreakSchedule_PutPostrollBreak, _
                  value.Ifc) = S_OK Then
        End If
    End If
End Property

Public Property Get PostrollBreak() As MediaBreak
    Dim Ret As MediaBreak
    If m_pIMediaBreakSchedule <> 0& Then
        Dim pIMediaBreak As Long
        If Invoke(m_pIMediaBreakSchedule, _
                  IMediaBreakSchedule_GetPostrollBreak, _
                  VarPtr(pIMediaBreak)) = S_OK Then
            If pIMediaBreak <> 0& Then
                Set Ret = New MediaBreak
                Ret.Ifc = pIMediaBreak
            End If
        End If
    End If
    Set PostrollBreak = Ret
End Property

Public Property Get PlaybackItem() As MediaPlaybackItem
    Dim Ret As MediaPlaybackItem
    If m_pIMediaBreakSchedule <> 0& Then
        Dim pIMediaPlaybackItem As Long
        If Invoke(m_pIMediaBreakSchedule, _
                  IMediaBreakSchedule_GetPlaybackItem, _
                  VarPtr(pIMediaPlaybackItem)) = S_OK Then
            If pIMediaPlaybackItem <> 0& Then
                Set Ret = New MediaPlaybackItem
                Ret.Ifc = pIMediaPlaybackItem
            End If
        End If
    End If
    Set PlaybackItem = Ret
End Property
          
' ----==== Functions ====----
' IMediaBreakSchedule
Public Function AddScheduleChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaBreakSchedule <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaBreakSchedule, _
                  IMediaBreakSchedule_AddScheduleChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddScheduleChanged = Ret
End Function

Public Function RemoveScheduleChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaBreakSchedule <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaBreakSchedule, _
                  IMediaBreakSchedule_RemoveScheduleChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveScheduleChanged = Ret
End Function

Public Function InsertMidrollBreak(ByVal InsertMediaBreak As MediaBreak) As Boolean
    Dim Ret As Boolean
    If m_pIMediaBreakSchedule <> 0& And _
       IsNotNothing(InsertMediaBreak) Then
        Dim value As Boolean
        If Invoke(m_pIMediaBreakSchedule, _
                  IMediaBreakSchedule_InsertMidrollBreak, _
                  InsertMediaBreak.Ifc, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    InsertMidrollBreak = Ret
End Function

Public Function RemoveMidrollBreak(ByVal RemoveMediaBreak As MediaBreak) As Boolean
    Dim Ret As Boolean
    If m_pIMediaBreakSchedule <> 0& And _
       IsNotNothing(RemoveMediaBreak) Then
        Dim value As Boolean
        If Invoke(m_pIMediaBreakSchedule, _
                  IMediaBreakSchedule_RemoveMidrollBreak, _
                  RemoveMediaBreak.Ifc, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    RemoveMidrollBreak = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



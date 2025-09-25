VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaBreakManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 11/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Playback_MediaBreakManager As String = "Windows.Media.Playback.MediaBreakManager"
Private Const IID_IMediaBreakManager As String = "{a854ddb1-feb4-4d9b-9d97-0fdbe58e5e39}"
 
' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IMediaBreakManager
    IMediaBreakManager_AddBreaksSeekedOver = 6
    IMediaBreakManager_RemoveBreaksSeekedOver = 7
    IMediaBreakManager_AddBreakStarted = 8
    IMediaBreakManager_RemoveBreakStarted = 9
    IMediaBreakManager_AddBreakEnded = 10
    IMediaBreakManager_RemoveBreakEnded = 11
    IMediaBreakManager_AddBreakSkipped = 12
    IMediaBreakManager_RemoveBreakSkipped = 13
    IMediaBreakManager_GetCurrentBreak = 14
    IMediaBreakManager_GetPlaybackSession = 15
    IMediaBreakManager_PlayBreak = 16
    IMediaBreakManager_SkipCurrentBreak = 17
                    
End Enum

' ----==== Variablen ====----
Private m_pIMediaBreakManager As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMediaBreakManager)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMediaBreakManager
End Property

Public Property Let Ifc(ByVal pIMediaBreakManager As Long)
    Call ReleaseIfc(m_pIMediaBreakManager)
    m_pIMediaBreakManager = pIMediaBreakManager
End Property
          
' IMediaBreakManager
Public Property Get CurrentBreak() As MediaBreak
    Dim Ret As MediaBreak
    If m_pIMediaBreakManager <> 0& Then
        Dim pIMediaBreak As Long
        If Invoke(m_pIMediaBreakManager, _
                  IMediaBreakManager_GetCurrentBreak, _
                  VarPtr(pIMediaBreak)) = S_OK Then
            If pIMediaBreak <> 0& Then
                Set Ret = New MediaBreak
                Ret.Ifc = pIMediaBreak
            End If
        End If
    End If
    Set CurrentBreak = Ret
End Property

Public Property Get PlaybackSession() As MediaPlaybackSession
    Dim Ret As MediaPlaybackSession
    If m_pIMediaBreakManager <> 0& Then
        Dim pIMediaPlaybackSession As Long
        If Invoke(m_pIMediaBreakManager, _
                  IMediaBreakManager_GetPlaybackSession, _
                  VarPtr(pIMediaPlaybackSession)) = S_OK Then
            If pIMediaPlaybackSession <> 0& Then
                Set Ret = New MediaPlaybackSession
                Ret.Ifc = pIMediaPlaybackSession
            End If
        End If
    End If
    Set PlaybackSession = Ret
End Property
          
' ----==== Functions ====----
' IMediaBreakManager
Public Function AddBreaksSeekedOver(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaBreakManager <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaBreakManager, _
                  IMediaBreakManager_AddBreaksSeekedOver, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddBreaksSeekedOver = Ret
End Function

Public Function RemoveBreaksSeekedOver(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaBreakManager <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaBreakManager, _
                  IMediaBreakManager_RemoveBreaksSeekedOver, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveBreaksSeekedOver = Ret
End Function

Public Function AddBreakStarted(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaBreakManager <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaBreakManager, _
                  IMediaBreakManager_AddBreakStarted, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddBreakStarted = Ret
End Function

Public Function RemoveBreakStarted(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaBreakManager <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaBreakManager, _
                  IMediaBreakManager_RemoveBreakStarted, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveBreakStarted = Ret
End Function

Public Function AddBreakEnded(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaBreakManager <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaBreakManager, _
                  IMediaBreakManager_AddBreakEnded, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddBreakEnded = Ret
End Function

Public Function RemoveBreakEnded(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaBreakManager <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaBreakManager, _
                  IMediaBreakManager_RemoveBreakEnded, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveBreakEnded = Ret
End Function

Public Function AddBreakSkipped(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaBreakManager <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaBreakManager, _
                  IMediaBreakManager_AddBreakSkipped, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddBreakSkipped = Ret
End Function

Public Function RemoveBreakSkipped(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaBreakManager <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaBreakManager, _
                  IMediaBreakManager_RemoveBreakSkipped, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveBreakSkipped = Ret
End Function

Public Function PlayBreak(ByVal value As MediaBreak) As Boolean
    Dim Ret As Boolean
    If m_pIMediaBreakManager <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pIMediaBreakManager, _
                  IMediaBreakManager_PlayBreak, _
                  value.Ifc) = S_OK Then
            Ret = True
        End If
    End If
    PlayBreak = Ret
End Function

Public Function SkipCurrentBreak() As Boolean
    Dim Ret As Boolean
    If m_pIMediaBreakManager <> 0& Then
        If Invoke(m_pIMediaBreakManager, _
                  IMediaBreakManager_SkipCurrentBreak) = S_OK Then
            Ret = True
        End If
    End If
    SkipCurrentBreak = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GlobalSystemMediaTransportControlsSessionMediaProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Control_GlobalSystemMediaTransportControlsSessionMediaProperties As String = "Windows.Media.Control.GlobalSystemMediaTransportControlsSessionMediaProperties"
Private Const IID_IGlobalSystemMediaTransportControlsSessionMediaProperties As String = "{68856cf6-adb4-54b2-ac16-05837907acb6}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
                        
    ' IGlobalSystemMediaTransportControlsSessionMediaProperties
    IGlobalSystemMediaTransportControlsSessionMediaProperties_GetTitle = 6
    IGlobalSystemMediaTransportControlsSessionMediaProperties_GetSubtitle = 7
    IGlobalSystemMediaTransportControlsSessionMediaProperties_GetAlbumArtist = 8
    IGlobalSystemMediaTransportControlsSessionMediaProperties_GetArtist = 9
    IGlobalSystemMediaTransportControlsSessionMediaProperties_GetAlbumTitle = 10
    IGlobalSystemMediaTransportControlsSessionMediaProperties_GetTrackNumber = 11
    IGlobalSystemMediaTransportControlsSessionMediaProperties_GetGenres = 12
    IGlobalSystemMediaTransportControlsSessionMediaProperties_GetAlbumTrackCount = 13
    IGlobalSystemMediaTransportControlsSessionMediaProperties_GetPlaybackType = 14
    IGlobalSystemMediaTransportControlsSessionMediaProperties_GetThumbnail = 15
                        
    ' IReference_xxx
    IReference_GetValue = 6
                        
End Enum

' ----==== Variablen ====----
Private m_pIGlobalSystemMediaTransportControlsSessionMediaProperties As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIGlobalSystemMediaTransportControlsSessionMediaProperties)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIGlobalSystemMediaTransportControlsSessionMediaProperties
End Property

Public Property Let Ifc(ByVal pIGlobalSystemMediaTransportControlsSessionMediaProperties As Long)
    Call ReleaseIfc(m_pIGlobalSystemMediaTransportControlsSessionMediaProperties)
    m_pIGlobalSystemMediaTransportControlsSessionMediaProperties = pIGlobalSystemMediaTransportControlsSessionMediaProperties
End Property

' IGlobalSystemMediaTransportControlsSessionMediaProperties
Public Property Get Title() As String
    Dim Ret As String
    If m_pIGlobalSystemMediaTransportControlsSessionMediaProperties <> 0& Then
        Dim hString As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionMediaProperties, _
                  IGlobalSystemMediaTransportControlsSessionMediaProperties_GetTitle, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    Title = Ret
End Property

Public Property Get Subtitle() As String
    Dim Ret As String
    If m_pIGlobalSystemMediaTransportControlsSessionMediaProperties <> 0& Then
        Dim hString As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionMediaProperties, _
                  IGlobalSystemMediaTransportControlsSessionMediaProperties_GetSubtitle, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    Subtitle = Ret
End Property

Public Property Get AlbumArtist() As String
    Dim Ret As String
    If m_pIGlobalSystemMediaTransportControlsSessionMediaProperties <> 0& Then
        Dim hString As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionMediaProperties, _
                  IGlobalSystemMediaTransportControlsSessionMediaProperties_GetAlbumArtist, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    AlbumArtist = Ret
End Property

Public Property Get Artist() As String
    Dim Ret As String
    If m_pIGlobalSystemMediaTransportControlsSessionMediaProperties <> 0& Then
        Dim hString As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionMediaProperties, _
                  IGlobalSystemMediaTransportControlsSessionMediaProperties_GetArtist, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    Artist = Ret
End Property

Public Property Get AlbumTitle() As String
    Dim Ret As String
    If m_pIGlobalSystemMediaTransportControlsSessionMediaProperties <> 0& Then
        Dim hString As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionMediaProperties, _
                  IGlobalSystemMediaTransportControlsSessionMediaProperties_GetAlbumTitle, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    AlbumTitle = Ret
End Property

Public Property Get TrackNumber() As Long
    Dim Ret As Long
    If m_pIGlobalSystemMediaTransportControlsSessionMediaProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionMediaProperties, _
                  IGlobalSystemMediaTransportControlsSessionMediaProperties_GetTrackNumber, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    TrackNumber = Ret
End Property

Public Property Get Genres() As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pIGlobalSystemMediaTransportControlsSessionMediaProperties <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionMediaProperties, _
                  IGlobalSystemMediaTransportControlsSessionMediaProperties_GetGenres, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_String, pIVectorView) 'ReadOnlyList_String
            End If
        End If
    End If
    Set Genres = Ret
End Property

Public Property Get AlbumTrackCount() As Long
    Dim Ret As Long
    If m_pIGlobalSystemMediaTransportControlsSessionMediaProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionMediaProperties, _
                  IGlobalSystemMediaTransportControlsSessionMediaProperties_GetAlbumTrackCount, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    AlbumTrackCount = Ret
End Property

Public Property Get PlaybackType() As MediaPlaybackType
    Dim Ret As MediaPlaybackType
    If m_pIGlobalSystemMediaTransportControlsSessionMediaProperties <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionMediaProperties, _
                  IGlobalSystemMediaTransportControlsSessionMediaProperties_GetPlaybackType, _
                  VarPtr(pIReference)) = S_OK Then
            If pIReference <> 0& Then
                Dim value As Long
                If Invoke(pIReference, _
                          IReference_GetValue, _
                          VarPtr(value)) = S_OK Then
                    Ret = value
                End If
                Call ReleaseIfc(pIReference)
            End If
        End If
    End If
    PlaybackType = Ret
End Property

Public Property Get Thumbnail() As RandomAccessStreamReference
    Dim Ret As RandomAccessStreamReference
    If m_pIGlobalSystemMediaTransportControlsSessionMediaProperties <> 0& Then
        Dim pIRandomAccessStreamReference As Long
        If Invoke(m_pIGlobalSystemMediaTransportControlsSessionMediaProperties, _
                  IGlobalSystemMediaTransportControlsSessionMediaProperties_GetThumbnail, _
                  VarPtr(pIRandomAccessStreamReference)) = S_OK Then
            If pIRandomAccessStreamReference <> 0& Then
                Set Ret = New RandomAccessStreamReference
                Ret.Ifc = pIRandomAccessStreamReference
            End If
        End If
    End If
    Set Thumbnail = Ret
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

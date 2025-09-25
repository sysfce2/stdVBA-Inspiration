VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MusicDisplayProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_MusicDisplayProperties As String = "Windows.Media.MusicDisplayProperties"
Private Const IID_IMusicDisplayProperties As String = "{6bbf0c59-d0a0-4d26-92a0-f978e1d18e7b}"
Private Const IID_IMusicDisplayProperties2 As String = "{00368462-97d3-44b9-b00f-008afcefaf18}"
Private Const IID_IMusicDisplayProperties3 As String = "{4db51ac1-0681-4e8c-9401-b8159d9eefc7}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IMusicDisplayProperties
    IMusicDisplayProperties_GetTitle = 6
    IMusicDisplayProperties_PutTitle = 7
    IMusicDisplayProperties_GetAlbumArtist = 8
    IMusicDisplayProperties_PutAlbumArtist = 9
    IMusicDisplayProperties_GetArtist = 10
    IMusicDisplayProperties_PutArtist = 11
                    
    ' IMusicDisplayProperties2
    IMusicDisplayProperties2_GetAlbumTitle = 6
    IMusicDisplayProperties2_PutAlbumTitle = 7
    IMusicDisplayProperties2_GetTrackNumber = 8
    IMusicDisplayProperties2_PutTrackNumber = 9
    IMusicDisplayProperties2_GetGenres = 10
                    
    ' IMusicDisplayProperties3
    IMusicDisplayProperties3_GetAlbumTrackCount = 6
    IMusicDisplayProperties3_PutAlbumTrackCount = 7

End Enum

' ----==== Variablen ====----
Private m_pIMusicDisplayProperties As Long
Private m_pIMusicDisplayProperties2 As Long
Private m_pIMusicDisplayProperties3 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMusicDisplayProperties3)
    Call ReleaseIfc(m_pIMusicDisplayProperties2)
    Call ReleaseIfc(m_pIMusicDisplayProperties)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMusicDisplayProperties
End Property

Public Property Let Ifc(ByVal pIMusicDisplayProperties As Long)
    Call ReleaseIfc(m_pIMusicDisplayProperties3)
    Call ReleaseIfc(m_pIMusicDisplayProperties2)
    Call ReleaseIfc(m_pIMusicDisplayProperties)
    m_pIMusicDisplayProperties = pIMusicDisplayProperties
    Call QueryIfc(m_pIMusicDisplayProperties, IID_IMusicDisplayProperties2, m_pIMusicDisplayProperties2)
    Call QueryIfc(m_pIMusicDisplayProperties, IID_IMusicDisplayProperties3, m_pIMusicDisplayProperties3)
End Property

' IMusicDisplayProperties
Public Property Get Title() As String
    Dim Ret As String
    If m_pIMusicDisplayProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIMusicDisplayProperties, _
                  IMusicDisplayProperties_GetTitle, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Title = Ret
End Property

Public Property Let Title(ByVal value As String)
    If m_pIMusicDisplayProperties <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIMusicDisplayProperties, _
                  IMusicDisplayProperties_PutTitle, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get AlbumArtist() As String
    Dim Ret As String
    If m_pIMusicDisplayProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIMusicDisplayProperties, _
                  IMusicDisplayProperties_GetAlbumArtist, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    AlbumArtist = Ret
End Property

Public Property Let AlbumArtist(ByVal value As String)
    If m_pIMusicDisplayProperties <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIMusicDisplayProperties, _
                  IMusicDisplayProperties_PutAlbumArtist, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Artist() As String
    Dim Ret As String
    If m_pIMusicDisplayProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIMusicDisplayProperties, _
                  IMusicDisplayProperties_GetArtist, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Artist = Ret
End Property

Public Property Let Artist(ByVal value As String)
    If m_pIMusicDisplayProperties <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIMusicDisplayProperties, _
                  IMusicDisplayProperties_PutArtist, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property
                    
' IMusicDisplayProperties2
Public Property Get AlbumTitle() As String
    Dim Ret As String
    If m_pIMusicDisplayProperties2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIMusicDisplayProperties2, _
                  IMusicDisplayProperties2_GetAlbumTitle, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    AlbumTitle = Ret
End Property

Public Property Let AlbumTitle(ByVal value As String)
    If m_pIMusicDisplayProperties2 <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIMusicDisplayProperties2, _
                  IMusicDisplayProperties2_PutAlbumTitle, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get TrackNumber() As Long
    Dim Ret As Long
    If m_pIMusicDisplayProperties2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIMusicDisplayProperties2, _
                  IMusicDisplayProperties2_GetTrackNumber, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    TrackNumber = Ret
End Property

Public Property Let TrackNumber(ByVal value As Long)
    If m_pIMusicDisplayProperties2 <> 0& Then
        If Invoke(m_pIMusicDisplayProperties2, _
                  IMusicDisplayProperties2_PutTrackNumber, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get Genres() As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pIMusicDisplayProperties2 <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIMusicDisplayProperties2, _
                  IMusicDisplayProperties2_GetGenres, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_String, pIVectorView) 'ReadOnlyList_String
            End If
        End If
    End If
    Set Genres = Ret
End Property
                    
' IMusicDisplayProperties3
Public Property Get AlbumTrackCount() As Long
    Dim Ret As Long
    If m_pIMusicDisplayProperties3 <> 0& Then
        Dim value As Long
        If Invoke(m_pIMusicDisplayProperties3, _
                  IMusicDisplayProperties3_GetAlbumTrackCount, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    AlbumTrackCount = Ret
End Property

Public Property Let AlbumTrackCount(ByVal value As Long)
    If m_pIMusicDisplayProperties3 <> 0& Then
        If Invoke(m_pIMusicDisplayProperties3, _
                  IMusicDisplayProperties3_PutAlbumTrackCount, _
                  value) = S_OK Then
        End If
    End If
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



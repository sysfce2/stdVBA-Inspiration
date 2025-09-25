VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MusicProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_FileProperties_MusicProperties As String = "Windows.Storage.FileProperties.MusicProperties"
Private Const IID_IMusicProperties As String = "{bc8aab62-66ec-419a-bc5d-ca65a4cb46da}"
Private Const IID_IStorageItemExtraProperties As String = "{c54361b2-54cd-432b-bdbc-4b19c4b470d7}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IMusicProperties
    IMusicProperties_GetAlbum = 6
    IMusicProperties_PutAlbum = 7
    IMusicProperties_GetArtist = 8
    IMusicProperties_PutArtist = 9
    IMusicProperties_GetGenre = 10
    IMusicProperties_GetTrackNumber = 11
    IMusicProperties_PutTrackNumber = 12
    IMusicProperties_GetTitle = 13
    IMusicProperties_PutTitle = 14
    IMusicProperties_GetRating = 15
    IMusicProperties_PutRating = 16
    IMusicProperties_GetDuration = 17
    IMusicProperties_GetBitrate = 18
    IMusicProperties_GetAlbumArtist = 19
    IMusicProperties_PutAlbumArtist = 20
    IMusicProperties_GetComposers = 21
    IMusicProperties_GetConductors = 22
    IMusicProperties_GetSubtitle = 23
    IMusicProperties_PutSubtitle = 24
    IMusicProperties_GetProducers = 25
    IMusicProperties_GetPublisher = 26
    IMusicProperties_PutPublisher = 27
    IMusicProperties_GetWriters = 28
    IMusicProperties_GetYear = 29
    IMusicProperties_PutYear = 30

    ' IStorageItemExtraProperties
'    IStorageItemExtraProperties_RetrievePropertiesAsync = 6
'    IStorageItemExtraProperties_SavePropertiesAsync = 7
    IStorageItemExtraProperties_SavePropertiesAsyncOverloadDefault = 8

End Enum

' ----==== Vars ====----
Private m_pIMusicProperties As Long
Private m_pIStorageItemExtraProperties As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMusicProperties)
    Call ReleaseIfc(m_pIStorageItemExtraProperties)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMusicProperties
End Property

Public Property Let Ifc(ByVal pIMusicProperties As Long)
    Call ReleaseIfc(m_pIMusicProperties)
    Call ReleaseIfc(m_pIStorageItemExtraProperties)
    m_pIMusicProperties = pIMusicProperties
    Call QueryIfc(m_pIMusicProperties, _
                  IID_IStorageItemExtraProperties, _
                  m_pIStorageItemExtraProperties)
End Property

' IMusicProperties
Public Property Get Album() As String
    Dim Ret As String
    If m_pIMusicProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIMusicProperties, _
                  IMusicProperties_GetAlbum, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Album = Ret
End Property

Public Property Let Album(ByVal value As String)
    If m_pIMusicProperties <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIMusicProperties, _
                  IMusicProperties_PutAlbum, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Artist() As String
    Dim Ret As String
    If m_pIMusicProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIMusicProperties, _
                  IMusicProperties_GetArtist, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Artist = Ret
End Property

Public Property Let Artist(ByVal value As String)
    If m_pIMusicProperties <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIMusicProperties, _
                  IMusicProperties_PutArtist, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Genre() As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pIMusicProperties <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIMusicProperties, _
                  IMusicProperties_GetGenre, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_String, pIVectorView) 'ReadOnlyList_String
            End If
        End If
    End If
    Set Genre = Ret
End Property

Public Property Get TrackNumber() As Long
    Dim Ret As Long
    If m_pIMusicProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIMusicProperties, _
                  IMusicProperties_GetTrackNumber, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    TrackNumber = Ret
End Property

Public Property Let TrackNumber(ByVal value As Long)
    If m_pIMusicProperties <> 0& Then
        If Invoke(m_pIMusicProperties, _
                  IMusicProperties_PutTrackNumber, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get Title() As String
    Dim Ret As String
    If m_pIMusicProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIMusicProperties, _
                  IMusicProperties_GetTitle, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Title = Ret
End Property

Public Property Let Title(ByVal value As String)
    If m_pIMusicProperties <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIMusicProperties, _
                  IMusicProperties_PutTitle, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Rating() As Long
    Dim Ret As Long
    If m_pIMusicProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIMusicProperties, _
                  IMusicProperties_GetRating, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Rating = Ret
End Property

Public Property Let Rating(ByVal value As Long)
    If m_pIMusicProperties <> 0& Then
        If Invoke(m_pIMusicProperties, _
                  IMusicProperties_PutRating, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get Duration() As TimeSpan
    Dim Ret As TimeSpan
    If m_pIMusicProperties <> 0& Then
        Dim value As Currency
        If Invoke(m_pIMusicProperties, _
                  IMusicProperties_GetDuration, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New TimeSpan
            Ret.TimeSpanValue = value
        End If
    End If
    Set Duration = Ret
End Property

Public Property Get Bitrate() As Long
    Dim Ret As Long
    If m_pIMusicProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIMusicProperties, _
                  IMusicProperties_GetBitrate, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Bitrate = Ret
End Property

Public Property Get AlbumArtist() As String
    Dim Ret As String
    If m_pIMusicProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIMusicProperties, _
                  IMusicProperties_GetAlbumArtist, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    AlbumArtist = Ret
End Property

Public Property Let AlbumArtist(ByVal value As String)
    If m_pIMusicProperties <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIMusicProperties, _
                  IMusicProperties_PutAlbumArtist, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Composers() As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pIMusicProperties <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIMusicProperties, _
                  IMusicProperties_GetComposers, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_String, pIVectorView) 'ReadOnlyList_String
            End If
        End If
    End If
    Set Composers = Ret
End Property

Public Property Get Conductors() As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pIMusicProperties <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIMusicProperties, _
                  IMusicProperties_GetConductors, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_String, pIVectorView) 'ReadOnlyList_String
            End If
        End If
    End If
    Set Conductors = Ret
End Property

Public Property Get Subtitle() As String
    Dim Ret As String
    If m_pIMusicProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIMusicProperties, _
                  IMusicProperties_GetSubtitle, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Subtitle = Ret
End Property

Public Property Let Subtitle(ByVal value As String)
    If m_pIMusicProperties <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIMusicProperties, _
                  IMusicProperties_PutSubtitle, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Producers() As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pIMusicProperties <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIMusicProperties, _
                  IMusicProperties_GetProducers, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_String, pIVectorView) 'ReadOnlyList_String
            End If
        End If
    End If
    Set Producers = Ret
End Property

Public Property Get Publisher() As String
    Dim Ret As String
    If m_pIMusicProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIMusicProperties, _
                  IMusicProperties_GetPublisher, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Publisher = Ret
End Property

Public Property Let Publisher(ByVal value As String)
    If m_pIMusicProperties <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIMusicProperties, _
                  IMusicProperties_PutPublisher, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Writers() As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pIMusicProperties <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIMusicProperties, _
                  IMusicProperties_GetWriters, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_String, pIVectorView) 'ReadOnlyList_String
            End If
        End If
    End If
    Set Writers = Ret
End Property

Public Property Get Year() As Long
    Dim Ret As Long
    If m_pIMusicProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIMusicProperties, _
                  IMusicProperties_GetYear, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Year = Ret
End Property

Public Property Let Year(ByVal value As Long)
    If m_pIMusicProperties <> 0& Then
        If Invoke(m_pIMusicProperties, _
                  IMusicProperties_PutYear, _
                  value) = S_OK Then
        End If
    End If
End Property

' ----==== Functions ====----
' IStorageItemExtraProperties
'    IStorageItemExtraProperties_RetrievePropertiesAsync = 6
'    IStorageItemExtraProperties_SavePropertiesAsync = 7

Public Function SavePropertiesAsyncOverloadDefault() As Boolean
    Dim Ret As Boolean
    If m_pIStorageItemExtraProperties <> 0& Then
        Dim operation As Long
        If Invoke(m_pIStorageItemExtraProperties, _
                  IStorageItemExtraProperties_SavePropertiesAsyncOverloadDefault, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation, False) Then
                    Ret = True
                End If
            End If
        End If
    End If
    SavePropertiesAsyncOverloadDefault = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



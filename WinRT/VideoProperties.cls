VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VideoProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Storage_FileProperties_VideoProperties As String = "Windows.Storage.FileProperties.VideoProperties"
Private Const IID_IVideoProperties As String = "{719ae507-68de-4db8-97de-49998c059f2f}"
Private Const IID_IStorageItemExtraProperties As String = "{c54361b2-54cd-432b-bdbc-4b19c4b470d7}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IVideoProperties
    IVideoProperties_GetRating = 6
    IVideoProperties_PutRating = 7
    IVideoProperties_GetKeywords = 8
    IVideoProperties_GetWidth = 9
    IVideoProperties_GetHeight = 10
    IVideoProperties_GetDuration = 11
    IVideoProperties_GetLatitude = 12
    IVideoProperties_GetLongitude = 13
    IVideoProperties_GetTitle = 14
    IVideoProperties_PutTitle = 15
    IVideoProperties_GetSubtitle = 16
    IVideoProperties_PutSubtitle = 17
    IVideoProperties_GetProducers = 18
    IVideoProperties_GetPublisher = 19
    IVideoProperties_PutPublisher = 20
    IVideoProperties_GetWriters = 21
    IVideoProperties_GetYear = 22
    IVideoProperties_PutYear = 23
    IVideoProperties_GetBitrate = 24
    IVideoProperties_GetDirectors = 25
    IVideoProperties_GetOrientation = 26

    ' IStorageItemExtraProperties
'    IStorageItemExtraProperties_RetrievePropertiesAsync = 6
'    IStorageItemExtraProperties_SavePropertiesAsync = 7
    IStorageItemExtraProperties_SavePropertiesAsyncOverloadDefault = 8

    ' IReference
    IReference_GetValue = 6

End Enum

' ----==== Vars ====----
Private m_pIVideoProperties As Long
Private m_pIStorageItemExtraProperties As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIVideoProperties)
    Call ReleaseIfc(m_pIStorageItemExtraProperties)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIVideoProperties
End Property

Public Property Let Ifc(ByVal pIVideoProperties As Long)
    Call ReleaseIfc(m_pIVideoProperties)
    Call ReleaseIfc(m_pIStorageItemExtraProperties)
    m_pIVideoProperties = pIVideoProperties
    Call QueryIfc(m_pIVideoProperties, _
                  IID_IStorageItemExtraProperties, _
                  m_pIStorageItemExtraProperties)
End Property

' IVideoProperties
Public Property Get Rating() As Long
    Dim Ret As Long
    If m_pIVideoProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIVideoProperties, _
                  IVideoProperties_GetRating, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Rating = Ret
End Property

Public Property Let Rating(ByVal value As Long)
    If m_pIVideoProperties <> 0& Then
        If Invoke(m_pIVideoProperties, _
                  IVideoProperties_PutRating, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get Keywords() As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pIVideoProperties <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIVideoProperties, _
                  IVideoProperties_GetKeywords, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_String, pIVectorView) 'ReadOnlyList_String
            End If
        End If
    End If
    Set Keywords = Ret
End Property

Public Property Get Width() As Long
    Dim Ret As Long
    If m_pIVideoProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIVideoProperties, _
                  IVideoProperties_GetWidth, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Width = Ret
End Property

Public Property Get Height() As Long
    Dim Ret As Long
    If m_pIVideoProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIVideoProperties, _
                  IVideoProperties_GetHeight, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Height = Ret
End Property

Public Property Get Duration() As TimeSpan
    Dim Ret As TimeSpan
    If m_pIVideoProperties <> 0& Then
        Dim value As Currency
        If Invoke(m_pIVideoProperties, _
                  IVideoProperties_GetDuration, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New TimeSpan
            Ret.TimeSpanValue = value
        End If
    End If
    Set Duration = Ret
End Property

Public Property Get Latitude() As Double
    Dim Ret As Double
    If m_pIVideoProperties <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIVideoProperties, _
                  IVideoProperties_GetLatitude, _
                  VarPtr(pIReference)) = S_OK Then
            Dim value As Double
            If Invoke(pIReference, _
                      IReference_GetValue, _
                      VarPtr(value)) = S_OK Then
                Ret = value
            End If
            Call ReleaseIfc(pIReference)
        End If
    End If
    Latitude = Ret
End Property

Public Property Get Longitude() As Double
    Dim Ret As Double
    If m_pIVideoProperties <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIVideoProperties, _
                  IVideoProperties_GetLongitude, _
                  VarPtr(pIReference)) = S_OK Then
            Dim value As Double
            If Invoke(pIReference, _
                      IReference_GetValue, _
                      VarPtr(value)) = S_OK Then
                Ret = value
            End If
            Call ReleaseIfc(pIReference)
        End If
    End If
    Longitude = Ret
End Property

Public Property Get Title() As String
    Dim Ret As String
    If m_pIVideoProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIVideoProperties, _
                  IVideoProperties_GetTitle, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Title = Ret
End Property

Public Property Let Title(ByVal value As String)
    If m_pIVideoProperties <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIVideoProperties, _
                  IVideoProperties_PutTitle, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Subtitle() As String
    Dim Ret As String
    If m_pIVideoProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIVideoProperties, _
                  IVideoProperties_GetSubtitle, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Subtitle = Ret
End Property

Public Property Let Subtitle(ByVal value As String)
    If m_pIVideoProperties <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIVideoProperties, _
                  IVideoProperties_PutSubtitle, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Producers() As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pIVideoProperties <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIVideoProperties, _
                  IVideoProperties_GetProducers, _
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
    If m_pIVideoProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIVideoProperties, _
                  IVideoProperties_GetPublisher, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Publisher = Ret
End Property

Public Property Let Publisher(ByVal value As String)
    If m_pIVideoProperties <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIVideoProperties, _
                  IVideoProperties_PutPublisher, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Writers() As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pIVideoProperties <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIVideoProperties, _
                  IVideoProperties_GetWriters, _
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
    If m_pIVideoProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIVideoProperties, _
                  IVideoProperties_GetYear, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Year = Ret
End Property

Public Property Let Year(ByVal value As Long)
    If m_pIVideoProperties <> 0& Then
        If Invoke(m_pIVideoProperties, _
                  IVideoProperties_PutYear, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get Bitrate() As Long
    Dim Ret As Long
    If m_pIVideoProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIVideoProperties, _
                  IVideoProperties_GetBitrate, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Bitrate = Ret
End Property

Public Property Get Directors() As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pIVideoProperties <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIVideoProperties, _
                  IVideoProperties_GetDirectors, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_String, pIVectorView) 'ReadOnlyList_String
            End If
        End If
    End If
    Set Directors = Ret
End Property

Public Property Get Orientation() As VideoOrientation
    Dim Ret As VideoOrientation
    If m_pIVideoProperties <> 0& Then
        Dim value As Long
        If Invoke(m_pIVideoProperties, _
                  IVideoProperties_GetOrientation, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Orientation = Ret
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



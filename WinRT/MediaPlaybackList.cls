VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaPlaybackList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 11/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Playback_MediaPlaybackList As String = "Windows.Media.Playback.MediaPlaybackList"
Private Const IID_IMediaPlaybackList As String = "{7f77ee9c-dc42-4e26-a98d-7850df8ec925}"
Private Const IID_IMediaPlaybackList2 As String = "{0e09b478-600a-4274-a14b-0b6723d0f48b}"
Private Const IID_IMediaPlaybackList3 As String = "{0d328220-b80a-4d09-9ff8-f87094a1c831}"
Private Const IID_IMediaPlaybackSource As String = "{dd24bba9-bc47-4463-aa90-c18b7e5ffde1}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    'IMediaPlaybackSource
    ' has no member
    
    ' IMediaPlaybackList
    IMediaPlaybackList_AddItemFailed = 6
    IMediaPlaybackList_RemoveItemFailed = 7
    IMediaPlaybackList_AddCurrentItemChanged = 8
    IMediaPlaybackList_RemoveCurrentItemChanged = 9
    IMediaPlaybackList_AddItemOpened = 10
    IMediaPlaybackList_RemoveItemOpened = 11
'    IMediaPlaybackList_GetItems = 12
    IMediaPlaybackList_GetAutoRepeatEnabled = 13
    IMediaPlaybackList_PutAutoRepeatEnabled = 14
    IMediaPlaybackList_GetShuffleEnabled = 15
    IMediaPlaybackList_PutShuffleEnabled = 16
    IMediaPlaybackList_GetCurrentItem = 17
    IMediaPlaybackList_GetCurrentItemIndex = 18
    IMediaPlaybackList_MoveNext = 19
    IMediaPlaybackList_MovePrevious = 20
    IMediaPlaybackList_MoveTo = 21
                        
    ' IMediaPlaybackList2
    IMediaPlaybackList2_GetMaxPrefetchTime = 6
'    IMediaPlaybackList2_PutMaxPrefetchTime = 7
    IMediaPlaybackList2_GetStartingItem = 8
    IMediaPlaybackList2_PutStartingItem = 9
    IMediaPlaybackList2_GetShuffledItems = 10
'    IMediaPlaybackList2_SetShuffledItems = 11
                        
    ' IMediaPlaybackList3
    IMediaPlaybackList3_GetMaxPlayedItemsToKeepOpen = 6
'    IMediaPlaybackList3_PutMaxPlayedItemsToKeepOpen = 7
    
    ' IReference
    IReference_GetValue = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIMediaPlaybackList As Long
Private m_pIMediaPlaybackList2 As Long
Private m_pIMediaPlaybackList3 As Long
Private m_pIMediaPlaybackSource As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    Dim pIMediaPlaybackList As Long
    If GetActivateInstance(Windows_Media_Playback_MediaPlaybackList, _
                           IID_IMediaPlaybackList, _
                           pIMediaPlaybackList) Then
        Me.Ifc = pIMediaPlaybackList
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMediaPlaybackSource)
    Call ReleaseIfc(m_pIMediaPlaybackList3)
    Call ReleaseIfc(m_pIMediaPlaybackList2)
    Call ReleaseIfc(m_pIMediaPlaybackList)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMediaPlaybackList
End Property

Public Property Let Ifc(ByVal pIMediaPlaybackList As Long)
    Call ReleaseIfc(m_pIMediaPlaybackSource)
    Call ReleaseIfc(m_pIMediaPlaybackList3)
    Call ReleaseIfc(m_pIMediaPlaybackList2)
    Call ReleaseIfc(m_pIMediaPlaybackList)
    m_pIMediaPlaybackList = pIMediaPlaybackList
    Call QueryIfc(m_pIMediaPlaybackList, IID_IMediaPlaybackList2, m_pIMediaPlaybackList2)
    Call QueryIfc(m_pIMediaPlaybackList, IID_IMediaPlaybackList3, m_pIMediaPlaybackList3)
    Call QueryIfc(m_pIMediaPlaybackList, IID_IMediaPlaybackSource, m_pIMediaPlaybackSource)
End Property

' IMediaPlaybackList
'    IMediaPlaybackList_GetItems = 12

Public Property Get AutoRepeatEnabled() As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackList <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlaybackList, _
                  IMediaPlaybackList_GetAutoRepeatEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    AutoRepeatEnabled = Ret
End Property

Public Property Let AutoRepeatEnabled(ByVal value As Boolean)
    If m_pIMediaPlaybackList <> 0& Then
        If Invoke(m_pIMediaPlaybackList, _
                  IMediaPlaybackList_PutAutoRepeatEnabled, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get ShuffleEnabled() As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackList <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlaybackList, _
                  IMediaPlaybackList_GetShuffleEnabled, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    ShuffleEnabled = Ret
End Property

Public Property Let ShuffleEnabled(ByVal value As Boolean)
    If m_pIMediaPlaybackList <> 0& Then
        If Invoke(m_pIMediaPlaybackList, _
                  IMediaPlaybackList_PutShuffleEnabled, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get CurrentItem() As MediaPlaybackItem
    Dim Ret As MediaPlaybackItem
    If m_pIMediaPlaybackList <> 0& Then
        Dim pIMediaPlaybackItem As Long
        If Invoke(m_pIMediaPlaybackList, _
                  IMediaPlaybackList_GetCurrentItem, _
                  VarPtr(pIMediaPlaybackItem)) = S_OK Then
            If pIMediaPlaybackItem <> 0& Then
                Set Ret = New MediaPlaybackItem
                Ret.Ifc = pIMediaPlaybackItem
            End If
        End If
    End If
    Set CurrentItem = Ret
End Property

Public Property Get CurrentItemIndex() As Long
    Dim Ret As Long
    If m_pIMediaPlaybackList <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlaybackList, _
                  IMediaPlaybackList_GetCurrentItemIndex, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    CurrentItemIndex = Ret
End Property
                        
' IMediaPlaybackList2
Public Property Get MaxPrefetchTime() As TimeSpan
    Dim Ret As TimeSpan
    If m_pIMediaPlaybackList2 <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIMediaPlaybackList2, _
                  IMediaPlaybackList2_GetMaxPrefetchTime, _
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
    Set MaxPrefetchTime = Ret
End Property

'    IMediaPlaybackList2_PutMaxPrefetchTime = 7

Public Property Get StartingItem() As MediaPlaybackItem
    Dim Ret As MediaPlaybackItem
    If m_pIMediaPlaybackList2 <> 0& Then
        Dim pIMediaPlaybackItem As Long
        If Invoke(m_pIMediaPlaybackList2, _
                  IMediaPlaybackList2_GetStartingItem, _
                  VarPtr(pIMediaPlaybackItem)) = S_OK Then
            If pIMediaPlaybackItem <> 0& Then
                Set Ret = New MediaPlaybackItem
                Ret.Ifc = pIMediaPlaybackItem
            End If
        End If
    End If
    Set StartingItem = Ret
End Property

Public Property Let StartingItem(ByVal value As MediaPlaybackItem)
    If m_pIMediaPlaybackList2 <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pIMediaPlaybackList2, _
                  IMediaPlaybackList2_PutStartingItem, _
                  value.Ifc) = S_OK Then
        End If
    End If
End Property

Public Property Get ShuffledItems() As ReadOnlyList_1 'ReadOnlyList_MediaPlaybackItem
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_MediaPlaybackItem
    If m_pIMediaPlaybackList2 <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIMediaPlaybackList2, _
                  IMediaPlaybackList2_GetShuffledItems, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_MediaPlaybackItem, pIVectorView) 'ReadOnlyList_MediaPlaybackItem
            End If
        End If
    End If
    Set ShuffledItems = Ret
End Property
                        
' IMediaPlaybackList3
Public Property Get MaxPlayedItemsToKeepOpen() As Long
    Dim Ret As Long
    If m_pIMediaPlaybackList3 <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIMediaPlaybackList3, _
                  IMediaPlaybackList3_GetMaxPlayedItemsToKeepOpen, _
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
    MaxPlayedItemsToKeepOpen = Ret
End Property

'    IMediaPlaybackList3_PutMaxPlayedItemsToKeepOpen = 7

' ----==== Functions ====----
' IMediaPlaybackList
Public Function AddItemFailed(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackList <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackList, _
                  IMediaPlaybackList_AddItemFailed, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddItemFailed = Ret
End Function

Public Function RemoveItemFailed(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackList <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackList, _
                  IMediaPlaybackList_RemoveItemFailed, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveItemFailed = Ret
End Function

Public Function AddCurrentItemChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackList <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackList, _
                  IMediaPlaybackList_AddCurrentItemChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddCurrentItemChanged = Ret
End Function

Public Function RemoveCurrentItemChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackList <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackList, _
                  IMediaPlaybackList_RemoveCurrentItemChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveCurrentItemChanged = Ret
End Function

Public Function AddItemOpened(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaPlaybackList <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaPlaybackList, _
                  IMediaPlaybackList_AddItemOpened, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddItemOpened = Ret
End Function

Public Function RemoveItemOpened(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaPlaybackList <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaPlaybackList, _
                  IMediaPlaybackList_RemoveItemOpened, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveItemOpened = Ret
End Function

Public Function MoveNext() As MediaPlaybackItem
    Dim Ret As MediaPlaybackItem
    If m_pIMediaPlaybackList <> 0& Then
        Dim pIMediaPlaybackItem As Long
        If Invoke(m_pIMediaPlaybackList, _
                  IMediaPlaybackList_MoveNext, _
                  VarPtr(pIMediaPlaybackItem)) = S_OK Then
            If pIMediaPlaybackItem <> 0& Then
                Set Ret = New MediaPlaybackItem
                Ret.Ifc = pIMediaPlaybackItem
            End If
        End If
    End If
    Set MoveNext = Ret
End Function

Public Function MovePrevious() As MediaPlaybackItem
    Dim Ret As MediaPlaybackItem
    If m_pIMediaPlaybackList <> 0& Then
        Dim pIMediaPlaybackItem As Long
        If Invoke(m_pIMediaPlaybackList, _
                  IMediaPlaybackList_MovePrevious, _
                  VarPtr(pIMediaPlaybackItem)) = S_OK Then
            If pIMediaPlaybackItem <> 0& Then
                Set Ret = New MediaPlaybackItem
                Ret.Ifc = pIMediaPlaybackItem
            End If
        End If
    End If
    Set MovePrevious = Ret
End Function

Public Function MoveTo(ByVal ItemIndex As Long) As MediaPlaybackItem
    Dim Ret As MediaPlaybackItem
    If m_pIMediaPlaybackList <> 0& Then
        Dim pIMediaPlaybackItem As Long
        If Invoke(m_pIMediaPlaybackList, _
                  IMediaPlaybackList_MoveTo, _
                  ItemIndex, _
                  VarPtr(pIMediaPlaybackItem)) = S_OK Then
            If pIMediaPlaybackItem <> 0& Then
                Set Ret = New MediaPlaybackItem
                Ret.Ifc = pIMediaPlaybackItem
            End If
        End If
    End If
    Set MoveTo = Ret
End Function
                        
' IMediaPlaybackList2
'    IMediaPlaybackList2_SetShuffledItems = 11
                        
' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



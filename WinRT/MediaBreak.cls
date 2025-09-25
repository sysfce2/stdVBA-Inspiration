VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaBreak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 11/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Playback_MediaBreak As String = "Windows.Media.Playback.MediaBreak"
Private Const IID_IMediaBreak As String = "{714be270-0def-4ebc-a489-6b34930e1558}"
Private Const IID_IMediaBreakFactory As String = "{4516e002-18e0-4079-8b5f-d33495c15d2e}"
 
' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IMediaBreakFactory
    IMediaBreakFactory_Create = 6
    IMediaBreakFactory_CreateWithPresentationPosition = 7
    
    ' IMediaBreak
    IMediaBreak_GetPlaybackList = 6
    IMediaBreak_GetPresentationPosition = 7
    IMediaBreak_GetInsertionMethod = 8
'    IMediaBreak_GetCustomProperties = 9
    IMediaBreak_GetCanStart = 10
    IMediaBreak_PutCanStart = 11
                        
    ' IReference
    IReference_GetValue = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIMediaBreak As Long
Private m_pIMediaBreakFactory As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Media_Playback_MediaBreak, _
                            IID_IMediaBreakFactory, _
                            m_pIMediaBreakFactory) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMediaBreak)
    Call ReleaseIfc(m_pIMediaBreakFactory)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMediaBreak
End Property

Public Property Let Ifc(ByVal pIMediaBreak As Long)
    Call ReleaseIfc(m_pIMediaBreak)
    m_pIMediaBreak = pIMediaBreak
End Property
          
' IMediaBreak
Public Property Get PlaybackList() As MediaPlaybackList
    Dim Ret As MediaPlaybackList
    If m_pIMediaBreak <> 0& Then
        Dim pIMediaPlaybackList As Long
        If Invoke(m_pIMediaBreak, _
                  IMediaBreak_GetPlaybackList, _
                  VarPtr(pIMediaPlaybackList)) = S_OK Then
            If pIMediaPlaybackList <> 0& Then
                Set Ret = New MediaPlaybackList
                Ret.Ifc = pIMediaPlaybackList
            End If
        End If
    End If
    Set PlaybackList = Ret
End Property

Public Property Get PresentationPosition() As TimeSpan
    Dim Ret As TimeSpan
    If m_pIMediaBreak <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIMediaBreak, _
                  IMediaBreak_GetPresentationPosition, _
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
    Set PresentationPosition = Ret
End Property

Public Property Get InsertionMethod() As MediaBreakInsertionMethod
    Dim Ret As MediaBreakInsertionMethod
    If m_pIMediaBreak <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaBreak, _
                  IMediaBreak_GetInsertionMethod, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    InsertionMethod = Ret
End Property

'    IMediaBreak_GetCustomProperties = 9

Public Property Get CanStart() As Boolean
    Dim Ret As Boolean
    If m_pIMediaBreak <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaBreak, _
                  IMediaBreak_GetCanStart, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    CanStart = Ret
End Property

Public Property Let CanStart(ByVal value As Boolean)
    If m_pIMediaBreak <> 0& Then
        If Invoke(m_pIMediaBreak, _
                  IMediaBreak_PutCanStart, _
                  value) = S_OK Then
        End If
    End If
End Property
          
' ----==== Functions ====----
' IMediaBreakFactory
Public Function Create(ByVal InsertionMethod As MediaBreakInsertionMethod) As MediaBreak
    Dim Ret As MediaBreak
    If m_pIMediaBreakFactory <> 0& Then
        Dim pIMediaBreak As Long
        If Invoke(m_pIMediaBreakFactory, _
                  IMediaBreakFactory_Create, _
                  InsertionMethod, _
                  VarPtr(pIMediaBreak)) = S_OK Then
            If pIMediaBreak <> 0& Then
                Set Ret = New MediaBreak
                Ret.Ifc = pIMediaBreak
            End If
        End If
    End If
    Set Create = Ret
End Function

Public Function CreateWithPresentationPosition(ByVal InsertionMethod As MediaBreakInsertionMethod, _
                                               ByVal PresentationPosition As TimeSpan) As MediaBreak
    Dim Ret As MediaBreak
    If m_pIMediaBreakFactory <> 0& Then
        Dim pIMediaBreak As Long
        If Invoke(m_pIMediaBreakFactory, _
                  IMediaBreakFactory_CreateWithPresentationPosition, _
                  InsertionMethod, _
                  PresentationPosition.TimeSpanValue, _
                  VarPtr(pIMediaBreak)) = S_OK Then
            If pIMediaBreak <> 0& Then
                Set Ret = New MediaBreak
                Ret.Ifc = pIMediaBreak
            End If
        End If
    End If
    Set CreateWithPresentationPosition = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



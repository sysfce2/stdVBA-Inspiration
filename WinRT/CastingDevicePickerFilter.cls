VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CastingDevicePickerFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Casting_CastingDevicePickerFilter As String = "Windows.Media.Casting.CastingDevicePickerFilter"
Private Const IID_ICastingDevicePickerFilter As String = "{91db92a2-57cb-48f1-9b59-a59b7a1f02a2}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ICastingDevicePickerFilter
    ICastingDevicePickerFilter_GetSupportsAudio = 6
    ICastingDevicePickerFilter_PutSupportsAudio = 7
    ICastingDevicePickerFilter_GetSupportsVideo = 8
    ICastingDevicePickerFilter_PutSupportsVideo = 9
    ICastingDevicePickerFilter_GetSupportsPictures = 10
    ICastingDevicePickerFilter_PutSupportsPictures = 11
    ICastingDevicePickerFilter_GetSupportedCastingSources = 12
                            
End Enum

' ----==== Variablen ====----
Private m_pICastingDevicePickerFilter As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pICastingDevicePickerFilter)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pICastingDevicePickerFilter
End Property

Public Property Let Ifc(ByVal pICastingDevicePickerFilterAppearance As Long)
    Call ReleaseIfc(m_pICastingDevicePickerFilter)
    m_pICastingDevicePickerFilter = pICastingDevicePickerFilterAppearance
End Property

' ICastingDevicePickerFilter
Public Property Get SupportsAudio() As Boolean
    Dim Ret As Boolean
    If m_pICastingDevicePickerFilter <> 0& Then
        Dim Value As Long
        If Invoke(m_pICastingDevicePickerFilter, _
                  ICastingDevicePickerFilter_GetSupportsAudio, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    SupportsAudio = Ret
End Property

Public Property Let SupportsAudio(ByVal Value As Boolean)
    If m_pICastingDevicePickerFilter <> 0& Then
        If Invoke(m_pICastingDevicePickerFilter, _
                  ICastingDevicePickerFilter_PutSupportsAudio, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get SupportsVideo() As Boolean
    Dim Ret As Boolean
    If m_pICastingDevicePickerFilter <> 0& Then
        Dim Value As Long
        If Invoke(m_pICastingDevicePickerFilter, _
                  ICastingDevicePickerFilter_GetSupportsVideo, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    SupportsVideo = Ret
End Property

Public Property Let SupportsVideo(ByVal Value As Boolean)
    If m_pICastingDevicePickerFilter <> 0& Then
        If Invoke(m_pICastingDevicePickerFilter, _
                  ICastingDevicePickerFilter_PutSupportsVideo, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get SupportsPictures() As Boolean
    Dim Ret As Boolean
    If m_pICastingDevicePickerFilter <> 0& Then
        Dim Value As Long
        If Invoke(m_pICastingDevicePickerFilter, _
                  ICastingDevicePickerFilter_GetSupportsPictures, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    SupportsPictures = Ret
End Property

Public Property Let SupportsPictures(ByVal Value As Boolean)
    If m_pICastingDevicePickerFilter <> 0& Then
        If Invoke(m_pICastingDevicePickerFilter, _
                  ICastingDevicePickerFilter_PutSupportsPictures, _
                  Value) = S_OK Then
        End If
    End If
End Property

'    ICastingDevicePickerFilter_GetSupportedCastingSources = 6
Public Property Get SupportedCastingSources() As List_CastingSource
    Dim Ret As List_CastingSource
    If m_pICastingDevicePickerFilter <> 0& Then
        Dim pIVector As Long
        If Invoke(m_pICastingDevicePickerFilter, _
                  ICastingDevicePickerFilter_GetSupportedCastingSources, _
                  VarPtr(pIVector)) = S_OK Then
            If pIVector <> 0& Then
                Set Ret = New List_CastingSource
                Ret.Ifc = pIVector
            End If
        End If
    End If
    Set SupportedCastingSources = Ret
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



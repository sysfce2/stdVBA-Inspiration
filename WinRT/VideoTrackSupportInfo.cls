VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VideoTrackSupportInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 11/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Core_VideoTrackSupportInfo As String = "Windows.Media.Core.VideoTrackSupportInfo"
Private Const IID_IVideoTrackSupportInfo As String = "{4bb534a0-fc5f-450d-8ff0-778d590486de}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IVideoTrackSupportInfo
    IVideoTrackSupportInfo_GetDecoderStatus = 6
    IVideoTrackSupportInfo_GetMediaSourceStatus = 7
                        
End Enum

' ----==== Vars ====----
Private m_pIVideoTrackSupportInfo As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIVideoTrackSupportInfo)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIVideoTrackSupportInfo
End Property

Public Property Let Ifc(ByVal pIVideoTrackSupportInfo As Long)
    Call ReleaseIfc(m_pIVideoTrackSupportInfo)
    m_pIVideoTrackSupportInfo = pIVideoTrackSupportInfo
End Property

' IVideoTrackSupportInfo
Public Property Get DecoderStatus() As MediaDecoderStatus
    Dim Ret As MediaDecoderStatus
    If m_pIVideoTrackSupportInfo <> 0& Then
        Dim value As Long
        If Invoke(m_pIVideoTrackSupportInfo, _
                  IVideoTrackSupportInfo_GetDecoderStatus, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    DecoderStatus = Ret
End Property

Public Property Get MediaSourceStatus() As MediaSourceStatus
    Dim Ret As MediaSourceStatus
    If m_pIVideoTrackSupportInfo <> 0& Then
        Dim value As Long
        If Invoke(m_pIVideoTrackSupportInfo, _
                  IVideoTrackSupportInfo_GetMediaSourceStatus, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    MediaSourceStatus = Ret
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



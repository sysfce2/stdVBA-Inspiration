VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaPlayerFailedEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Playback_MediaPlayer_MediaPlayerFailedEventArgs As String = "Windows.Media.Playback.MediaPlayer.MediaPlayerFailedEventArgs"
Private Const IID_IMediaPlayerFailedEventArgs As String = "{2744e9b9-a7e3-4f16-bac4-7914ebc08301}"
 
' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IMediaPlayerFailedEventArgs
    IMediaPlayerFailedEventArgs_GetError = 6
    IMediaPlayerFailedEventArgs_GetExtendedErrorCode = 7
    IMediaPlayerFailedEventArgs_GetErrorMessage = 8
    
End Enum

' ----==== Variablen ====----
Private m_pIMediaPlayerFailedEventArgs As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMediaPlayerFailedEventArgs)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMediaPlayerFailedEventArgs
End Property

Public Property Let Ifc(ByVal pIMediaPlayerFailedEventArgs As Long)
    Call ReleaseIfc(m_pIMediaPlayerFailedEventArgs)
    m_pIMediaPlayerFailedEventArgs = pIMediaPlayerFailedEventArgs
End Property
                        
' IMediaPlayerFailedEventArgs
Public Property Get Error() As MediaPlayerError
    Dim Ret As MediaPlayerError
    If m_pIMediaPlayerFailedEventArgs <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlayerFailedEventArgs, _
                  IMediaPlayerFailedEventArgs_GetError, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Error = Ret
End Property

Public Property Get ExtendedErrorCode() As Long
    Dim Ret As Long
    If m_pIMediaPlayerFailedEventArgs <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlayerFailedEventArgs, _
                  IMediaPlayerFailedEventArgs_GetExtendedErrorCode, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    ExtendedErrorCode = Ret
End Property

Public Property Get ErrorMessage() As String
    Dim Ret As String
    If m_pIMediaPlayerFailedEventArgs <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaPlayerFailedEventArgs, _
                  IMediaPlayerFailedEventArgs_GetErrorMessage, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    ErrorMessage = Ret
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


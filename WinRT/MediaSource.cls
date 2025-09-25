VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 09/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Core_MediaSource As String = "Windows.Media.Core.MediaSource"
Private Const IID_IMediaSourceStatics As String = "{f77d6fa4-4652-410e-b1d8-e9a5e245a45c}"
Private Const IID_IMediaSourceStatics2 As String = "{eee161a4-7f13-4896-b8cb-df0de5bcb9f1}"
Private Const IID_IMediaSourceStatics3 As String = "{453a30d6-2bea-4122-9f73-eace04526e35}"
Private Const IID_IMediaSourceStatics4 As String = "{281b3bfc-e50a-4428-a500-9c4ed918d3f0}"
Private Const IID_IMediaSource2 As String = "{2eb61048-655f-4c37-b813-b4e45dfa0abe}"
Private Const IID_IMediaSource3 As String = "{b59f0d9b-4b6e-41ed-bbb4-7c7509a994ad}"
Private Const IID_IMediaSource4 As String = "{bdafad57-8eff-4c63-85a6-84de0ae3e4f2}"
Private Const IID_IMediaSource5 As String = "{331a22ae-ed2e-4a22-94c8-b743a92b3022}"
Private Const IID_IMediaPlaybackSource As String = "{ef9dc2bc-9317-4696-b051-2bad643177b5}"

' * Class implements the following interface:
' *    Windows.Foundation.IClosable
 
' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IMediaSourceStatics
'    IMediaSourceStatics_CreateFromAdaptiveMediaSource = 6
'    IMediaSourceStatics_CreateFromMediaStreamSource = 7
'    IMediaSourceStatics_CreateFromMseStreamSource = 8
'    IMediaSourceStatics_CreateFromIMediaSource = 9
    IMediaSourceStatics_CreateFromStorageFile = 10
    IMediaSourceStatics_CreateFromStream = 11
    IMediaSourceStatics_CreateFromStreamReference = 12
    IMediaSourceStatics_CreateFromUri = 13
                        
    ' IMediaSourceStatics2
'    IMediaSourceStatics2_CreateFromMediaBinder = 6

    ' IMediaSourceStatics3
'    IMediaSourceStatics3_CreateFromMediaFrameSource = 6

    ' IMediaSourceStatics4
'    IMediaSourceStatics4_CreateFromDownloadOperation = 6

    ' IMediaSource2
    IMediaSource2_AddOpenOperationCompleted = 6
    IMediaSource2_RemoveOpenOperationCompleted = 7
'    IMediaSource2_GetCustomProperties = 8
    IMediaSource2_GetDuration = 9
    IMediaSource2_GetIsOpen = 10
'    IMediaSource2_GetExternalTimedTextSources = 11
'    IMediaSource2_GetExternalTimedMetadataTracks = 12
    
    ' IMediaSource3
    IMediaSource3_AddStateChanged = 6
    IMediaSource3_RemoveStateChanged = 7
'    IMediaSource3_GetState = 8
    IMediaSource3_Reset = 9
                    
    ' IMediaSource4
'    IMediaSource4_GetAdaptiveMediaSource = 6
'    IMediaSource4_GetMediaStreamSource = 6
'    IMediaSource4_GetMseStreamSource = 6
'    IMediaSource4_GetUri = 6
'    IMediaSource4_OpenAsync = 6
                        
    ' IMediaSource5
'    IMediaSource5_GetDownloadOperation = 6
    
    ' IMediaPlaybackSource
    ' has no member
    
    ' IReference
    IReference_GetValue = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIMediaSourceStatics As Long
Private m_pIMediaSourceStatics2 As Long
Private m_pIMediaSourceStatics3 As Long
Private m_pIMediaSourceStatics4 As Long
Private m_pIMediaSource2 As Long
Private m_pIMediaSource3 As Long
Private m_pIMediaSource4 As Long
Private m_pIMediaSource5 As Long
Private m_pIMediaPlaybackSource As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Media_Core_MediaSource, _
                           IID_IMediaSourceStatics, _
                           m_pIMediaSourceStatics) Then
        Call QueryIfc(m_pIMediaSourceStatics, IID_IMediaSourceStatics2, m_pIMediaSourceStatics2)
        Call QueryIfc(m_pIMediaSourceStatics, IID_IMediaSourceStatics3, m_pIMediaSourceStatics3)
        Call QueryIfc(m_pIMediaSourceStatics, IID_IMediaSourceStatics4, m_pIMediaSourceStatics4)
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIMediaSourceStatics4)
    Call ReleaseIfc(m_pIMediaSourceStatics3)
    Call ReleaseIfc(m_pIMediaSourceStatics2)
    Call ReleaseIfc(m_pIMediaSourceStatics)
    Call ReleaseIfc(m_pIMediaPlaybackSource)
    Call ReleaseIfc(m_pIMediaSource5)
    Call ReleaseIfc(m_pIMediaSource4)
    Call ReleaseIfc(m_pIMediaSource3)
    Call DisposeIfc(m_pIMediaSource2) ' <- IClosable
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIMediaSource2
End Property

Public Property Let Ifc(ByVal pIMediaSource2 As Long)
    Call ReleaseIfc(m_pIMediaPlaybackSource)
    Call ReleaseIfc(m_pIMediaSource5)
    Call ReleaseIfc(m_pIMediaSource4)
    Call ReleaseIfc(m_pIMediaSource3)
    Call DisposeIfc(m_pIMediaSource2) ' <- IClosable
    m_pIMediaSource2 = pIMediaSource2
    Call QueryIfc(m_pIMediaSource2, IID_IMediaSource3, m_pIMediaSource3)
    Call QueryIfc(m_pIMediaSource2, IID_IMediaSource4, m_pIMediaSource4)
    Call QueryIfc(m_pIMediaSource2, IID_IMediaSource5, m_pIMediaSource5)
    Call QueryIfc(m_pIMediaSource2, IID_IMediaPlaybackSource, m_pIMediaPlaybackSource)
End Property
                        
' IMediaSource2
'    IMediaSource2_GetCustomProperties = 8

Public Property Get Duration() As TimeSpan
    Dim Ret As TimeSpan
    If m_pIMediaSource2 <> 0& Then
        Dim pIReference As Long
        If Invoke(m_pIMediaSource2, _
                  IMediaSource2_GetDuration, _
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
    Set Duration = Ret
End Property

Public Property Get IsOpen() As Boolean
    Dim Ret As Boolean
    If m_pIMediaSource2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaSource2, _
                  IMediaSource2_GetIsOpen, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsOpen = Ret
End Property

'    IMediaSource2_GetExternalTimedTextSources = 11
'    IMediaSource2_GetExternalTimedMetadataTracks = 12
    
' IMediaSource3
'    IMediaSource3_GetState = 8
                    
' IMediaSource4
'    IMediaSource4_GetAdaptiveMediaSource = 6
'    IMediaSource4_GetMediaStreamSource = 6
'    IMediaSource4_GetMseStreamSource = 6
'    IMediaSource4_GetUri = 6
                        
' IMediaSource5
'    IMediaSource5_GetDownloadOperation = 6
    
' ----==== Functions ====----
' IMediaSourceStatics
'    IMediaSourceStatics_CreateFromAdaptiveMediaSource = 6
'    IMediaSourceStatics_CreateFromMediaStreamSource = 7
'    IMediaSourceStatics_CreateFromMseStreamSource = 8
'    IMediaSourceStatics_CreateFromIMediaSource = 9

Public Function CreateFromStorageFile(ByVal file As StorageFile) As MediaSource
    Dim Ret As MediaSource
    If m_pIMediaSourceStatics <> 0& And _
       IsNotNothing(file) Then
        Dim result As Long
        If Invoke(m_pIMediaSourceStatics, _
                  IMediaSourceStatics_CreateFromStorageFile, _
                  file.Ifc, _
                  VarPtr(result)) = S_OK Then
            If result <> 0& Then
                Set Ret = New MediaSource
                Ret.Ifc = result
            End If
        End If
    End If
    Set CreateFromStorageFile = Ret
End Function

Public Function CreateFromStream(ByVal stream As RandomAccessStream, _
                                 ByVal ContentType As String) As MediaSource
    Dim Ret As MediaSource
    If m_pIMediaSourceStatics <> 0& And _
       IsNotNothing(stream) Then
        Dim result As Long
        Dim hString As Long
        hString = CreateWindowsString(ContentType)
        If Invoke(m_pIMediaSourceStatics, _
                  IMediaSourceStatics_CreateFromStream, _
                  stream.Ifc, _
                  hString, _
                  VarPtr(result)) = S_OK Then
            If result <> 0& Then
                Set Ret = New MediaSource
                Ret.Ifc = result
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set CreateFromStream = Ret
End Function

Public Function CreateFromStreamReference(ByVal stream As RandomAccessStreamReference, _
                                          ByVal ContentType As String) As MediaSource
    Dim Ret As MediaSource
    If m_pIMediaSourceStatics <> 0& And _
       IsNotNothing(stream) Then
        Dim result As Long
        Dim hString As Long
        hString = CreateWindowsString(ContentType)
        If Invoke(m_pIMediaSourceStatics, _
                  IMediaSourceStatics_CreateFromStreamReference, _
                  stream.Ifc, _
                  hString, _
                  VarPtr(result)) = S_OK Then
            If result <> 0& Then
                Set Ret = New MediaSource
                Ret.Ifc = result
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set CreateFromStreamReference = Ret
End Function

Public Function CreateFromUri(ByVal value As Uri) As MediaSource
    Dim Ret As MediaSource
    If m_pIMediaSourceStatics <> 0& And _
       IsNotNothing(value) Then
        Dim result As Long
        If Invoke(m_pIMediaSourceStatics, _
                  IMediaSourceStatics_CreateFromUri, _
                  value.Ifc, _
                  VarPtr(result)) = S_OK Then
            If result <> 0& Then
                Set Ret = New MediaSource
                Ret.Ifc = result
            End If
        End If
    End If
    Set CreateFromUri = Ret
End Function
                        
' IMediaSourceStatics2
'    IMediaSourceStatics2_CreateFromMediaBinder = 6

' IMediaSourceStatics3
'    IMediaSourceStatics3_CreateFromMediaFrameSource = 6

' IMediaSourceStatics4
'    IMediaSourceStatics4_CreateFromDownloadOperation = 6

' IMediaSource2
Public Function AddOpenOperationCompleted(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaSource2 <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaSource2, _
                  IMediaSource2_AddOpenOperationCompleted, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddOpenOperationCompleted = Ret
End Function

Public Function RemoveOpenOperationCompleted(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaSource2 <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaSource2, _
                  IMediaSource2_RemoveOpenOperationCompleted, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveOpenOperationCompleted = Ret
End Function
    
' IMediaSource3
Public Function AddStateChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIMediaSource3 <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIMediaSource3, _
                  IMediaSource3_AddStateChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddStateChanged = Ret
End Function

Public Function RemoveStateChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIMediaSource3 <> 0& And token <> 0@ Then
        If Invoke(m_pIMediaSource3, _
                  IMediaSource3_RemoveStateChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveStateChanged = Ret
End Function

Public Function Reset() As Boolean
    Dim Ret As Boolean
    If m_pIMediaSource3 <> 0& Then
        Dim value As Long
        If Invoke(m_pIMediaSource3, _
                  IMediaSource3_Reset) = S_OK Then
            Ret = True
        End If
    End If
    Reset = Ret
End Function
                    
' IMediaSource4
'    IMediaSource4_OpenAsync = 6
    
' IMediaPlaybackSource
' has no member

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AudioStateMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 11/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Audio_AudioStateMonitor As String = "Windows.Media.Audio.AudioStateMonitor"
Private Const IID_IAudioStateMonitor As String = "{1d13d136-0199-4cdc-b84e-e72c2b581ece}"
Private Const IID_IAudioStateMonitorStatics As String = "{6374ea4c-1b3b-4001-94d9-dd225330fa40}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IAudioStateMonitorStatics
    IAudioStateMonitorStatics_CreateForRenderMonitoring = 6
    IAudioStateMonitorStatics_CreateForRenderMonitoringWithCategory = 7
    IAudioStateMonitorStatics_CreateForRenderMonitoringWithCategoryAndDeviceRole = 8
    IAudioStateMonitorStatics_CreateForRenderMonitoringWithCategoryAndDeviceId = 9
    IAudioStateMonitorStatics_CreateForCaptureMonitoring = 10
    IAudioStateMonitorStatics_CreateForCaptureMonitoringWithCategory = 11
    IAudioStateMonitorStatics_CreateForCaptureMonitoringWithCategoryAndDeviceRole = 12
    IAudioStateMonitorStatics_CreateForCaptureMonitoringWithCategoryAndDeviceId = 13
                        
    ' IAudioStateMonitor
    IAudioStateMonitor_AddSoundLevelChanged = 6
    IAudioStateMonitor_RemoveSoundLevelChanged = 7
    IAudioStateMonitor_GetSoundLevel = 8

End Enum

' ----==== Variablen ====----
Private m_pIAudioStateMonitor As Long
Private m_pIAudioStateMonitorStatics As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Media_Audio_AudioStateMonitor, _
                            IID_IAudioStateMonitorStatics, _
                            m_pIAudioStateMonitorStatics) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIAudioStateMonitor)
    Call ReleaseIfc(m_pIAudioStateMonitorStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIAudioStateMonitor
End Property

Public Property Let Ifc(ByVal pIAudioStateMonitor As Long)
    Call ReleaseIfc(m_pIAudioStateMonitor)
    m_pIAudioStateMonitor = pIAudioStateMonitor
End Property

' IAudioStateMonitor
Public Property Get SoundLevel() As Long
    Dim Ret As Long
    If m_pIAudioStateMonitor <> 0& Then
        Dim value As Long
        If Invoke(m_pIAudioStateMonitor, _
                  IAudioStateMonitor_GetSoundLevel, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    SoundLevel = Ret
End Property

' ----==== Functions ====----
' IAudioStateMonitorStatics
Public Function CreateForRenderMonitoring() As AudioStateMonitor
    Dim Ret As AudioStateMonitor
    If m_pIAudioStateMonitorStatics <> 0& Then
        Dim pIAudioStateMonitor As Long
        If Invoke(m_pIAudioStateMonitorStatics, _
                  IAudioStateMonitorStatics_CreateForRenderMonitoring, _
                  VarPtr(pIAudioStateMonitor)) = S_OK Then
            If pIAudioStateMonitor <> 0& Then
                Set Ret = New AudioStateMonitor
                Ret.Ifc = pIAudioStateMonitor
            End If
        End If
    End If
    Set CreateForRenderMonitoring = Ret
End Function

Public Function CreateForRenderMonitoringWithCategory(ByVal Category As AudioRenderCategory) As AudioStateMonitor
    Dim Ret As AudioStateMonitor
    If m_pIAudioStateMonitorStatics <> 0& Then
        Dim pIAudioStateMonitor As Long
        If Invoke(m_pIAudioStateMonitorStatics, _
                  IAudioStateMonitorStatics_CreateForRenderMonitoringWithCategory, _
                  Category, _
                  VarPtr(pIAudioStateMonitor)) = S_OK Then
            If pIAudioStateMonitor <> 0& Then
                Set Ret = New AudioStateMonitor
                Ret.Ifc = pIAudioStateMonitor
            End If
        End If
    End If
    Set CreateForRenderMonitoringWithCategory = Ret
End Function

Public Function CreateForRenderMonitoringWithCategoryAndDeviceRole(ByVal Category As AudioRenderCategory, _
                                                                   ByVal Role As AudioDeviceRole) As AudioStateMonitor
    Dim Ret As AudioStateMonitor
    If m_pIAudioStateMonitorStatics <> 0& Then
        Dim pIAudioStateMonitor As Long
        If Invoke(m_pIAudioStateMonitorStatics, _
                  IAudioStateMonitorStatics_CreateForRenderMonitoringWithCategoryAndDeviceRole, _
                  Category, _
                  Role, _
                  VarPtr(pIAudioStateMonitor)) = S_OK Then
            If pIAudioStateMonitor <> 0& Then
                Set Ret = New AudioStateMonitor
                Ret.Ifc = pIAudioStateMonitor
            End If
        End If
    End If
    Set CreateForRenderMonitoringWithCategoryAndDeviceRole = Ret
End Function

Public Function CreateForRenderMonitoringWithCategoryAndDeviceId(ByVal Category As AudioRenderCategory, _
                                                                 ByVal DeviceId As String) As AudioStateMonitor
    Dim Ret As AudioStateMonitor
    If m_pIAudioStateMonitorStatics <> 0& Then
        Dim pIAudioStateMonitor As Long
        Dim hDeviceId As Long
        hDeviceId = CreateWindowsString(DeviceId)
        If Invoke(m_pIAudioStateMonitorStatics, _
                  IAudioStateMonitorStatics_CreateForRenderMonitoringWithCategoryAndDeviceId, _
                  Category, _
                  hDeviceId, _
                  VarPtr(pIAudioStateMonitor)) = S_OK Then
            If pIAudioStateMonitor <> 0& Then
                Set Ret = New AudioStateMonitor
                Ret.Ifc = pIAudioStateMonitor
            End If
        End If
        Call DeleteWindowsString(hDeviceId)
    End If
    Set CreateForRenderMonitoringWithCategoryAndDeviceId = Ret
End Function

Public Function CreateForCaptureMonitoring() As AudioStateMonitor
    Dim Ret As AudioStateMonitor
    If m_pIAudioStateMonitorStatics <> 0& Then
        Dim pIAudioStateMonitor As Long
        If Invoke(m_pIAudioStateMonitorStatics, _
                  IAudioStateMonitorStatics_CreateForCaptureMonitoring, _
                  VarPtr(pIAudioStateMonitor)) = S_OK Then
            If pIAudioStateMonitor <> 0& Then
                Set Ret = New AudioStateMonitor
                Ret.Ifc = pIAudioStateMonitor
            End If
        End If
    End If
    Set CreateForCaptureMonitoring = Ret
End Function

Public Function CreateForCaptureMonitoringWithCategory(ByVal Category As MediaCategory) As AudioStateMonitor
    Dim Ret As AudioStateMonitor
    If m_pIAudioStateMonitorStatics <> 0& Then
        Dim pIAudioStateMonitor As Long
        If Invoke(m_pIAudioStateMonitorStatics, _
                  IAudioStateMonitorStatics_CreateForCaptureMonitoringWithCategory, _
                  Category, _
                  VarPtr(pIAudioStateMonitor)) = S_OK Then
            If pIAudioStateMonitor <> 0& Then
                Set Ret = New AudioStateMonitor
                Ret.Ifc = pIAudioStateMonitor
            End If
        End If
    End If
    Set CreateForCaptureMonitoringWithCategory = Ret
End Function

Public Function CreateForCaptureMonitoringWithCategoryAndDeviceRole(ByVal Category As MediaCategory, _
                                                                    ByVal Role As AudioDeviceRole) As AudioStateMonitor
    Dim Ret As AudioStateMonitor
    If m_pIAudioStateMonitorStatics <> 0& Then
        Dim pIAudioStateMonitor As Long
        If Invoke(m_pIAudioStateMonitorStatics, _
                  IAudioStateMonitorStatics_CreateForCaptureMonitoringWithCategoryAndDeviceRole, _
                  Category, _
                  Role, _
                  VarPtr(pIAudioStateMonitor)) = S_OK Then
            If pIAudioStateMonitor <> 0& Then
                Set Ret = New AudioStateMonitor
                Ret.Ifc = pIAudioStateMonitor
            End If
        End If
    End If
    Set CreateForCaptureMonitoringWithCategoryAndDeviceRole = Ret
End Function

Public Function CreateForCaptureMonitoringWithCategoryAndDeviceId(ByVal Category As MediaCategory, _
                                                                  ByVal DeviceId As String) As AudioStateMonitor
    Dim Ret As AudioStateMonitor
    If m_pIAudioStateMonitorStatics <> 0& Then
        Dim pIAudioStateMonitor As Long
        Dim hDeviceId As Long
        hDeviceId = CreateWindowsString(DeviceId)
        If Invoke(m_pIAudioStateMonitorStatics, _
                  IAudioStateMonitorStatics_CreateForCaptureMonitoringWithCategoryAndDeviceId, _
                  Category, _
                  hDeviceId, _
                  VarPtr(pIAudioStateMonitor)) = S_OK Then
            If pIAudioStateMonitor <> 0& Then
                Set Ret = New AudioStateMonitor
                Ret.Ifc = pIAudioStateMonitor
            End If
        End If
        Call DeleteWindowsString(hDeviceId)
    End If
    Set CreateForCaptureMonitoringWithCategoryAndDeviceId = Ret
End Function
                        
' IAudioStateMonitor
Public Function AddSoundLevelChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pIAudioStateMonitor <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pIAudioStateMonitor, _
                  IAudioStateMonitor_AddSoundLevelChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddSoundLevelChanged = Ret
End Function

Public Function RemoveSoundLevelChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pIAudioStateMonitor <> 0& And token <> 0@ Then
        If Invoke(m_pIAudioStateMonitor, _
                  IAudioStateMonitor_RemoveSoundLevelChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveSoundLevelChanged = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


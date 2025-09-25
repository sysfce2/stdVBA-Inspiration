VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppResourceGroupInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_AppResourceGroupInfo As String = "Windows.System.AppResourceGroupInfo"
Private Const IID_IAppResourceGroupInfo As String = "{b913f77a-e807-49f4-845e-7b8bdcfe8ee7}"
Private Const IID_IAppResourceGroupInfo2 As String = "{ee9b236d-d305-4d6b-92f7-6afdad72dedc}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IAppResourceGroupInfo
    IAppResourceGroupInfo_GetInstanceId = 6
    IAppResourceGroupInfo_GetIsShared = 7
'    IAppResourceGroupInfo_GetBackgroundTaskReports = 8
'    IAppResourceGroupInfo_GetMemoryReport = 9
    IAppResourceGroupInfo_GetProcessDiagnosticInfos = 10
'    IAppResourceGroupInfo_GetStateReport = 11
    
    ' IAppResourceGroupInfo2
'    IAppResourceGroupInfo2_StartSuspendAsync = 6
'    IAppResourceGroupInfo2_StartResumeAsync = 7
'    IAppResourceGroupInfo2_StartTerminateAsync = 8
  
End Enum

' ----==== Variablen ====----
Private m_pIAppResourceGroupInfo As Long
Private m_pIAppResourceGroupInfo2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIAppResourceGroupInfo2)
    Call ReleaseIfc(m_pIAppResourceGroupInfo)
End Sub

' ----==== Properties ====----
Public Property Let Ifc(ByVal pIAppResourceGroupInfo As Long)
    Call ReleaseIfc(m_pIAppResourceGroupInfo2)
    Call ReleaseIfc(m_pIAppResourceGroupInfo)
    m_pIAppResourceGroupInfo = pIAppResourceGroupInfo
    Call QueryIfc(m_pIAppResourceGroupInfo, _
                  IID_IAppResourceGroupInfo2, _
                  m_pIAppResourceGroupInfo2)
End Property

Public Property Get Ifc() As Long
    Ifc = m_pIAppResourceGroupInfo
End Property

' IAppResourceGroupInfo
Friend Property Get InstanceId() As GUID
    Dim Ret As GUID
    If m_pIAppResourceGroupInfo <> 0& Then
        Dim Value As GUID
        If Invoke(m_pIAppResourceGroupInfo, _
                  IAppResourceGroupInfo_GetInstanceId, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    InstanceId = Ret
End Property

Public Property Get IsShared() As Boolean
    Dim Ret As Boolean
    If m_pIAppResourceGroupInfo <> 0& Then
        Dim Value As Long
        If Invoke(m_pIAppResourceGroupInfo, _
                  IAppResourceGroupInfo_GetIsShared, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    IsShared = Ret
End Property

' ----==== Functions ====----
' IAppResourceGroupInfo
'    IAppResourceGroupInfo_GetBackgroundTaskReports = 8
'    IAppResourceGroupInfo_GetMemoryReport = 9

Public Function GetProcessDiagnosticInfos() As ProcessDiagnosticInfo
    Dim Ret As ProcessDiagnosticInfo
    If m_pIAppResourceGroupInfo <> 0& Then
        Dim pIProcessDiagnosticInfo As Long
        If Invoke(m_pIAppResourceGroupInfo, _
                  IAppResourceGroupInfo_GetProcessDiagnosticInfos, _
                  VarPtr(pIProcessDiagnosticInfo)) = S_OK Then
            If pIProcessDiagnosticInfo <> 0& Then
                Set Ret = New ProcessDiagnosticInfo
                Ret.Ifc = pIProcessDiagnosticInfo
            End If
        End If
    End If
    Set GetProcessDiagnosticInfos = Ret
End Function

'    IAppResourceGroupInfo_GetStateReport = 11

' IAppResourceGroupInfo2
'    IAppResourceGroupInfo2_StartSuspendAsync = 6
'    IAppResourceGroupInfo2_StartResumeAsync = 7
'    IAppResourceGroupInfo2_StartTerminateAsync = 8

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



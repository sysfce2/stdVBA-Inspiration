VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SystemDiagnosticInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_Diagnostics_SystemDiagnosticInfo As String = "Windows.System.Diagnostics.SystemDiagnosticInfo"
Private Const IID_ISystemDiagnosticInfo As String = "{a290fe05-dff3-407f-9a1b-0b2b317ca800}"
Private Const IID_ISystemDiagnosticInfoStatics As String = "{d404ac21-fc7d-45f0-9a3f-39203aed9f7e}"
Private Const IID_ISystemDiagnosticInfoStatics2 As String = "{79ded189-6af9-4da9-a422-15f73255b3eb}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ISystemDiagnosticInfoStatics
    ISystemDiagnosticInfoStatics_GetForCurrentSystem = 6
    
    ' ISystemDiagnosticInfoStatics2
    ISystemDiagnosticInfoStatics2_IsArchitectureSupported = 6
    ISystemDiagnosticInfoStatics2_GetPreferredArchitecture = 7
    
    ' ISystemDiagnosticInfo
    ISystemDiagnosticInfo_GetMemoryUsage = 6
    ISystemDiagnosticInfo_GetCpuUsage = 7
    
End Enum

' ----==== Variablen ====----
Private m_pISystemDiagnosticInfo As Long
Private m_pISystemDiagnosticInfoStatics As Long
Private m_pISystemDiagnosticInfoStatics2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_System_Diagnostics_SystemDiagnosticInfo, _
                            IID_ISystemDiagnosticInfoStatics, _
                            m_pISystemDiagnosticInfoStatics) Then
        Call QueryIfc(m_pISystemDiagnosticInfoStatics, _
                      IID_ISystemDiagnosticInfoStatics2, _
                      m_pISystemDiagnosticInfoStatics2)
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pISystemDiagnosticInfoStatics2)
    Call ReleaseIfc(m_pISystemDiagnosticInfoStatics)
    Call ReleaseIfc(m_pISystemDiagnosticInfo)
End Sub

' ----==== Properties ====----
Public Property Let Ifc(ByVal pISystemDiagnosticInfo As Long)
    Call ReleaseIfc(m_pISystemDiagnosticInfo)
    m_pISystemDiagnosticInfo = pISystemDiagnosticInfo
End Property

Public Property Get Ifc() As Long
    Ifc = m_pISystemDiagnosticInfo
End Property

' ISystemDiagnosticInfoStatics2
Public Property Get PreferredArchitecture() As ProcessorArchitecture
    Dim Ret As ProcessorArchitecture
    If m_pISystemDiagnosticInfoStatics2 <> 0& Then
        Dim Value As Long
        If Invoke(m_pISystemDiagnosticInfoStatics2, _
                  ISystemDiagnosticInfoStatics2_GetPreferredArchitecture, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    PreferredArchitecture = Ret
End Property

' ISystemDiagnosticInfo
Public Property Get MemoryUsage() As SystemMemoryUsage
    Dim Ret As SystemMemoryUsage
    If m_pISystemDiagnosticInfo <> 0& Then
        Dim pISystemMemoryUsage As Long
        If Invoke(m_pISystemDiagnosticInfo, _
                  ISystemDiagnosticInfo_GetMemoryUsage, _
                  VarPtr(pISystemMemoryUsage)) = S_OK Then
            If pISystemMemoryUsage <> 0& Then
                Set Ret = New SystemMemoryUsage
                Ret.Ifc = pISystemMemoryUsage
            End If
        End If
    End If
    Set MemoryUsage = Ret
End Property

Public Property Get CpuUsage() As SystemCpuUsage
    Dim Ret As SystemCpuUsage
    If m_pISystemDiagnosticInfo <> 0& Then
        Dim pISystemCpuUsage As Long
        If Invoke(m_pISystemDiagnosticInfo, _
                  ISystemDiagnosticInfo_GetCpuUsage, _
                  VarPtr(pISystemCpuUsage)) = S_OK Then
            If pISystemCpuUsage <> 0& Then
                Set Ret = New SystemCpuUsage
                Ret.Ifc = pISystemCpuUsage
            End If
        End If
    End If
    Set CpuUsage = Ret
End Property

' ----==== Functions ====----
' ISystemDiagnosticInfoStatics
Public Function GetForCurrentSystem() As SystemDiagnosticInfo
    Dim Ret As SystemDiagnosticInfo
    If m_pISystemDiagnosticInfoStatics <> 0& Then
        Dim pISystemDiagnosticInfo As Long
        If Invoke(m_pISystemDiagnosticInfoStatics, _
                  ISystemDiagnosticInfoStatics_GetForCurrentSystem, _
                  VarPtr(pISystemDiagnosticInfo)) = S_OK Then
            If pISystemDiagnosticInfo <> 0& Then
                Set Ret = New SystemDiagnosticInfo
                Ret.Ifc = pISystemDiagnosticInfo
            End If
        End If
    End If
    Set GetForCurrentSystem = Ret
End Function

' ISystemDiagnosticInfoStatics2
Public Function IsArchitectureSupported(ByVal Architecture As ProcessorArchitecture) As Boolean
    Dim Ret As Boolean
    If m_pISystemDiagnosticInfoStatics2 <> 0& Then
        Dim Value As Long
        If Invoke(m_pISystemDiagnosticInfoStatics2, _
                  ISystemDiagnosticInfoStatics2_IsArchitectureSupported, _
                  Architecture, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    IsArchitectureSupported = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function

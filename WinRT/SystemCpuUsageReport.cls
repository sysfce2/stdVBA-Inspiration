VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SystemCpuUsageReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_Diagnostics_SystemCpuUsageReport As String = "Windows.System.Diagnostics.SystemCpuUsageReport"
Private Const IID_ISystemCpuUsageReport As String = "{2c26d0b2-9483-4f62-ab57-82b29d9719b8}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ISystemCpuUsageReport
    ISystemCpuUsageReport_GetKernelTime = 6
    ISystemCpuUsageReport_GetUserTime = 7
    ISystemCpuUsageReport_GetIdleTime = 8
    
End Enum

' ----==== Variablen ====----
Private m_pISystemCpuUsageReport As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pISystemCpuUsageReport)
End Sub

' ----==== Properties ====----
Public Property Let Ifc(ByVal pISystemCpuUsageReport As Long)
    Call ReleaseIfc(m_pISystemCpuUsageReport)
    m_pISystemCpuUsageReport = pISystemCpuUsageReport
End Property

Public Property Get Ifc() As Long
    Ifc = m_pISystemCpuUsageReport
End Property

' ISystemCpuUsageReport
Public Property Get KernelTime() As TimeSpan
    Dim Ret As TimeSpan
    If m_pISystemCpuUsageReport <> 0& Then
        Dim Value As Currency
        If Invoke(m_pISystemCpuUsageReport, _
                  ISystemCpuUsageReport_GetKernelTime, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New TimeSpan
            Ret.TimeSpanValue = Value
        End If
    End If
    Set KernelTime = Ret
End Property

Public Property Get UserTime() As TimeSpan
    Dim Ret As TimeSpan
    If m_pISystemCpuUsageReport <> 0& Then
        Dim Value As Currency
        If Invoke(m_pISystemCpuUsageReport, _
                  ISystemCpuUsageReport_GetUserTime, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New TimeSpan
            Ret.TimeSpanValue = Value
        End If
    End If
    Set UserTime = Ret
End Property

Public Property Get IdleTime() As TimeSpan
    Dim Ret As TimeSpan
    If m_pISystemCpuUsageReport <> 0& Then
        Dim Value As Currency
        If Invoke(m_pISystemCpuUsageReport, _
                  ISystemCpuUsageReport_GetIdleTime, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New TimeSpan
            Ret.TimeSpanValue = Value
        End If
    End If
    Set IdleTime = Ret
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



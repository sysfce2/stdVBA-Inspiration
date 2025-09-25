VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProcessCpuUsageReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_Diagnostics_ProcessCpuUsageReport As String = "Windows.System.Diagnostics.ProcessCpuUsageReport"
Private Const IID_IProcessCpuUsageReport As String = "{8a6d9cac-3987-4e2f-a119-6b5fa214f1b4}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IProcessCpuUsageReport
    IProcessCpuUsageReport_GetKernelTime = 6
    IProcessCpuUsageReport_GetUserTime = 7
    
End Enum

' ----==== Variablen ====----
Private m_pIProcessCpuUsageReport As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIProcessCpuUsageReport)
End Sub

' ----==== Properties ====----
Public Property Let Ifc(ByVal pIProcessCpuUsageReport As Long)
    Call ReleaseIfc(m_pIProcessCpuUsageReport)
    m_pIProcessCpuUsageReport = pIProcessCpuUsageReport
End Property

Public Property Get Ifc() As Long
    Ifc = m_pIProcessCpuUsageReport
End Property

' IProcessCpuUsageReport
Public Property Get KernelTime() As TimeSpan
    Dim Ret As TimeSpan
    If m_pIProcessCpuUsageReport <> 0& Then
        Dim Value As Currency
        If Invoke(m_pIProcessCpuUsageReport, _
                  IProcessCpuUsageReport_GetKernelTime, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New TimeSpan
            Ret.TimeSpanValue = Value
        End If
    End If
    Set KernelTime = Ret
End Property

Public Property Get UserTime() As TimeSpan
    Dim Ret As TimeSpan
    If m_pIProcessCpuUsageReport <> 0& Then
        Dim Value As Currency
        If Invoke(m_pIProcessCpuUsageReport, _
                  IProcessCpuUsageReport_GetUserTime, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New TimeSpan
            Ret.TimeSpanValue = Value
        End If
    End If
    Set UserTime = Ret
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



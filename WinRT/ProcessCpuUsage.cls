VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProcessCpuUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_Diagnostics_ProcessCpuUsage As String = "Windows.System.Diagnostics.ProcessCpuUsage"
Private Const IID_IProcessCpuUsage As String = "{0bbb2472-c8bf-423a-a810-b559ae4354e2}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IProcessCpuUsage
    IProcessCpuUsage_GetReport = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIProcessCpuUsage As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIProcessCpuUsage)
End Sub

' ----==== Properties ====----
Public Property Let Ifc(ByVal pIProcessCpuUsage As Long)
    Call ReleaseIfc(m_pIProcessCpuUsage)
    m_pIProcessCpuUsage = pIProcessCpuUsage
End Property

Public Property Get Ifc() As Long
    Ifc = m_pIProcessCpuUsage
End Property

' ----==== Functions ====----
' IProcessCpuUsage
Public Function GetReport() As ProcessCpuUsageReport
    Dim Ret As ProcessCpuUsageReport
    If m_pIProcessCpuUsage <> 0& Then
        Dim pIProcessCpuUsageReport As Long
        If Invoke(m_pIProcessCpuUsage, _
                  IProcessCpuUsage_GetReport, _
                  VarPtr(pIProcessCpuUsageReport)) = S_OK Then
            If pIProcessCpuUsageReport <> 0& Then
                Set Ret = New ProcessCpuUsageReport
                Ret.Ifc = pIProcessCpuUsageReport
            End If
        End If
    End If
    Set GetReport = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



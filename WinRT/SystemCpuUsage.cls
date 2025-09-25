VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SystemCpuUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_Diagnostics_SystemCpuUsage As String = "Windows.System.Diagnostics.SystemCpuUsage"
Private Const IID_ISystemCpuUsage As String = "{6037b3ac-02d6-4234-8362-7fe3adc81f5f}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ISystemCpuUsage
    ISystemCpuUsage_GetReport = 6
    
End Enum

' ----==== Variablen ====----
Private m_pISystemCpuUsage As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pISystemCpuUsage)
End Sub

' ----==== Properties ====----
Public Property Let Ifc(ByVal pISystemCpuUsage As Long)
    Call ReleaseIfc(m_pISystemCpuUsage)
    m_pISystemCpuUsage = pISystemCpuUsage
End Property

Public Property Get Ifc() As Long
    Ifc = m_pISystemCpuUsage
End Property

' ----==== Functions ====----
' ISystemCpuUsage
Public Function GetReport() As SystemCpuUsageReport
    Dim Ret As SystemCpuUsageReport
    If m_pISystemCpuUsage <> 0& Then
        Dim pISystemCpuUsageReport As Long
        If Invoke(m_pISystemCpuUsage, _
                  ISystemCpuUsage_GetReport, _
                  VarPtr(pISystemCpuUsageReport)) = S_OK Then
            If pISystemCpuUsageReport <> 0& Then
                Set Ret = New SystemCpuUsageReport
                Ret.Ifc = pISystemCpuUsageReport
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



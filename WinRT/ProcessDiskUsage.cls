VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProcessDiskUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 01/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_Diagnostics_ProcessDiskUsage As String = "Windows.System.Diagnostics.ProcessDiskUsage"
Private Const IID_IProcessDiskUsage As String = "{5ad78bfd-7e51-4e53-bfaa-5a6ee1aabbf8}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IProcessDiskUsage
    IProcessDiskUsage_GetReport = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIProcessDiskUsage As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIProcessDiskUsage)
End Sub

' ----==== Properties ====----
Public Property Let Ifc(ByVal pIProcessDiskUsage As Long)
    Call ReleaseIfc(m_pIProcessDiskUsage)
    m_pIProcessDiskUsage = pIProcessDiskUsage
End Property

Public Property Get Ifc() As Long
    Ifc = m_pIProcessDiskUsage
End Property

' ----==== Functions ====----
' IProcessDiskUsage
Public Function GetReport() As ProcessDiskUsageReport
    Dim Ret As ProcessDiskUsageReport
    If m_pIProcessDiskUsage <> 0& Then
        Dim pIProcessDiskUsageReport As Long
        If Invoke(m_pIProcessDiskUsage, _
                  IProcessDiskUsage_GetReport, _
                  VarPtr(pIProcessDiskUsageReport)) = S_OK Then
            If pIProcessDiskUsageReport <> 0& Then
                Set Ret = New ProcessDiskUsageReport
                Ret.Ifc = pIProcessDiskUsageReport
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



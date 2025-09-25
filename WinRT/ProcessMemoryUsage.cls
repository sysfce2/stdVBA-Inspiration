VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProcessMemoryUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_Diagnostics_ProcessMemoryUsage As String = "Windows.System.Diagnostics.ProcessMemoryUsage"
Private Const IID_IProcessMemoryUsage As String = "{f50b229b-827c-42b7-b07c-0e32627e6b3e}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IProcessMemoryUsage
    IProcessMemoryUsage_GetReport = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIProcessMemoryUsage As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIProcessMemoryUsage)
End Sub

' ----==== Properties ====----
Public Property Let Ifc(ByVal pIProcessMemoryUsage As Long)
    Call ReleaseIfc(m_pIProcessMemoryUsage)
    m_pIProcessMemoryUsage = pIProcessMemoryUsage
End Property

Public Property Get Ifc() As Long
    Ifc = m_pIProcessMemoryUsage
End Property

' ----==== Functions ====----
' IProcessMemoryUsage
Public Function GetReport() As ProcessMemoryUsageReport
    Dim Ret As ProcessMemoryUsageReport
    If m_pIProcessMemoryUsage <> 0& Then
        Dim pIProcessMemoryUsageReport As Long
        If Invoke(m_pIProcessMemoryUsage, _
                  IProcessMemoryUsage_GetReport, _
                  VarPtr(pIProcessMemoryUsageReport)) = S_OK Then
            If pIProcessMemoryUsageReport <> 0& Then
                Set Ret = New ProcessMemoryUsageReport
                Ret.Ifc = pIProcessMemoryUsageReport
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



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SystemMemoryUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_Diagnostics_SystemMemoryUsage As String = "Windows.System.Diagnostics.SystemMemoryUsage"
Private Const IID_ISystemMemoryUsage As String = "{17ffc595-1702-49cf-aa27-2f0a32591404}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ISystemMemoryUsage
    ISystemMemoryUsage_GetReport = 6
    
End Enum

' ----==== Variablen ====----
Private m_pISystemMemoryUsage As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pISystemMemoryUsage)
End Sub

' ----==== Properties ====----
Public Property Let Ifc(ByVal pISystemMemoryUsage As Long)
    Call ReleaseIfc(m_pISystemMemoryUsage)
    m_pISystemMemoryUsage = pISystemMemoryUsage
End Property

Public Property Get Ifc() As Long
    Ifc = m_pISystemMemoryUsage
End Property

' ----==== Functions ====----
' ISystemMemoryUsage
Public Function GetReport() As SystemMemoryUsageReport
    Dim Ret As SystemMemoryUsageReport
    If m_pISystemMemoryUsage <> 0& Then
        Dim pISystemMemoryUsageReport As Long
        If Invoke(m_pISystemMemoryUsage, _
                  ISystemMemoryUsage_GetReport, _
                  VarPtr(pISystemMemoryUsageReport)) = S_OK Then
            If pISystemMemoryUsageReport <> 0& Then
                Set Ret = New SystemMemoryUsageReport
                Ret.Ifc = pISystemMemoryUsageReport
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



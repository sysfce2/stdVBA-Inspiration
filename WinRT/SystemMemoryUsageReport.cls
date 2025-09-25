VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SystemMemoryUsageReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_Diagnostics_ISystemMemoryUsageReport As String = "Windows.System.Diagnostics.ISystemMemoryUsageReport"
Private Const IID_ISystemMemoryUsageReport As String = "{38663c87-2a9f-403a-bd19-2cf3e8169500}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ISystemMemoryUsageReport
    ISystemMemoryUsageReport_GetTotalPhysicalSizeInBytes = 6
    ISystemMemoryUsageReport_GetAvailableSizeInBytes = 7
    ISystemMemoryUsageReport_GetCommittedSizeInBytes = 8

End Enum

' ----==== Variablen ====----
Private m_pISystemMemoryUsageReport As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pISystemMemoryUsageReport)
End Sub

' ----==== Properties ====----
Public Property Let Ifc(ByVal pISystemMemoryUsageReport As Long)
    Call ReleaseIfc(m_pISystemMemoryUsageReport)
    m_pISystemMemoryUsageReport = pISystemMemoryUsageReport
End Property

Public Property Get Ifc() As Long
    Ifc = m_pISystemMemoryUsageReport
End Property

' ISystemMemoryUsageReport
Public Property Get TotalPhysicalSizeInBytes() As Currency
    Dim Ret As Currency
    If m_pISystemMemoryUsageReport <> 0& Then
        Dim Value As Currency
        If Invoke(m_pISystemMemoryUsageReport, _
                  ISystemMemoryUsageReport_GetTotalPhysicalSizeInBytes, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value * 10000#
        End If
    End If
    TotalPhysicalSizeInBytes = Ret
End Property

Public Property Get AvailableSizeInBytes() As Currency
    Dim Ret As Currency
    If m_pISystemMemoryUsageReport <> 0& Then
        Dim Value As Currency
        If Invoke(m_pISystemMemoryUsageReport, _
                  ISystemMemoryUsageReport_GetAvailableSizeInBytes, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value * 10000#
        End If
    End If
    AvailableSizeInBytes = Ret
End Property

Public Property Get CommittedSizeInBytes() As Currency
    Dim Ret As Currency
    If m_pISystemMemoryUsageReport <> 0& Then
        Dim Value As Currency
        If Invoke(m_pISystemMemoryUsageReport, _
                  ISystemMemoryUsageReport_GetCommittedSizeInBytes, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value * 10000#
        End If
    End If
    CommittedSizeInBytes = Ret
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




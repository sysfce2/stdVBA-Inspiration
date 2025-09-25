VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProcessDiskUsageReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_Diagnostics_ProcessDiskUsageReport As String = "Windows.System.Diagnostics.ProcessDiskUsageReport"
Private Const IID_IProcessDiskUsageReport As String = "{401627fd-535d-4c1f-81b8-da54e1be635e}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IProcessDiskUsageReport
    IProcessDiskUsageReport_GetReadOperationCount = 6
    IProcessDiskUsageReport_GetWriteOperationCount = 7
    IProcessDiskUsageReport_GetOtherOperationCount = 8
    IProcessDiskUsageReport_GetBytesReadCount = 9
    IProcessDiskUsageReport_GetBytesWrittenCount = 10
    IProcessDiskUsageReport_GetOtherBytesCount = 11

End Enum

' ----==== Variablen ====----
Private m_pIProcessDiskUsageReport As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIProcessDiskUsageReport)
End Sub

' ----==== Properties ====----
Public Property Let Ifc(ByVal pIProcessDiskUsageReport As Long)
    Call ReleaseIfc(m_pIProcessDiskUsageReport)
    m_pIProcessDiskUsageReport = pIProcessDiskUsageReport
End Property

Public Property Get Ifc() As Long
    Ifc = m_pIProcessDiskUsageReport
End Property

' IProcessDiskUsageReport
Public Property Get ReadOperationCount() As Currency
    Dim Ret As Currency
    If m_pIProcessDiskUsageReport <> 0& Then
        Dim Value As Currency
        If Invoke(m_pIProcessDiskUsageReport, _
                  IProcessDiskUsageReport_GetReadOperationCount, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value * 10000#
        End If
    End If
    ReadOperationCount = Ret
End Property

Public Property Get WriteOperationCount() As Currency
    Dim Ret As Currency
    If m_pIProcessDiskUsageReport <> 0& Then
        Dim Value As Currency
        If Invoke(m_pIProcessDiskUsageReport, _
                  IProcessDiskUsageReport_GetWriteOperationCount, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value * 10000#
        End If
    End If
    WriteOperationCount = Ret
End Property

Public Property Get OtherOperationCount() As Currency
    Dim Ret As Currency
    If m_pIProcessDiskUsageReport <> 0& Then
        Dim Value As Currency
        If Invoke(m_pIProcessDiskUsageReport, _
                  IProcessDiskUsageReport_GetOtherOperationCount, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value * 10000#
        End If
    End If
    OtherOperationCount = Ret
End Property

Public Property Get BytesReadCount() As Currency
    Dim Ret As Currency
    If m_pIProcessDiskUsageReport <> 0& Then
        Dim Value As Currency
        If Invoke(m_pIProcessDiskUsageReport, _
                  IProcessDiskUsageReport_GetBytesReadCount, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value * 10000#
        End If
    End If
    BytesReadCount = Ret
End Property

Public Property Get BytesWrittenCount() As Currency
    Dim Ret As Currency
    If m_pIProcessDiskUsageReport <> 0& Then
        Dim Value As Currency
        If Invoke(m_pIProcessDiskUsageReport, _
                  IProcessDiskUsageReport_GetBytesWrittenCount, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value * 10000#
        End If
    End If
    BytesWrittenCount = Ret
End Property

Public Property Get OtherBytesCount() As Currency
    Dim Ret As Currency
    If m_pIProcessDiskUsageReport <> 0& Then
        Dim Value As Currency
        If Invoke(m_pIProcessDiskUsageReport, _
                  IProcessDiskUsageReport_GetOtherBytesCount, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value * 10000#
        End If
    End If
    OtherBytesCount = Ret
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




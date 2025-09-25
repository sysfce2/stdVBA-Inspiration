VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProcessMemoryUsageReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 01/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_Diagnostics_ProcessMemoryUsageReport As String = "Windows.System.Diagnostics.ProcessMemoryUsageReport"
Private Const IID_IProcessMemoryUsageReport As String = "{c2c77cba-1951-4685-8532-7e749ecf8eeb}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IProcessMemoryUsageReport
    IProcessMemoryUsageReport_GetNonPagedPoolSizeInBytes = 6
    IProcessMemoryUsageReport_GetPageFaultCount = 7
    IProcessMemoryUsageReport_GetPageFileSizeInBytes = 8
    IProcessMemoryUsageReport_GetPagedPoolSizeInBytes = 9
    IProcessMemoryUsageReport_GetPeakNonPagedPoolSizeInBytes = 10
    IProcessMemoryUsageReport_GetPeakPageFileSizeInBytes = 11
    IProcessMemoryUsageReport_GetPeakPagedPoolSizeInBytes = 12
    IProcessMemoryUsageReport_GetPeakVirtualMemorySizeInBytes = 13
    IProcessMemoryUsageReport_GetPeakWorkingSetSizeInBytes = 14
    IProcessMemoryUsageReport_GetPrivatePageCount = 15
    IProcessMemoryUsageReport_GetVirtualMemorySizeInBytes = 16
    IProcessMemoryUsageReport_GetWorkingSetSizeInBytes = 17

End Enum

' ----==== Variablen ====----
Private m_pIProcessMemoryUsageReport As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIProcessMemoryUsageReport)
End Sub

' ----==== Properties ====----
Public Property Let Ifc(ByVal pIProcessMemoryUsageReport As Long)
    Call ReleaseIfc(m_pIProcessMemoryUsageReport)
    m_pIProcessMemoryUsageReport = pIProcessMemoryUsageReport
End Property

Public Property Get Ifc() As Long
    Ifc = m_pIProcessMemoryUsageReport
End Property

Public Property Get NonPagedPoolSizeInBytes() As Currency
    Dim Ret As Currency
    If m_pIProcessMemoryUsageReport <> 0& Then
        Dim Value As Currency
        If Invoke(m_pIProcessMemoryUsageReport, _
                  IProcessMemoryUsageReport_GetNonPagedPoolSizeInBytes, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value * 10000#
        End If
    End If
    NonPagedPoolSizeInBytes = Ret
End Property

Public Property Get PageFaultCount() As Long
    Dim Ret As Long
    If m_pIProcessMemoryUsageReport <> 0& Then
        Dim Value As Long
        If Invoke(m_pIProcessMemoryUsageReport, _
                  IProcessMemoryUsageReport_GetPageFaultCount, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    PageFaultCount = Ret
End Property

Public Property Get PageFileSizeInBytes() As Currency
    Dim Ret As Currency
    If m_pIProcessMemoryUsageReport <> 0& Then
        Dim Value As Currency
        If Invoke(m_pIProcessMemoryUsageReport, _
                  IProcessMemoryUsageReport_GetPageFileSizeInBytes, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value * 10000#
        End If
    End If
    PageFileSizeInBytes = Ret
End Property

Public Property Get PagedPoolSizeInBytes() As Currency
    Dim Ret As Currency
    If m_pIProcessMemoryUsageReport <> 0& Then
        Dim Value As Currency
        If Invoke(m_pIProcessMemoryUsageReport, _
                  IProcessMemoryUsageReport_GetPagedPoolSizeInBytes, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value * 10000#
        End If
    End If
    PagedPoolSizeInBytes = Ret
End Property

Public Property Get PeakNonPagedPoolSizeInBytes() As Currency
    Dim Ret As Currency
    If m_pIProcessMemoryUsageReport <> 0& Then
        Dim Value As Currency
        If Invoke(m_pIProcessMemoryUsageReport, _
                  IProcessMemoryUsageReport_GetPeakNonPagedPoolSizeInBytes, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value * 10000#
        End If
    End If
    PeakNonPagedPoolSizeInBytes = Ret
End Property

Public Property Get PeakPageFileSizeInBytes() As Currency
    Dim Ret As Currency
    If m_pIProcessMemoryUsageReport <> 0& Then
        Dim Value As Currency
        If Invoke(m_pIProcessMemoryUsageReport, _
                  IProcessMemoryUsageReport_GetPeakPageFileSizeInBytes, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value * 10000#
        End If
    End If
    PeakPageFileSizeInBytes = Ret
End Property

Public Property Get PeakPagedPoolSizeInBytes() As Currency
    Dim Ret As Currency
    If m_pIProcessMemoryUsageReport <> 0& Then
        Dim Value As Currency
        If Invoke(m_pIProcessMemoryUsageReport, _
                  IProcessMemoryUsageReport_GetPeakPagedPoolSizeInBytes, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value * 10000#
        End If
    End If
    PeakPagedPoolSizeInBytes = Ret
End Property

Public Property Get PeakVirtualMemorySizeInBytes() As Currency
    Dim Ret As Currency
    If m_pIProcessMemoryUsageReport <> 0& Then
        Dim Value As Currency
        If Invoke(m_pIProcessMemoryUsageReport, _
                  IProcessMemoryUsageReport_GetPeakVirtualMemorySizeInBytes, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value * 10000#
        End If
    End If
    PeakVirtualMemorySizeInBytes = Ret
End Property

Public Property Get PeakWorkingSetSizeInBytes() As Currency
    Dim Ret As Currency
    If m_pIProcessMemoryUsageReport <> 0& Then
        Dim Value As Currency
        If Invoke(m_pIProcessMemoryUsageReport, _
                  IProcessMemoryUsageReport_GetPeakWorkingSetSizeInBytes, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value * 10000#
        End If
    End If
    PeakWorkingSetSizeInBytes = Ret
End Property

Public Property Get PrivatePageCount() As Currency
    Dim Ret As Currency
    If m_pIProcessMemoryUsageReport <> 0& Then
        Dim Value As Currency
        If Invoke(m_pIProcessMemoryUsageReport, _
                  IProcessMemoryUsageReport_GetPrivatePageCount, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value * 10000#
        End If
    End If
    PrivatePageCount = Ret
End Property

Public Property Get VirtualMemorySizeInBytes() As Currency
    Dim Ret As Currency
    If m_pIProcessMemoryUsageReport <> 0& Then
        Dim Value As Currency
        If Invoke(m_pIProcessMemoryUsageReport, _
                  IProcessMemoryUsageReport_GetVirtualMemorySizeInBytes, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value * 10000#
        End If
    End If
    VirtualMemorySizeInBytes = Ret
End Property

Public Property Get WorkingSetSizeInBytes() As Currency
    Dim Ret As Currency
    If m_pIProcessMemoryUsageReport <> 0& Then
        Dim Value As Currency
        If Invoke(m_pIProcessMemoryUsageReport, _
                  IProcessMemoryUsageReport_GetWorkingSetSizeInBytes, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value * 10000#
        End If
    End If
    WorkingSetSizeInBytes = Ret
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



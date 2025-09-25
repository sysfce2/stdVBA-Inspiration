VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProcessDiagnosticInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_Diagnostics_ProcessDiagnosticInfo As String = "Windows.System.Diagnostics.ProcessDiagnosticInfo"
Private Const IID_IProcessDiagnosticInfo As String = "{e830b04b-300e-4ee6-a0ab-5b5f5231b434}"
Private Const IID_IProcessDiagnosticInfo2 As String = "{9558cb1a-3d0b-49ec-ab70-4f7a112805de}"
Private Const IID_IProcessDiagnosticInfoStatics As String = "{2f41b260-b49f-428c-aa0e-84744f49ca95}"
Private Const IID_IProcessDiagnosticInfoStatics2 As String = "{4a869897-9899-4a44-a29b-091663be09b6}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IProcessDiagnosticInfoStatics
    IProcessDiagnosticInfoStatics_GetForProcesses = 6
    IProcessDiagnosticInfoStatics_GetForCurrentProcess = 7
    
    ' IProcessDiagnosticInfoStatics2
    IProcessDiagnosticInfoStatics2_TryGetForProcessId = 6
    
    ' IProcessDiagnosticInfo
    IProcessDiagnosticInfo_GetProcessId = 6
    IProcessDiagnosticInfo_GetExecutableFileName = 7
    IProcessDiagnosticInfo_GetParent = 8
    IProcessDiagnosticInfo_GetProcessStartTime = 9
    IProcessDiagnosticInfo_GetDiskUsage = 10
    IProcessDiagnosticInfo_GetMemoryUsage = 11
    IProcessDiagnosticInfo_GetCpuUsage = 12

    ' IProcessDiagnosticInfo2
    IProcessDiagnosticInfo2_GetAppDiagnosticInfos = 6
    IProcessDiagnosticInfo2_GetIsPackaged = 7

End Enum

' ----==== Variablen ====----
Private m_pIProcessDiagnosticInfo As Long
Private m_pIProcessDiagnosticInfo2 As Long
Private m_pIProcessDiagnosticInfoStatics As Long
Private m_pIProcessDiagnosticInfoStatics2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_System_Diagnostics_ProcessDiagnosticInfo, _
                            IID_IProcessDiagnosticInfoStatics, _
                            m_pIProcessDiagnosticInfoStatics) Then
        Call QueryIfc(m_pIProcessDiagnosticInfoStatics, _
                      IID_IProcessDiagnosticInfoStatics2, _
                      m_pIProcessDiagnosticInfoStatics2)
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIProcessDiagnosticInfo2)
    Call ReleaseIfc(m_pIProcessDiagnosticInfo)
    Call ReleaseIfc(m_pIProcessDiagnosticInfoStatics2)
    Call ReleaseIfc(m_pIProcessDiagnosticInfoStatics)
End Sub

' ----==== Properties ====----
Public Property Let Ifc(ByVal pIProcessDiagnosticInfo As Long)
    Call ReleaseIfc(m_pIProcessDiagnosticInfo2)
    Call ReleaseIfc(m_pIProcessDiagnosticInfo)
    m_pIProcessDiagnosticInfo = pIProcessDiagnosticInfo
    Call QueryIfc(m_pIProcessDiagnosticInfo, _
                  IID_IProcessDiagnosticInfo2, _
                  m_pIProcessDiagnosticInfo2)
End Property

Public Property Get Ifc() As Long
    Ifc = m_pIProcessDiagnosticInfo
End Property

' IProcessDiagnosticInfo
Public Property Get ProcessId() As Long
    Dim Ret As Long
    If m_pIProcessDiagnosticInfo <> 0& Then
        Dim value As Long
        If Invoke(m_pIProcessDiagnosticInfo, _
                  IProcessDiagnosticInfo_GetProcessId, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    ProcessId = Ret
End Property

Public Property Get ExecutableFileName() As String
    Dim Ret As String
    If m_pIProcessDiagnosticInfo <> 0& Then
        Dim value As Long
        If Invoke(m_pIProcessDiagnosticInfo, _
                  IProcessDiagnosticInfo_GetExecutableFileName, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    ExecutableFileName = Ret
End Property

Public Property Get Parent() As ProcessDiagnosticInfo
    Dim Ret As ProcessDiagnosticInfo
    If m_pIProcessDiagnosticInfo <> 0& Then
        Dim pIProcessDiagnosticInfo As Long
        If Invoke(m_pIProcessDiagnosticInfo, _
                  IProcessDiagnosticInfo_GetParent, _
                  VarPtr(pIProcessDiagnosticInfo)) = S_OK Then
            If pIProcessDiagnosticInfo <> 0& Then
                Set Ret = New ProcessDiagnosticInfo
                Ret.Ifc = pIProcessDiagnosticInfo
            End If
        End If
    End If
    Set Parent = Ret
End Property

Public Property Get ProcessStartTime() As DateTime
    Dim Ret As DateTime
    If m_pIProcessDiagnosticInfo <> 0& Then
        Dim value As Currency
        If Invoke(m_pIProcessDiagnosticInfo, _
                  IProcessDiagnosticInfo_GetProcessStartTime, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New DateTime
            Ret.DateTimeValue = value
        End If
    End If
    Set ProcessStartTime = Ret
End Property

Public Property Get DiskUsage() As ProcessDiskUsage
    Dim Ret As ProcessDiskUsage
    If m_pIProcessDiagnosticInfo <> 0& Then
        Dim pIProcessDiskUsage As Long
        If Invoke(m_pIProcessDiagnosticInfo, _
                  IProcessDiagnosticInfo_GetDiskUsage, _
                  VarPtr(pIProcessDiskUsage)) = S_OK Then
            If pIProcessDiskUsage <> 0& Then
                Set Ret = New ProcessDiskUsage
                Ret.Ifc = pIProcessDiskUsage
            End If
        End If
    End If
    Set DiskUsage = Ret
End Property

Public Property Get MemoryUsage() As ProcessMemoryUsage
    Dim Ret As ProcessMemoryUsage
    If m_pIProcessDiagnosticInfo <> 0& Then
        Dim pIProcessMemoryUsage As Long
        If Invoke(m_pIProcessDiagnosticInfo, _
                  IProcessDiagnosticInfo_GetMemoryUsage, _
                  VarPtr(pIProcessMemoryUsage)) = S_OK Then
            If pIProcessMemoryUsage <> 0& Then
                Set Ret = New ProcessMemoryUsage
                Ret.Ifc = pIProcessMemoryUsage
            End If
        End If
    End If
    Set MemoryUsage = Ret
End Property

Public Property Get CpuUsage() As ProcessCpuUsage
    Dim Ret As ProcessCpuUsage
    If m_pIProcessDiagnosticInfo <> 0& Then
        Dim pIProcessCpuUsage As Long
        If Invoke(m_pIProcessDiagnosticInfo, _
                  IProcessDiagnosticInfo_GetCpuUsage, _
                  VarPtr(pIProcessCpuUsage)) = S_OK Then
            If pIProcessCpuUsage <> 0& Then
                Set Ret = New ProcessCpuUsage
                Ret.Ifc = pIProcessCpuUsage
            End If
        End If
    End If
    Set CpuUsage = Ret
End Property

' IProcessDiagnosticInfo2
Public Property Get IsPackaged() As Boolean
    Dim Ret As Boolean
    If m_pIProcessDiagnosticInfo2 <> 0& Then
        Dim value As Long
        If Invoke(m_pIProcessDiagnosticInfo2, _
                  IProcessDiagnosticInfo2_GetIsPackaged, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsPackaged = Ret
End Property

' ----==== Functions ====----
' IProcessDiagnosticInfoStatics
Public Function GetForProcesses() As ReadOnlyList_1 'ReadOnlyList_ProcessDiagnosticInfo
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_ProcessDiagnosticInfo
    If m_pIProcessDiagnosticInfoStatics <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIProcessDiagnosticInfoStatics, _
                  IProcessDiagnosticInfoStatics_GetForProcesses, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_ProcessDiagnosticInfo, pIVectorView) 'ReadOnlyList_ProcessDiagnosticInfo
            End If
        End If
    End If
    Set GetForProcesses = Ret
End Function

Public Function GetForCurrentProcess() As ProcessDiagnosticInfo
    Dim Ret As ProcessDiagnosticInfo
    If m_pIProcessDiagnosticInfoStatics <> 0& Then
        Dim pIProcessDiagnosticInfo As Long
        If Invoke(m_pIProcessDiagnosticInfoStatics, _
                  IProcessDiagnosticInfoStatics_GetForCurrentProcess, _
                  VarPtr(pIProcessDiagnosticInfo)) = S_OK Then
            If pIProcessDiagnosticInfo <> 0& Then
                Set Ret = New ProcessDiagnosticInfo
                Ret.Ifc = pIProcessDiagnosticInfo
            End If
        End If
    End If
    Set GetForCurrentProcess = Ret
End Function
    
' IProcessDiagnosticInfoStatics2
Public Function TryGetForProcessId(ByVal ProcessId As Long) As ProcessDiagnosticInfo
    Dim Ret As ProcessDiagnosticInfo
    If m_pIProcessDiagnosticInfoStatics2 <> 0& Then
        Dim pIProcessDiagnosticInfo As Long
        If Invoke(m_pIProcessDiagnosticInfoStatics2, _
                  IProcessDiagnosticInfoStatics2_TryGetForProcessId, _
                  ProcessId, _
                  VarPtr(pIProcessDiagnosticInfo)) = S_OK Then
            If pIProcessDiagnosticInfo <> 0& Then
                Set Ret = New ProcessDiagnosticInfo
                Ret.Ifc = pIProcessDiagnosticInfo
            End If
        End If
    End If
    Set TryGetForProcessId = Ret
End Function

' IProcessDiagnosticInfo2
Public Function GetAppDiagnosticInfos() As ReadOnlyList_1 'ReadOnlyList_AppDiagnosticInfo
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_AppDiagnosticInfo
    If m_pIProcessDiagnosticInfo2 <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIProcessDiagnosticInfo2, _
                  IProcessDiagnosticInfo2_GetAppDiagnosticInfos, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_AppDiagnosticInfo, pIVectorView) 'ReadOnlyList_AppDiagnosticInfo
            End If
        End If
    End If
    Set GetAppDiagnosticInfos = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


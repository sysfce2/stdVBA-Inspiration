VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Diagnostics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

' Namespace Diagnostics

Option Explicit

' ----==== Enums Windows_System.Diagnostics.xxx ====----

' ----==== Namespace Diagnostics.SystemCpuUsage ====----
Public Property Get SystemCpuUsage() As SystemCpuUsage
    Set SystemCpuUsage = New SystemCpuUsage
End Property

' ----==== Namespace Diagnostics.SystemMemoryUsage ====----
Public Property Get SystemMemoryUsage() As SystemMemoryUsage
    Set SystemMemoryUsage = New SystemMemoryUsage
End Property

' ----==== Namespace Diagnostics.SystemDiagnosticInfo ====----
Public Property Get SystemDiagnosticInfo() As SystemDiagnosticInfo
    Set SystemDiagnosticInfo = New SystemDiagnosticInfo
End Property

' ----==== Namespace Diagnostics.SystemCpuUsageReport ====----
Public Property Get SystemCpuUsageReport() As SystemCpuUsageReport
    Set SystemCpuUsageReport = New SystemCpuUsageReport
End Property

' ----==== Namespace Diagnostics.SystemMemoryUsageReport ====----
Public Property Get SystemMemoryUsageReport() As SystemMemoryUsageReport
    Set SystemMemoryUsageReport = New SystemMemoryUsageReport
End Property

' ----==== Namespace Diagnostics.ProcessCpuUsage ====----
Public Property Get ProcessCpuUsage() As ProcessCpuUsage
    Set ProcessCpuUsage = New ProcessCpuUsage
End Property

' ----==== Namespace Diagnostics.ProcessDiagnosticInfo ====----
Public Property Get ProcessDiagnosticInfo() As ProcessDiagnosticInfo
    Set ProcessDiagnosticInfo = New ProcessDiagnosticInfo
End Property

' ----==== Namespace Diagnostics.ProcessMemoryUsageReport ====----
Public Property Get ProcessMemoryUsageReport() As ProcessMemoryUsageReport
    Set ProcessMemoryUsageReport = New ProcessMemoryUsageReport
End Property

' ----==== Namespace Diagnostics.ProcessCpuUsageReport ====----
Public Property Get ProcessCpuUsageReport() As ProcessCpuUsageReport
    Set ProcessCpuUsageReport = New ProcessCpuUsageReport
End Property

' ----==== Namespace Diagnostics.ProcessMemoryUsage ====----
Public Property Get ProcessMemoryUsage() As ProcessMemoryUsage
    Set ProcessMemoryUsage = New ProcessMemoryUsage
End Property

' ----==== Namespace Diagnostics.ProcessDiskUsage ====----
Public Property Get ProcessDiskUsage() As ProcessDiskUsage
    Set ProcessDiskUsage = New ProcessDiskUsage
End Property

' ----==== Namespace Diagnostics.ProcessDiskUsageReport ====----
Public Property Get ProcessDiskUsageReport() As ProcessDiskUsageReport
    Set ProcessDiskUsageReport = New ProcessDiskUsageReport
End Property


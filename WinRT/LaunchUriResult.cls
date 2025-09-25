VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LaunchUriResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_LaunchUriResult As String = "Windows.System.LaunchUriResult"
Private Const IID_ILaunchUriResult As String = "{ec27a8df-f6d5-45ca-913a-70a40c5c8221}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ILaunchUriResult
    ILaunchUriResult_GetStatus = 6
'    ILaunchUriResult_GetResult = 7
       
End Enum

' ----==== Variablen ====----
Private m_pILaunchUriResult As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pILaunchUriResult)
End Sub

' ----==== Properties ====----
Public Property Let Ifc(ByVal pILaunchUriResult As Long)
    Call ReleaseIfc(m_pILaunchUriResult)
    m_pILaunchUriResult = pILaunchUriResult
End Property

Public Property Get Ifc() As Long
    Ifc = m_pILaunchUriResult
End Property

' ILaunchUriResult
Public Property Get Status() As LaunchUriStatus
    Dim Ret As LaunchUriStatus
    If m_pILaunchUriResult <> 0& Then
        Dim Value As Long
        If Invoke(m_pILaunchUriResult, _
                  ILaunchUriResult_GetStatus, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Status = Ret
End Property

'    ILaunchUriResult_GetResult = 7

' ----==== Functions ====----

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


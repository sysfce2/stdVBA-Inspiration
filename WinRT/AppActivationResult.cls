VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppActivationResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_AppActivationResult As String = "Windows.System.AppActivationResult"
Private Const IID_IAppActivationResult As String = "{6b528900-f46e-4eb0-aa6c-38af557cf9ed}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IAppActivationResult
    IAppActivationResult_GetExtendedError = 6
    IAppActivationResult_GetAppResourceGroupInfo = 7
  
End Enum

' ----==== Variablen ====----
Private m_pIAppActivationResult As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIAppActivationResult)
End Sub

' ----==== Properties ====----
Public Property Let Ifc(ByVal pIAppActivationResult As Long)
    Call ReleaseIfc(m_pIAppActivationResult)
    m_pIAppActivationResult = pIAppActivationResult
End Property

Public Property Get Ifc() As Long
    Ifc = m_pIAppActivationResult
End Property

Public Property Get ExtendedError() As Long
    Dim Ret As Long
    If m_pIAppActivationResult <> 0& Then
        Dim Value As Long
        If Invoke(m_pIAppActivationResult, _
                  IAppActivationResult_GetExtendedError, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    ExtendedError = Ret
End Property

Public Property Get AppResourceGroupInfo() As AppResourceGroupInfo
    Dim Ret As AppResourceGroupInfo
    If m_pIAppActivationResult <> 0& Then
        Dim pIAppResourceGroupInfo As Long
        If Invoke(m_pIAppActivationResult, _
                  IAppActivationResult_GetAppResourceGroupInfo, _
                  VarPtr(pIAppResourceGroupInfo)) = S_OK Then
            If pIAppResourceGroupInfo <> 0& Then
                Set Ret = New AppResourceGroupInfo
                Ret.Ifc = pIAppResourceGroupInfo
            End If
        End If
    End If
    Set AppResourceGroupInfo = Ret
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



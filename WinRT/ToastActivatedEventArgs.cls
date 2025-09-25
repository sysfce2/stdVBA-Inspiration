VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToastActivatedEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 02/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_UI_Notifications_ToastActivatedEventArgs As String = "Windows.UI.Notifications.ToastActivatedEventArgs"
Private Const IID_IToastActivatedEventArgs As String = "{e3bf92f3-c197-436f-8265-0625824f8dac}"
Private Const IID_IToastActivatedEventArgs2 As String = "{ab7da512-cc61-568e-81be-304ac31038fa}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IToastActivatedEventArgs
    IToastActivatedEventArgs_GetArguments = 6
    
    ' IToastActivatedEventArgs2
    IToastActivatedEventArgs2_GetUserInput = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIToastActivatedEventArgs As Long
Private m_pIToastActivatedEventArgs2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIToastActivatedEventArgs2)
    Call ReleaseIfc(m_pIToastActivatedEventArgs)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIToastActivatedEventArgs
End Property

Public Property Let Ifc(ByVal pIToastActivatedEventArgs As Long)
    Call ReleaseIfc(m_pIToastActivatedEventArgs2)
    Call ReleaseIfc(m_pIToastActivatedEventArgs)
    Call QueryIfc(pIToastActivatedEventArgs, _
                  IID_IToastActivatedEventArgs, _
                  m_pIToastActivatedEventArgs)
    Call QueryIfc(m_pIToastActivatedEventArgs, _
                  IID_IToastActivatedEventArgs2, _
                  m_pIToastActivatedEventArgs2)
End Property

' IToastActivatedEventArgs
Public Property Get Arguments() As String
    Dim Ret As String
    If m_pIToastActivatedEventArgs <> 0& Then
        Dim value As Long
        If Invoke(m_pIToastActivatedEventArgs, _
                  IToastActivatedEventArgs_GetArguments, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Arguments = Ret
End Property
    
' IToastActivatedEventArgs2
Public Property Get UserInput() As ValueSet
    Dim Ret As ValueSet
    If m_pIToastActivatedEventArgs2 <> 0& Then
        Dim pIPropertySet As Long
        If Invoke(m_pIToastActivatedEventArgs2, _
                  IToastActivatedEventArgs2_GetUserInput, _
                  VarPtr(pIPropertySet)) = S_OK Then
            If pIPropertySet <> 0& Then
                Set Ret = New ValueSet
                Ret.Ifc = pIPropertySet
            End If
        End If
    End If
    Set UserInput = Ret
End Property

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



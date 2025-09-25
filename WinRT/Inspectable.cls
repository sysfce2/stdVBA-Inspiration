VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Inspectable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const IID_IInspectable As String = "{af86e2e0-b12d-4c6a-9c5a-d7aa65101e90}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IInspectable
    IInspectable_GetIids = 3
    IInspectable_GetRuntimeClassName = 4
    IInspectable_GetTrustLevel = 5
    
End Enum

' ----==== Vars ====----
Private m_pIInspectable As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIInspectable)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIInspectable
End Property

Public Property Let Ifc(ByVal pIInspectable As Long)
    Call ReleaseIfc(m_pIInspectable)
    m_pIInspectable = pIInspectable
End Property

' ----==== Functions ====----
Friend Function GetIids(ByRef iids() As GUID) As Long
    Dim Ret As Long
    If m_pIInspectable <> 0& Then
        Dim count As Long
        Dim pIIDs As Long
        If Invoke(m_pIInspectable, _
                  IInspectable_GetIids, _
                  VarPtr(count), _
                  VarPtr(pIIDs)) = S_OK Then
            If count > 0& Then
                Dim bytes As Long
                Dim Item As Long
                ReDim iids(count - 1)
                For Item = 0 To count - 1
                    bytes = Len(iids(Item))
                    Call CopyMemory(iids(Item), ByVal pIIDs + (bytes * Item), bytes)
                Next
                Ret = count
            End If
        End If
    End If
    GetIids = Ret
End Function

Public Function GetRuntimeClassName() As String
    Dim Ret As String
    If m_pIInspectable <> 0& Then
        Dim hString As Long
        If Invoke(m_pIInspectable, _
                  IInspectable_GetRuntimeClassName, _
                  VarPtr(hString)) = S_OK Then
            If hString <> 0& Then
                Ret = GetWindowsString(hString)
            End If
        End If
    End If
    GetRuntimeClassName = Ret
End Function

Public Function GetTrustLevel() As TrustLevel
    Dim Ret As TrustLevel
    If m_pIInspectable <> 0& Then
        Dim value As Long
        If Invoke(m_pIInspectable, _
                  IInspectable_GetTrustLevel, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    GetTrustLevel = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


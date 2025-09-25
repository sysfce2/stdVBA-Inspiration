VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CastingSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Casting_CastingSource As String = "Windows.Media.Casting.CastingSource"
Private Const IID_ICastingSource As String = "{f429ea72-3467-47e6-a027-522923e9d727}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ICastingSource
    ICastingSource_GetPreferredSourceUri = 6
    ICastingSource_PutPreferredSourceUri = 7
                        
End Enum

' ----==== Variablen ====----
Private m_pICastingSource As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pICastingSource)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pICastingSource
End Property

Public Property Let Ifc(ByVal pICastingSource As Long)
    Call ReleaseIfc(m_pICastingSource)
    m_pICastingSource = pICastingSource
End Property

' ICastingSource
Public Property Get PreferredSourceUri() As Uri
    Dim Ret As Uri
    If m_pICastingSource <> 0& Then
        Dim pIUri As Long
        If Invoke(m_pICastingSource, _
                  ICastingSource_GetPreferredSourceUri, _
                  VarPtr(pIUri)) = S_OK Then
            If pIUri <> 0& Then
                Set Ret = New Uri
                Ret.Ifc = pIUri
            End If
        End If
    End If
    Set PreferredSourceUri = Ret
End Property

Public Property Let PreferredSourceUri(ByVal value As Uri)
    If m_pICastingSource <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pICastingSource, _
                  ICastingSource_PutPreferredSourceUri, _
                  value.Ifc) = S_OK Then
        End If
    End If
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


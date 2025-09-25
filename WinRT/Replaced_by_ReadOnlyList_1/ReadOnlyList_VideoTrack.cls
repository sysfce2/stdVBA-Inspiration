VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReadOnlyList_VideoTrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 11/2023

Option Explicit

' ----==== Const ====----

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IVectorView_IVideoTrack
    IVectorView_IVideoTrack_GetAt = 6
    IVectorView_IVideoTrack_GetSize = 7
    
End Enum

' ----==== Variablen ====----
Private m_pIVectorView_IVideoTrack As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIVectorView_IVideoTrack)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIVectorView_IVideoTrack
End Property

Public Property Let Ifc(ByVal pIVectorView_IVideoTrack As Long)
    Call ReleaseIfc(m_pIVectorView_IVideoTrack)
    m_pIVectorView_IVideoTrack = pIVectorView_IVideoTrack
End Property

'IVectorView_IVideoTrack
Public Property Get Size() As Long
    Dim Ret As Long
    If m_pIVectorView_IVideoTrack <> 0& Then
        Dim value As Long
        If Invoke(m_pIVectorView_IVideoTrack, _
                  IVectorView_IVideoTrack_GetSize, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Size = Ret
End Property

' ----==== Functions ====----
'IVectorView_IVideoTrack
Public Function GetAt(ByVal Index As Long) As VideoTrack
    Dim Ret As VideoTrack
    If m_pIVectorView_IVideoTrack <> 0& Then
        Dim pIVideoTrack As Long
        If Invoke(m_pIVectorView_IVideoTrack, _
                  IVectorView_IVideoTrack_GetAt, _
                  Index, _
                  VarPtr(pIVideoTrack)) = S_OK Then
            If pIVideoTrack <> 0& Then
                Set Ret = New VideoTrack
                Ret.Ifc = pIVideoTrack
            End If
        End If
    End If
    Set GetAt = Ret
End Function

' ----==== Extra Functions ====----
Public Function GetAll() As VideoTrack()
    Dim Ret() As VideoTrack
    Dim count As Long
    count = Me.Size
    If count > 0& Then
        ReDim Ret(count - 1)
        Dim Item As Long
        For Item = 0 To count - 1
            Set Ret(Item) = Me.GetAt(Item)
        Next
    End If
    GetAll = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



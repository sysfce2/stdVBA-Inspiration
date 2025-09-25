VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReadOnlyList_ScheduledToastNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 08/2023

Option Explicit

' ----==== Const ====----

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IVectorView_IScheduledToastNotification
    IVectorView_IScheduledToastNotification_GetAt = 6
    IVectorView_IScheduledToastNotification_GetSize = 7
    
End Enum

' ----==== Variablen ====----
Private m_pIVectorView_IScheduledToastNotification As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIVectorView_IScheduledToastNotification)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIVectorView_IScheduledToastNotification
End Property

Public Property Let Ifc(ByVal pIVectorView_IScheduledToastNotification As Long)
    Call ReleaseIfc(m_pIVectorView_IScheduledToastNotification)
    m_pIVectorView_IScheduledToastNotification = pIVectorView_IScheduledToastNotification
End Property

'IVectorView_IScheduledToastNotification
Public Property Get Size() As Long
    Dim Ret As Long
    If m_pIVectorView_IScheduledToastNotification <> 0& Then
        Dim value As Long
        If Invoke(m_pIVectorView_IScheduledToastNotification, _
                  IVectorView_IScheduledToastNotification_GetSize, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Size = Ret
End Property

' ----==== Functions ====----
'IVectorView_IScheduledToastNotification
Public Function GetAt(ByVal index As Long) As ScheduledToastNotification
    Dim Ret As ScheduledToastNotification
    If m_pIVectorView_IScheduledToastNotification <> 0& Then
        Dim pIScheduledToastNotification As Long
        If Invoke(m_pIVectorView_IScheduledToastNotification, _
                  IVectorView_IScheduledToastNotification_GetAt, _
                  index, _
                  VarPtr(pIScheduledToastNotification)) = S_OK Then
            If pIScheduledToastNotification <> 0& Then
                Set Ret = New ScheduledToastNotification
                Ret.Ifc = pIScheduledToastNotification
            End If
        End If
    End If
    Set GetAt = Ret
End Function

' ----==== Extra Functions ====----
Public Function GetAll() As ScheduledToastNotification()
    Dim Ret() As ScheduledToastNotification
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



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReadOnlyList_MediaTimeRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IVectorView_xxx
    IVectorView_GetAt = 6
    IVectorView_GetSize = 7
    
End Enum

Private Type MediaTimeRange_
    Start As Currency
    End As Currency
End Type

' ----==== Variablen ====----
Private m_pIVectorView As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIVectorView)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIVectorView
End Property

Public Property Let Ifc(ByVal pIVectorView As Long)
    Call ReleaseIfc(m_pIVectorView)
    m_pIVectorView = pIVectorView
End Property

'IVectorView_
Public Property Get Size() As Long
    Dim Ret As Long
    If m_pIVectorView <> 0& Then
        Dim value As Long
        If Invoke(m_pIVectorView, _
                  IVectorView_GetSize, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Size = Ret
End Property

' ----==== Functions ====----
Public Function GetAt(ByVal Index As Long) As MediaTimeRange
    Dim Ret As MediaTimeRange
    If m_pIVectorView <> 0& Then
        Dim tMediaTimeRange As MediaTimeRange_
        If Invoke(m_pIVectorView, _
                  IVectorView_GetAt, _
                  Index, _
                  VarPtr(tMediaTimeRange)) = S_OK Then
            Dim TimeRangeStart As New TimeSpan
            Dim TimeRangeEnd As New TimeSpan
            TimeRangeStart.TimeSpanValue = tMediaTimeRange.Start
            TimeRangeEnd.TimeSpanValue = tMediaTimeRange.End
            Set Ret = New MediaTimeRange
            Ret.StartTime = TimeRangeStart
            Ret.EndTime = TimeRangeEnd
        End If
    End If
    Set GetAt = Ret
End Function

' ----==== Extra Functions ====----
Public Function GetAll() As MediaTimeRange()
    Dim Ret() As MediaTimeRange
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


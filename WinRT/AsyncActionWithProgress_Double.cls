VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AsyncActionWithProgress_Double"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 04/2023

Option Explicit

' ----==== Const ====----
Private Const IAsyncActionWithProgress_Double As String = "{4f1430a6-a825-56ca-b047-1a9bad52ba67}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IAsyncActionWithProgress_Double
    IAsyncActionWithProgress_Double_PutProgress = 6
'    IAsyncActionWithProgress_Double_GetProgress = 7
    IAsyncActionWithProgress_Double_PutCompleted = 8
'    IAsyncActionWithProgress_Double_GetCompleted = 9
    IAsyncActionWithProgress_Double_GetResults = 10
                        
End Enum

' ----==== Variablen ====----
Private m_pIAsyncActionWithProgress_Double As Long

' ----==== Events ====----
Public Event Progress(ByVal asyncInfo As Long, ByVal progressInfo As Double)
Public Event Completed(ByVal asyncInfo As Long, ByVal asyncStat As Long)

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIAsyncActionWithProgress_Double)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIAsyncActionWithProgress_Double
End Property

Public Property Let Ifc(ByVal pIAsyncActionWithProgress_Double As Long)
    Call ReleaseIfc(m_pIAsyncActionWithProgress_Double)
    m_pIAsyncActionWithProgress_Double = pIAsyncActionWithProgress_Double
    Call Invoke(m_pIAsyncActionWithProgress_Double, _
                IAsyncActionWithProgress_Double_PutProgress, _
                AAPH_Double.Create(Me))
    Call Invoke(m_pIAsyncActionWithProgress_Double, _
                IAsyncActionWithProgress_Double_PutCompleted, _
                AAWPCH_Double.Create(Me))
End Property

' ----==== Functions ====----
' IAsyncActionWithProgress_Double
Public Property Get GetResults() As Boolean
    Dim Ret As Boolean
    If m_pIAsyncActionWithProgress_Double <> 0& Then
        If Invoke(IAsyncActionWithProgress_Double, _
                  IAsyncActionWithProgress_Double_GetResults) = S_OK Then
            Ret = True
        End If
    End If
    GetResults = Ret
End Property

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function

' ----==== Invoke Handler ====----
Public Sub Invoke_AsyncActionProgressHandler_Double(ByVal asyncInfo As Long, ByVal progressInfo As Double)
    RaiseEvent Progress(asyncInfo, progressInfo)
End Sub

Public Sub Invoke_AsyncActionWithProgressCompletedHandler_Double(ByVal asyncInfo As Long, ByVal asyncStat As Long)
    RaiseEvent Completed(asyncInfo, asyncStat)
End Sub



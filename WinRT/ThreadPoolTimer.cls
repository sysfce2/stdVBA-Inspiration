VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThreadPoolTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_Threading_ThreadPoolTimer As String = "Windows.System.Threading.ThreadPoolTimer"
Private Const IID_IThreadPoolTimer As String = "{594ebe78-55ea-4a88-a50d-3402ae1f9cf2}"
Private Const IID_IThreadPoolTimerStatics As String = "{1a8a9d02-e482-461b-b8c7-8efad1cce590}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IThreadPoolTimerStatics
    IThreadPoolTimerStatics_CreatePeriodicTimer = 6
    IThreadPoolTimerStatics_CreateTimer = 7
    IThreadPoolTimerStatics_CreatePeriodicTimerWithCompletion = 8
    IThreadPoolTimerStatics_CreateTimerWithCompletion = 9

    ' IThreadPoolTimer
    IThreadPoolTimer_GetPeriod = 6
    IThreadPoolTimer_GetDelay = 7
    IThreadPoolTimer_Cancel = 8
       
End Enum

' ----==== Vars ====----
Private m_pIThreadPoolTimer As Long
Private m_pIThreadPoolTimerStatics As Long

' ----==== Events ====----
Public Event TimerElapsed(ByVal pTimer As Long)
Public Event TimerDestroyed(ByVal pTimer As Long)

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_System_Threading_ThreadPoolTimer, _
                            IID_IThreadPoolTimerStatics, _
                            m_pIThreadPoolTimerStatics) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIThreadPoolTimer)
    Call ReleaseIfc(m_pIThreadPoolTimerStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIThreadPoolTimer
End Property

Public Property Let Ifc(ByVal pIThreadPoolTimer As Long)
    Call ReleaseIfc(m_pIThreadPoolTimer)
    m_pIThreadPoolTimer = pIThreadPoolTimer
End Property

' IThreadPoolTimer
Public Property Get Period() As TimeSpan
    Dim Ret As TimeSpan
    If m_pIThreadPoolTimer <> 0& Then
        Dim Value As Currency
        If Invoke(m_pIThreadPoolTimer, _
                  IThreadPoolTimer_GetPeriod, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New TimeSpan
            Ret.TimeSpanValue = Value
        End If
    End If
    Set Period = Ret
End Property

Public Property Get Delay() As TimeSpan
    Dim Ret As TimeSpan
    If m_pIThreadPoolTimer <> 0& Then
        Dim Value As Currency
        If Invoke(m_pIThreadPoolTimer, _
                  IThreadPoolTimer_GetDelay, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New TimeSpan
            Ret.TimeSpanValue = Value
        End If
    End If
    Set Delay = Ret
End Property

' ----==== Functions ====----
' IThreadPoolTimerStatics
Public Function CreatePeriodicTimer(ByVal Period As Currency) As ThreadPoolTimer
    Dim Ret As ThreadPoolTimer
    If m_pIThreadPoolTimerStatics <> 0& Then
        Set Ret = New ThreadPoolTimer
        Dim pIThreadPoolTimer As Long
        If Invoke(m_pIThreadPoolTimerStatics, _
                  IThreadPoolTimerStatics_CreatePeriodicTimer, _
                  TEH.Create(Ret), _
                  Period, _
                  VarPtr(pIThreadPoolTimer)) = S_OK Then
            If pIThreadPoolTimer <> 0& Then
                Ret.Ifc = pIThreadPoolTimer
            Else
                Set Ret = Nothing
            End If
        Else
            Set Ret = Nothing
        End If
    End If
    Set CreatePeriodicTimer = Ret
End Function

Public Function CreateTimer(ByVal Delay As Currency) As ThreadPoolTimer
    Dim Ret As ThreadPoolTimer
    If m_pIThreadPoolTimerStatics <> 0& Then
        Set Ret = New ThreadPoolTimer
        Dim pIThreadPoolTimer As Long
        If Invoke(m_pIThreadPoolTimerStatics, _
                  IThreadPoolTimerStatics_CreateTimer, _
                  TEH.Create(Ret), _
                  Delay, _
                  VarPtr(pIThreadPoolTimer)) = S_OK Then
            If pIThreadPoolTimer <> 0& Then
                Ret.Ifc = pIThreadPoolTimer
            Else
                Set Ret = Nothing
            End If
        Else
            Set Ret = Nothing
        End If
    End If
    Set CreateTimer = Ret
End Function

Public Function CreatePeriodicTimerWithCompletion(ByVal Period As Currency) As ThreadPoolTimer
    Dim Ret As ThreadPoolTimer
    If m_pIThreadPoolTimerStatics <> 0& Then
        Set Ret = New ThreadPoolTimer
        Dim pIThreadPoolTimer As Long
        If Invoke(m_pIThreadPoolTimerStatics, _
                  IThreadPoolTimerStatics_CreatePeriodicTimerWithCompletion, _
                  TEH.Create(Ret), _
                  Period, _
                  TDH.Create(Ret), _
                  VarPtr(pIThreadPoolTimer)) = S_OK Then
            If pIThreadPoolTimer <> 0& Then
                Ret.Ifc = pIThreadPoolTimer
            Else
                Set Ret = Nothing
            End If
        Else
            Set Ret = Nothing
        End If
    End If
    Set CreatePeriodicTimerWithCompletion = Ret
End Function

Public Function CreateTimerWithCompletion(ByVal Delay As Currency) As ThreadPoolTimer
    Dim Ret As ThreadPoolTimer
    If m_pIThreadPoolTimerStatics <> 0& Then
        Set Ret = New ThreadPoolTimer
        Dim pIThreadPoolTimer As Long
        If Invoke(m_pIThreadPoolTimerStatics, _
                  IThreadPoolTimerStatics_CreateTimerWithCompletion, _
                  TEH.Create(Ret), _
                  Delay, _
                  TDH.Create(Ret), _
                  VarPtr(pIThreadPoolTimer)) = S_OK Then
            If pIThreadPoolTimer <> 0& Then
                Ret.Ifc = pIThreadPoolTimer
            Else
                Set Ret = Nothing
            End If
        Else
            Set Ret = Nothing
        End If
    End If
    Set CreateTimerWithCompletion = Ret
End Function

' IThreadPoolTimer
Public Function Cancel() As Boolean
    Dim Ret As Boolean
    If m_pIThreadPoolTimer <> 0& Then
        If Invoke(m_pIThreadPoolTimer, _
                  IThreadPoolTimer_Cancel) = S_OK Then
            Ret = True
        End If
    End If
    Cancel = Ret
End Function

' ----==== Invoke TimerHandler ====----
Public Sub Invoke_TimerElapsed(ByVal pTimer As Long)
    RaiseEvent TimerElapsed(pTimer)
End Sub

Public Sub Invoke_TimerDestroyed(ByVal pTimer As Long)
    RaiseEvent TimerDestroyed(pTimer)
End Sub

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



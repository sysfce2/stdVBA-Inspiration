VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CastingConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Casting_CastingConnection As String = "Windows.Media.Casting.CastingConnection"
Private Const IID_ICastingConnection As String = "{cd951653-c2f1-4498-8b78-5fb4cd3640dd}"

' * Class implements the following interfaces:
' *    Windows.Foundation.IClosable
 
' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ICastingConnection
    ICastingConnection_GetState = 6
    ICastingConnection_GetDevice = 7
    ICastingConnection_GetSource = 8
    ICastingConnection_PutSource = 9
    ICastingConnection_AddStateChanged = 10
    ICastingConnection_RemoveStateChanged = 11
    ICastingConnection_AddErrorOccurred = 12
    ICastingConnection_RemoveErrorOccurred = 13
    ICastingConnection_RequestStartCastingAsync = 14
    ICastingConnection_DisconnectAsync = 15
                        
End Enum

' ----==== Variablen ====----
Private m_pICastingConnection As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call DisposeIfc(m_pICastingConnection) ' <- IClosable
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pICastingConnection
End Property

Public Property Let Ifc(ByVal pICastingConnection As Long)
    Call DisposeIfc(m_pICastingConnection) ' <- IClosable
    m_pICastingConnection = pICastingConnection
End Property

' ICastingConnection
Public Property Get State() As CastingConnectionState
    Dim Ret As CastingConnectionState
    If m_pICastingConnection <> 0& Then
        Dim value As Long
        If Invoke(m_pICastingConnection, _
                  ICastingConnection_GetState, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    State = Ret
End Property

Public Property Get Device() As CastingDevice
    Dim Ret As CastingDevice
    If m_pICastingConnection <> 0& Then
        Dim pICastingDevice As Long
        If Invoke(m_pICastingConnection, _
                  ICastingConnection_GetDevice, _
                  VarPtr(pICastingDevice)) = S_OK Then
            If pICastingDevice <> 0& Then
                Set Ret = New CastingDevice
                Ret.Ifc = pICastingDevice
            End If
        End If
    End If
    Set Device = Ret
End Property

Public Property Get Source() As CastingSource
    Dim Ret As CastingSource
    If m_pICastingConnection <> 0& Then
        Dim pICastingSource As Long
        If Invoke(m_pICastingConnection, _
                  ICastingConnection_GetSource, _
                  VarPtr(pICastingSource)) = S_OK Then
            If pICastingSource <> 0& Then
                Set Ret = New CastingSource
                Ret.Ifc = pICastingSource
            End If
        End If
    End If
    Set Source = Ret
End Property

Public Property Let Source(ByVal value As CastingSource)
    Dim Ret As CastingSource
    If m_pICastingConnection <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pICastingConnection, _
                  ICastingConnection_PutSource, _
                  value.Ifc) = S_OK Then
        End If
    End If
End Property

' ----==== Functions ====----
' ICastingConnection
Public Function AddStateChanged(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pICastingConnection <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pICastingConnection, _
                  ICastingConnection_AddStateChanged, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddStateChanged = Ret
End Function

Public Function RemoveStateChanged(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pICastingConnection <> 0& And token <> 0@ Then
        If Invoke(m_pICastingConnection, _
                  ICastingConnection_RemoveStateChanged, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveStateChanged = Ret
End Function

Public Function AddErrorOccurred(ByVal handler As Long) As Currency
    Dim Ret As Currency
    If m_pICastingConnection <> 0& And handler <> 0& Then
        Dim token As Currency
        If Invoke(m_pICastingConnection, _
                  ICastingConnection_AddErrorOccurred, _
                  handler, _
                  VarPtr(token)) = S_OK Then
            Ret = token
        End If
    End If
    AddErrorOccurred = Ret
End Function

Public Function RemoveErrorOccurred(ByVal token As Currency) As Boolean
    Dim Ret As Boolean
    If m_pICastingConnection <> 0& And token <> 0@ Then
        If Invoke(m_pICastingConnection, _
                  ICastingConnection_RemoveErrorOccurred, _
                  token) = S_OK Then
            Ret = True
        End If
    End If
    RemoveErrorOccurred = Ret
End Function

Public Function RequestStartCastingAsync(ByVal value As CastingSource) As CastingConnectionErrorStatus
    Dim Ret As CastingConnectionErrorStatus
    If m_pICastingConnection <> 0& And _
       IsNotNothing(value) Then
        Dim operation As Long
        If Invoke(m_pICastingConnection, _
                  ICastingConnection_RequestStartCastingAsync, _
                  value.Ifc, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Ret = operation
                End If
            End If
        End If
    End If
    RequestStartCastingAsync = Ret
End Function

Public Function DisconnectAsync() As CastingConnectionErrorStatus
    Dim Ret As CastingConnectionErrorStatus
    If m_pICastingConnection <> 0& Then
        Dim operation As Long
        If Invoke(m_pICastingConnection, _
                  ICastingConnection_DisconnectAsync, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Ret = operation
                End If
            End If
        End If
    End If
    DisconnectAsync = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


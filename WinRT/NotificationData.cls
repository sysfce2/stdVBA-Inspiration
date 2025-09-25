VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NotificationData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 08/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_UI_Notifications_NotificationData As String = "Windows.UI.Notifications.NotificationData"
Private Const IID_INotificationData As String = "{9ffd2312-9d6a-4aaf-b6ac-ff17f0c1f280}"
Private Const IID_INotificationDataFactory As String = "{23c1e33a-1c10-46fb-8040-dec384621cf8}"
 
' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' INotificationDataFactory
    INotificationDataFactory_CreateNotificationDataWithValuesAndSequenceNumber = 6
    INotificationDataFactory_CreateNotificationDataWithValues = 7
    
    ' INotificationData
    INotificationData_GetValues = 6
    INotificationData_GetSequenceNumber = 7
    INotificationData_PutSequenceNumber = 8
                        
    ' IMap_HSTRING_HSTRING
'    IMap_HSTRING_HSTRING_Lookup = 6
'    IMap_HSTRING_HSTRING_GetSize = 7
'    IMap_HSTRING_HSTRING_HasKey = 8
'    IMap_HSTRING_HSTRING_GetView = 9
    IMap_HSTRING_HSTRING_Insert = 10
'    IMap_HSTRING_HSTRING_Remove = 11
'    IMap_HSTRING_HSTRING_Clear = 12
                        
End Enum

' ----==== Variablen ====----
Private m_pINotificationData As Long
Private m_pINotificationDataFactory As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_UI_Notifications_NotificationData, _
                            IID_INotificationDataFactory, _
                            m_pINotificationDataFactory) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pINotificationData)
    Call ReleaseIfc(m_pINotificationDataFactory)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pINotificationData
End Property

Public Property Let Ifc(ByVal pINotificationData As Long)
    Call ReleaseIfc(m_pINotificationData)
    m_pINotificationData = pINotificationData
End Property

' INotificationData
'INotificationData_GetValues

Public Property Get SequenceNumber() As Long
    Dim Ret As Long
    If m_pINotificationData <> 0& Then
        Dim value As Long
        If Invoke(m_pINotificationData, _
                  INotificationData_GetSequenceNumber, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    SequenceNumber = Ret
End Property

Public Property Let SequenceNumber(ByVal value As Long)
    If m_pINotificationData <> 0& Then
        If Invoke(m_pINotificationData, _
                  INotificationData_PutSequenceNumber, _
                  value) = S_OK Then
        End If
    End If
End Property

' ----==== Functions ====----
' INotificationDataFactory
Friend Function CreateNotificationDataWithValuesAndSequenceNumber(ByRef initialValues() As KeyValuePair_String_String, _
                                                                  ByVal SequenceNumber As Long) As Boolean
    Dim Ret As Boolean
    If m_pINotificationDataFactory <> 0& Then
        Dim value As Long
        If Invoke(m_pINotificationDataFactory, _
                  INotificationDataFactory_CreateNotificationDataWithValues, _
                  0&, _
                  SequenceNumber, _
                  VarPtr(value)) = S_OK Then
            Me.Ifc = value
            Ret = AddValues(initialValues)
        End If
    End If
    CreateNotificationDataWithValuesAndSequenceNumber = Ret
End Function

Friend Function CreateNotificationDataWithValues(ByRef initialValues() As KeyValuePair_String_String) As Boolean
    Dim Ret As Boolean
    If m_pINotificationDataFactory <> 0& Then
        Dim value As Long
        If Invoke(m_pINotificationDataFactory, _
                  INotificationDataFactory_CreateNotificationDataWithValues, _
                  0&, _
                  VarPtr(value)) = S_OK Then
            Me.Ifc = value
            Ret = AddValues(initialValues)
        End If
    End If
    CreateNotificationDataWithValues = Ret
End Function

Private Function AddValues(ByRef initialValues() As KeyValuePair_String_String) As Boolean
    Dim Ret As Boolean
    If m_pINotificationData <> 0& Then
        Dim pIMap_HSTRING_HSTRING As Long
        If Invoke(m_pINotificationData, _
                  INotificationData_GetValues, _
                  VarPtr(pIMap_HSTRING_HSTRING)) = S_OK Then
            Dim count As Long
            count = UBound(initialValues)
            If count >= 0 Then
                Dim Item As Long
                For Item = 0 To count
                    Dim hKey As Long
                    Dim hValue As Long
                    Dim replaced As Long
                    hKey = CreateWindowsString(initialValues(Item).Key)
                    hValue = CreateWindowsString(initialValues(Item).value)
                    If Invoke(pIMap_HSTRING_HSTRING, _
                              IMap_HSTRING_HSTRING_Insert, _
                              hKey, _
                              hValue, _
                              VarPtr(replaced)) = S_OK Then
                    End If
                    Call DeleteWindowsString(hKey)
                    Call DeleteWindowsString(hValue)
                Next
                Ret = True
            End If
            Call ReleaseIfc(pIMap_HSTRING_HSTRING)
        End If
    End If
    AddValues = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


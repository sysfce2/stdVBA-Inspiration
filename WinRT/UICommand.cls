VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UICommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_UI_Popups_UICommand As String = "Windows.UI.Popups.UICommand"
Private Const IID_IUICommand As String = "{4ff93a75-4145-47ff-ac7f-dff1c1fa5b0f}"
Private Const IID_IUICommandFactory As String = "{a21a8189-26b0-4676-ae94-54041bc125e8}"

Private Const Windows_Foundation_PropertyValue As String = "Windows.Foundation.PropertyValue"
Private Const IID_IPropertyValue As String = "{4bd682dd-7554-40e9-9a9b-82654ede7e62}"
Private Const IID_IPropertyValueStatics As String = "{629bdbc8-d932-4ff4-96b9-8d96c5c1e858}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IUICommandFactory
    IUICommandFactory_Create = 6
    IUICommandFactory_CreateWithHandler = 7
    IUICommandFactory_CreateWithHandlerAndId = 8

    ' IUICommand
    IUICommand_GetLabel = 6
    IUICommand_PutLabel = 7
    IUICommand_GetInvoked = 8
    IUICommand_PutInvoked = 9
    IUICommand_GetId = 10
    IUICommand_PutId = 11

    ' IPropertyValueStatics
'    IPropertyValueStatics_CreateEmpty = 6
'    IPropertyValueStatics_CreateUInt8 = 7
    IPropertyValueStatics_CreateInt16 = 8
'    IPropertyValueStatics_CreateUInt16 = 9
    IPropertyValueStatics_CreateInt32 = 10
'    IPropertyValueStatics_CreateUInt32 = 11
'    IPropertyValueStatics_CreateInt64 = 12
'    IPropertyValueStatics_CreateUInt64 = 13
'    IPropertyValueStatics_CreateSingle = 14
'    IPropertyValueStatics_CreateDouble = 15
'    IPropertyValueStatics_CreateChar16 = 16
'    IPropertyValueStatics_CreateBoolean = 17
    IPropertyValueStatics_CreateString = 18
'    IPropertyValueStatics_CreateInspectable = 19
'    IPropertyValueStatics_CreateGuid = 20
'    IPropertyValueStatics_CreateDateTime = 21
'    IPropertyValueStatics_CreateTimeSpan = 22
'    IPropertyValueStatics_CreatePoint = 23
'    IPropertyValueStatics_CreateSize = 24
'    IPropertyValueStatics_CreateRect = 25
'    IPropertyValueStatics_CreateUInt8Array = 26
'    IPropertyValueStatics_CreateInt16Array = 27
'    IPropertyValueStatics_CreateUInt16Array = 28
'    IPropertyValueStatics_CreateInt32Array = 29
'    IPropertyValueStatics_CreateUInt32Array = 30
'    IPropertyValueStatics_CreateInt64Array = 31
'    IPropertyValueStatics_CreateUInt64Array = 32
'    IPropertyValueStatics_CreateSingleArray = 33
'    IPropertyValueStatics_CreateDoubleArray = 34
'    IPropertyValueStatics_CreateChar16Array = 35
'    IPropertyValueStatics_CreateBooleanArray = 36
'    IPropertyValueStatics_CreateStringArray = 37
'    IPropertyValueStatics_CreateInspectableArray = 38
'    IPropertyValueStatics_CreateGuidArray = 39
'    IPropertyValueStatics_CreateDateTimeArray = 40
'    IPropertyValueStatics_CreateTimeSpanArray = 41
'    IPropertyValueStatics_CreatePointArray = 42
'    IPropertyValueStatics_CreateSizeArray = 43
'    IPropertyValueStatics_CreateRectArray = 44
                    
    ' IPropertyValue
    IPropertyValue_GetType = 6
'    IPropertyValue_GetIsNumericScalar = 7
'    IPropertyValue_GetUInt8 = 8
    IPropertyValue_GetInt16 = 9
'    IPropertyValue_GetUInt16 = 10
    IPropertyValue_GetInt32 = 11
'    IPropertyValue_GetUInt32 = 12
'    IPropertyValue_GetInt64 = 13
'    IPropertyValue_GetUInt64 = 14
'    IPropertyValue_GetSingle = 15
'    IPropertyValue_GetDouble = 16
'    IPropertyValue_GetChar16 = 17
'    IPropertyValue_GetBoolean = 18
    IPropertyValue_GetString = 19
'    IPropertyValue_GetGuid = 20
'    IPropertyValue_GetDateTime = 21
'    IPropertyValue_GetTimeSpan = 22
'    IPropertyValue_GetPoint = 23
'    IPropertyValue_GetSize = 24
'    IPropertyValue_GetRect = 25
'    IPropertyValue_GetUInt8Array = 26
'    IPropertyValue_GetInt16Array = 27
'    IPropertyValue_GetUInt16Array = 28
'    IPropertyValue_GetInt32Array = 29
'    IPropertyValue_GetUInt32Array = 30
'    IPropertyValue_GetInt64Array = 31
'    IPropertyValue_GetUInt64Array = 32
'    IPropertyValue_GetSingleArray = 33
'    IPropertyValue_GetDoubleArray = 34
'    IPropertyValue_GetChar16Array = 35
'    IPropertyValue_GetBooleanArray = 36
'    IPropertyValue_GetStringArray = 37
'    IPropertyValue_GetInspectableArray = 38
'    IPropertyValue_GetGuidArray = 39
'    IPropertyValue_GetDateTimeArray = 40
'    IPropertyValue_GetTimeSpanArray = 41
'    IPropertyValue_GetPointArray = 42
'    IPropertyValue_GetSizeArray = 43
'    IPropertyValue_GetRectArray = 44

End Enum

' ----==== Variablen ====----
Private m_pIUICommand As Long
Private m_pIUICommandFactory As Long
Private m_pIPropertyValueStatics As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_UI_Popups_UICommand, _
                            IID_IUICommandFactory, _
                            m_pIUICommandFactory) Then
    End If
    If GetActivationFactory(Windows_Foundation_PropertyValue, _
                            IID_IPropertyValueStatics, _
                            m_pIPropertyValueStatics) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIUICommand)
    Call ReleaseIfc(m_pIUICommandFactory)
    Call ReleaseIfc(m_pIPropertyValueStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIUICommand
End Property

Public Property Let Ifc(ByVal pIUICommand As Long)
    Call ReleaseIfc(m_pIUICommand)
    m_pIUICommand = pIUICommand
End Property

' IUICommand
Public Property Get Label() As String
    Dim Ret As String
    If m_pIUICommand <> 0& Then
        Dim hString As Long
        If Invoke(m_pIUICommand, _
                  IUICommand_GetLabel, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    Label = Ret
End Property

Public Property Let Label(ByVal Value As String)
    If m_pIUICommand <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Value)
        If Invoke(m_pIUICommand, _
                  IUICommand_PutLabel, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get Invoked() As Long
    Dim Ret As Long
    If m_pIUICommand <> 0& Then
        Dim Value As Long
        If Invoke(m_pIUICommand, _
                  IUICommand_GetInvoked, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Invoked = Ret
End Property

Public Property Let Invoked(ByVal Value As Long)
    If m_pIUICommand <> 0& Then
        If Invoke(m_pIUICommand, _
                  IUICommand_PutInvoked, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get Id() As Variant
    Dim Ret As Variant
    If m_pIUICommand <> 0& Then
        Dim pIInspectable As Long
        If Invoke(m_pIUICommand, _
                  IUICommand_GetId, _
                  VarPtr(pIInspectable)) = S_OK Then
            Ret = GetIInspectable(pIInspectable)
            Call ReleaseIfc(pIInspectable)
        End If
    End If
    Id = Ret
End Property

Public Property Let Id(ByVal Value As Variant)
    If m_pIUICommand <> 0& Then
        Dim pIInspectable As Long
        pIInspectable = CreateIInspectable(Value)
        If Invoke(m_pIUICommand, _
                  IUICommand_PutId, _
                  pIInspectable) = S_OK Then
        End If
        Call ReleaseIfc(pIInspectable)
    End If
End Property

' ----==== Functions ====----
' IUICommandFactory
Public Function Create(ByVal Label As String) As UICommand
    Dim Ret As UICommand
    If m_pIUICommandFactory <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Label)
        Dim pIUICommand As Long
        If Invoke(m_pIUICommandFactory, _
                  IUICommandFactory_Create, _
                  hString, _
                  VarPtr(pIUICommand)) = S_OK Then
            If pIUICommand <> 0& Then
                Set Ret = New UICommand
                Ret.Ifc = pIUICommand
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set Create = Ret
End Function

Public Function CreateWithHandler(ByVal Label As String, _
                                  ByVal action As UICommandInvokedHandler) As UICommand
    Dim Ret As UICommand
    If m_pIUICommandFactory <> 0& And _
       IsNotNothing(action) <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Label)
        Dim pIUICommand As Long
        If Invoke(m_pIUICommandFactory, _
                  IUICommandFactory_CreateWithHandler, _
                  hString, _
                  action.Ifc, _
                  VarPtr(pIUICommand)) = S_OK Then
            If pIUICommand <> 0& Then
                Set Ret = New UICommand
                Ret.Ifc = pIUICommand
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set CreateWithHandler = Ret
End Function

Public Function CreateWithHandlerAndId(ByVal Label As String, _
                                       ByVal action As UICommandInvokedHandler, _
                                       ByVal commandId As Variant) As UICommand
    Dim Ret As UICommand
    If m_pIUICommandFactory <> 0& And _
       IsNotNothing(action) Then
        Dim hString As Long
        hString = CreateWindowsString(Label)
        Dim pIUICommand As Long
        Dim pIInspectable As Long
        pIInspectable = CreateIInspectable(commandId)
        If Invoke(m_pIUICommandFactory, _
                  IUICommandFactory_CreateWithHandlerAndId, _
                  hString, _
                  action.Ifc, _
                  pIInspectable, _
                  VarPtr(pIUICommand)) = S_OK Then
            If pIUICommand <> 0& Then
                Set Ret = New UICommand
                Ret.Ifc = pIUICommand
            End If
        End If
        Call ReleaseIfc(pIInspectable)
        Call DeleteWindowsString(hString)
    End If
    Set CreateWithHandlerAndId = Ret
End Function

Private Function CreateIInspectable(ByVal Value As Variant) As Long
    Dim Ret As Long
    If m_pIPropertyValueStatics <> 0& Then
        Select Case VarType(Value)
        Case vbInteger
            Dim int16 As Long
            If Invoke(m_pIPropertyValueStatics, _
                      IPropertyValueStatics_CreateInt16, _
                      CInt(Value), _
                      VarPtr(int16)) = S_OK Then
                Ret = int16
            End If
        Case vbLong
            Dim int32 As Long
            If Invoke(m_pIPropertyValueStatics, _
                      IPropertyValueStatics_CreateInt32, _
                      CInt(Value), _
                      VarPtr(int32)) = S_OK Then
                Ret = int32
            End If
        Case vbString
            Dim hString As Long
            hString = CreateWindowsString(CStr(Value))
            Dim pRet As Long
            If Invoke(m_pIPropertyValueStatics, _
                      IPropertyValueStatics_CreateString, _
                      hString, _
                      VarPtr(pRet)) = S_OK Then
                Ret = pRet
            End If
            Call DeleteWindowsString(hString)
        End Select
    End If
    CreateIInspectable = Ret
End Function

Private Function GetIInspectable(ByVal pIInspectable As Long) As Variant
    Dim Ret As Variant
    If pIInspectable <> 0& Then
        Dim PropType As PropertyType
        If Invoke(pIInspectable, _
                  IPropertyValue_GetType, _
                  VarPtr(PropType)) = S_OK Then
            Select Case PropType
                Case PropertyType.PropertyType_Int16
                    Dim int16 As Integer
                    If Invoke(pIInspectable, _
                              IPropertyValue_GetInt16, _
                              VarPtr(int16)) = S_OK Then
                        Ret = int16
                    End If
                Case PropertyType.PropertyType_Int32
                    Dim int32 As Long
                    If Invoke(pIInspectable, _
                              IPropertyValue_GetInt32, _
                              VarPtr(int32)) = S_OK Then
                        Ret = int32
                    End If
                Case PropertyType.PropertyType_String
                    Dim hString As Long
                    If Invoke(pIInspectable, _
                              IPropertyValue_GetString, _
                              VarPtr(hString)) = S_OK Then
                        Ret = GetWindowsString(hString)
                    End If
            End Select
        End If
    End If
    GetIInspectable = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



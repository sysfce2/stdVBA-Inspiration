VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

' see: https://learn.microsoft.com/en-us/uwp/api/windows.foundation.metadata.apiinformation?view=winrt-22621

Option Explicit

' ----==== Const ====----
Private Const Windows_Foundation_Metadata_ApiInformation As String = "Windows.Foundation.Metadata.ApiInformation"
Private Const IID_IApiInformationStatics As String = "{997439fe-f681-4a11-b416-c13a47e8ba36}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IID_IApiInformationStatics
    IApiInformationStatics_IsTypePresent = 6
    IApiInformationStatics_IsMethodPresent = 7
    IApiInformationStatics_IsMethodPresentWithArity = 8
    IApiInformationStatics_IsEventPresent = 9
    IApiInformationStatics_IsPropertyPresent = 10
    IApiInformationStatics_IsReadOnlyPropertyPresent = 11
    IApiInformationStatics_IsWriteablePropertyPresent = 12
    IApiInformationStatics_IsEnumNamedValuePresent = 13
    IApiInformationStatics_IsApiContractPresentByMajor = 14
    IApiInformationStatics_IsApiContractPresentByMajorAndMinor = 15
                        
End Enum

' ----==== Vars ====----
Private m_pIApiInformationStatics As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Foundation_Metadata_ApiInformation, _
                            IID_IApiInformationStatics, _
                            m_pIApiInformationStatics) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIApiInformationStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIApiInformationStatics
End Property

' ----==== Functions ====----
' IID_IApiInformationStatics
Public Function IsTypePresent(ByVal typeName As String) As Boolean
    Dim Ret As Boolean
    If m_pIApiInformationStatics <> 0& Then
        Dim result As Long
        Dim hTypeName As Long
        hTypeName = CreateWindowsString(typeName)
        If Invoke(m_pIApiInformationStatics, _
                  IApiInformationStatics_IsTypePresent, _
                  hTypeName, _
                  VarPtr(result)) = S_OK Then
            Ret = CBool(result)
        End If
        Call DeleteWindowsString(hTypeName)
    End If
    IsTypePresent = Ret
End Function

Public Function IsMethodPresent(ByVal typeName As String, _
                                ByVal methodName As String) As Boolean
    Dim Ret As Boolean
    If m_pIApiInformationStatics <> 0& Then
        Dim result As Long
        Dim hTypeName As Long
        Dim hMethodName As Long
        hTypeName = CreateWindowsString(typeName)
        hMethodName = CreateWindowsString(methodName)
        If Invoke(m_pIApiInformationStatics, _
                  IApiInformationStatics_IsMethodPresent, _
                  hTypeName, _
                  hMethodName, _
                  VarPtr(result)) = S_OK Then
            Ret = CBool(result)
        End If
        Call DeleteWindowsString(hMethodName)
        Call DeleteWindowsString(hTypeName)
    End If
    IsMethodPresent = Ret
End Function

Public Function IsMethodPresentWithArity(ByVal typeName As String, _
                                         ByVal methodName As String, _
                                         ByVal inputParameterCount As Long) As Boolean
    Dim Ret As Boolean
    If m_pIApiInformationStatics <> 0& Then
        Dim result As Long
        Dim hTypeName As Long
        Dim hMethodName As Long
        hTypeName = CreateWindowsString(typeName)
        hMethodName = CreateWindowsString(methodName)
        If Invoke(m_pIApiInformationStatics, _
                  IApiInformationStatics_IsMethodPresentWithArity, _
                  hTypeName, _
                  hMethodName, _
                  inputParameterCount, _
                  VarPtr(result)) = S_OK Then
            Ret = CBool(result)
        End If
        Call DeleteWindowsString(hMethodName)
        Call DeleteWindowsString(hTypeName)
    End If
    IsMethodPresentWithArity = Ret
End Function

Public Function IsEventPresent(ByVal typeName As String, _
                               ByVal eventName As String) As Boolean
    Dim Ret As Boolean
    If m_pIApiInformationStatics <> 0& Then
        Dim result As Long
        Dim hTypeName As Long
        Dim hEventName As Long
        hTypeName = CreateWindowsString(typeName)
        hEventName = CreateWindowsString(eventName)
        If Invoke(m_pIApiInformationStatics, _
                  IApiInformationStatics_IsEventPresent, _
                  hTypeName, _
                  hEventName, _
                  VarPtr(result)) = S_OK Then
            Ret = CBool(result)
        End If
        Call DeleteWindowsString(hEventName)
        Call DeleteWindowsString(hTypeName)
    End If
    IsEventPresent = Ret
End Function

Public Function IsPropertyPresent(ByVal typeName As String, _
                                  ByVal PropertyName As String) As Boolean
    Dim Ret As Boolean
    If m_pIApiInformationStatics <> 0& Then
        Dim result As Long
        Dim hTypeName As Long
        Dim hPropertyName As Long
        hTypeName = CreateWindowsString(typeName)
        hPropertyName = CreateWindowsString(PropertyName)
        If Invoke(m_pIApiInformationStatics, _
                  IApiInformationStatics_IsPropertyPresent, _
                  hTypeName, _
                  hPropertyName, _
                  VarPtr(result)) = S_OK Then
            Ret = CBool(result)
        End If
        Call DeleteWindowsString(hPropertyName)
        Call DeleteWindowsString(hTypeName)
    End If
    IsPropertyPresent = Ret
End Function

Public Function IsReadOnlyPropertyPresent(ByVal typeName As String, _
                                          ByVal PropertyName As String) As Boolean
    Dim Ret As Boolean
    If m_pIApiInformationStatics <> 0& Then
        Dim result As Long
        Dim hTypeName As Long
        Dim hPropertyName As Long
        hTypeName = CreateWindowsString(typeName)
        hPropertyName = CreateWindowsString(PropertyName)
        If Invoke(m_pIApiInformationStatics, _
                  IApiInformationStatics_IsReadOnlyPropertyPresent, _
                  hTypeName, _
                  hPropertyName, _
                  VarPtr(result)) = S_OK Then
            Ret = CBool(result)
        End If
        Call DeleteWindowsString(hPropertyName)
        Call DeleteWindowsString(hTypeName)
    End If
    IsReadOnlyPropertyPresent = Ret
End Function

Public Function IsWriteablePropertyPresent(ByVal typeName As String, _
                                           ByVal PropertyName As String) As Boolean
    Dim Ret As Boolean
    If m_pIApiInformationStatics <> 0& Then
        Dim result As Long
        Dim hTypeName As Long
        Dim hPropertyName As Long
        hTypeName = CreateWindowsString(typeName)
        hPropertyName = CreateWindowsString(PropertyName)
        If Invoke(m_pIApiInformationStatics, _
                  IApiInformationStatics_IsWriteablePropertyPresent, _
                  hTypeName, _
                  hPropertyName, _
                  VarPtr(result)) = S_OK Then
            Ret = CBool(result)
        End If
        Call DeleteWindowsString(hPropertyName)
        Call DeleteWindowsString(hTypeName)
    End If
    IsWriteablePropertyPresent = Ret
End Function

Public Function IsEnumNamedValuePresent(ByVal enumTypeName As String, _
                                        ByVal valueName As String) As Boolean
    Dim Ret As Boolean
    If m_pIApiInformationStatics <> 0& Then
        Dim result As Long
        Dim hEnumTypeName As Long
        Dim hValueName As Long
        hEnumTypeName = CreateWindowsString(enumTypeName)
        hValueName = CreateWindowsString(valueName)
        If Invoke(m_pIApiInformationStatics, _
                  IApiInformationStatics_IsEnumNamedValuePresent, _
                  hEnumTypeName, _
                  hValueName, _
                  VarPtr(result)) = S_OK Then
            Ret = CBool(result)
        End If
        Call DeleteWindowsString(hEnumTypeName)
        Call DeleteWindowsString(hValueName)
    End If
    IsEnumNamedValuePresent = Ret
End Function

Public Function IsApiContractPresentByMajor(ByVal contractName As String, _
                                            ByVal majorVersion As Integer) As Boolean
    Dim Ret As Boolean
    If m_pIApiInformationStatics <> 0& Then
        Dim result As Long
        Dim hContractName As Long
        hContractName = CreateWindowsString(contractName)
        If Invoke(m_pIApiInformationStatics, _
                  IApiInformationStatics_IsApiContractPresentByMajor, _
                  hContractName, _
                  majorVersion, _
                  VarPtr(result)) = S_OK Then
            Ret = CBool(result)
        End If
        Call DeleteWindowsString(hContractName)
    End If
    IsApiContractPresentByMajor = Ret
End Function

Public Function IsApiContractPresentByMajorAndMinor(ByVal contractName As String, _
                                                    ByVal majorVersion As Integer, _
                                                    ByVal minorVersion As Integer) As Boolean
    Dim Ret As Boolean
    If m_pIApiInformationStatics <> 0& Then
        Dim result As Long
        Dim hContractName As Long
        hContractName = CreateWindowsString(contractName)
        If Invoke(m_pIApiInformationStatics, _
                  IApiInformationStatics_IsApiContractPresentByMajorAndMinor, _
                  hContractName, _
                  majorVersion, _
                  minorVersion, _
                  VarPtr(result)) = S_OK Then
            Ret = CBool(result)
        End If
        Call DeleteWindowsString(hContractName)
    End If
    IsApiContractPresentByMajorAndMinor = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


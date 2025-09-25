VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GeographicRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Globalization_GeographicRegion As String = "Windows.Globalization.GeographicRegion"
Private Const IID_IGeographicRegion As String = "{01e9a621-4a64-4ed9-954f-9edeb07bd903}"
Private Const IID_IGeographicRegionStatics As String = "{29e28974-7ad9-4ef4-8799-b3b44fadec08}"
Private Const IID_IGeographicRegionFactory As String = "{53425270-77b4-426b-859f-81e19d512546}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IGeographicRegionFactory
    IGeographicRegionFactory_CreateGeographicRegion = 6
    
    ' IGeographicRegionStatics
    IGeographicRegionStatics_IsSupported = 6
                        
    ' IGeographicRegion
    IGeographicRegion_GetCode = 6
    IGeographicRegion_GetCodeTwoLetter = 7
    IGeographicRegion_GetCodeThreeLetter = 8
    IGeographicRegion_GetCodeThreeDigit = 9
    IGeographicRegion_GetDisplayName = 10
    IGeographicRegion_GetNativeName = 11
    IGeographicRegion_GetCurrenciesInUse = 12

End Enum

' ----==== Variablen ====----
Private m_pIGeographicRegion As Long
Private m_pIGeographicRegionStatics As Long
Private m_pIGeographicRegionFactory As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Globalization_GeographicRegion, _
                            IID_IGeographicRegionFactory, _
                            m_pIGeographicRegionFactory) Then
    End If
    If GetActivationFactory(Windows_Globalization_GeographicRegion, _
                            IID_IGeographicRegionStatics, _
                            m_pIGeographicRegionStatics) Then
    End If
    Dim pIGeographicRegion As Long
    If GetActivateInstance(Windows_Globalization_GeographicRegion, _
                           IID_IGeographicRegion, _
                           pIGeographicRegion) Then
        Me.Ifc = pIGeographicRegion
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIGeographicRegion)
    Call ReleaseIfc(m_pIGeographicRegionStatics)
    Call ReleaseIfc(m_pIGeographicRegionFactory)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIGeographicRegion
End Property

Public Property Let Ifc(ByVal pIGeographicRegion As Long)
    Call ReleaseIfc(m_pIGeographicRegion)
    m_pIGeographicRegion = pIGeographicRegion
End Property

' IGeographicRegion
Public Property Get Code() As String
    Dim Ret As String
    If m_pIGeographicRegion <> 0& Then
        Dim value As Long
        If Invoke(m_pIGeographicRegion, _
                  IGeographicRegion_GetCode, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Code = Ret
End Property

Public Property Get CodeTwoLetter() As String
    Dim Ret As String
    If m_pIGeographicRegion <> 0& Then
        Dim value As Long
        If Invoke(m_pIGeographicRegion, _
                  IGeographicRegion_GetCodeTwoLetter, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    CodeTwoLetter = Ret
End Property

Public Property Get CodeThreeLetter() As String
    Dim Ret As String
    If m_pIGeographicRegion <> 0& Then
        Dim value As Long
        If Invoke(m_pIGeographicRegion, _
                  IGeographicRegion_GetCodeThreeLetter, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    CodeThreeLetter = Ret
End Property

Public Property Get CodeThreeDigit() As String
    Dim Ret As String
    If m_pIGeographicRegion <> 0& Then
        Dim value As Long
        If Invoke(m_pIGeographicRegion, _
                  IGeographicRegion_GetCodeThreeDigit, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    CodeThreeDigit = Ret
End Property

Public Property Get DisplayName() As String
    Dim Ret As String
    If m_pIGeographicRegion <> 0& Then
        Dim value As Long
        If Invoke(m_pIGeographicRegion, _
                  IGeographicRegion_GetDisplayName, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    DisplayName = Ret
End Property

Public Property Get NativeName() As String
    Dim Ret As String
    If m_pIGeographicRegion <> 0& Then
        Dim value As Long
        If Invoke(m_pIGeographicRegion, _
                  IGeographicRegion_GetNativeName, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    NativeName = Ret
End Property

Public Property Get CurrenciesInUse() As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pIGeographicRegion <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIGeographicRegion, _
                  IGeographicRegion_GetCurrenciesInUse, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_String, pIVectorView) 'ReadOnlyList_String
            End If
        End If
    End If
    Set CurrenciesInUse = Ret
End Property

' ----==== Functions ====----
' IGeographicRegionFactory
Public Function CreateGeographicRegion(ByVal GeographicRegionCode As String) As GeographicRegion
    Dim Ret As GeographicRegion
    If m_pIGeographicRegionFactory <> 0& Then
        Dim pIGeographicRegion As Long
        Dim hGeographicRegionCode As Long
        hGeographicRegionCode = CreateWindowsString(GeographicRegionCode)
        If Invoke(m_pIGeographicRegionFactory, _
                  IGeographicRegionFactory_CreateGeographicRegion, _
                  hGeographicRegionCode, _
                  VarPtr(pIGeographicRegion)) = S_OK Then
            If pIGeographicRegion <> 0& Then
                Set Ret = New GeographicRegion
                Ret.Ifc = pIGeographicRegion
            End If
        End If
        Call DeleteWindowsString(hGeographicRegionCode)
    End If
    Set CreateGeographicRegion = Ret
End Function

' IGeographicRegionStatics
Public Function IsSupported(ByVal GeographicRegionCode As String) As Boolean
    Dim Ret As Boolean
    If m_pIGeographicRegionStatics <> 0& Then
        Dim Result As Long
        Dim hGeographicRegionCode As Long
        hGeographicRegionCode = CreateWindowsString(GeographicRegionCode)
        If Invoke(m_pIGeographicRegionStatics, _
                  IGeographicRegionStatics_IsSupported, _
                  hGeographicRegionCode, _
                  VarPtr(Result)) = S_OK Then
            Ret = CBool(Result)
        End If
        Call DeleteWindowsString(hGeographicRegionCode)
    End If
    IsSupported = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


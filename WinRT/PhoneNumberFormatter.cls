VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PhoneNumberFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Globalization_PhoneNumberFormatting_PhoneNumberFormatter As String = "Windows.Globalization.PhoneNumberFormatting.PhoneNumberFormatter"
Private Const IID_IPhoneNumberFormatter As String = "{1556b49e-bad4-4b4a-900d-4407adb7c981}"
Private Const IID_IPhoneNumberFormatterStatics As String = "{5ca6f931-84d9-414b-ab4e-a0552c878602}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IPhoneNumberFormatterStatics
    IPhoneNumberFormatterStatics_TryCreate = 6
    IPhoneNumberFormatterStatics_GetCountryCodeForRegion = 7
    IPhoneNumberFormatterStatics_GetNationalDirectDialingPrefixForRegion = 8
    IPhoneNumberFormatterStatics_WrapWithLeftToRightMarkers = 9
                        
    ' IPhoneNumberFormatter
    IPhoneNumberFormatter_Format = 6
    IPhoneNumberFormatter_FormatWithOutputFormat = 7
    IPhoneNumberFormatter_FormatPartialString = 8
    IPhoneNumberFormatter_FormatString = 9
    IPhoneNumberFormatter_FormatStringWithLeftToRightMarkers = 10
                        
End Enum

' ----==== Variablen ====----
Private m_pIPhoneNumberFormatter As Long
Private m_pIPhoneNumberFormatterStatics As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Globalization_PhoneNumberFormatting_PhoneNumberFormatter, _
                            IID_IPhoneNumberFormatterStatics, _
                            m_pIPhoneNumberFormatterStatics) Then
    End If
    Dim pIPhoneNumberFormatter As Long
    If GetActivateInstance(Windows_Globalization_PhoneNumberFormatting_PhoneNumberFormatter, _
                           IID_IPhoneNumberFormatter, _
                           pIPhoneNumberFormatter) Then
        Me.Ifc = pIPhoneNumberFormatter
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIPhoneNumberFormatter)
    Call ReleaseIfc(m_pIPhoneNumberFormatterStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIPhoneNumberFormatter
End Property

Public Property Let Ifc(ByVal pIPhoneNumberFormatter As Long)
    Call ReleaseIfc(m_pIPhoneNumberFormatter)
    m_pIPhoneNumberFormatter = pIPhoneNumberFormatter
End Property

' ----==== Functions ====----
' IPhoneNumberFormatterStatics
Public Function TryCreate(ByVal RegionCode As String) As PhoneNumberFormatter
    Dim Ret As PhoneNumberFormatter
    If m_pIPhoneNumberFormatterStatics <> 0& Then
        Dim pIPhoneNumberFormatter As Long
        Dim hRegionCode As Long
        hRegionCode = CreateWindowsString(RegionCode)
        If Invoke(m_pIPhoneNumberFormatterStatics, _
                  IPhoneNumberFormatterStatics_TryCreate, _
                  hRegionCode, _
                  VarPtr(pIPhoneNumberFormatter)) = S_OK Then
            If pIPhoneNumberFormatter <> 0& Then
                Set Ret = New PhoneNumberFormatter
                Ret.Ifc = pIPhoneNumberFormatter
            End If
        End If
        Call DeleteWindowsString(hRegionCode)
    End If
    Set TryCreate = Ret
End Function

Public Function GetCountryCodeForRegion(ByVal RegionCode As String) As Long
    Dim Ret As Long
    If m_pIPhoneNumberFormatterStatics <> 0& Then
        Dim result As Long
        Dim hRegionCode As Long
        hRegionCode = CreateWindowsString(RegionCode)
        If Invoke(m_pIPhoneNumberFormatterStatics, _
                  IPhoneNumberFormatterStatics_GetCountryCodeForRegion, _
                  hRegionCode, _
                  VarPtr(result)) = S_OK Then
            Ret = result
        End If
        Call DeleteWindowsString(hRegionCode)
    End If
    GetCountryCodeForRegion = Ret
End Function

Public Function GetNationalDirectDialingPrefixForRegion(ByVal RegionCode As String, _
                                                        ByVal StripNonDigit As Boolean) As String
    Dim Ret As String
    If m_pIPhoneNumberFormatterStatics <> 0& Then
        Dim result As Long
        Dim hRegionCode As Long
        hRegionCode = CreateWindowsString(RegionCode)
        If Invoke(m_pIPhoneNumberFormatterStatics, _
                  IPhoneNumberFormatterStatics_GetNationalDirectDialingPrefixForRegion, _
                  hRegionCode, _
                  StripNonDigit, _
                  VarPtr(result)) = S_OK Then
            Ret = GetWindowsString(result)
        End If
        Call DeleteWindowsString(hRegionCode)
    End If
    GetNationalDirectDialingPrefixForRegion = Ret
End Function

Public Function WrapWithLeftToRightMarkers(ByVal Number As String) As String
    Dim Ret As String
    If m_pIPhoneNumberFormatterStatics <> 0& Then
        Dim result As Long
        Dim hNumber As Long
        hNumber = CreateWindowsString(Number)
        If Invoke(m_pIPhoneNumberFormatterStatics, _
                  IPhoneNumberFormatterStatics_WrapWithLeftToRightMarkers, _
                  hNumber, _
                  VarPtr(result)) = S_OK Then
            Ret = GetWindowsString(result)
        End If
        Call DeleteWindowsString(hNumber)
    End If
    WrapWithLeftToRightMarkers = Ret
End Function
                        
' IPhoneNumberFormatter
Public Function Format(ByVal Number As PhoneNumberInfo) As String
    Dim Ret As String
    If m_pIPhoneNumberFormatter <> 0& And _
       IsNotNothing(Number) Then
        Dim result As Long
        If Invoke(m_pIPhoneNumberFormatter, _
                  IPhoneNumberFormatter_Format, _
                  Number.Ifc, _
                  VarPtr(result)) = S_OK Then
            Ret = GetWindowsString(result)
        End If
    End If
    Format = Ret
End Function

Public Function FormatWithOutputFormat(ByVal Number As PhoneNumberInfo, _
                                       ByVal NumberFormat As PhoneNumberFormat) As String
    Dim Ret As String
    If m_pIPhoneNumberFormatter <> 0& And _
       IsNotNothing(Number) Then
        Dim result As Long
        If Invoke(m_pIPhoneNumberFormatter, _
                  IPhoneNumberFormatter_FormatWithOutputFormat, _
                  Number.Ifc, _
                  NumberFormat, _
                  VarPtr(result)) = S_OK Then
            Ret = GetWindowsString(result)
        End If
    End If
    FormatWithOutputFormat = Ret
End Function

Public Function FormatPartialString(ByVal Number As String) As String
    Dim Ret As String
    If m_pIPhoneNumberFormatter <> 0& Then
        Dim result As Long
        Dim hNumber As Long
        hNumber = CreateWindowsString(Number)
        If Invoke(m_pIPhoneNumberFormatter, _
                  IPhoneNumberFormatter_FormatPartialString, _
                  hNumber, _
                  VarPtr(result)) = S_OK Then
            Ret = GetWindowsString(result)
        End If
        Call DeleteWindowsString(hNumber)
    End If
    FormatPartialString = Ret
End Function

Public Function FormatString(ByVal Number As String) As String
    Dim Ret As String
    If m_pIPhoneNumberFormatter <> 0& Then
        Dim result As Long
        Dim hNumber As Long
        hNumber = CreateWindowsString(Number)
        If Invoke(m_pIPhoneNumberFormatter, _
                  IPhoneNumberFormatter_FormatString, _
                  hNumber, _
                  VarPtr(result)) = S_OK Then
            Ret = GetWindowsString(result)
        End If
        Call DeleteWindowsString(hNumber)
    End If
    FormatString = Ret
End Function

Public Function FormatStringWithLeftToRightMarkers(ByVal Number As String) As String
    Dim Ret As String
    If m_pIPhoneNumberFormatter <> 0& Then
        Dim result As Long
        Dim hNumber As Long
        hNumber = CreateWindowsString(Number)
        If Invoke(m_pIPhoneNumberFormatter, _
                  IPhoneNumberFormatter_FormatStringWithLeftToRightMarkers, _
                  hNumber, _
                  VarPtr(result)) = S_OK Then
            Ret = GetWindowsString(result)
        End If
        Call DeleteWindowsString(hNumber)
    End If
    FormatStringWithLeftToRightMarkers = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PhoneNumberInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Globalization_PhoneNumberFormatting_PhoneNumberInfo As String = "Windows.Globalization.PhoneNumberFormatting.PhoneNumberInfo"
Private Const IID_IStringable As String = "{96369f54-8eb6-48f0-abce-c1b211e627c3}"
Private Const IID_IPhoneNumberInfo As String = "{1c7ce4dd-c8b4-4ea3-9aef-b342e2c5b417}"
Private Const IID_IPhoneNumberInfoStatics As String = "{5b3f4f6a-86a9-40e9-8649-6d61161928d4}"
Private Const IID_IPhoneNumberInfoFactory As String = "{8202b964-adaa-4cff-8fcf-17e7516a28ff}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IPhoneNumberInfoFactory
    IPhoneNumberInfoFactory_Create = 6

    ' IPhoneNumberInfoStatics
    IPhoneNumberInfoStatics_TryParse = 6
    IPhoneNumberInfoStatics_TryParseWithRegion = 7
                        
    ' IPhoneNumberInfo
    IPhoneNumberInfo_GetCountryCode = 6
    IPhoneNumberInfo_GetPhoneNumber = 7
    IPhoneNumberInfo_GetLengthOfGeographicalAreaCode = 8
    IPhoneNumberInfo_GetNationalSignificantNumber = 9
    IPhoneNumberInfo_GetLengthOfNationalDestinationCode = 10
    IPhoneNumberInfo_PredictNumberKind = 11
    IPhoneNumberInfo_GetGeographicRegionCode = 12
    IPhoneNumberInfo_CheckNumberMatch = 13
    
    ' IStringable
    IStringable_ToString = 6
                        
End Enum

' ----==== Variablen ====----
Private m_pIStringable As Long
Private m_pIPhoneNumberInfo As Long
Private m_pIPhoneNumberInfoStatics As Long
Private m_pIPhoneNumberInfoFactory As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Globalization_PhoneNumberFormatting_PhoneNumberInfo, _
                            IID_IPhoneNumberInfoFactory, _
                            m_pIPhoneNumberInfoFactory) Then
    End If
    If GetActivationFactory(Windows_Globalization_PhoneNumberFormatting_PhoneNumberInfo, _
                            IID_IPhoneNumberInfoStatics, _
                            m_pIPhoneNumberInfoStatics) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIStringable)
    Call ReleaseIfc(m_pIPhoneNumberInfo)
    Call ReleaseIfc(m_pIPhoneNumberInfoStatics)
    Call ReleaseIfc(m_pIPhoneNumberInfoFactory)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIPhoneNumberInfo
End Property

Public Property Let Ifc(ByVal pIPhoneNumberInfo As Long)
    Call ReleaseIfc(m_pIStringable)
    Call ReleaseIfc(m_pIPhoneNumberInfo)
    m_pIPhoneNumberInfo = pIPhoneNumberInfo
    Call QueryIfc(m_pIPhoneNumberInfo, IID_IStringable, m_pIStringable)
End Property

' IPhoneNumberInfo
Public Property Get CountryCode() As Long
    Dim Ret As Long
    If m_pIPhoneNumberInfo <> 0& Then
        Dim value As Long
        If Invoke(m_pIPhoneNumberInfo, _
                  IPhoneNumberInfo_GetCountryCode, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    CountryCode = Ret
End Property

Public Property Get PhoneNumber() As String
    Dim Ret As String
    If m_pIPhoneNumberInfo <> 0& Then
        Dim value As Long
        If Invoke(m_pIPhoneNumberInfo, _
                  IPhoneNumberInfo_GetPhoneNumber, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    PhoneNumber = Ret
End Property

' ----==== Functions ====----
' IPhoneNumberInfoFactory
Public Function Create(ByVal Number As String) As PhoneNumberInfo
    Dim Ret As PhoneNumberInfo
    If m_pIPhoneNumberInfoFactory <> 0& Then
        Dim pIPhoneNumberInfo As Long
        Dim hNumber As Long
        hNumber = CreateWindowsString(Number)
        If Invoke(m_pIPhoneNumberInfoFactory, _
                  IPhoneNumberInfoFactory_Create, _
                  hNumber, _
                  VarPtr(pIPhoneNumberInfo)) = S_OK Then
            If pIPhoneNumberInfo <> 0& Then
                Set Ret = New PhoneNumberInfo
                Ret.Ifc = pIPhoneNumberInfo
            End If
        End If
        Call DeleteWindowsString(hNumber)
    End If
    Set Create = Ret
End Function

' IPhoneNumberInfoStatics
Public Function TryParse(ByVal InputStr As String, _
                         ByVal PhoneNumber As PhoneNumberInfo) As PhoneNumberParseResult
    Dim Ret As PhoneNumberParseResult
    If m_pIPhoneNumberInfoStatics <> 0& And _
       IsNotNothing(PhoneNumber) Then
        Dim result As Long
        Dim hInputStr As Long
        hInputStr = CreateWindowsString(InputStr)
        If Invoke(m_pIPhoneNumberInfoStatics, _
                  IPhoneNumberInfoStatics_TryParse, _
                  hInputStr, _
                  PhoneNumber.Ifc, _
                  VarPtr(result)) = S_OK Then
            Ret = result
        End If
        Call DeleteWindowsString(hInputStr)
    End If
    TryParse = Ret
End Function

Public Function TryParseWithRegion(ByVal InputStr As String, _
                                   ByVal RegionCode As String, _
                                   ByVal PhoneNumber As PhoneNumberInfo) As PhoneNumberParseResult
    Dim Ret As PhoneNumberParseResult
    If m_pIPhoneNumberInfoStatics <> 0& And _
       IsNotNothing(PhoneNumber) Then
        Dim result As Long
        Dim hInputStr As Long
        Dim hRegionCode As Long
        hInputStr = CreateWindowsString(InputStr)
        hRegionCode = CreateWindowsString(RegionCode)
        If Invoke(m_pIPhoneNumberInfoStatics, _
                  IPhoneNumberInfoStatics_TryParseWithRegion, _
                  hInputStr, _
                  hRegionCode, _
                  PhoneNumber.Ifc, _
                  VarPtr(result)) = S_OK Then
            Ret = result
        End If
        Call DeleteWindowsString(hRegionCode)
        Call DeleteWindowsString(hInputStr)
    End If
    TryParseWithRegion = Ret
End Function

' IPhoneNumberInfo
Public Function GetLengthOfGeographicalAreaCode() As Long
    Dim Ret As Long
    If m_pIPhoneNumberInfo <> 0& Then
        Dim value As Long
        If Invoke(m_pIPhoneNumberInfo, _
                  IPhoneNumberInfo_GetLengthOfGeographicalAreaCode, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    GetLengthOfGeographicalAreaCode = Ret
End Function

Public Function GetNationalSignificantNumber() As String
    Dim Ret As String
    If m_pIPhoneNumberInfo <> 0& Then
        Dim value As Long
        If Invoke(m_pIPhoneNumberInfo, _
                  IPhoneNumberInfo_GetNationalSignificantNumber, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    GetNationalSignificantNumber = Ret
End Function

Public Function GetLengthOfNationalDestinationCode() As Long
    Dim Ret As Long
    If m_pIPhoneNumberInfo <> 0& Then
        Dim value As Long
        If Invoke(m_pIPhoneNumberInfo, _
                  IPhoneNumberInfo_GetLengthOfNationalDestinationCode, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    GetLengthOfNationalDestinationCode = Ret
End Function

Public Function PredictNumberKind() As PredictedPhoneNumberKind
    Dim Ret As PredictedPhoneNumberKind
    If m_pIPhoneNumberInfo <> 0& Then
        Dim value As Long
        If Invoke(m_pIPhoneNumberInfo, _
                  IPhoneNumberInfo_PredictNumberKind, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    PredictNumberKind = Ret
End Function

Public Function GetGeographicRegionCode() As String
    Dim Ret As String
    If m_pIPhoneNumberInfo <> 0& Then
        Dim value As Long
        If Invoke(m_pIPhoneNumberInfo, _
                  IPhoneNumberInfo_GetGeographicRegionCode, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    GetGeographicRegionCode = Ret
End Function

Public Function CheckNumberMatch(ByVal OtherNumber As PhoneNumberInfo) As PhoneNumberMatchResult
    Dim Ret As PhoneNumberMatchResult
    If m_pIPhoneNumberInfo <> 0& And _
       IsNotNothing(OtherNumber) Then
        Dim value As Long
        If Invoke(m_pIPhoneNumberInfo, _
                  IPhoneNumberInfo_CheckNumberMatch, _
                  OtherNumber.Ifc, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    CheckNumberMatch = Ret
End Function

' IStringable
Public Function ToString() As String
    Dim Ret As String
    If m_pIStringable <> 0& Then
        Dim hString As Long
        If Invoke(m_pIStringable, _
                  IStringable_ToString, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    ToString = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



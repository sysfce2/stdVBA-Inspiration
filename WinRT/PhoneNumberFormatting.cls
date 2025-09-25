VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PhoneNumberFormatting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

' Namespace PhoneNumberFormatting

Option Explicit

' ----==== Enums Windows.Globalization.PhoneNumberFormatting.xxx ====----
Public Enum PhoneNumberFormat
    PhoneNumberFormat_E164 = 0
    PhoneNumberFormat_International = 1
    PhoneNumberFormat_National = 2
    PhoneNumberFormat_Rfc3966 = 3
End Enum

Public Enum PhoneNumberMatchResult
    PhoneNumberMatchResult_NoMatch = 0
    PhoneNumberMatchResult_ShortNationalSignificantNumberMatch = 1
    PhoneNumberMatchResult_NationalSignificantNumberMatch = 2
    PhoneNumberMatchResult_ExactMatch = 3
End Enum

Public Enum PhoneNumberParseResult
    PhoneNumberParseResult_Valid = 0
    PhoneNumberParseResult_NotANumber = 1
    PhoneNumberParseResult_InvalidCountryCode = 2
    PhoneNumberParseResult_TooShort = 3
    PhoneNumberParseResult_TooLong = 4
End Enum

Public Enum PredictedPhoneNumberKind
    PredictedPhoneNumberKind_FixedLine = 0
    PredictedPhoneNumberKind_Mobile = 1
    PredictedPhoneNumberKind_FixedLineOrMobile = 2
    PredictedPhoneNumberKind_TollFree = 3
    PredictedPhoneNumberKind_PremiumRate = 4
    PredictedPhoneNumberKind_SharedCost = 5
    PredictedPhoneNumberKind_Voip = 6
    PredictedPhoneNumberKind_PersonalNumber = 7
    PredictedPhoneNumberKind_Pager = 8
    PredictedPhoneNumberKind_UniversalAccountNumber = 9
    PredictedPhoneNumberKind_Voicemail = 10
    PredictedPhoneNumberKind_Unknown = 11
End Enum

' ----==== Namespace PhoneNumberFormatting.PhoneNumberFormatter ====----
Public Property Get PhoneNumberFormatter() As PhoneNumberFormatter
    Set PhoneNumberFormatter = New PhoneNumberFormatter
End Property

' ----==== Namespace PhoneNumberFormatting.PhoneNumberInfo ====----
Public Property Get PhoneNumberInfo() As PhoneNumberInfo
    Set PhoneNumberInfo = New PhoneNumberInfo
End Property


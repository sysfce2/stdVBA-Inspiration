VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CurrencyFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Globalization_NumberFormatting_CurrencyFormatter As String = "Windows.Globalization.NumberFormatting.CurrencyFormatter"
Private Const IID_ICurrencyFormatter As String = "{11730ca5-4b00-41b2-b332-73b12a497d54}"
Private Const IID_ICurrencyFormatter2 As String = "{072c2f1d-e7ba-4197-920e-247c92f7dea6}"
Private Const IID_INumberFormatter As String = "{a5007c49-7676-4db7-8631-1b6ff265caa9}"
Private Const IID_INumberFormatter2 As String = "{d4a8c1f0-80d0-4b0d-a89e-882c1e8f8310}"
Private Const IID_INumberFormatterOptions As String = "{80332d21-aee1-4a39-baa2-07ed8c96daf6}"
Private Const IID_INumberParser As String = "{e6659412-4a13-4a53-83a1-392fbe4cff9f}"
Private Const IID_INumberRounderOption As String = "{3b088433-646f-4efe-8d48-66eb2e49e736}"
Private Const IID_ISignedZeroOption As String = "{fd1cdd31-0a3c-49c4-a642-96a1564f4f30}"
Private Const IID_ISignificantDigitsOption As String = "{1d4dfcdd-2d43-4ee8-bbf1-c1b26a711a58}"
Private Const IID_ICurrencyFormatterFactory As String = "{86c7537e-b938-4aa2-84b0-2c33dc5b1450}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' ICurrencyFormatterFactory
    ICurrencyFormatterFactory_CreateCurrencyFormatterCode = 6
    ICurrencyFormatterFactory_CreateCurrencyFormatterCodeContext = 7
                        
    ' ICurrencyFormatter
    ICurrencyFormatter_GetCurrency = 6
'    DEPRECATED ("Currency may be read-only for releases after Windows 8.1. Instead, use a new CurrencyFormatter.")
'    ICurrencyFormatter_PutCurrency = 7
                        
    ' ICurrencyFormatter2
    ICurrencyFormatter2_GetMode = 6
    ICurrencyFormatter2_PutMode = 7
    ICurrencyFormatter2_ApplyRoundingForCurrency = 8
    
    ' INumberFormatter
    INumberFormatter_FormatInt = 6
    INumberFormatter_FormatUInt = 7
    INumberFormatter_FormatDouble = 8
                        
    ' INumberFormatter2
    INumberFormatter2_FormatInt = 6
    INumberFormatter2_FormatUInt = 7
    INumberFormatter2_FormatDouble = 8
                        
    ' INumberFormatterOptions
    INumberFormatterOptions_GetLanguages = 6
    INumberFormatterOptions_GetGeographicRegion = 7
    INumberFormatterOptions_GetIntegerDigits = 8
    INumberFormatterOptions_PutIntegerDigits = 9
    INumberFormatterOptions_GetFractionDigits = 10
    INumberFormatterOptions_PutFractionDigits = 11
    INumberFormatterOptions_GetIsGrouped = 12
    INumberFormatterOptions_PutIsGrouped = 13
    INumberFormatterOptions_GetIsDecimalPointAlwaysDisplayed = 14
    INumberFormatterOptions_PutIsDecimalPointAlwaysDisplayed = 15
    INumberFormatterOptions_GetNumeralSystem = 16
    INumberFormatterOptions_PutNumeralSystem = 17
    INumberFormatterOptions_GetResolvedLanguage = 18
    INumberFormatterOptions_GetResolvedGeographicRegion = 19
                        
    ' INumberParser
    INumberParser_ParseInt = 6
    INumberParser_ParseUInt = 7
    INumberParser_ParseDouble = 8
                        
    ' INumberRounderOption
    INumberRounderOption_GetNumberRounder = 6
    INumberRounderOption_PutNumberRounder = 7
    
    ' ISignedZeroOption
    ISignedZeroOption_GetIsZeroSigned = 6
    ISignedZeroOption_PutIsZeroSigned = 7
                        
    ' ISignificantDigitsOption
    ISignificantDigitsOption_GetSignificantDigits = 6
    ISignificantDigitsOption_PutSignificantDigits = 7
    
    ' IReference
    IReference_GetValue = 6
    
End Enum

' ----==== Variablen ====----
Private m_pISignedZeroOption As Long
Private m_pISignificantDigitsOption As Long
Private m_pINumberParser As Long
Private m_pINumberRounderOption As Long
Private m_pINumberFormatter As Long
Private m_pINumberFormatter2 As Long
Private m_pINumberFormatterOptions As Long
Private m_pICurrencyFormatter As Long
Private m_pICurrencyFormatter2 As Long
Private m_pICurrencyFormatterFactory As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Globalization_NumberFormatting_CurrencyFormatter, _
                            IID_ICurrencyFormatterFactory, _
                            m_pICurrencyFormatterFactory) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pISignedZeroOption)
    Call ReleaseIfc(m_pISignificantDigitsOption)
    Call ReleaseIfc(m_pINumberRounderOption)
    Call ReleaseIfc(m_pINumberParser)
    Call ReleaseIfc(m_pINumberFormatterOptions)
    Call ReleaseIfc(m_pINumberFormatter2)
    Call ReleaseIfc(m_pINumberFormatter)
    Call ReleaseIfc(m_pICurrencyFormatter2)
    Call ReleaseIfc(m_pICurrencyFormatter)
    Call ReleaseIfc(m_pICurrencyFormatterFactory)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pICurrencyFormatter
End Property

Public Property Let Ifc(ByVal pICurrencyFormatter As Long)
    Call ReleaseIfc(m_pISignedZeroOption)
    Call ReleaseIfc(m_pISignificantDigitsOption)
    Call ReleaseIfc(m_pINumberRounderOption)
    Call ReleaseIfc(m_pINumberParser)
    Call ReleaseIfc(m_pINumberFormatterOptions)
    Call ReleaseIfc(m_pINumberFormatter2)
    Call ReleaseIfc(m_pINumberFormatter)
    Call ReleaseIfc(m_pICurrencyFormatter2)
    Call ReleaseIfc(m_pICurrencyFormatter)
    m_pICurrencyFormatter = pICurrencyFormatter
    Call QueryIfc(m_pICurrencyFormatter, IID_ICurrencyFormatter2, m_pICurrencyFormatter2)
    Call QueryIfc(m_pICurrencyFormatter, IID_INumberFormatter, m_pINumberFormatter)
    Call QueryIfc(m_pICurrencyFormatter, IID_INumberFormatter2, m_pINumberFormatter2)
    Call QueryIfc(m_pICurrencyFormatter, IID_INumberFormatterOptions, m_pINumberFormatterOptions)
    Call QueryIfc(m_pICurrencyFormatter, IID_INumberParser, m_pINumberParser)
    Call QueryIfc(m_pICurrencyFormatter, IID_INumberRounderOption, m_pINumberRounderOption)
    Call QueryIfc(m_pICurrencyFormatter, IID_ISignificantDigitsOption, m_pISignificantDigitsOption)
    Call QueryIfc(m_pICurrencyFormatter, IID_ISignedZeroOption, m_pISignedZeroOption)
End Property

' ICurrencyFormatter
Public Property Get CurrencyFormat() As String
    Dim Ret As String
    If m_pICurrencyFormatter <> 0& Then
        Dim value As Long
        If Invoke(m_pICurrencyFormatter, _
                  ICurrencyFormatter_GetCurrency, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    CurrencyFormat = Ret
End Property

'    DEPRECATED ("Currency may be read-only for releases after Windows 8.1. Instead, use a new CurrencyFormatter.")
'    ICurrencyFormatter_PutCurrency = 7

' ICurrencyFormatter2
Public Property Get Mode() As CurrencyFormatterMode
    Dim Ret As CurrencyFormatterMode
    If m_pICurrencyFormatter2 <> 0& Then
        Dim value As Long
        If Invoke(m_pICurrencyFormatter2, _
                  ICurrencyFormatter2_GetMode, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Mode = Ret
End Property

Public Property Let Mode(ByVal value As CurrencyFormatterMode)
    If m_pICurrencyFormatter2 <> 0& Then
        If Invoke(m_pICurrencyFormatter2, _
                  ICurrencyFormatter2_PutMode, _
                  value) = S_OK Then
        End If
    End If
End Property
    
' INumberFormatterOptions
Public Property Get Languages() As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pINumberFormatterOptions <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pINumberFormatterOptions, _
                  INumberFormatterOptions_GetLanguages, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_String, pIVectorView) 'ReadOnlyList_String
            End If
        End If
    End If
    Set Languages = Ret
End Property

Public Property Get GeographicRegion() As String
    Dim Ret As String
    If m_pINumberFormatterOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pINumberFormatterOptions, _
                  INumberFormatterOptions_GetGeographicRegion, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    GeographicRegion = Ret
End Property

Public Property Get IntegerDigits() As Long
    Dim Ret As Long
    If m_pINumberFormatterOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pINumberFormatterOptions, _
                  INumberFormatterOptions_GetIntegerDigits, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    IntegerDigits = Ret
End Property

Public Property Let IntegerDigits(ByVal value As Long)
    If m_pINumberFormatterOptions <> 0& Then
        If Invoke(m_pINumberFormatterOptions, _
                  INumberFormatterOptions_PutIntegerDigits, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get FractionDigits() As Long
    Dim Ret As Long
    If m_pINumberFormatterOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pINumberFormatterOptions, _
                  INumberFormatterOptions_GetFractionDigits, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    FractionDigits = Ret
End Property

Public Property Let FractionDigits(ByVal value As Long)
    If m_pINumberFormatterOptions <> 0& Then
        If Invoke(m_pINumberFormatterOptions, _
                  INumberFormatterOptions_PutFractionDigits, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get IsGrouped() As Boolean
    Dim Ret As Boolean
    If m_pINumberFormatterOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pINumberFormatterOptions, _
                  INumberFormatterOptions_GetIsGrouped, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsGrouped = Ret
End Property

Public Property Let IsGrouped(ByVal value As Boolean)
    If m_pINumberFormatterOptions <> 0& Then
        If Invoke(m_pINumberFormatterOptions, _
                  INumberFormatterOptions_PutIsGrouped, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get IsDecimalPointAlwaysDisplayed() As Boolean
    Dim Ret As Boolean
    If m_pINumberFormatterOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pINumberFormatterOptions, _
                  INumberFormatterOptions_GetIsDecimalPointAlwaysDisplayed, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsDecimalPointAlwaysDisplayed = Ret
End Property

Public Property Let IsDecimalPointAlwaysDisplayed(ByVal value As Boolean)
    If m_pINumberFormatterOptions <> 0& Then
        If Invoke(m_pINumberFormatterOptions, _
                  INumberFormatterOptions_PutIsDecimalPointAlwaysDisplayed, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get NumeralSystem() As String
    Dim Ret As String
    If m_pINumberFormatterOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pINumberFormatterOptions, _
                  INumberFormatterOptions_GetNumeralSystem, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    NumeralSystem = Ret
End Property

Public Property Let NumeralSystem(ByVal value As String)
    If m_pINumberFormatterOptions <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pINumberFormatterOptions, _
                  INumberFormatterOptions_PutNumeralSystem, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get ResolvedLanguage() As String
    Dim Ret As String
    If m_pINumberFormatterOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pINumberFormatterOptions, _
                  INumberFormatterOptions_GetResolvedLanguage, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    ResolvedLanguage = Ret
End Property

Public Property Get ResolvedGeographicRegion() As String
    Dim Ret As String
    If m_pINumberFormatterOptions <> 0& Then
        Dim value As Long
        If Invoke(m_pINumberFormatterOptions, _
                  INumberFormatterOptions_GetResolvedGeographicRegion, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    ResolvedGeographicRegion = Ret
End Property
                        
Public Property Get NumberRounder() As NumberRounder
    Dim Ret As NumberRounder
    If m_pINumberRounderOption <> 0& Then
        Dim pINumberRounder As Long
        If Invoke(m_pINumberRounderOption, _
                  INumberRounderOption_GetNumberRounder, _
                  VarPtr(pINumberRounder)) = S_OK Then
            If pINumberRounder <> 0& Then
                Set Ret = New NumberRounder
                Ret.Ifc = pINumberRounder
            End If
        End If
    End If
    Set NumberRounder = Ret
End Property

Public Property Let NumberRounder(ByVal value As NumberRounder)
    If m_pINumberRounderOption <> 0& Then
        If Invoke(m_pINumberRounderOption, _
                  INumberRounderOption_PutNumberRounder, _
                  value.Ifc) = S_OK Then
        End If
    End If
End Property
    
' ISignedZeroOption
Public Property Get IsZeroSigned() As Boolean
    Dim Ret As Boolean
    If m_pISignedZeroOption <> 0& Then
        Dim value As Long
        If Invoke(m_pISignedZeroOption, _
                  ISignedZeroOption_GetIsZeroSigned, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsZeroSigned = Ret
End Property

Public Property Let IsZeroSigned(ByVal value As Boolean)
    If m_pISignedZeroOption <> 0& Then
        If Invoke(m_pISignedZeroOption, _
                  ISignedZeroOption_PutIsZeroSigned, _
                  value) = S_OK Then
        End If
    End If
End Property
                        
' ISignificantDigitsOption
Public Property Get SignificantDigits() As Long
    Dim Ret As Long
    If m_pISignificantDigitsOption <> 0& Then
        Dim value As Long
        If Invoke(m_pISignificantDigitsOption, _
                  ISignificantDigitsOption_GetSignificantDigits, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    SignificantDigits = Ret
End Property

Public Property Let SignificantDigits(ByVal value As Long)
    If m_pISignificantDigitsOption <> 0& Then
        If Invoke(m_pISignificantDigitsOption, _
                  ISignificantDigitsOption_PutSignificantDigits, _
                  value) = S_OK Then
        End If
    End If
End Property

' ----==== Functions ====----
' ICurrencyFormatterFactory
Public Function CreateCurrencyFormatterCode(ByVal CurrencyCode As String) As CurrencyFormatter
    Dim Ret As CurrencyFormatter
    If m_pICurrencyFormatterFactory <> 0& Then
        Dim pICurrencyFormatter As Long
        Dim hCurrencyCode As Long
        hCurrencyCode = CreateWindowsString(CurrencyCode)
        If Invoke(m_pICurrencyFormatterFactory, _
                  ICurrencyFormatterFactory_CreateCurrencyFormatterCode, _
                  hCurrencyCode, _
                  VarPtr(pICurrencyFormatter)) = S_OK Then
            If pICurrencyFormatter <> 0& Then
                Set Ret = New CurrencyFormatter
                Ret.Ifc = pICurrencyFormatter
            End If
        End If
        Call DeleteWindowsString(hCurrencyCode)
    End If
    Set CreateCurrencyFormatterCode = Ret
End Function

Public Function CreateCurrencyFormatterCodeContext(ByVal CurrencyCode As String, _
                                                   ByVal Languages As List_String, _
                                                   ByVal GeographicRegion As String) As CurrencyFormatter
    Dim Ret As CurrencyFormatter
    If m_pICurrencyFormatterFactory <> 0& And _
       IsNotNothing(Languages) Then
        Dim pICurrencyFormatter As Long
        Dim hCurrencyCode As Long
        Dim hGeographicRegion As Long
        hCurrencyCode = CreateWindowsString(CurrencyCode)
        hGeographicRegion = CreateWindowsString(GeographicRegion)
        If Invoke(m_pICurrencyFormatterFactory, _
                  ICurrencyFormatterFactory_CreateCurrencyFormatterCodeContext, _
                  hCurrencyCode, _
                  Languages.Ifc, _
                  hGeographicRegion, _
                  VarPtr(pICurrencyFormatter)) = S_OK Then
            If pICurrencyFormatter <> 0& Then
                Set Ret = New CurrencyFormatter
                Ret.Ifc = pICurrencyFormatter
            End If
        End If
        Call DeleteWindowsString(hGeographicRegion)
        Call DeleteWindowsString(hCurrencyCode)
    End If
    Set CreateCurrencyFormatterCodeContext = Ret
End Function
                        
' ICurrencyFormatter2
Public Function ApplyRoundingForCurrency(ByVal value As RoundingAlgorithm) As Boolean
    Dim Ret As Boolean
    If m_pICurrencyFormatter2 <> 0& Then
        If Invoke(m_pICurrencyFormatter2, _
                  ICurrencyFormatter2_ApplyRoundingForCurrency, _
                  value) = S_OK Then
            Ret = True
        End If
    End If
    ApplyRoundingForCurrency = Ret
End Function
    
' INumberFormatter
Public Function FormatInt(ByVal value As Currency) As String
    Dim Ret As String
    If m_pINumberFormatter <> 0& Then
        Dim Result As Long
        If Invoke(m_pINumberFormatter, _
                  INumberFormatter_FormatInt, _
                  value, _
                  VarPtr(Result)) = S_OK Then
            Ret = GetWindowsString(Result)
        End If
    End If
    FormatInt = Ret
End Function

Public Function FormatUInt(ByVal value As Currency) As String
    Dim Ret As String
    If m_pINumberFormatter <> 0& Then
        Dim Result As Long
        If Invoke(m_pINumberFormatter, _
                  INumberFormatter_FormatUInt, _
                  value, _
                  VarPtr(Result)) = S_OK Then
            Ret = GetWindowsString(Result)
        End If
    End If
    FormatUInt = Ret
End Function

Public Function FormatDouble(ByVal value As Double) As String
    Dim Ret As String
    If m_pINumberFormatter <> 0& Then
        Dim Result As Long
        If Invoke(m_pINumberFormatter, _
                  INumberFormatter_FormatDouble, _
                  value, _
                  VarPtr(Result)) = S_OK Then
            Ret = GetWindowsString(Result)
        End If
    End If
    FormatDouble = Ret
End Function
                        
' INumberFormatter2
Public Function FormatInt2(ByVal value As Currency) As String
    Dim Ret As String
    If m_pINumberFormatter2 <> 0& Then
        Dim Result As Long
        If Invoke(m_pINumberFormatter2, _
                  INumberFormatter2_FormatInt, _
                  value, _
                  VarPtr(Result)) = S_OK Then
            Ret = GetWindowsString(Result)
        End If
    End If
    FormatInt2 = Ret
End Function

Public Function FormatUInt2(ByVal value As Currency) As String
    Dim Ret As String
    If m_pINumberFormatter2 <> 0& Then
        Dim Result As Long
        If Invoke(m_pINumberFormatter2, _
                  INumberFormatter2_FormatUInt, _
                  value, _
                  VarPtr(Result)) = S_OK Then
            Ret = GetWindowsString(Result)
        End If
    End If
    FormatUInt2 = Ret
End Function

Public Function FormatDouble2(ByVal value As Double) As String
    Dim Ret As String
    If m_pINumberFormatter2 <> 0& Then
        Dim Result As Long
        If Invoke(m_pINumberFormatter2, _
                  INumberFormatter2_FormatDouble, _
                  value, _
                  VarPtr(Result)) = S_OK Then
            Ret = GetWindowsString(Result)
        End If
    End If
    FormatDouble2 = Ret
End Function
                                                
' INumberParser
Public Function ParseInt(ByVal Text As String) As Currency
    Dim Ret As Currency
    If m_pINumberParser <> 0& Then
        Dim pIReference As Long
        Dim hText As Long
        hText = CreateWindowsString(Text)
        If Invoke(m_pINumberParser, _
                  INumberParser_ParseInt, _
                  hText, _
                  VarPtr(pIReference)) = S_OK Then
            If pIReference <> 0& Then
                Dim value As Currency
                If Invoke(pIReference, IReference_GetValue, VarPtr(value)) = S_OK Then
                    Ret = value
                End If
                Call ReleaseIfc(pIReference)
            End If
        End If
        Call DeleteWindowsString(hText)
    End If
    ParseInt = Ret
End Function

Public Function ParseUInt(ByVal Text As String) As Currency
    Dim Ret As Currency
    If m_pINumberParser <> 0& Then
        Dim pIReference As Long
        Dim hText As Long
        hText = CreateWindowsString(Text)
        If Invoke(m_pINumberParser, _
                  INumberParser_ParseUInt, _
                  hText, _
                  VarPtr(pIReference)) = S_OK Then
            If pIReference <> 0& Then
                Dim value As Currency
                If Invoke(pIReference, IReference_GetValue, VarPtr(value)) = S_OK Then
                    Ret = value
                End If
                Call ReleaseIfc(pIReference)
            End If
        End If
        Call DeleteWindowsString(hText)
    End If
    ParseUInt = Ret
End Function

Public Function ParseDouble(ByVal Text As String) As Double
    Dim Ret As Double
    If m_pINumberParser <> 0& Then
        Dim pIReference As Long
        Dim hText As Long
        hText = CreateWindowsString(Text)
        If Invoke(m_pINumberParser, _
                  INumberParser_ParseDouble, _
                  hText, _
                  VarPtr(pIReference)) = S_OK Then
            If pIReference <> 0& Then
                Dim value As Double
                If Invoke(pIReference, IReference_GetValue, VarPtr(value)) = S_OK Then
                    Ret = value
                End If
                Call ReleaseIfc(pIReference)
            End If
        End If
        Call DeleteWindowsString(hText)
    End If
    ParseDouble = Ret
End Function
                        
' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



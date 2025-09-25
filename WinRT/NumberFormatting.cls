VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NumberFormatting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

' Namespace NumberFormatting

Option Explicit

' ----==== Enums Windows.Globalization.NumberFormatting.xxx ====----
Public Enum CurrencyFormatterMode
    CurrencyFormatterMode_UseSymbol = 0
    CurrencyFormatterMode_UseCurrencyCode = 1
End Enum

Public Enum RoundingAlgorithm
    RoundingAlgorithm_None = 0
    RoundingAlgorithm_RoundDown = 1
    RoundingAlgorithm_RoundUp = 2
    RoundingAlgorithm_RoundTowardsZero = 3
    RoundingAlgorithm_RoundAwayFromZero = 4
    RoundingAlgorithm_RoundHalfDown = 5
    RoundingAlgorithm_RoundHalfUp = 6
    RoundingAlgorithm_RoundHalfTowardsZero = 7
    RoundingAlgorithm_RoundHalfAwayFromZero = 8
    RoundingAlgorithm_RoundHalfToEven = 9
    RoundingAlgorithm_RoundHalfToOdd = 10
End Enum

' ----==== Namespace NumberFormatting.CurrencyFormatter ====----
Public Property Get CurrencyFormatter() As CurrencyFormatter
    Set CurrencyFormatter = New CurrencyFormatter
End Property

' ----==== Namespace NumberFormatting.DecimalFormatter ====----
Public Property Get DecimalFormatter() As DecimalFormatter
    Set DecimalFormatter = New DecimalFormatter
End Property

' ----==== Namespace NumberFormatting.PercentFormatter ====----
Public Property Get PercentFormatter() As PercentFormatter
    Set PercentFormatter = New PercentFormatter
End Property

' ----==== Namespace NumberFormatting.PermilleFormatter ====----
Public Property Get PermilleFormatter() As PermilleFormatter
    Set PermilleFormatter = New PermilleFormatter
End Property

' ----==== Namespace NumberFormatting.NumeralSystemTranslator ====----
Public Property Get NumeralSystemTranslator() As NumeralSystemTranslator
    Set NumeralSystemTranslator = New NumeralSystemTranslator
End Property

' ----==== Namespace NumberFormatting.SignificantDigitsNumberRounder ====----
Public Property Get SignificantDigitsNumberRounder() As SignificantDigitsNumberRounder
    Set SignificantDigitsNumberRounder = New SignificantDigitsNumberRounder
End Property

' ----==== Namespace NumberFormatting.IncrementNumberRounder ====----
Public Property Get IncrementNumberRounder() As IncrementNumberRounder
    Set IncrementNumberRounder = New IncrementNumberRounder
End Property


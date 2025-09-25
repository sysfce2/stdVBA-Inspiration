VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Globalization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

' Namespace Globalization

Option Explicit

' ----==== Enums Windows.Globalization.xxx ====----
Public Enum DayOfWeek
    DayOfWeek_Sunday = 0
    DayOfWeek_Monday = 1
    DayOfWeek_Tuesday = 2
    DayOfWeek_Wednesday = 3
    DayOfWeek_Thursday = 4
    DayOfWeek_Friday = 5
    DayOfWeek_Saturday = 6
End Enum

Public Enum LanguageLayoutDirection
    LanguageLayoutDirection_Ltr = 0
    LanguageLayoutDirection_Rtl = 1
    LanguageLayoutDirection_TtbLtr = 2
    LanguageLayoutDirection_TtbRtl = 3
End Enum

' ----==== Namespace Globalization.Calendar ====----
Public Property Get Calendar() As Calendar
    Set Calendar = New Calendar
End Property

' ----==== Namespace Globalization.Language ====----
Public Property Get Language() As Language
    Set Language = New Language
End Property

' ----==== Namespace Globalization.DateTimeFormatting ====----
Public Property Get DateTimeFormatting() As DateTimeFormatting
    Set DateTimeFormatting = New DateTimeFormatting
End Property

' ----==== Namespace Globalization.NumberFormatting ====----
Public Property Get NumberFormatting() As NumberFormatting
    Set NumberFormatting = New NumberFormatting
End Property

' ----==== Namespace Globalization.PhoneNumberFormatting ====----
Public Property Get PhoneNumberFormatting() As PhoneNumberFormatting
    Set PhoneNumberFormatting = New PhoneNumberFormatting
End Property

' ----==== Namespace Globalization.GeographicRegion ====----
Public Property Get GeographicRegion() As GeographicRegion
    Set GeographicRegion = New GeographicRegion
End Property

' ----==== Namespace Globalization.Fonts ====----
Public Property Get Fonts() As Fonts
    Set Fonts = New Fonts
End Property


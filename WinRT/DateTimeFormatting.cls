VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DateTimeFormatting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

' Namespace DateTimeFormatting

Option Explicit

' ----==== Enums Windows.Globalization.DateTimeFormatting.xxx ====----
Public Enum DayFormat
    DayFormat_None = 0
    DayFormat_Default = 1
End Enum

Public Enum DayOfWeekFormat
    DayOfWeekFormat_None = 0
    DayOfWeekFormat_Default = 1
    DayOfWeekFormat_Abbreviated = 2
    DayOfWeekFormat_Full = 3
End Enum

Public Enum HourFormat
    HourFormat_None = 0
    HourFormat_Default = 1
End Enum

Public Enum MinuteFormat
    MinuteFormat_None = 0
    MinuteFormat_Default = 1
End Enum

Public Enum MonthFormat
    MonthFormat_None = 0
    MonthFormat_Default = 1
    MonthFormat_Abbreviated = 2
    MonthFormat_Full = 3
    MonthFormat_Numeric = 4
End Enum

Public Enum SecondFormat
    SecondFormat_None = 0
    SecondFormat_Default = 1
End Enum

Public Enum YearFormat
    YearFormat_None = 0
    YearFormat_Default = 1
    YearFormat_Abbreviated = 2
    YearFormat_Full = 3
End Enum

' ----==== Namespace DateTimeFormatting.DateTimeFormatter ====----
Public Property Get DateTimeFormatter() As DateTimeFormatter
    Set DateTimeFormatter = New DateTimeFormatter
End Property


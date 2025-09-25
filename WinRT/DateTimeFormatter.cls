VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DateTimeFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Globalization_DateTimeFormatting_DateTimeFormatter As String = "Windows.Globalization.DateTimeFormatting.DateTimeFormatter"
Private Const IID_IDateTimeFormatter As String = "{95eeca10-73e0-4e4b-a183-3d6ad0ba35ec}"
Private Const IID_IDateTimeFormatter2 As String = "{27c91a86-bdaa-4fd0-9e36-671d5aa5ee03}"
Private Const IID_IDateTimeFormatterFactory As String = "{ec8d8a53-1a2e-412d-8815-3b745fb1a2a0}"
Private Const IID_IDateTimeFormatterStatics As String = "{bfcde7c0-df4c-4a2e-9012-f47daf3f1212}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IDateTimeFormatterFactory
    IDateTimeFormatterFactory_CreateDateTimeFormatter = 6
    IDateTimeFormatterFactory_CreateDateTimeFormatterLanguages = 7
    IDateTimeFormatterFactory_CreateDateTimeFormatterContext = 8
    IDateTimeFormatterFactory_CreateDateTimeFormatterDate = 9
    IDateTimeFormatterFactory_CreateDateTimeFormatterTime = 10
    IDateTimeFormatterFactory_CreateDateTimeFormatterDateTimeLanguages = 11
    IDateTimeFormatterFactory_CreateDateTimeFormatterDateTimeContext = 12
    
    ' IDateTimeFormatterStatics
    IDateTimeFormatterStatics_GetLongDate = 6
    IDateTimeFormatterStatics_GetLongTime = 7
    IDateTimeFormatterStatics_GetShortDate = 8
    IDateTimeFormatterStatics_GetShortTime = 9

    ' IDateTimeFormatter
    IDateTimeFormatter_GetLanguages = 6
    IDateTimeFormatter_GetGeographicRegion = 7
    IDateTimeFormatter_GetCalendar = 8
    IDateTimeFormatter_GetClock = 9
    IDateTimeFormatter_GetNumeralSystem = 10
    IDateTimeFormatter_PutNumeralSystem = 11
    IDateTimeFormatter_GetPatterns = 12
    IDateTimeFormatter_GetTemplate = 13
    IDateTimeFormatter_Format = 14
    IDateTimeFormatter_GetIncludeYear = 15
    IDateTimeFormatter_GetIncludeMonth = 16
    IDateTimeFormatter_GetIncludeDayOfWeek = 17
    IDateTimeFormatter_GetIncludeDay = 18
    IDateTimeFormatter_GetIncludeHour = 19
    IDateTimeFormatter_GetIncludeMinute = 20
    IDateTimeFormatter_GetIncludeSecond = 21
    IDateTimeFormatter_GetResolvedLanguage = 22
    IDateTimeFormatter_GetResolvedGeographicRegion = 23

    ' IDateTimeFormatter2
    IDateTimeFormatter2_FormatUsingTimeZone = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIDateTimeFormatter As Long
Private m_pIDateTimeFormatter2 As Long
Private m_pIDateTimeFormatterFactory As Long
Private m_pIDateTimeFormatterStatics As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Globalization_DateTimeFormatting_DateTimeFormatter, _
                            IID_IDateTimeFormatterFactory, _
                            m_pIDateTimeFormatterFactory) Then
    End If
    If GetActivationFactory(Windows_Globalization_DateTimeFormatting_DateTimeFormatter, _
                            IID_IDateTimeFormatterStatics, _
                            m_pIDateTimeFormatterStatics) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIDateTimeFormatter2)
    Call ReleaseIfc(m_pIDateTimeFormatter)
    Call ReleaseIfc(m_pIDateTimeFormatterStatics)
    Call ReleaseIfc(m_pIDateTimeFormatterFactory)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIDateTimeFormatter
End Property

Public Property Let Ifc(ByVal pIDateTimeFormatter As Long)
    Call ReleaseIfc(m_pIDateTimeFormatter2)
    Call ReleaseIfc(m_pIDateTimeFormatter)
    m_pIDateTimeFormatter = pIDateTimeFormatter
    Call QueryIfc(m_pIDateTimeFormatter, IID_IDateTimeFormatter2, m_pIDateTimeFormatter2)
End Property

' IDateTimeFormatterStatics
Public Property Get LongDate() As DateTimeFormatter
    Dim Ret As DateTimeFormatter
    If m_pIDateTimeFormatterStatics <> 0& Then
        Dim pIDateTimeFormatter As Long
        If Invoke(m_pIDateTimeFormatterStatics, _
                  IDateTimeFormatterStatics_GetLongDate, _
                  VarPtr(pIDateTimeFormatter)) = S_OK Then
            If pIDateTimeFormatter <> 0& Then
                Set Ret = New DateTimeFormatter
                Ret.Ifc = pIDateTimeFormatter
            End If
        End If
    End If
    Set LongDate = Ret
End Property

Public Property Get LongTime() As DateTimeFormatter
    Dim Ret As DateTimeFormatter
    If m_pIDateTimeFormatterStatics <> 0& Then
        Dim pIDateTimeFormatter As Long
        If Invoke(m_pIDateTimeFormatterStatics, _
                  IDateTimeFormatterStatics_GetLongTime, _
                  VarPtr(pIDateTimeFormatter)) = S_OK Then
            If pIDateTimeFormatter <> 0& Then
                Set Ret = New DateTimeFormatter
                Ret.Ifc = pIDateTimeFormatter
            End If
        End If
    End If
    Set LongTime = Ret
End Property

Public Property Get ShortDate() As DateTimeFormatter
    Dim Ret As DateTimeFormatter
    If m_pIDateTimeFormatterStatics <> 0& Then
        Dim pIDateTimeFormatter As Long
        If Invoke(m_pIDateTimeFormatterStatics, _
                  IDateTimeFormatterStatics_GetShortDate, _
                  VarPtr(pIDateTimeFormatter)) = S_OK Then
            If pIDateTimeFormatter <> 0& Then
                Set Ret = New DateTimeFormatter
                Ret.Ifc = pIDateTimeFormatter
            End If
        End If
    End If
    Set ShortDate = Ret
End Property

Public Property Get ShortTime() As DateTimeFormatter
    Dim Ret As DateTimeFormatter
    If m_pIDateTimeFormatterStatics <> 0& Then
        Dim pIDateTimeFormatter As Long
        If Invoke(m_pIDateTimeFormatterStatics, _
                  IDateTimeFormatterStatics_GetShortTime, _
                  VarPtr(pIDateTimeFormatter)) = S_OK Then
            If pIDateTimeFormatter <> 0& Then
                Set Ret = New DateTimeFormatter
                Ret.Ifc = pIDateTimeFormatter
            End If
        End If
    End If
    Set ShortTime = Ret
End Property

' IDateTimeFormatter
Public Property Get Languages() As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pIDateTimeFormatter <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIDateTimeFormatter, _
                  IDateTimeFormatter_GetLanguages, _
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
    If m_pIDateTimeFormatter <> 0& Then
        Dim value As Long
        If Invoke(m_pIDateTimeFormatter, _
                  IDateTimeFormatter_GetGeographicRegion, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    GeographicRegion = Ret
End Property

Public Property Get Calendar() As String
    Dim Ret As String
    If m_pIDateTimeFormatter <> 0& Then
        Dim value As Long
        If Invoke(m_pIDateTimeFormatter, _
                  IDateTimeFormatter_GetCalendar, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Calendar = Ret
End Property

Public Property Get Clock() As String
    Dim Ret As String
    If m_pIDateTimeFormatter <> 0& Then
        Dim value As Long
        If Invoke(m_pIDateTimeFormatter, _
                  IDateTimeFormatter_GetClock, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Clock = Ret
End Property

Public Property Get NumeralSystem() As String
    Dim Ret As String
    If m_pIDateTimeFormatter <> 0& Then
        Dim value As Long
        If Invoke(m_pIDateTimeFormatter, _
                  IDateTimeFormatter_GetNumeralSystem, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    NumeralSystem = Ret
End Property

Public Property Let NumeralSystem(ByVal value As String)
    If m_pIDateTimeFormatter <> 0& Then
        Dim hValue As Long
        hValue = CreateWindowsString(value)
        If Invoke(m_pIDateTimeFormatter, _
                  IDateTimeFormatter_PutNumeralSystem, _
                  hValue) = S_OK Then
        End If
        Call DeleteWindowsString(hValue)
    End If
End Property

'    IDateTimeFormatter_GetPatterns = 12
Public Property Get Patterns() As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pIDateTimeFormatter <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIDateTimeFormatter, _
                  IDateTimeFormatter_GetPatterns, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_String, pIVectorView) 'ReadOnlyList_String
            End If
        End If
    End If
    Set Patterns = Ret
End Property

Public Property Get Template() As String
    Dim Ret As String
    If m_pIDateTimeFormatter <> 0& Then
        Dim value As Long
        If Invoke(m_pIDateTimeFormatter, _
                  IDateTimeFormatter_GetTemplate, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    Template = Ret
End Property

Public Property Get IncludeYear() As YearFormat
    Dim Ret As YearFormat
    If m_pIDateTimeFormatter <> 0& Then
        Dim value As Long
        If Invoke(m_pIDateTimeFormatter, _
                  IDateTimeFormatter_GetIncludeYear, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    IncludeYear = Ret
End Property

Public Property Get IncludeMonth() As MonthFormat
    Dim Ret As MonthFormat
    If m_pIDateTimeFormatter <> 0& Then
        Dim value As Long
        If Invoke(m_pIDateTimeFormatter, _
                  IDateTimeFormatter_GetIncludeMonth, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    IncludeMonth = Ret
End Property

Public Property Get IncludeDayOfWeek() As DayOfWeekFormat
    Dim Ret As DayOfWeekFormat
    If m_pIDateTimeFormatter <> 0& Then
        Dim value As Long
        If Invoke(m_pIDateTimeFormatter, _
                  IDateTimeFormatter_GetIncludeDayOfWeek, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    IncludeDayOfWeek = Ret
End Property

Public Property Get IncludeDay() As DayFormat
    Dim Ret As DayFormat
    If m_pIDateTimeFormatter <> 0& Then
        Dim value As Long
        If Invoke(m_pIDateTimeFormatter, _
                  IDateTimeFormatter_GetIncludeDay, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    IncludeDay = Ret
End Property

Public Property Get IncludeHour() As HourFormat
    Dim Ret As HourFormat
    If m_pIDateTimeFormatter <> 0& Then
        Dim value As Long
        If Invoke(m_pIDateTimeFormatter, _
                  IDateTimeFormatter_GetIncludeHour, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    IncludeHour = Ret
End Property

Public Property Get IncludeMinute() As MinuteFormat
    Dim Ret As MinuteFormat
    If m_pIDateTimeFormatter <> 0& Then
        Dim value As Long
        If Invoke(m_pIDateTimeFormatter, _
                  IDateTimeFormatter_GetIncludeMinute, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    IncludeMinute = Ret
End Property

Public Property Get IncludeSecond() As SecondFormat
    Dim Ret As SecondFormat
    If m_pIDateTimeFormatter <> 0& Then
        Dim value As Long
        If Invoke(m_pIDateTimeFormatter, _
                  IDateTimeFormatter_GetIncludeSecond, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    IncludeSecond = Ret
End Property

Public Property Get ResolvedLanguage() As String
    Dim Ret As String
    If m_pIDateTimeFormatter <> 0& Then
        Dim value As Long
        If Invoke(m_pIDateTimeFormatter, _
                  IDateTimeFormatter_GetResolvedLanguage, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    ResolvedLanguage = Ret
End Property

Public Property Get ResolvedGeographicRegion() As String
    Dim Ret As String
    If m_pIDateTimeFormatter <> 0& Then
        Dim value As Long
        If Invoke(m_pIDateTimeFormatter, _
                  IDateTimeFormatter_GetResolvedGeographicRegion, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    ResolvedGeographicRegion = Ret
End Property

' ----==== Functions ====----
' IDateTimeFormatterFactory
Public Function CreateDateTimeFormatter(ByVal FormatTemplate As String) As DateTimeFormatter
    Dim Ret As DateTimeFormatter
    If m_pIDateTimeFormatterFactory <> 0& Then
        Dim pIDateTimeFormatter As Long
        Dim hFormatTemplate As Long
        hFormatTemplate = CreateWindowsString(FormatTemplate)
        If Invoke(m_pIDateTimeFormatterFactory, _
                  IDateTimeFormatterFactory_CreateDateTimeFormatter, _
                  hFormatTemplate, _
                  VarPtr(pIDateTimeFormatter)) = S_OK Then
            If pIDateTimeFormatter <> 0& Then
                Set Ret = New DateTimeFormatter
                Ret.Ifc = pIDateTimeFormatter
            End If
        End If
        Call DeleteWindowsString(hFormatTemplate)
    End If
    Set CreateDateTimeFormatter = Ret
End Function

Public Function CreateDateTimeFormatterLanguages(ByVal FormatTemplate As String, _
                                                 ByVal Languages As List_String) As DateTimeFormatter
    Dim Ret As DateTimeFormatter
    If m_pIDateTimeFormatterFactory <> 0& And _
       IsNotNothing(Languages) Then
        Dim pIDateTimeFormatter As Long
        Dim hFormatTemplate As Long
        hFormatTemplate = CreateWindowsString(FormatTemplate)
        If Invoke(m_pIDateTimeFormatterFactory, _
                  IDateTimeFormatterFactory_CreateDateTimeFormatterLanguages, _
                  hFormatTemplate, _
                  Languages.Ifc, _
                  VarPtr(pIDateTimeFormatter)) = S_OK Then
            If pIDateTimeFormatter <> 0& Then
                Set Ret = New DateTimeFormatter
                Ret.Ifc = pIDateTimeFormatter
            End If
        End If
        Call DeleteWindowsString(hFormatTemplate)
    End If
    Set CreateDateTimeFormatterLanguages = Ret
End Function

Public Function CreateDateTimeFormatterContext(ByVal FormatTemplate As String, _
                                               ByVal Languages As List_String, _
                                               ByVal GeographicRegion As String, _
                                               ByVal Calendar As String, _
                                               ByVal Clock As String) As DateTimeFormatter
    Dim Ret As DateTimeFormatter
    If m_pIDateTimeFormatterFactory <> 0& And _
       IsNotNothing(Languages) Then
        Dim pIDateTimeFormatter As Long
        Dim hFormatTemplate As Long
        Dim hGeographicRegion As Long
        Dim hCalendar As Long
        Dim hClock As Long
        hFormatTemplate = CreateWindowsString(FormatTemplate)
        hGeographicRegion = CreateWindowsString(GeographicRegion)
        hCalendar = CreateWindowsString(Calendar)
        hClock = CreateWindowsString(Clock)
        If Invoke(m_pIDateTimeFormatterFactory, _
                  IDateTimeFormatterFactory_CreateDateTimeFormatterContext, _
                  hFormatTemplate, _
                  Languages.Ifc, _
                  hGeographicRegion, _
                  hCalendar, _
                  hClock, _
                  VarPtr(pIDateTimeFormatter)) = S_OK Then
            If pIDateTimeFormatter <> 0& Then
                Set Ret = New DateTimeFormatter
                Ret.Ifc = pIDateTimeFormatter
            End If
        End If
        Call DeleteWindowsString(hClock)
        Call DeleteWindowsString(hCalendar)
        Call DeleteWindowsString(hGeographicRegion)
        Call DeleteWindowsString(hFormatTemplate)
    End If
    Set CreateDateTimeFormatterContext = Ret
End Function

Public Function CreateDateTimeFormatterDate(ByVal YearFormat As YearFormat, _
                                            ByVal MonthFormat As MonthFormat, _
                                            ByVal DayFormat As DayFormat, _
                                            ByVal DayOfWeekFormat As DayOfWeekFormat) As DateTimeFormatter
    Dim Ret As DateTimeFormatter
    If m_pIDateTimeFormatterFactory <> 0& Then
        Dim pIDateTimeFormatter As Long
        If Invoke(m_pIDateTimeFormatterFactory, _
                  IDateTimeFormatterFactory_CreateDateTimeFormatterDate, _
                  YearFormat, _
                  MonthFormat, _
                  DayFormat, _
                  DayOfWeekFormat, _
                  VarPtr(pIDateTimeFormatter)) = S_OK Then
            If pIDateTimeFormatter <> 0& Then
                Set Ret = New DateTimeFormatter
                Ret.Ifc = pIDateTimeFormatter
            End If
        End If
    End If
    Set CreateDateTimeFormatterDate = Ret
End Function

Public Function CreateDateTimeFormatterTime(ByVal HourFormat As HourFormat, _
                                            ByVal MinuteFormat As MinuteFormat, _
                                            ByVal SecondFormat As SecondFormat) As DateTimeFormatter
    Dim Ret As DateTimeFormatter
    If m_pIDateTimeFormatterFactory <> 0& Then
        Dim pIDateTimeFormatter As Long
        If Invoke(m_pIDateTimeFormatterFactory, _
                  IDateTimeFormatterFactory_CreateDateTimeFormatterTime, _
                  HourFormat, _
                  MinuteFormat, _
                  SecondFormat, _
                  VarPtr(pIDateTimeFormatter)) = S_OK Then
            If pIDateTimeFormatter <> 0& Then
                Set Ret = New DateTimeFormatter
                Ret.Ifc = pIDateTimeFormatter
            End If
        End If
    End If
    Set CreateDateTimeFormatterTime = Ret
End Function

Public Function CreateDateTimeFormatterDateTimeLanguages(ByVal YearFormat As YearFormat, _
                                                         ByVal MonthFormat As MonthFormat, _
                                                         ByVal DayFormat As DayFormat, _
                                                         ByVal DayOfWeekFormat As DayOfWeekFormat, _
                                                         ByVal HourFormat As HourFormat, _
                                                         ByVal MinuteFormat As MinuteFormat, _
                                                         ByVal SecondFormat As SecondFormat, _
                                                         ByVal Languages As List_String) As DateTimeFormatter
    Dim Ret As DateTimeFormatter
    If m_pIDateTimeFormatterFactory <> 0& And _
       IsNotNothing(Languages) Then
        Dim pIDateTimeFormatter As Long
        If Invoke(m_pIDateTimeFormatterFactory, _
                  IDateTimeFormatterFactory_CreateDateTimeFormatterDateTimeLanguages, _
                  YearFormat, _
                  MonthFormat, _
                  DayFormat, _
                  DayOfWeekFormat, _
                  HourFormat, _
                  MinuteFormat, _
                  SecondFormat, _
                  Languages.Ifc, _
                  VarPtr(pIDateTimeFormatter)) = S_OK Then
            If pIDateTimeFormatter <> 0& Then
                Set Ret = New DateTimeFormatter
                Ret.Ifc = pIDateTimeFormatter
            End If
        End If
    End If
    Set CreateDateTimeFormatterDateTimeLanguages = Ret
End Function

Public Function CreateDateTimeFormatterDateTimeContext(ByVal YearFormat As YearFormat, _
                                                       ByVal MonthFormat As MonthFormat, _
                                                       ByVal DayFormat As DayFormat, _
                                                       ByVal DayOfWeekFormat As DayOfWeekFormat, _
                                                       ByVal HourFormat As HourFormat, _
                                                       ByVal MinuteFormat As MinuteFormat, _
                                                       ByVal SecondFormat As SecondFormat, _
                                                       ByVal Languages As List_String, _
                                                       ByVal GeographicRegion As String, _
                                                       ByVal Calendar As String, _
                                                       ByVal Clock As String) As DateTimeFormatter
    Dim Ret As DateTimeFormatter
    If m_pIDateTimeFormatterFactory <> 0& And _
       IsNotNothing(Languages) Then
        Dim pIDateTimeFormatter As Long
        Dim hGeographicRegion As Long
        Dim hCalendar As Long
        Dim hClock As Long
        hGeographicRegion = CreateWindowsString(GeographicRegion)
        hCalendar = CreateWindowsString(Calendar)
        hClock = CreateWindowsString(Clock)
        If Invoke(m_pIDateTimeFormatterFactory, _
                  IDateTimeFormatterFactory_CreateDateTimeFormatterDateTimeContext, _
                  YearFormat, _
                  MonthFormat, _
                  DayFormat, _
                  DayOfWeekFormat, _
                  HourFormat, _
                  MinuteFormat, _
                  SecondFormat, _
                  Languages.Ifc, _
                  hGeographicRegion, _
                  hCalendar, _
                  hClock, _
                  VarPtr(pIDateTimeFormatter)) = S_OK Then
            If pIDateTimeFormatter <> 0& Then
                Set Ret = New DateTimeFormatter
                Ret.Ifc = pIDateTimeFormatter
            End If
        End If
        Call DeleteWindowsString(hClock)
        Call DeleteWindowsString(hCalendar)
        Call DeleteWindowsString(hGeographicRegion)
    End If
    Set CreateDateTimeFormatterDateTimeContext = Ret
End Function
    
' IDateTimeFormatter
Public Function Format(ByVal DateTime As DateTime) As String
    Dim Ret As String
    If m_pIDateTimeFormatter <> 0& Then
        Dim Result As Long
        If Invoke(m_pIDateTimeFormatter, _
                  IDateTimeFormatter_Format, _
                  DateTime.DateTimeValue, _
                  VarPtr(Result)) = S_OK Then
            Ret = GetWindowsString(Result)
        End If
    End If
    Format = Ret
End Function

' IDateTimeFormatter2
Public Function FormatUsingTimeZone(ByVal DateTime As DateTime, _
                                    ByVal TimeZoneId As String) As String
    Dim Ret As String
    If m_pIDateTimeFormatter2 <> 0& Then
        Dim Result As Long
        Dim hTimeZoneId As Long
        hTimeZoneId = CreateWindowsString(TimeZoneId)
        If Invoke(m_pIDateTimeFormatter2, _
                  IDateTimeFormatter2_FormatUsingTimeZone, _
                  DateTime.DateTimeValue, _
                  hTimeZoneId, _
                  VarPtr(Result)) = S_OK Then
            Ret = GetWindowsString(Result)
        End If
        Call DeleteWindowsString(hTimeZoneId)
    End If
    FormatUsingTimeZone = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


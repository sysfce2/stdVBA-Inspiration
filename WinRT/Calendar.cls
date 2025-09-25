VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Globalization_Calendar As String = "Windows.Globalization.Calendar"
Private Const IID_ICalendar As String = "{ca30221d-86d9-40fb-a26b-d44eb7cf08ea}"
Private Const IID_ICalendarFactory As String = "{83f58412-e56b-4c75-a66e-0f63d57758a6}"
Private Const IID_ICalendarFactory2 As String = "{b44b378c-ca7e-4590-9e72-ea2bec1a5115}"
Private Const IID_ITimeZoneOnCalendar As String = "{bb3c25e5-46cf-4317-a3f5-02621ad54478}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' ICalendarFactory
    ICalendarFactory_CreateCalendarDefaultCalendarAndClock = 6
    ICalendarFactory_CreateCalendar = 7

    ' ICalendarFactory2
    ICalendarFactory2_CreateCalendarWithTimeZone = 6

    ' ICalendar
    ICalendar_Clone = 6
    ICalendar_SetToMin = 7
    ICalendar_SetToMax = 8
    ICalendar_GetLanguages = 9
    ICalendar_GetNumeralSystem = 10
    ICalendar_PutNumeralSystem = 11
    ICalendar_GetCalendarSystem = 12
    ICalendar_ChangeCalendarSystem = 13
    ICalendar_GetClock = 14
    ICalendar_ChangeClock = 15
    ICalendar_GetDateTime = 16
    ICalendar_SetDateTime = 17
    ICalendar_SetToNow = 18
    ICalendar_GetFirstEra = 19
    ICalendar_GetLastEra = 20
    ICalendar_GetNumberOfEras = 21
    ICalendar_GetEra = 22
    ICalendar_PutEra = 23
    ICalendar_AddEras = 24
    ICalendar_EraAsFullString = 25
    ICalendar_EraAsString = 26
    ICalendar_GetFirstYearInThisEra = 27
    ICalendar_GetLastYearInThisEra = 28
    ICalendar_GetNumberOfYearsInThisEra = 29
    ICalendar_GetYear = 30
    ICalendar_PutYear = 31
    ICalendar_AddYears = 32
    ICalendar_YearAsString = 33
    ICalendar_YearAsTruncatedString = 34
    ICalendar_YearAsPaddedString = 35
    ICalendar_GetFirstMonthInThisYear = 36
    ICalendar_GetLastMonthInThisYear = 37
    ICalendar_GetNumberOfMonthsInThisYear = 38
    ICalendar_GetMonth = 39
    ICalendar_PutMonth = 40
    ICalendar_AddMonths = 41
    ICalendar_MonthAsFullString = 42
    ICalendar_MonthAsString = 43
    ICalendar_MonthAsFullSoloString = 44
    ICalendar_MonthAsSoloString = 45
    ICalendar_MonthAsNumericString = 46
    ICalendar_MonthAsPaddedNumericString = 47
    ICalendar_AddWeeks = 48
    ICalendar_GetFirstDayInThisMonth = 49
    ICalendar_GetLastDayInThisMonth = 50
    ICalendar_GetNumberOfDaysInThisMonth = 51
    ICalendar_GetDay = 52
    ICalendar_PutDay = 53
    ICalendar_AddDays = 54
    ICalendar_DayAsString = 55
    ICalendar_DayAsPaddedString = 56
    ICalendar_GetDayOfWeek = 57
    ICalendar_DayOfWeekAsFullString = 58
    ICalendar_DayOfWeekAsString = 59
    ICalendar_DayOfWeekAsFullSoloString = 60
    ICalendar_DayOfWeekAsSoloString = 61
    ICalendar_GetFirstPeriodInThisDay = 62
    ICalendar_GetLastPeriodInThisDay = 63
    ICalendar_GetNumberOfPeriodsInThisDay = 64
    ICalendar_GetPeriod = 65
    ICalendar_PutPeriod = 66
    ICalendar_AddPeriods = 67
    ICalendar_PeriodAsFullString = 68
    ICalendar_PeriodAsString = 69
    ICalendar_GetFirstHourInThisPeriod = 70
    ICalendar_GetLastHourInThisPeriod = 71
    ICalendar_GetNumberOfHoursInThisPeriod = 72
    ICalendar_GetHour = 73
    ICalendar_PutHour = 74
    ICalendar_AddHours = 75
    ICalendar_HourAsString = 76
    ICalendar_HourAsPaddedString = 77
    ICalendar_GetMinute = 78
    ICalendar_PutMinute = 79
    ICalendar_AddMinutes = 80
    ICalendar_MinuteAsString = 81
    ICalendar_MinuteAsPaddedString = 82
    ICalendar_GetSecond = 83
    ICalendar_PutSecond = 84
    ICalendar_AddSeconds = 85
    ICalendar_SecondAsString = 86
    ICalendar_SecondAsPaddedString = 87
    ICalendar_GetNanosecond = 88
    ICalendar_PutNanosecond = 89
    ICalendar_AddNanoseconds = 90
    ICalendar_NanosecondAsString = 91
    ICalendar_NanosecondAsPaddedString = 92
    ICalendar_Compare = 93
    ICalendar_CompareDateTime = 94
    ICalendar_CopyTo = 95
    ICalendar_GetFirstMinuteInThisHour = 96
    ICalendar_GetLastMinuteInThisHour = 97
    ICalendar_GetNumberOfMinutesInThisHour = 98
    ICalendar_GetFirstSecondInThisMinute = 99
    ICalendar_GetLastSecondInThisMinute = 100
    ICalendar_GetNumberOfSecondsInThisMinute = 101
    ICalendar_GetResolvedLanguage = 102
    ICalendar_GetIsDaylightSavingTime = 103
                    
    ' ITimeZoneOnCalendar
    ITimeZoneOnCalendar_GetTimeZone = 6
    ITimeZoneOnCalendar_ChangeTimeZone = 7
    ITimeZoneOnCalendar_TimeZoneAsFullString = 8
    ITimeZoneOnCalendar_TimeZoneAsString = 9
    
End Enum

' ----==== Variablen ====----
Private m_pICalendar As Long
Private m_pICalendarFactory As Long
Private m_pICalendarFactory2 As Long
Private m_pITimeZoneOnCalendar As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Globalization_Calendar, _
                            IID_ICalendarFactory, _
                            m_pICalendarFactory) Then
        Call QueryIfc(m_pICalendarFactory, _
                      IID_ICalendarFactory2, _
                      m_pICalendarFactory2)
    End If
    Dim pICalendar As Long
    If GetActivateInstance(Windows_Globalization_Calendar, _
                           IID_ICalendar, _
                           pICalendar) Then
        Me.Ifc = pICalendar
    End If

End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pITimeZoneOnCalendar)
    Call ReleaseIfc(m_pICalendar)
    Call ReleaseIfc(m_pICalendarFactory2)
    Call ReleaseIfc(m_pICalendarFactory)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pICalendar
End Property

Public Property Let Ifc(ByVal pICalendar As Long)
    Call ReleaseIfc(m_pITimeZoneOnCalendar)
    Call ReleaseIfc(m_pICalendar)
    m_pICalendar = pICalendar
    Call QueryIfc(m_pICalendar, _
                  IID_ITimeZoneOnCalendar, _
                  m_pITimeZoneOnCalendar)
End Property

' ICalendar
Public Property Get Languages() As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pICalendar <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetLanguages, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_String, pIVectorView) 'ReadOnlyList_String
            End If
        End If
    End If
    Set Languages = Ret
End Property

Public Property Get NumeralSystem() As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetNumeralSystem, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    NumeralSystem = Ret
End Property

Public Property Let NumeralSystem(ByVal value As String)
    If m_pICalendar <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pICalendar, _
                  ICalendar_PutNumeralSystem, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

Public Property Get FirstEra() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetFirstEra, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    FirstEra = Ret
End Property

Public Property Get LastEra() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetLastEra, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    LastEra = Ret
End Property

Public Property Get NumberOfEras() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetNumberOfEras, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    NumberOfEras = Ret
End Property

Public Property Get Era() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetEra, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Era = Ret
End Property

Public Property Let Era(ByVal value As Long)
    If m_pICalendar <> 0& Then
        If Invoke(m_pICalendar, _
                  ICalendar_PutEra, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get FirstYearInThisEra() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetFirstYearInThisEra, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    FirstYearInThisEra = Ret
End Property

Public Property Get LastYearInThisEra() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetLastYearInThisEra, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    LastYearInThisEra = Ret
End Property

Public Property Get NumberOfYearsInThisEra() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetNumberOfYearsInThisEra, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    NumberOfYearsInThisEra = Ret
End Property

Public Property Get Year() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetYear, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Year = Ret
End Property

Public Property Let Year(ByVal value As Long)
    If m_pICalendar <> 0& Then
        If Invoke(m_pICalendar, _
                  ICalendar_PutYear, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get FirstMonthInThisYear() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetFirstMonthInThisYear, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    FirstMonthInThisYear = Ret
End Property

Public Property Get LastMonthInThisYear() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetLastMonthInThisYear, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    LastMonthInThisYear = Ret
End Property

Public Property Get NumberOfMonthsInThisYear() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetNumberOfMonthsInThisYear, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    NumberOfMonthsInThisYear = Ret
End Property

Public Property Get Month() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetMonth, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Month = Ret
End Property

Public Property Let Month(ByVal value As Long)
    If m_pICalendar <> 0& Then
        If Invoke(m_pICalendar, _
                  ICalendar_PutMonth, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get FirstDayInThisMonth() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetFirstDayInThisMonth, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    FirstDayInThisMonth = Ret
End Property

Public Property Get LastDayInThisMonth() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetLastDayInThisMonth, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    LastDayInThisMonth = Ret
End Property

Public Property Get NumberOfDaysInThisMonth() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetNumberOfDaysInThisMonth, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    NumberOfDaysInThisMonth = Ret
End Property

Public Property Get Day() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetDay, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Day = Ret
End Property

Public Property Let Day(ByVal value As Long)
    If m_pICalendar <> 0& Then
        If Invoke(m_pICalendar, _
                  ICalendar_PutDay, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get DayOfWeek() As DayOfWeek
    Dim Ret As DayOfWeek
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetDayOfWeek, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    DayOfWeek = Ret
End Property

Public Property Get FirstPeriodInThisDay() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetFirstPeriodInThisDay, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    FirstPeriodInThisDay = Ret
End Property

Public Property Get LastPeriodInThisDay() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetLastPeriodInThisDay, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    LastPeriodInThisDay = Ret
End Property

Public Property Get NumberOfPeriodsInThisDay() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetNumberOfPeriodsInThisDay, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    NumberOfPeriodsInThisDay = Ret
End Property

Public Property Get Period() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetPeriod, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Period = Ret
End Property

Public Property Let Period(ByVal value As Long)
    If m_pICalendar <> 0& Then
        If Invoke(m_pICalendar, _
                  ICalendar_PutPeriod, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get FirstHourInThisPeriod() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetFirstHourInThisPeriod, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    FirstHourInThisPeriod = Ret
End Property

Public Property Get LastHourInThisPeriod() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetLastHourInThisPeriod, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    LastHourInThisPeriod = Ret
End Property

Public Property Get NumberOfHoursInThisPeriod() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetNumberOfHoursInThisPeriod, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    NumberOfHoursInThisPeriod = Ret
End Property

Public Property Get Hour() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetHour, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Hour = Ret
End Property

Public Property Let Hour(ByVal value As Long)
    If m_pICalendar <> 0& Then
        If Invoke(m_pICalendar, _
                  ICalendar_PutHour, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get Minute() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetMinute, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Minute = Ret
End Property

Public Property Let Minute(ByVal value As Long)
    If m_pICalendar <> 0& Then
        If Invoke(m_pICalendar, _
                  ICalendar_PutMinute, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get Second() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetSecond, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Second = Ret
End Property

Public Property Let Second(ByVal value As Long)
    If m_pICalendar <> 0& Then
        If Invoke(m_pICalendar, _
                  ICalendar_PutSecond, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get Nanosecond() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetNanosecond, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Nanosecond = Ret
End Property

Public Property Let Nanosecond(ByVal value As Long)
    If m_pICalendar <> 0& Then
        If Invoke(m_pICalendar, _
                  ICalendar_PutNanosecond, _
                  value) = S_OK Then
        End If
    End If
End Property

Public Property Get FirstMinuteInThisHour() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetFirstMinuteInThisHour, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    FirstMinuteInThisHour = Ret
End Property

Public Property Get LastMinuteInThisHour() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetLastMinuteInThisHour, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    LastMinuteInThisHour = Ret
End Property

Public Property Get NumberOfMinutesInThisHour() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetNumberOfMinutesInThisHour, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    NumberOfMinutesInThisHour = Ret
End Property

Public Property Get FirstSecondInThisMinute() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetFirstSecondInThisMinute, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    FirstSecondInThisMinute = Ret
End Property

Public Property Get LastSecondInThisMinute() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetLastSecondInThisMinute, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    LastSecondInThisMinute = Ret
End Property

Public Property Get NumberOfSecondsInThisMinute() As Long
    Dim Ret As Long
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetNumberOfSecondsInThisMinute, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    NumberOfSecondsInThisMinute = Ret
End Property

Public Property Get ResolvedLanguage() As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetResolvedLanguage, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    ResolvedLanguage = Ret
End Property

Public Property Get IsDaylightSavingTime() As Boolean
    Dim Ret As Boolean
    If m_pICalendar <> 0& Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetIsDaylightSavingTime, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    IsDaylightSavingTime = Ret
End Property

' ----==== Functions ====----
' ICalendarFactory
Public Function CreateCalendarDefaultCalendarAndClock(ByVal Languages As List_String) As Calendar
    Dim Ret As Calendar
    If m_pICalendarFactory <> 0& And _
       IsNotNothing(Languages) Then
        Dim pICalendar As Long
        If Invoke(m_pICalendarFactory, _
                  ICalendarFactory_CreateCalendarDefaultCalendarAndClock, _
                  Languages.Ifc, _
                  VarPtr(pICalendar)) = S_OK Then
            If pICalendar <> 0& Then
                Set Ret = New Calendar
                Ret.Ifc = pICalendar
            End If
        End If
    End If
    Set CreateCalendarDefaultCalendarAndClock = Ret
End Function

Public Function CreateCalendar(ByVal Languages As List_String, _
                               ByVal Calendar As String, _
                               ByVal Clock As String) As Calendar
    Dim Ret As Calendar
    If m_pICalendarFactory <> 0& And _
       IsNotNothing(Languages) Then
        Dim pICalendar As Long
        Dim hCalendar As Long
        Dim hClock As Long
        hCalendar = CreateWindowsString(Calendar)
        hClock = CreateWindowsString(Clock)
        If Invoke(m_pICalendarFactory, _
                  ICalendarFactory_CreateCalendar, _
                  Languages.Ifc, _
                  hCalendar, _
                  hClock, _
                  VarPtr(pICalendar)) = S_OK Then
            If pICalendar <> 0& Then
                Set Ret = New Calendar
                Ret.Ifc = pICalendar
            End If
        End If
        Call DeleteWindowsString(hClock)
        Call DeleteWindowsString(hCalendar)
    End If
    Set CreateCalendar = Ret
End Function

' ICalendarFactory2
Public Function CreateCalendarWithTimeZone(ByVal Languages As List_String, _
                                           ByVal Calendar As String, _
                                           ByVal Clock As String, _
                                           ByVal TimeZoneId As String) As Calendar
    Dim Ret As Calendar
    If m_pICalendarFactory2 <> 0& And _
       IsNotNothing(Languages) Then
        Dim pICalendar As Long
        Dim hCalendar As Long
        Dim hClock As Long
        Dim hTimeZoneId As Long
        hCalendar = CreateWindowsString(Calendar)
        hClock = CreateWindowsString(Clock)
        hTimeZoneId = CreateWindowsString(TimeZoneId)
        If Invoke(m_pICalendarFactory2, _
                  ICalendarFactory2_CreateCalendarWithTimeZone, _
                  Languages.Ifc, _
                  hCalendar, _
                  hClock, _
                  hTimeZoneId, _
                  VarPtr(pICalendar)) = S_OK Then
            If pICalendar <> 0& Then
                Set Ret = New Calendar
                Ret.Ifc = pICalendar
            End If
        End If
        Call DeleteWindowsString(hTimeZoneId)
        Call DeleteWindowsString(hClock)
        Call DeleteWindowsString(hCalendar)
    End If
    Set CreateCalendarWithTimeZone = Ret
End Function

' ICalendar
Public Function Clone() As Calendar
    Dim Ret As Calendar
    If m_pICalendar <> 0& Then
        Dim pICalendar As Long
        If Invoke(m_pICalendar, _
                  ICalendar_Clone, _
                  VarPtr(pICalendar)) = S_OK Then
            If pICalendar <> 0& Then
                Set Ret = New Calendar
                Ret.Ifc = pICalendar
            End If
        End If
    End If
    Set Clone = Ret
End Function

Public Function SetToMin() As Boolean
    Dim Ret As Boolean
    If m_pICalendar <> 0& Then
        If Invoke(m_pICalendar, _
                  ICalendar_SetToMin) = S_OK Then
            Ret = True
        End If
    End If
    SetToMin = Ret
End Function

Public Function SetToMax() As Boolean
    Dim Ret As Boolean
    If m_pICalendar <> 0& Then
        If Invoke(m_pICalendar, _
                  ICalendar_SetToMax) = S_OK Then
            Ret = True
        End If
    End If
    SetToMax = Ret
End Function

Public Function GetCalendarSystem() As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetCalendarSystem, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    GetCalendarSystem = Ret
End Function

Public Function ChangeCalendarSystem(ByVal value As String) As Boolean
    Dim Ret As Boolean
    If m_pICalendar <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pICalendar, _
                  ICalendar_ChangeCalendarSystem, _
                  hString) = S_OK Then
            Ret = True
        End If
        Call DeleteWindowsString(hString)
    End If
    ChangeCalendarSystem = Ret
End Function

Public Function GetClock() As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_GetClock, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    GetClock = Ret
End Function

Public Function ChangeClock(ByVal value As String) As Boolean
    Dim Ret As Boolean
    If m_pICalendar <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pICalendar, _
                  ICalendar_ChangeClock, _
                  hString) = S_OK Then
            Ret = True
        Call DeleteWindowsString(hString)
        End If
    End If
    ChangeClock = Ret
End Function

Public Function GetDateTime() As DateTime
    Dim Ret As DateTime
    If m_pICalendar <> 0& Then
        Dim value As Currency
        If Invoke(m_pICalendar, _
                  ICalendar_GetDateTime, _
                  VarPtr(value)) = S_OK Then
            Set Ret = New DateTime
            Ret.DateTimeValue = value
        End If
    End If
    Set GetDateTime = Ret
End Function

Public Function SetDateTime(ByVal value As DateTime) As Boolean
    Dim Ret As Boolean
    If m_pICalendar <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pICalendar, _
                  ICalendar_SetDateTime, _
                  value.DateTimeValue) = S_OK Then
            Ret = True
        End If
    End If
    SetDateTime = Ret
End Function

Public Function SetToNow() As Boolean
    Dim Ret As Boolean
    If m_pICalendar <> 0& Then
        If Invoke(m_pICalendar, _
                  ICalendar_SetToNow) = S_OK Then
            Ret = True
        End If
    End If
    SetToNow = Ret
End Function

Public Function AddEras(ByVal eras As Long) As Boolean
    Dim Ret As Boolean
    If m_pICalendar <> 0& Then
        If Invoke(m_pICalendar, _
                  ICalendar_AddEras, _
                  eras) = S_OK Then
            Ret = True
        End If
    End If
    AddEras = Ret
End Function

Public Function EraAsFullString() As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_EraAsFullString, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    EraAsFullString = Ret
End Function

Public Function EraAsString(Optional ByVal idealLength As Long = 0&) As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_EraAsString, _
                  Abs(idealLength), _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    EraAsString = Ret
End Function

Public Function AddYears(ByVal years As Long) As Boolean
    Dim Ret As Boolean
    If m_pICalendar <> 0& Then
        If Invoke(m_pICalendar, _
                  ICalendar_AddYears, _
                  years) = S_OK Then
            Ret = True
        End If
    End If
    AddYears = Ret
End Function

Public Function YearAsString() As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_YearAsString, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    YearAsString = Ret
End Function

Public Function YearAsTruncatedString(Optional ByVal remainingDigits As Long = 4&) As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_YearAsTruncatedString, _
                  Abs(remainingDigits), _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    YearAsTruncatedString = Ret
End Function

Public Function YearAsPaddedString(Optional ByVal minDigits As Long = 4&) As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_YearAsPaddedString, _
                  Abs(minDigits), _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    YearAsPaddedString = Ret
End Function

Public Function AddMonths(ByVal months As Long) As Boolean
    Dim Ret As Boolean
    If m_pICalendar <> 0& Then
        If Invoke(m_pICalendar, _
                  ICalendar_AddMonths, _
                  months) = S_OK Then
            Ret = True
        End If
    End If
    AddMonths = Ret
End Function

Public Function MonthAsFullString() As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_MonthAsFullString, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    MonthAsFullString = Ret
End Function

Public Function MonthAsString(Optional ByVal idealLength As Long = 0&) As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_MonthAsString, _
                  Abs(idealLength), _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    MonthAsString = Ret
End Function

Public Function MonthAsFullSoloString() As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_MonthAsFullSoloString, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    MonthAsFullSoloString = Ret
End Function

Public Function MonthAsSoloString(Optional ByVal idealLength As Long = 0&) As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_MonthAsSoloString, _
                  Abs(idealLength), _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    MonthAsSoloString = Ret
End Function

Public Function MonthAsNumericString() As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_MonthAsNumericString, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    MonthAsNumericString = Ret
End Function

Public Function MonthAsPaddedNumericString(Optional ByVal minDigits As Long = 2&) As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_MonthAsPaddedNumericString, _
                  Abs(minDigits), _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    MonthAsPaddedNumericString = Ret
End Function

Public Function AddWeeks(ByVal weeks As Long) As Boolean
    Dim Ret As Boolean
    If m_pICalendar <> 0& Then
        If Invoke(m_pICalendar, _
                  ICalendar_AddWeeks, _
                  weeks) = S_OK Then
            Ret = True
        End If
    End If
    AddWeeks = Ret
End Function

Public Function AddDays(ByVal days As Long) As Boolean
    Dim Ret As Boolean
    If m_pICalendar <> 0& Then
        If Invoke(m_pICalendar, _
                  ICalendar_AddDays, _
                  days) = S_OK Then
            Ret = True
        End If
    End If
    AddDays = Ret
End Function

Public Function DayAsString() As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_DayAsString, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    DayAsString = Ret
End Function

Public Function DayAsPaddedString(Optional ByVal minDigits As Long = 2&) As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_DayAsPaddedString, _
                  Abs(minDigits), _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    DayAsPaddedString = Ret
End Function

Public Function DayOfWeekAsFullString() As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_DayOfWeekAsFullString, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    DayOfWeekAsFullString = Ret
End Function

Public Function DayOfWeekAsString(Optional ByVal idealLength As Long = 0&) As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_DayOfWeekAsString, _
                  Abs(idealLength), _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    DayOfWeekAsString = Ret
End Function

Public Function DayOfWeekAsFullSoloString() As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_DayOfWeekAsFullSoloString, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    DayOfWeekAsFullSoloString = Ret
End Function

Public Function DayOfWeekAsSoloString(Optional ByVal idealLength As Long = 0&) As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_DayOfWeekAsSoloString, _
                  Abs(idealLength), _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    DayOfWeekAsSoloString = Ret
End Function

Public Function AddPeriods(ByVal periods As Long) As Boolean
    Dim Ret As Boolean
    If m_pICalendar <> 0& Then
        If Invoke(m_pICalendar, _
                  ICalendar_AddPeriods, _
                  periods) = S_OK Then
            Ret = True
        End If
    End If
    AddPeriods = Ret
End Function

Public Function PeriodAsFullString() As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_PeriodAsFullString, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    PeriodAsFullString = Ret
End Function

Public Function PeriodAsString(Optional ByVal idealLength As Long = 0&) As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_PeriodAsString, _
                  Abs(idealLength), _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    PeriodAsString = Ret
End Function

Public Function AddHours(ByVal hours As Long) As Boolean
    Dim Ret As Boolean
    If m_pICalendar <> 0& Then
        If Invoke(m_pICalendar, _
                  ICalendar_AddHours, _
                  hours) = S_OK Then
            Ret = True
        End If
    End If
    AddHours = Ret
End Function

Public Function HourAsString() As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_HourAsString, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    HourAsString = Ret
End Function

Public Function HourAsPaddedString(Optional ByVal minDigits As Long = 2&) As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_HourAsPaddedString, _
                  Abs(minDigits), _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    HourAsPaddedString = Ret
End Function

Public Function AddMinutes(ByVal minutes As Long) As Boolean
    Dim Ret As Boolean
    If m_pICalendar <> 0& Then
        If Invoke(m_pICalendar, _
                  ICalendar_AddMinutes, _
                  minutes) = S_OK Then
            Ret = True
        End If
    End If
    AddMinutes = Ret
End Function

Public Function MinuteAsString() As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_MinuteAsString, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    MinuteAsString = Ret
End Function

Public Function MinuteAsPaddedString(Optional ByVal minDigits As Long = 2&) As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_MinuteAsPaddedString, _
                  Abs(minDigits), _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    MinuteAsPaddedString = Ret
End Function

Public Function AddSeconds(ByVal seconds As Long) As Boolean
    Dim Ret As Boolean
    If m_pICalendar <> 0& Then
        If Invoke(m_pICalendar, _
                  ICalendar_AddSeconds, _
                  seconds) = S_OK Then
            Ret = True
        End If
    End If
    AddSeconds = Ret
End Function

Public Function SecondAsString() As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_SecondAsString, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    SecondAsString = Ret
End Function

Public Function SecondAsPaddedString(Optional ByVal minDigits As Long = 2&) As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_SecondAsPaddedString, _
                  Abs(minDigits), _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    SecondAsPaddedString = Ret
End Function

Public Function AddNanoseconds(ByVal nanoseconds As Long) As Boolean
    Dim Ret As Boolean
    If m_pICalendar <> 0& Then
        If Invoke(m_pICalendar, _
                  ICalendar_AddNanoseconds, _
                  nanoseconds) = S_OK Then
            Ret = True
        End If
    End If
    AddNanoseconds = Ret
End Function

Public Function NanosecondAsString() As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_NanosecondAsString, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    NanosecondAsString = Ret
End Function

Public Function NanosecondAsPaddedString(Optional ByVal minDigits As Long = 2&) As String
    Dim Ret As String
    If m_pICalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pICalendar, _
                  ICalendar_NanosecondAsPaddedString, _
                  Abs(minDigits), _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    NanosecondAsPaddedString = Ret
End Function

Public Function Compare(ByVal other As Calendar) As Long
    Dim Ret As Long
    If m_pICalendar <> 0& And _
       IsNotNothing(other) Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_Compare, _
                  other.Ifc, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Compare = Ret
End Function

Public Function CompareDateTime(ByVal other As DateTime) As Long
    Dim Ret As Long
    If m_pICalendar <> 0& And _
       IsNotNothing(other) Then
        Dim value As Long
        If Invoke(m_pICalendar, _
                  ICalendar_CompareDateTime, _
                  other.DateTimeValue, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    CompareDateTime = Ret
End Function

Public Function CopyTo(ByVal other As Calendar) As Boolean
    Dim Ret As Boolean
    If m_pICalendar <> 0& And _
       IsNotNothing(other) Then
        If Invoke(m_pICalendar, _
                  ICalendar_CopyTo, _
                  other.Ifc) = S_OK Then
            Ret = True
        End If
    End If
    CopyTo = Ret
End Function

' ITimeZoneOnCalendar
Public Function GetTimeZone() As String
    Dim Ret As String
    If m_pITimeZoneOnCalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pITimeZoneOnCalendar, _
                  ITimeZoneOnCalendar_GetTimeZone, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    GetTimeZone = Ret
End Function

Public Function ChangeTimeZone(ByVal TimeZoneId As String) As Boolean
    Dim Ret As Boolean
    If m_pITimeZoneOnCalendar <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(TimeZoneId)
        If Invoke(m_pITimeZoneOnCalendar, _
                  ITimeZoneOnCalendar_ChangeTimeZone, _
                  hString) = S_OK Then
            Ret = True
        End If
        Call DeleteWindowsString(hString)
    End If
    ChangeTimeZone = Ret
End Function

Public Function TimeZoneAsFullString() As String
    Dim Ret As String
    If m_pITimeZoneOnCalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pITimeZoneOnCalendar, _
                  ITimeZoneOnCalendar_TimeZoneAsFullString, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    TimeZoneAsFullString = Ret
End Function

Public Function TimeZoneAsString(Optional ByVal idealLength As Long = 0&) As String
    Dim Ret As String
    If m_pITimeZoneOnCalendar <> 0& Then
        Dim hString As Long
        If Invoke(m_pITimeZoneOnCalendar, _
                  ITimeZoneOnCalendar_TimeZoneAsString, _
                  Abs(idealLength), _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    TimeZoneAsString = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



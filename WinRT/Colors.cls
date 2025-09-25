VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Colors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----

Private Const Windows_UI_Colors As String = "Windows.UI.Colors"
Private Const IID_IColors As String = "{9b8c9326-4ca6-4ce5-8994-9eff65cabdcc}"
Private Const IID_IColorsStatics As String = "{cff52e04-cca6-4614-a17e-754910c84a99}"

Private Const Windows_UI_ColorHelper As String = "Windows.UI.ColorHelper"
Private Const IID_IColorHelperStatics As String = "{8504dbea-fb6a-4144-a6c2-33499c9284f5}"
Private Const IID_IColorHelperStatics2 As String = "{24d9af02-6eb0-4b94-855c-fcf0818d9a16}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IColors
    ' has no member

    ' IColorsStatics
    IColorsStatics_GetAliceBlue = 6
    IColorsStatics_GetAntiqueWhite = 7
    IColorsStatics_GetAqua = 8
    IColorsStatics_GetAquamarine = 9
    IColorsStatics_GetAzure = 10
    IColorsStatics_GetBeige = 11
    IColorsStatics_GetBisque = 12
    IColorsStatics_GetBlack = 13
    IColorsStatics_GetBlanchedAlmond = 14
    IColorsStatics_GetBlue = 15
    IColorsStatics_GetBlueViolet = 16
    IColorsStatics_GetBrown = 17
    IColorsStatics_GetBurlyWood = 18
    IColorsStatics_GetCadetBlue = 19
    IColorsStatics_GetChartreuse = 20
    IColorsStatics_GetChocolate = 21
    IColorsStatics_GetCoral = 22
    IColorsStatics_GetCornflowerBlue = 23
    IColorsStatics_GetCornsilk = 24
    IColorsStatics_GetCrimson = 25
    IColorsStatics_GetCyan = 26
    IColorsStatics_GetDarkBlue = 27
    IColorsStatics_GetDarkCyan = 28
    IColorsStatics_GetDarkGoldenrod = 29
    IColorsStatics_GetDarkGray = 30
    IColorsStatics_GetDarkGreen = 31
    IColorsStatics_GetDarkKhaki = 32
    IColorsStatics_GetDarkMagenta = 33
    IColorsStatics_GetDarkOliveGreen = 34
    IColorsStatics_GetDarkOrange = 35
    IColorsStatics_GetDarkOrchid = 36
    IColorsStatics_GetDarkRed = 37
    IColorsStatics_GetDarkSalmon = 38
    IColorsStatics_GetDarkSeaGreen = 39
    IColorsStatics_GetDarkSlateBlue = 40
    IColorsStatics_GetDarkSlateGray = 41
    IColorsStatics_GetDarkTurquoise = 42
    IColorsStatics_GetDarkViolet = 43
    IColorsStatics_GetDeepPink = 44
    IColorsStatics_GetDeepSkyBlue = 45
    IColorsStatics_GetDimGray = 46
    IColorsStatics_GetDodgerBlue = 47
    IColorsStatics_GetFirebrick = 48
    IColorsStatics_GetFloralWhite = 49
    IColorsStatics_GetForestGreen = 50
    IColorsStatics_GetFuchsia = 51
    IColorsStatics_GetGainsboro = 52
    IColorsStatics_GetGhostWhite = 53
    IColorsStatics_GetGold = 54
    IColorsStatics_GetGoldenrod = 55
    IColorsStatics_GetGray = 56
    IColorsStatics_GetGreen = 57
    IColorsStatics_GetGreenYellow = 58
    IColorsStatics_GetHoneydew = 59
    IColorsStatics_GetHotPink = 60
    IColorsStatics_GetIndianRed = 61
    IColorsStatics_GetIndigo = 62
    IColorsStatics_GetIvory = 63
    IColorsStatics_GetKhaki = 64
    IColorsStatics_GetLavender = 65
    IColorsStatics_GetLavenderBlush = 66
    IColorsStatics_GetLawnGreen = 67
    IColorsStatics_GetLemonChiffon = 68
    IColorsStatics_GetLightBlue = 69
    IColorsStatics_GetLightCoral = 70
    IColorsStatics_GetLightCyan = 71
    IColorsStatics_GetLightGoldenrodYellow = 72
    IColorsStatics_GetLightGreen = 73
    IColorsStatics_GetLightGray = 74
    IColorsStatics_GetLightPink = 75
    IColorsStatics_GetLightSalmon = 76
    IColorsStatics_GetLightSeaGreen = 77
    IColorsStatics_GetLightSkyBlue = 78
    IColorsStatics_GetLightSlateGray = 79
    IColorsStatics_GetLightSteelBlue = 80
    IColorsStatics_GetLightYellow = 81
    IColorsStatics_GetLime = 82
    IColorsStatics_GetLimeGreen = 83
    IColorsStatics_GetLinen = 84
    IColorsStatics_GetMagenta = 85
    IColorsStatics_GetMaroon = 86
    IColorsStatics_GetMediumAquamarine = 87
    IColorsStatics_GetMediumBlue = 88
    IColorsStatics_GetMediumOrchid = 89
    IColorsStatics_GetMediumPurple = 90
    IColorsStatics_GetMediumSeaGreen = 91
    IColorsStatics_GetMediumSlateBlue = 92
    IColorsStatics_GetMediumSpringGreen = 93
    IColorsStatics_GetMediumTurquoise = 94
    IColorsStatics_GetMediumVioletRed = 95
    IColorsStatics_GetMidnightBlue = 96
    IColorsStatics_GetMintCream = 97
    IColorsStatics_GetMistyRose = 98
    IColorsStatics_GetMoccasin = 99
    IColorsStatics_GetNavajoWhite = 100
    IColorsStatics_GetNavy = 101
    IColorsStatics_GetOldLace = 102
    IColorsStatics_GetOlive = 103
    IColorsStatics_GetOliveDrab = 104
    IColorsStatics_GetOrange = 105
    IColorsStatics_GetOrangeRed = 106
    IColorsStatics_GetOrchid = 107
    IColorsStatics_GetPaleGoldenrod = 108
    IColorsStatics_GetPaleGreen = 109
    IColorsStatics_GetPaleTurquoise = 110
    IColorsStatics_GetPaleVioletRed = 111
    IColorsStatics_GetPapayaWhip = 112
    IColorsStatics_GetPeachPuff = 113
    IColorsStatics_GetPeru = 114
    IColorsStatics_GetPink = 115
    IColorsStatics_GetPlum = 116
    IColorsStatics_GetPowderBlue = 117
    IColorsStatics_GetPurple = 118
    IColorsStatics_GetRed = 119
    IColorsStatics_GetRosyBrown = 120
    IColorsStatics_GetRoyalBlue = 121
    IColorsStatics_GetSaddleBrown = 122
    IColorsStatics_GetSalmon = 123
    IColorsStatics_GetSandyBrown = 124
    IColorsStatics_GetSeaGreen = 125
    IColorsStatics_GetSeaShell = 126
    IColorsStatics_GetSienna = 127
    IColorsStatics_GetSilver = 128
    IColorsStatics_GetSkyBlue = 129
    IColorsStatics_GetSlateBlue = 130
    IColorsStatics_GetSlateGray = 131
    IColorsStatics_GetSnow = 132
    IColorsStatics_GetSpringGreen = 133
    IColorsStatics_GetSteelBlue = 134
    IColorsStatics_GetTan = 135
    IColorsStatics_GetTeal = 136
    IColorsStatics_GetThistle = 137
    IColorsStatics_GetTomato = 138
    IColorsStatics_GetTransparent = 139
    IColorsStatics_GetTurquoise = 140
    IColorsStatics_GetViolet = 141
    IColorsStatics_GetWheat = 142
    IColorsStatics_GetWhite = 143
    IColorsStatics_GetWhiteSmoke = 144
    IColorsStatics_GetYellow = 145
    IColorsStatics_GetYellowGreen = 146
                    
    ' IColorHelperStatics
    IColorHelperStatics_FromArgb = 6
                        
    ' IColorHelperStatics2
    IColorHelperStatics2_ToDisplayName = 6
                        
End Enum

' ----==== Vars ====----
Private m_pIColors As Long
Private m_pIColorsStatics As Long
Private m_pIColorHelperStatics As Long
Private m_pIColorHelperStatics2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_UI_Colors, _
                            IID_IColorsStatics, _
                            m_pIColorsStatics) Then
    End If
    If GetActivationFactory(Windows_UI_ColorHelper, _
                            IID_IColorHelperStatics, _
                            m_pIColorHelperStatics) Then
        Call QueryIfc(m_pIColorHelperStatics, _
                      IID_IColorHelperStatics2, _
                      m_pIColorHelperStatics2)
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIColors)
    Call ReleaseIfc(m_pIColorsStatics)
    Call ReleaseIfc(m_pIColorHelperStatics)
    Call ReleaseIfc(m_pIColorHelperStatics2)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIColors
End Property

Public Property Let Ifc(ByVal pIColors As Long)
    Call ReleaseIfc(m_pIColors)
    m_pIColors = pIColors
End Property

' IColorsStatic
Public Property Get AliceBlue() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetAliceBlue, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set AliceBlue = Ret
End Property

Public Property Get AntiqueWhite() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetAntiqueWhite, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set AntiqueWhite = Ret
End Property

Public Property Get Aqua() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetAqua, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Aqua = Ret
End Property

Public Property Get Aquamarine() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetAquamarine, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Aquamarine = Ret
End Property

Public Property Get Azure() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetAzure, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Azure = Ret
End Property

Public Property Get Beige() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetBeige, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Beige = Ret
End Property

Public Property Get Bisque() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetBisque, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Bisque = Ret
End Property

Public Property Get Black() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetBlack, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Black = Ret
End Property

Public Property Get BlanchedAlmond() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetBlanchedAlmond, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set BlanchedAlmond = Ret
End Property

Public Property Get Blue() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetBlue, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Blue = Ret
End Property

Public Property Get BlueViolet() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetBlueViolet, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set BlueViolet = Ret
End Property

Public Property Get Brown() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetBrown, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Brown = Ret
End Property

Public Property Get BurlyWood() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetBurlyWood, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set BurlyWood = Ret
End Property

Public Property Get CadetBlue() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetCadetBlue, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set CadetBlue = Ret
End Property

Public Property Get Chartreuse() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetChartreuse, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Chartreuse = Ret
End Property

Public Property Get Chocolate() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetChocolate, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Chocolate = Ret
End Property

Public Property Get Coral() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetCoral, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Coral = Ret
End Property

Public Property Get CornflowerBlue() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetCornflowerBlue, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set CornflowerBlue = Ret
End Property

Public Property Get Cornsilk() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetCornsilk, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Cornsilk = Ret
End Property

Public Property Get Crimson() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetCrimson, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Crimson = Ret
End Property

Public Property Get Cyan() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetCyan, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Cyan = Ret
End Property

Public Property Get DarkBlue() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetDarkBlue, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set DarkBlue = Ret
End Property

Public Property Get DarkCyan() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetDarkCyan, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set DarkCyan = Ret
End Property

Public Property Get DarkGoldenrod() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetDarkGoldenrod, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set DarkGoldenrod = Ret
End Property

Public Property Get DarkGray() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetDarkGray, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set DarkGray = Ret
End Property

Public Property Get DarkGreen() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetDarkGreen, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set DarkGreen = Ret
End Property

Public Property Get DarkKhaki() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetDarkKhaki, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set DarkKhaki = Ret
End Property

Public Property Get DarkMagenta() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetDarkMagenta, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set DarkMagenta = Ret
End Property

Public Property Get DarkOliveGreen() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetDarkOliveGreen, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set DarkOliveGreen = Ret
End Property

Public Property Get DarkOrange() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetDarkOrange, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set DarkOrange = Ret
End Property

Public Property Get DarkOrchid() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetDarkOrchid, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set DarkOrchid = Ret
End Property

Public Property Get DarkRed() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetDarkRed, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set DarkRed = Ret
End Property

Public Property Get DarkSalmon() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetDarkSalmon, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set DarkSalmon = Ret
End Property

Public Property Get DarkSeaGreen() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetDarkSeaGreen, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set DarkSeaGreen = Ret
End Property

Public Property Get DarkSlateBlue() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetDarkSlateBlue, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set DarkSlateBlue = Ret
End Property

Public Property Get DarkSlateGray() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetDarkSlateGray, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set DarkSlateGray = Ret
End Property

Public Property Get DarkTurquoise() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetDarkTurquoise, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set DarkTurquoise = Ret
End Property

Public Property Get DarkViolet() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetDarkViolet, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set DarkViolet = Ret
End Property

Public Property Get DeepPink() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetDeepPink, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set DeepPink = Ret
End Property

Public Property Get DeepSkyBlue() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetDeepSkyBlue, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set DeepSkyBlue = Ret
End Property

Public Property Get DimGray() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetDimGray, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set DimGray = Ret
End Property

Public Property Get DodgerBlue() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetDodgerBlue, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set DodgerBlue = Ret
End Property

Public Property Get Firebrick() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetFirebrick, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Firebrick = Ret
End Property

Public Property Get FloralWhite() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetFloralWhite, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set FloralWhite = Ret
End Property

Public Property Get ForestGreen() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetForestGreen, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set ForestGreen = Ret
End Property

Public Property Get Fuchsia() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetFuchsia, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Fuchsia = Ret
End Property

Public Property Get Gainsboro() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetGainsboro, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Gainsboro = Ret
End Property

Public Property Get GhostWhite() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetGhostWhite, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set GhostWhite = Ret
End Property

Public Property Get Gold() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetGold, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Gold = Ret
End Property

Public Property Get Goldenrod() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetGoldenrod, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Goldenrod = Ret
End Property

Public Property Get Gray() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetGray, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Gray = Ret
End Property

Public Property Get Green() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetGreen, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Green = Ret
End Property

Public Property Get GreenYellow() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetGreenYellow, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set GreenYellow = Ret
End Property

Public Property Get Honeydew() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetHoneydew, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Honeydew = Ret
End Property

Public Property Get HotPink() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetHotPink, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set HotPink = Ret
End Property

Public Property Get IndianRed() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetIndianRed, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set IndianRed = Ret
End Property

Public Property Get Indigo() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetIndigo, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Indigo = Ret
End Property

Public Property Get Ivory() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetIvory, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Ivory = Ret
End Property

Public Property Get Khaki() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetKhaki, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Khaki = Ret
End Property

Public Property Get Lavender() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetLavender, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Lavender = Ret
End Property

Public Property Get LavenderBlush() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetLavenderBlush, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set LavenderBlush = Ret
End Property

Public Property Get LawnGreen() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetLawnGreen, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set LawnGreen = Ret
End Property

Public Property Get LemonChiffon() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetLemonChiffon, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set LemonChiffon = Ret
End Property

Public Property Get LightBlue() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetLightBlue, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set LightBlue = Ret
End Property

Public Property Get LightCoral() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetLightCoral, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set LightCoral = Ret
End Property

Public Property Get LightCyan() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetLightCyan, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set LightCyan = Ret
End Property

Public Property Get LightGoldenrodYellow() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetLightGoldenrodYellow, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set LightGoldenrodYellow = Ret
End Property

Public Property Get LightGreen() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetLightGreen, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set LightGreen = Ret
End Property

Public Property Get LightGray() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetLightGray, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set LightGray = Ret
End Property

Public Property Get LightPink() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetLightPink, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set LightPink = Ret
End Property

Public Property Get LightSalmon() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetLightSalmon, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set LightSalmon = Ret
End Property

Public Property Get LightSeaGreen() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetLightSeaGreen, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set LightSeaGreen = Ret
End Property

Public Property Get LightSkyBlue() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetLightSkyBlue, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set LightSkyBlue = Ret
End Property

Public Property Get LightSlateGray() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetLightSlateGray, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set LightSlateGray = Ret
End Property

Public Property Get LightSteelBlue() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetLightSteelBlue, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set LightSteelBlue = Ret
End Property

Public Property Get LightYellow() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetLightYellow, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set LightYellow = Ret
End Property

Public Property Get Lime() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetLime, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Lime = Ret
End Property

Public Property Get LimeGreen() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetLimeGreen, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set LimeGreen = Ret
End Property

Public Property Get Linen() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetLinen, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Linen = Ret
End Property

Public Property Get Magenta() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetMagenta, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Magenta = Ret
End Property

Public Property Get Maroon() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetMaroon, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Maroon = Ret
End Property

Public Property Get MediumAquamarine() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetMediumAquamarine, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set MediumAquamarine = Ret
End Property

Public Property Get MediumBlue() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetMediumBlue, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set MediumBlue = Ret
End Property

Public Property Get MediumOrchid() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetMediumOrchid, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set MediumOrchid = Ret
End Property

Public Property Get MediumPurple() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetMediumPurple, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set MediumPurple = Ret
End Property

Public Property Get MediumSeaGreen() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetMediumSeaGreen, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set MediumSeaGreen = Ret
End Property

Public Property Get MediumSlateBlue() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetMediumSlateBlue, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set MediumSlateBlue = Ret
End Property

Public Property Get MediumSpringGreen() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetMediumSpringGreen, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set MediumSpringGreen = Ret
End Property

Public Property Get MediumTurquoise() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetMediumTurquoise, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set MediumTurquoise = Ret
End Property

Public Property Get MediumVioletRed() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetMediumVioletRed, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set MediumVioletRed = Ret
End Property

Public Property Get MidnightBlue() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetMidnightBlue, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set MidnightBlue = Ret
End Property

Public Property Get MintCream() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetMintCream, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set MintCream = Ret
End Property

Public Property Get MistyRose() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetMistyRose, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set MistyRose = Ret
End Property

Public Property Get Moccasin() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetMoccasin, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Moccasin = Ret
End Property

Public Property Get NavajoWhite() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetNavajoWhite, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set NavajoWhite = Ret
End Property

Public Property Get Navy() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetNavy, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Navy = Ret
End Property

Public Property Get OldLace() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetOldLace, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set OldLace = Ret
End Property

Public Property Get Olive() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetOlive, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Olive = Ret
End Property

Public Property Get OliveDrab() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetOliveDrab, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set OliveDrab = Ret
End Property

Public Property Get Orange() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetOrange, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Orange = Ret
End Property

Public Property Get OrangeRed() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetOrangeRed, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set OrangeRed = Ret
End Property

Public Property Get Orchid() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetOrchid, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Orchid = Ret
End Property

Public Property Get PaleGoldenrod() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetPaleGoldenrod, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set PaleGoldenrod = Ret
End Property

Public Property Get PaleGreen() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetPaleGreen, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set PaleGreen = Ret
End Property

Public Property Get PaleTurquoise() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetPaleTurquoise, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set PaleTurquoise = Ret
End Property

Public Property Get PaleVioletRed() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetPaleVioletRed, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set PaleVioletRed = Ret
End Property

Public Property Get PapayaWhip() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetPapayaWhip, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set PapayaWhip = Ret
End Property

Public Property Get PeachPuff() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetPeachPuff, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set PeachPuff = Ret
End Property

Public Property Get Peru() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetPeru, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Peru = Ret
End Property

Public Property Get Pink() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetPink, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Pink = Ret
End Property

Public Property Get Plum() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetPlum, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Plum = Ret
End Property

Public Property Get PowderBlue() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetPowderBlue, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set PowderBlue = Ret
End Property

Public Property Get Purple() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetPurple, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Purple = Ret
End Property

Public Property Get Red() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetRed, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Red = Ret
End Property

Public Property Get RosyBrown() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetRosyBrown, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set RosyBrown = Ret
End Property

Public Property Get RoyalBlue() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetRoyalBlue, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set RoyalBlue = Ret
End Property

Public Property Get SaddleBrown() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetSaddleBrown, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set SaddleBrown = Ret
End Property

Public Property Get Salmon() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetSalmon, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Salmon = Ret
End Property

Public Property Get SandyBrown() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetSandyBrown, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set SandyBrown = Ret
End Property

Public Property Get SeaGreen() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetSeaGreen, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set SeaGreen = Ret
End Property

Public Property Get SeaShell() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetSeaShell, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set SeaShell = Ret
End Property

Public Property Get Sienna() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetSienna, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Sienna = Ret
End Property

Public Property Get Silver() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetSilver, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Silver = Ret
End Property

Public Property Get SkyBlue() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetSkyBlue, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set SkyBlue = Ret
End Property

Public Property Get SlateBlue() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetSlateBlue, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set SlateBlue = Ret
End Property

Public Property Get SlateGray() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetSlateGray, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set SlateGray = Ret
End Property

Public Property Get Snow() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetSnow, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Snow = Ret
End Property

Public Property Get SpringGreen() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetSpringGreen, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set SpringGreen = Ret
End Property

Public Property Get SteelBlue() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetSteelBlue, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set SteelBlue = Ret
End Property

Public Property Get Tan() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetTan, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Tan = Ret
End Property

Public Property Get Teal() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetTeal, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Teal = Ret
End Property

Public Property Get Thistle() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetThistle, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Thistle = Ret
End Property

Public Property Get Tomato() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetTomato, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Tomato = Ret
End Property

Public Property Get Transparent() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetTransparent, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Transparent = Ret
End Property

Public Property Get Turquoise() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetTurquoise, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Turquoise = Ret
End Property

Public Property Get Violet() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetViolet, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Violet = Ret
End Property

Public Property Get Wheat() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetWheat, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Wheat = Ret
End Property

Public Property Get White() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetWhite, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set White = Ret
End Property

Public Property Get WhiteSmoke() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetWhiteSmoke, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set WhiteSmoke = Ret
End Property

Public Property Get Yellow() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetYellow, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set Yellow = Ret
End Property

Public Property Get YellowGreen() As Color
    Dim Ret As Color
    If m_pIColorsStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorsStatics, _
                  IColorsStatics_GetYellowGreen, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set YellowGreen = Ret
End Property

' ----==== Functions ====----
' IColorHelperStatics
Public Function FromArgb(ByVal A As Byte, _
                         ByVal R As Byte, _
                         ByVal G As Byte, _
                         ByVal B As Byte) As Color
    Dim Ret As Color
    If m_pIColorHelperStatics <> 0& Then
        Dim Value As Long
        If Invoke(m_pIColorHelperStatics, _
                  IColorHelperStatics_FromArgb, _
                  A, R, G, B, _
                  VarPtr(Value)) = S_OK Then
            Set Ret = New Color
            Ret.ColorValue = Value
        End If
    End If
    Set FromArgb = Ret
End Function
                        
' IColorHelperStatics2
Public Function ToDisplayName(ByVal Value As Color) As String
    Dim Ret As String
    If m_pIColorHelperStatics2 <> 0& And _
       IsNotNothing(Value) Then
        Dim hString As Long
        If Invoke(m_pIColorHelperStatics2, _
                  IColorHelperStatics2_ToDisplayName, _
                  Value.ColorValue, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    ToDisplayName = Ret
End Function

' ----==== extra Functions ====----
Public Function FromRgb(ByVal R As Byte, _
                        ByVal G As Byte, _
                        ByVal B As Byte) As Color
    Set FromRgb = FromArgb(255, R, G, B)
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function

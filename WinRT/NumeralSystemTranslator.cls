VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NumeralSystemTranslator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 10/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Globalization_NumberFormatting_NumeralSystemTranslator As String = "Windows.Globalization.NumberFormatting.NumeralSystemTranslator"
Private Const IID_INumeralSystemTranslator As String = "{28f5bc2c-8c23-4234-ad2e-fa5a3a426e9b}"
Private Const IID_INumeralSystemTranslatorFactory As String = "{9630c8da-36ef-4d88-a85c-6f0d98d620a6}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' INumeralSystemTranslatorFactory
    INumeralSystemTranslatorFactory_Create = 6
    
    ' INumeralSystemTranslator
    INumeralSystemTranslator_GetLanguages = 6
    INumeralSystemTranslator_GetResolvedLanguage = 7
    INumeralSystemTranslator_GetNumeralSystem = 8
    INumeralSystemTranslator_PutNumeralSystem = 9
    INumeralSystemTranslator_TranslateNumerals = 10
                        
End Enum

' ----==== Variablen ====----
Private m_pINumeralSystemTranslator As Long
Private m_pINumeralSystemTranslatorFactory As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Globalization_NumberFormatting_NumeralSystemTranslator, _
                            IID_INumeralSystemTranslatorFactory, _
                            m_pINumeralSystemTranslatorFactory) Then
    End If
    Dim pINumeralSystemTranslator As Long
    If GetActivateInstance(Windows_Globalization_NumberFormatting_NumeralSystemTranslator, _
                           IID_INumeralSystemTranslator, _
                           pINumeralSystemTranslator) Then
        Me.Ifc = pINumeralSystemTranslator
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pINumeralSystemTranslator)
    Call ReleaseIfc(m_pINumeralSystemTranslatorFactory)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pINumeralSystemTranslator
End Property

Public Property Let Ifc(ByVal pINumeralSystemTranslator As Long)
    Call ReleaseIfc(m_pINumeralSystemTranslator)
    m_pINumeralSystemTranslator = pINumeralSystemTranslator
End Property

' INumeralSystemTranslator
Public Property Get Languages() As ReadOnlyList_1 'ReadOnlyList_String
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_String
    If m_pINumeralSystemTranslator <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pINumeralSystemTranslator, _
                  INumeralSystemTranslator_GetLanguages, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_String, pIVectorView) 'ReadOnlyList_String
            End If
        End If
    End If
    Set Languages = Ret
End Property

Public Property Get ResolvedLanguage() As String
    Dim Ret As String
    If m_pINumeralSystemTranslator <> 0& Then
        Dim value As Long
        If Invoke(m_pINumeralSystemTranslator, _
                  INumeralSystemTranslator_GetResolvedLanguage, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    ResolvedLanguage = Ret
End Property

Public Property Get NumeralSystem() As String
    Dim Ret As String
    If m_pINumeralSystemTranslator <> 0& Then
        Dim value As Long
        If Invoke(m_pINumeralSystemTranslator, _
                  INumeralSystemTranslator_GetNumeralSystem, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    NumeralSystem = Ret
End Property

Public Property Let NumeralSystem(ByVal value As String)
    If m_pINumeralSystemTranslator <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pINumeralSystemTranslator, _
                  INumeralSystemTranslator_PutNumeralSystem, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

' ----==== Functions ====----
' INumeralSystemTranslatorFactory
Public Function Create(ByVal Languages As List_String) As NumeralSystemTranslator
    Dim Ret As NumeralSystemTranslator
    If m_pINumeralSystemTranslatorFactory <> 0& Then
        Dim pINumeralSystemTranslator As Long
        If Invoke(m_pINumeralSystemTranslatorFactory, _
                  INumeralSystemTranslatorFactory_Create, _
                  VarPtr(pINumeralSystemTranslator)) = S_OK Then
            If pINumeralSystemTranslator <> 0& Then
                Set Ret = New NumeralSystemTranslator
                Ret.Ifc = pINumeralSystemTranslator
            End If
        End If
    End If
    Set Create = Ret
End Function

' INumeralSystemTranslator
Public Function TranslateNumerals(ByVal value As String) As String
    Dim Ret As String
    If m_pINumeralSystemTranslator <> 0& Then
        Dim Result As Long
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pINumeralSystemTranslator, _
                  INumeralSystemTranslator_TranslateNumerals, _
                  hString, _
                  VarPtr(Result)) = S_OK Then
            Ret = GetWindowsString(Result)
        End If
        Call DeleteWindowsString(hString)
    End If
    TranslateNumerals = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OcrEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 06/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Media_Ocr_OcrEngine As String = "Windows.Media.Ocr.OcrEngine"
Private Const IID_IOcrEngine As String = "{5a14bc41-5b76-3140-b680-8825562683ac}"
Private Const IID_IOcrEngineStatics As String = "{5bffa85a-3384-3540-9940-699120d428a8}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IOcrEngineStatics
    IOcrEngineStatics_GetMaxImageDimension = 6
    IOcrEngineStatics_GetAvailableRecognizerLanguages = 7
    IOcrEngineStatics_IsLanguageSupported = 8
    IOcrEngineStatics_TryCreateFromLanguage = 9
    IOcrEngineStatics_TryCreateFromUserProfileLanguages = 10
    
    ' IOcrEngine
    IOcrEngine_RecognizeAsync = 6
    IOcrEngine_GetRecognizerLanguage = 7
    
End Enum

' ----==== Variablen ====----
Private m_pIOcrEngine As Long
Private m_pIOcrEngineStatics As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Media_Ocr_OcrEngine, _
                            IID_IOcrEngineStatics, _
                            m_pIOcrEngineStatics) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIOcrEngine)
    Call ReleaseIfc(m_pIOcrEngineStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIOcrEngine
End Property

Public Property Let Ifc(ByVal pIOcrEngine As Long)
    Call ReleaseIfc(m_pIOcrEngine)
    m_pIOcrEngine = pIOcrEngine
End Property

' IOcrEngineStatics
Public Property Get MaxImageDimension() As Long
    Dim Ret As Long
    If m_pIOcrEngineStatics <> 0& Then
        Dim value As Long
        If Invoke(m_pIOcrEngineStatics, _
                  IOcrEngineStatics_GetMaxImageDimension, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    MaxImageDimension = Ret
End Property

Public Property Get AvailableRecognizerLanguages() As ReadOnlyList_1 'ReadOnlyList_Language
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_Language
    If m_pIOcrEngineStatics <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIOcrEngineStatics, _
                  IOcrEngineStatics_GetAvailableRecognizerLanguages, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                Set Ret = New_ReadOnlyList_1(Of_Language, pIVectorView) 'ReadOnlyList_Language
            End If
        End If
    End If
    Set AvailableRecognizerLanguages = Ret
End Property

' IOcrEngine
Public Property Get RecognizerLanguage() As Language
    Dim Ret As Language
    If m_pIOcrEngine <> 0& Then
        Dim pILanguage As Long
        If Invoke(m_pIOcrEngine, _
                  IOcrEngine_GetRecognizerLanguage, _
                  VarPtr(pILanguage)) = S_OK Then
            If pILanguage <> 0& Then
                Set Ret = New Language
                Ret.Ifc = pILanguage
            End If
        End If
    End If
    Set RecognizerLanguage = Ret
End Property

' ----==== Functions ====----
' IOcrEngineStatics
Public Function IsLanguageSupported(ByVal lang As Language) As Boolean
    Dim Ret As Boolean
    If m_pIOcrEngineStatics <> 0& And _
       IsNotNothing(lang) Then
        Dim Result As Long
        If Invoke(m_pIOcrEngineStatics, _
                  IOcrEngineStatics_IsLanguageSupported, _
                  lang.Ifc, _
                  VarPtr(Result)) = S_OK Then
            Ret = CBool(Result)
        End If
    End If
    IsLanguageSupported = Ret
End Function

Public Function TryCreateFromLanguage(ByVal lang As Language) As OcrEngine
    Dim Ret As OcrEngine
    If m_pIOcrEngineStatics <> 0& And _
       IsNotNothing(lang) Then
        Dim pIOcrEngine As Long
        If Invoke(m_pIOcrEngineStatics, _
                  IOcrEngineStatics_TryCreateFromLanguage, _
                  lang.Ifc, _
                  VarPtr(pIOcrEngine)) = S_OK Then
            If pIOcrEngine <> 0& Then
                Set Ret = New OcrEngine
                Ret.Ifc = pIOcrEngine
            End If
        End If
    End If
    Set TryCreateFromLanguage = Ret
End Function

Public Function TryCreateFromUserProfileLanguages() As OcrEngine
    Dim Ret As OcrEngine
    If m_pIOcrEngineStatics <> 0& Then
        Dim pIOcrEngine As Long
        If Invoke(m_pIOcrEngineStatics, _
                  IOcrEngineStatics_TryCreateFromUserProfileLanguages, _
                  VarPtr(pIOcrEngine)) = S_OK Then
            If pIOcrEngine <> 0& Then
                Set Ret = New OcrEngine
                Ret.Ifc = pIOcrEngine
            End If
        End If
    End If
    Set TryCreateFromUserProfileLanguages = Ret
End Function

' IOcrEngine
Public Function RecognizeAsync(ByVal bitmap As SoftwareBitmap) As OcrResult
    Dim Ret As OcrResult
    If m_pIOcrEngine <> 0& And _
       IsNotNothing(bitmap) Then
        Dim pIOcrResult As Long
        If Invoke(m_pIOcrEngine, _
                  IOcrEngine_RecognizeAsync, _
                  bitmap.Ifc, _
                  VarPtr(pIOcrResult)) = S_OK Then
            If pIOcrResult <> 0& Then
                If Await(pIOcrResult) Then
                    Set Ret = New OcrResult
                    Ret.Ifc = pIOcrResult
                End If
            End If
        End If
    End If
    Set RecognizeAsync = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



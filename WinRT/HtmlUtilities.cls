VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HtmlUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Data_Html_HtmlUtilities As String = "Windows.Data.Html.HtmlUtilities"
Private Const IID_IHtmlUtilities As String = "{fec00add-2399-4fac-b5a7-05e9acd7181d}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IHtmlUtilities
    IHtmlUtilities_ConvertToText = 6
    
End Enum

' ----==== Vars ====----
Private m_pIHtmlUtilities As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Data_Html_HtmlUtilities, _
                            IID_IHtmlUtilities, _
                            m_pIHtmlUtilities) Then
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIHtmlUtilities)
End Sub

' ----==== Properties ====----

' ----==== Functions ====----
' IHtmlUtilities
Public Function ConvertToText(ByVal htmlStr As String) As String
    Dim Ret As String
    If m_pIHtmlUtilities <> 0& Then
        Dim hHtml As Long
        Dim hText As Long
        hHtml = CreateWindowsString(htmlStr)
        If Invoke(m_pIHtmlUtilities, _
                  IHtmlUtilities_ConvertToText, _
                  hHtml, _
                  VarPtr(hText)) = S_OK Then
            Ret = GetWindowsString(hText)
        End If
        Call DeleteWindowsString(hHtml)
    End If
    ConvertToText = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XmlLoadSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Data_Xml_Dom_XmlLoadSettings As String = "Windows.Data.Xml.Dom.XmlLoadSettings"
Private Const IID_IXmlLoadSettings As String = "{58aa07a8-fed6-46f7-b4c5-fb1ba72108d6}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IXmlLoadSettings
    IXmlLoadSettings_GetMaxElementDepth = 6
    IXmlLoadSettings_PutMaxElementDepth = 7
    IXmlLoadSettings_GetProhibitDtd = 8
    IXmlLoadSettings_PutProhibitDtd = 9
    IXmlLoadSettings_GetResolveExternals = 10
    IXmlLoadSettings_PutResolveExternals = 11
    IXmlLoadSettings_GetValidateOnParse = 12
    IXmlLoadSettings_PutValidateOnParse = 13
    IXmlLoadSettings_GetElementContentWhiteSpace = 14
    IXmlLoadSettings_PutElementContentWhiteSpace = 15
                            
End Enum

' ----==== Variablen ====----
Private m_pIXmlLoadSettings As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    Dim pIXmlLoadSettings As Long
    If GetActivateInstance(Windows_Data_Xml_Dom_XmlLoadSettings, _
                           IID_IXmlLoadSettings, _
                           pIXmlLoadSettings) Then
        Me.Ifc = pIXmlLoadSettings
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIXmlLoadSettings)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIXmlLoadSettings
End Property

Public Property Let Ifc(ByVal pIXmlLoadSettings As Long)
    Call ReleaseIfc(m_pIXmlLoadSettings)
    m_pIXmlLoadSettings = pIXmlLoadSettings
End Property

' IXmlLoadSettings
Public Property Get MaxElementDepth() As Long
    Dim Ret As Long
    If m_pIXmlLoadSettings <> 0& Then
        Dim Value As Long
        If Invoke(m_pIXmlLoadSettings, _
                  IXmlLoadSettings_GetMaxElementDepth, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    MaxElementDepth = Ret
End Property

Public Property Let MaxElementDepth(ByVal Value As Long)
    If m_pIXmlLoadSettings <> 0& Then
        If Invoke(m_pIXmlLoadSettings, _
                  IXmlLoadSettings_PutMaxElementDepth, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get ProhibitDtd() As Boolean
    Dim Ret As Boolean
    If m_pIXmlLoadSettings <> 0& Then
        Dim Value As Long
        If Invoke(m_pIXmlLoadSettings, _
                  IXmlLoadSettings_GetProhibitDtd, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    ProhibitDtd = Ret
End Property

Public Property Let ProhibitDtd(ByVal Value As Boolean)
    If m_pIXmlLoadSettings <> 0& Then
        If Invoke(m_pIXmlLoadSettings, _
                  IXmlLoadSettings_PutProhibitDtd, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get ResolveExternals() As Boolean
    Dim Ret As Boolean
    If m_pIXmlLoadSettings <> 0& Then
        Dim Value As Long
        If Invoke(m_pIXmlLoadSettings, _
                  IXmlLoadSettings_GetResolveExternals, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    ResolveExternals = Ret
End Property

Public Property Let ResolveExternals(ByVal Value As Boolean)
    If m_pIXmlLoadSettings <> 0& Then
        If Invoke(m_pIXmlLoadSettings, _
                  IXmlLoadSettings_PutResolveExternals, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get ValidateOnParse() As Boolean
    Dim Ret As Boolean
    If m_pIXmlLoadSettings <> 0& Then
        Dim Value As Long
        If Invoke(m_pIXmlLoadSettings, _
                  IXmlLoadSettings_GetValidateOnParse, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    ValidateOnParse = Ret
End Property

Public Property Let ValidateOnParse(ByVal Value As Boolean)
    If m_pIXmlLoadSettings <> 0& Then
        If Invoke(m_pIXmlLoadSettings, _
                  IXmlLoadSettings_PutValidateOnParse, _
                  Value) = S_OK Then
        End If
    End If
End Property

Public Property Get ElementContentWhiteSpace() As Boolean
    Dim Ret As Boolean
    If m_pIXmlLoadSettings <> 0& Then
        Dim Value As Long
        If Invoke(m_pIXmlLoadSettings, _
                  IXmlLoadSettings_GetElementContentWhiteSpace, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    ElementContentWhiteSpace = Ret
End Property

Public Property Let ElementContentWhiteSpace(ByVal Value As Boolean)
    If m_pIXmlLoadSettings <> 0& Then
        If Invoke(m_pIXmlLoadSettings, _
                  IXmlLoadSettings_PutElementContentWhiteSpace, _
                  Value) = S_OK Then
        End If
    End If
End Property

' ----==== Functions ====----

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


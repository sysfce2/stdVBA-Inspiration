VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XmlDomImplementation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 08/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Data_Xml_Dom_XmlDomImplementation As String = "Windows.Data.Xml.Dom.XmlDomImplementation"
Private Const IID_IXmlDomImplementation As String = "{6de58132-f11d-4fbb-8cc6-583cba93112f}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IXmlDomImplementation
    IXmlDomImplementation_HasFeature = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIXmlDomImplementation As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIXmlDomImplementation)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIXmlDomImplementation
End Property

Public Property Let Ifc(ByVal pIXmlDomImplementation As Long)
    Call ReleaseIfc(m_pIXmlDomImplementation)
    m_pIXmlDomImplementation = pIXmlDomImplementation
End Property

' ----==== Functions ====----
' IXmlDomImplementation
Public Function HasFeature(ByVal feature As String, _
                           ByVal version As Inspectable) As Boolean
    Dim Ret As Boolean
    If m_pIXmlDomImplementation <> 0& And IsNotNothing(version) Then
        Dim hFeature As Long
        Dim featureSupported As Long
        hFeature = CreateWindowsString(feature)
        If Invoke(m_pIXmlDomImplementation, _
                  IXmlDomImplementation_HasFeature, _
                  hFeature, _
                  version.Ifc, _
                  VarPtr(featureSupported)) = S_OK Then
            Ret = CBool(featureSupported)
        End If
        Call DeleteWindowsString(hFeature)
    End If
    HasFeature = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


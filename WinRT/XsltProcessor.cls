VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XsltProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 08/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Data_Xml_Xsl_XsltProcessor As String = "Windows.Data.Xml.Xsl.XsltProcessor"
Private Const IID_IXsltProcessorFactory As String = "{274146c0-9a51-4663-bf30-0ef742146f20}"
Private Const IID_IXsltProcessor As String = "{7b64703f-550c-48c6-a90f-93a5b964518f}"
Private Const IID_IXsltProcessor2 As String = "{8da45c56-97a5-44cb-a8be-27d86280c70a}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IXsltProcessorFactory
    IXsltProcessorFactory_CreateInstance = 6
                            
    ' IXsltProcessor
    IXsltProcessor_TransformToString = 6
    
    ' IXsltProcessor2
    IXsltProcessor2_TransformToDocument = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIXsltProcessorFactory As Long
Private m_pIXsltProcessor As Long
Private m_pIXsltProcessor2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Data_Xml_Xsl_XsltProcessor, _
                            IID_IXsltProcessorFactory, _
                            m_pIXsltProcessorFactory) Then
    End If
    Dim pIXsltProcessor As Long
    If GetActivateInstance(Windows_Data_Xml_Xsl_XsltProcessor, _
                           IID_IXsltProcessor, _
                           pIXsltProcessor) Then
        Me.Ifc = pIXsltProcessor
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIXsltProcessorFactory)
    Call ReleaseIfc(m_pIXsltProcessor2)
    Call ReleaseIfc(m_pIXsltProcessor)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIXsltProcessor
End Property

Public Property Let Ifc(ByVal pIXsltProcessor As Long)
    Call ReleaseIfc(m_pIXsltProcessor2)
    Call ReleaseIfc(m_pIXsltProcessor)
    m_pIXsltProcessor = pIXsltProcessor
    Call QueryIfc(m_pIXsltProcessor, IID_IXsltProcessor2, m_pIXsltProcessor2)
End Property

' ----==== Functions ====----
' IXsltProcessorFactory
Public Function CreateInstance(ByVal document As XmlDocument) As XsltProcessor
    Dim Ret As XsltProcessor
    If m_pIXsltProcessorFactory <> 0& And _
       IsNotNothing(document) Then
        Dim pIXsltProcessor As Long
        If Invoke(m_pIXsltProcessorFactory, _
                  IXsltProcessorFactory_CreateInstance, _
                  document.Ifc, _
                  VarPtr(pIXsltProcessor)) = S_OK Then
            If pIXsltProcessor <> 0& Then
                Set Ret = New XsltProcessor
                Ret.Ifc = pIXsltProcessor
            End If
        End If
    End If
    Set CreateInstance = Ret
End Function
                            
' IXsltProcessor
Public Function TransformToString(ByVal inputNode As XmlNode) As String
    Dim Ret As String
    If m_pIXsltProcessor <> 0& And _
       IsNotNothing(inputNode) Then
        Dim hString As Long
        If Invoke(m_pIXsltProcessor, _
                  IXsltProcessor_TransformToString, _
                  inputNode.Ifc, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    TransformToString = Ret
End Function
    
' IXsltProcessor2
Public Function TransformToDocument(ByVal inputNode As XmlNode) As XmlDocument
    Dim Ret As XmlDocument
    If m_pIXsltProcessor2 <> 0& And _
       IsNotNothing(inputNode) Then
        Dim pIXmlDocument As Long
        If Invoke(m_pIXsltProcessor2, _
                  IXsltProcessor2_TransformToDocument, _
                  inputNode.Ifc, _
                  VarPtr(pIXmlDocument)) = S_OK Then
            If pIXmlDocument <> 0& Then
                Set Ret = New XmlDocument
                Ret.Ifc = pIXmlDocument
            End If
        End If
    End If
    Set TransformToDocument = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


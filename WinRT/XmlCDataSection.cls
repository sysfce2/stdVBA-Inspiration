VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XmlCDataSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 08/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Data_Xml_Dom_XmlComment As String = "Windows.Data.Xml.Dom.XmlComment"
Private Const IID_IXmlCDataSection As String = "{4d04b46f-c8bd-45b4-8899-0400d7c2c60f}"
Private Const IID_IXmlText As String = "{f931a4cb-308d-4760-a1d5-43b67450ac7e}"
Private Const IID_IXmlCharacterData As String = "{132e42ab-4e36-4df6-b1c8-0ce62fd88b26}"
Private Const IID_IXmlNode As String = "{1c741d59-2122-47d5-a856-83f3d4214875}"
Private Const IID_IXmlNodeSelector As String = "{63dbba8b-d0db-4fe1-b745-f9433afdc25b}"
Private Const IID_IXmlNodeSerializer As String = "{5cc5b382-e6dd-4991-abef-06d8d2e7bd0c}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IXmlCDataSection
    ' has no member
                            
    ' IXmlText
    IXmlText_SplitText = 6
                            
    ' IXmlCharacterData
    IXmlCharacterData_GetData = 6
    IXmlCharacterData_PutData = 7
    IXmlCharacterData_GetLength = 8
    IXmlCharacterData_SubstringData = 9
    IXmlCharacterData_AppendData = 10
    IXmlCharacterData_InsertData = 11
    IXmlCharacterData_DeleteData = 12
    IXmlCharacterData_ReplaceData = 13
                            
    ' IXmlNode
    IXmlNode_GetNodeValue = 6
    IXmlNode_PutNodeValue = 7
    IXmlNode_GetNodeType = 8
    IXmlNode_GetNodeName = 9
    IXmlNode_GetParentNode = 10
    IXmlNode_GetChildNodes = 11
    IXmlNode_GetFirstChild = 12
    IXmlNode_GetLastChild = 13
    IXmlNode_GetPreviousSibling = 14
    IXmlNode_GetNextSibling = 15
    IXmlNode_GetAttributes = 16
    IXmlNode_HasChildNodes = 17
    IXmlNode_GetOwnerDocument = 18
    IXmlNode_InsertBefore = 19
    IXmlNode_ReplaceChild = 20
    IXmlNode_RemoveChild = 21
    IXmlNode_AppendChild = 22
    IXmlNode_CloneNode = 23
    IXmlNode_GetNamespaceUri = 24
    IXmlNode_GetLocalName = 25
    IXmlNode_GetPrefix = 26
    IXmlNode_Normalize = 27
    IXmlNode_PutPrefix = 28
    
    ' IXmlNodeSelector
    IXmlNodeSelector_SelectSingleNode = 6
    IXmlNodeSelector_SelectNodes = 7
    IXmlNodeSelector_SelectSingleNodeNS = 8
    IXmlNodeSelector_SelectNodesNS = 9

    ' IXmlNodeSerializer
    IXmlNodeSerializer_GetXml = 6
    IXmlNodeSerializer_GetInnerText = 7
    IXmlNodeSerializer_PutInnerText = 8
    
End Enum
' ----==== Variablen ====----
Private m_pIXmlCDataSection As Long
Private m_pIXmlText As Long
Private m_pIXmlNode As Long
Private m_pIXmlNodeSelector As Long
Private m_pIXmlCharacterData As Long
Private m_pIXmlNodeSerializer As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIXmlText)
    Call ReleaseIfc(m_pIXmlNode)
    Call ReleaseIfc(m_pIXmlCharacterData)
    Call ReleaseIfc(m_pIXmlNodeSelector)
    Call ReleaseIfc(m_pIXmlNodeSerializer)
    Call ReleaseIfc(m_pIXmlCDataSection)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIXmlCDataSection
End Property

Public Property Let Ifc(ByVal pIXmlCDataSection As Long)
    Call ReleaseIfc(m_pIXmlText)
    Call ReleaseIfc(m_pIXmlNode)
    Call ReleaseIfc(m_pIXmlCharacterData)
    Call ReleaseIfc(m_pIXmlNodeSelector)
    Call ReleaseIfc(m_pIXmlNodeSerializer)
    Call ReleaseIfc(m_pIXmlCDataSection)
    m_pIXmlCDataSection = pIXmlCDataSection
    Call QueryIfc(m_pIXmlCDataSection, IID_IXmlText, m_pIXmlText)
    Call QueryIfc(m_pIXmlCDataSection, IID_IXmlNode, m_pIXmlNode)
    Call QueryIfc(m_pIXmlCDataSection, IID_IXmlCharacterData, m_pIXmlCharacterData)
    Call QueryIfc(m_pIXmlCDataSection, IID_IXmlNodeSelector, m_pIXmlNodeSelector)
    Call QueryIfc(m_pIXmlCDataSection, IID_IXmlNodeSerializer, m_pIXmlNodeSerializer)
End Property

' IXmlCharacterData
Public Property Get data() As String
    Dim Ret As String
    If m_pIXmlCharacterData <> 0& Then
        Dim value As Long
        If Invoke(m_pIXmlCharacterData, _
                  IXmlCharacterData_GetData, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    data = Ret
End Property

Public Property Let data(ByVal value As String)
    Dim Ret As String
    If m_pIXmlCharacterData <> 0& Then
        Dim hValue As Long
        hValue = CreateWindowsString(value)
        If Invoke(m_pIXmlCharacterData, _
                  IXmlCharacterData_PutData, _
                  hValue) = S_OK Then
        End If
        Call DeleteWindowsString(hValue)
    End If
End Property

Public Property Get Length() As Long
    Dim Ret As Long
    If m_pIXmlCharacterData <> 0& Then
        Dim value As Long
        If Invoke(m_pIXmlCharacterData, _
                  IXmlCharacterData_GetLength, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    Length = Ret
End Property

' IXmlNode
Public Property Get NodeValue() As Inspectable
    Dim Ret As Inspectable
    If m_pIXmlNode <> 0& Then
        Dim value As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_GetNodeValue, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                Set Ret = New Inspectable
                Ret.Ifc = value
            End If
        End If
    End If
    Set NodeValue = Ret
End Property

Public Property Let NodeValue(ByVal value As Inspectable)
    If m_pIXmlNode <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pIXmlNode, _
                  IXmlNode_PutNodeValue, _
                  value.Ifc) = S_OK Then
        End If
    End If
End Property

Public Property Get NodeType() As NodeType
    Dim Ret As NodeType
    If m_pIXmlNode <> 0& Then
        Dim value As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_GetNodeType, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    NodeType = Ret
End Property

Public Property Get NodeName() As String
    Dim Ret As String
    If m_pIXmlNode <> 0& Then
        Dim value As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_GetNodeName, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    NodeName = Ret
End Property

Public Property Get ParentNode() As XmlNode
    Dim Ret As XmlNode
    If m_pIXmlNode <> 0& Then
        Dim pIXmlNode As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_GetParentNode, _
                  VarPtr(pIXmlNode)) = S_OK Then
            If pIXmlNode <> 0& Then
                Set Ret = New XmlNode
                Ret.Ifc = pIXmlNode
            End If
        End If
    End If
    Set ParentNode = Ret
End Property

Public Property Get ChildNodes() As XmlNodeList
    Dim Ret As XmlNodeList
    If m_pIXmlNode <> 0& Then
        Dim pIXmlNodeList As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_GetChildNodes, _
                  VarPtr(pIXmlNodeList)) = S_OK Then
            If pIXmlNodeList <> 0& Then
                Set Ret = New XmlNodeList
                Ret.Ifc = pIXmlNodeList
            End If
        End If
    End If
    Set ChildNodes = Ret
End Property

Public Property Get FirstChild() As XmlNode
    Dim Ret As XmlNode
    If m_pIXmlNode <> 0& Then
        Dim pIXmlNode As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_GetFirstChild, _
                  VarPtr(pIXmlNode)) = S_OK Then
            If pIXmlNode <> 0& Then
                Set Ret = New XmlNode
                Ret.Ifc = pIXmlNode
            End If
        End If
    End If
    Set FirstChild = Ret
End Property

Public Property Get LastChild() As XmlNode
    Dim Ret As XmlNode
    If m_pIXmlNode <> 0& Then
        Dim pIXmlNode As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_GetLastChild, _
                  VarPtr(pIXmlNode)) = S_OK Then
            If pIXmlNode <> 0& Then
                Set Ret = New XmlNode
                Ret.Ifc = pIXmlNode
            End If
        End If
    End If
    Set LastChild = Ret
End Property

Public Property Get PreviousSibling() As XmlNode
    Dim Ret As XmlNode
    If m_pIXmlNode <> 0& Then
        Dim pIXmlNode As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_GetPreviousSibling, _
                  VarPtr(pIXmlNode)) = S_OK Then
            If pIXmlNode <> 0& Then
                Set Ret = New XmlNode
                Ret.Ifc = pIXmlNode
            End If
        End If
    End If
    Set PreviousSibling = Ret
End Property

Public Property Get NextSibling() As XmlNode
    Dim Ret As XmlNode
    If m_pIXmlNode <> 0& Then
        Dim pIXmlNode As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_GetNextSibling, _
                  VarPtr(pIXmlNode)) = S_OK Then
            If pIXmlNode <> 0& Then
                Set Ret = New XmlNode
                Ret.Ifc = pIXmlNode
            End If
        End If
    End If
    Set NextSibling = Ret
End Property

Public Property Get Attributes() As XmlNamedNodeMap
    Dim Ret As XmlNamedNodeMap
    If m_pIXmlNode <> 0& Then
        Dim pIXmlNamedNodeMap As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_GetAttributes, _
                  VarPtr(pIXmlNamedNodeMap)) = S_OK Then
            If pIXmlNamedNodeMap <> 0& Then
                Set Ret = New XmlNamedNodeMap
                Ret.Ifc = pIXmlNamedNodeMap
            End If
        End If
    End If
    Set Attributes = Ret
End Property

Public Property Get OwnerDocument() As XmlDocument
    Dim Ret As XmlDocument
    If m_pIXmlNode <> 0& Then
        Dim pIXmlDocument As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_GetOwnerDocument, _
                  VarPtr(pIXmlDocument)) = S_OK Then
            If pIXmlDocument <> 0& Then
                Set Ret = New XmlDocument
                Ret.Ifc = pIXmlDocument
            End If
        End If
    End If
    Set OwnerDocument = Ret
End Property

Public Property Get namespaceUri() As Inspectable
    Dim Ret As Inspectable
    If m_pIXmlNode <> 0& Then
        Dim value As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_GetNamespaceUri, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                Set Ret = New Inspectable
                Ret.Ifc = value
            End If
        End If
    End If
    Set namespaceUri = Ret
End Property

Public Property Get localName() As Inspectable
    Dim Ret As Inspectable
    If m_pIXmlNode <> 0& Then
        Dim value As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_GetLocalName, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                Set Ret = New Inspectable
                Ret.Ifc = value
            End If
        End If
    End If
    Set localName = Ret
End Property

Public Property Get Prefix() As Inspectable
    Dim Ret As Inspectable
    If m_pIXmlNode <> 0& Then
        Dim value As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_GetPrefix, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                Set Ret = New Inspectable
                Ret.Ifc = value
            End If
        End If
    End If
    Set Prefix = Ret
End Property

Public Property Let Prefix(ByVal value As Inspectable)
    If m_pIXmlNode <> 0& And _
       IsNotNothing(value) Then
        If Invoke(m_pIXmlNode, _
                  IXmlNode_PutPrefix, _
                  value.Ifc) = S_OK Then
        End If
    End If
End Property

' IXmlNodeSerializer
Public Property Get InnerText() As String
    Dim Ret As String
    If m_pIXmlNodeSerializer <> 0& Then
        Dim value As Long
        If Invoke(m_pIXmlNodeSerializer, _
                  IXmlNodeSerializer_GetInnerText, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    InnerText = Ret
End Property

Public Property Let InnerText(ByVal value As String)
    If m_pIXmlNodeSerializer <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(value)
        If Invoke(m_pIXmlNodeSerializer, _
                  IXmlNodeSerializer_PutInnerText, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

' ----==== Functions ====----
' IXmlText
Public Function SplitText(ByVal offset As Long) As XmlText
    Dim Ret As XmlText
        If m_pIXmlText <> 0& Then
            Dim pIXmlText As Long
            If Invoke(m_pIXmlText, _
                      IXmlText_SplitText, _
                      offset, _
                      VarPtr(pIXmlText)) = S_OK Then
                If pIXmlText <> 0& Then
                    Set Ret = New XmlText
                    Ret.Ifc = pIXmlText
                End If
            End If
        End If
    Set SplitText = Ret
End Function

' IXmlCharacterData
Public Function SubstringData(ByVal offset As Long, _
                              ByVal count As Long) As String
    Dim Ret As String
    If m_pIXmlCharacterData <> 0& Then
        Dim value As Long
        If Invoke(m_pIXmlCharacterData, _
                  IXmlCharacterData_SubstringData, _
                  offset, _
                  count, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    SubstringData = Ret
End Function

Public Function AppendData(ByVal data As String) As Boolean
    Dim Ret As Boolean
    If m_pIXmlCharacterData <> 0& Then
        Dim hData As Long
        hData = CreateWindowsString(data)
        If Invoke(m_pIXmlCharacterData, _
                  IXmlCharacterData_AppendData, _
                  hData) = S_OK Then
            Ret = True
        End If
        Call DeleteWindowsString(hData)
    End If
    AppendData = Ret
End Function

Public Function InsertData(ByVal offset As Long, _
                           ByVal data As String) As Boolean
    Dim Ret As Boolean
    If m_pIXmlCharacterData <> 0& Then
        Dim hData As Long
        hData = CreateWindowsString(data)
        If Invoke(m_pIXmlCharacterData, _
                  IXmlCharacterData_InsertData, _
                  offset, _
                  hData) = S_OK Then
            Ret = True
        End If
        Call DeleteWindowsString(hData)
    End If
    InsertData = Ret
End Function

Public Function DeleteData(ByVal offset As Long, _
                           ByVal count As Long) As Boolean
    Dim Ret As Boolean
    If m_pIXmlCharacterData <> 0& Then
        If Invoke(m_pIXmlCharacterData, _
                  IXmlCharacterData_DeleteData, _
                  offset, _
                  count) = S_OK Then
            Ret = True
        End If
    End If
    DeleteData = Ret
End Function

Public Function ReplaceData(ByVal offset As Long, _
                            ByVal count As Long, _
                            ByVal data As String) As Boolean
    Dim Ret As Boolean
    If m_pIXmlCharacterData <> 0& Then
        Dim hData As Long
        hData = CreateWindowsString(data)
        If Invoke(m_pIXmlCharacterData, _
                  IXmlCharacterData_ReplaceData, _
                  offset, _
                  count, _
                  hData) = S_OK Then
            Ret = True
        End If
        Call DeleteWindowsString(hData)
    End If
    ReplaceData = Ret
End Function

' IXmlNode
Public Function HasChildNodes() As Boolean
    Dim Ret As Boolean
    If m_pIXmlNode <> 0& Then
        Dim value As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_HasChildNodes, _
                  VarPtr(value)) = S_OK Then
            Ret = CBool(value)
        End If
    End If
    HasChildNodes = Ret
End Function

Public Function InsertBefore(ByVal newChild As Object, _
                             ByVal referenceChild As Object) As XmlNode
    Dim Ret As XmlNode
    If m_pIXmlNode <> 0& And _
       IsNotNothing(newChild) And _
       IsNotNothing(referenceChild) Then
        Dim pIXmlNode As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_InsertBefore, _
                  newChild.Ifc, _
                  referenceChild.Ifc, _
                  VarPtr(pIXmlNode)) = S_OK Then
            If pIXmlNode <> 0& Then
                Set Ret = New XmlNode
                Ret.Ifc = pIXmlNode
            End If
        End If
    End If
    Set InsertBefore = Ret
End Function

Public Function ReplaceChild(ByVal newChild As Object, _
                             ByVal referenceChild As Object) As XmlNode
    Dim Ret As XmlNode
    If m_pIXmlNode <> 0& And _
       IsNotNothing(newChild) And _
       IsNotNothing(referenceChild) Then
        Dim pIXmlNode As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_ReplaceChild, _
                  newChild.Ifc, _
                  referenceChild.Ifc, _
                  VarPtr(pIXmlNode)) = S_OK Then
            If pIXmlNode <> 0& Then
                Set Ret = New XmlNode
                Ret.Ifc = pIXmlNode
            End If
        End If
    End If
    Set ReplaceChild = Ret
End Function

Public Function RemoveChild(ByVal childNode As Object) As XmlNode
    Dim Ret As XmlNode
    If m_pIXmlNode <> 0& And _
       IsNotNothing(childNode) Then
        Dim pIXmlNode As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_RemoveChild, _
                  childNode.Ifc, _
                  VarPtr(pIXmlNode)) = S_OK Then
            If pIXmlNode <> 0& Then
                Set Ret = New XmlNode
                Ret.Ifc = pIXmlNode
            End If
        End If
    End If
    Set RemoveChild = Ret
End Function

Public Function AppendChild(ByVal newChild As Object) As XmlNode
    Dim Ret As XmlNode
    If m_pIXmlNode <> 0& And _
       IsNotNothing(newChild) Then
        Dim pIXmlNode As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_AppendChild, _
                  newChild.Ifc, _
                  VarPtr(pIXmlNode)) = S_OK Then
            If pIXmlNode <> 0& Then
                Set Ret = New XmlNode
                Ret.Ifc = pIXmlNode
            End If
        End If
    End If
    Set AppendChild = Ret
End Function

Public Function CloneNode(ByVal deep As Boolean) As XmlNode
    Dim Ret As XmlNode
    If m_pIXmlNode <> 0& Then
        Dim pIXmlNode As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_CloneNode, _
                  deep, _
                  VarPtr(pIXmlNode)) = S_OK Then
            If pIXmlNode <> 0& Then
                Set Ret = New XmlNode
                Ret.Ifc = pIXmlNode
            End If
        End If
    End If
    Set CloneNode = Ret
End Function

Public Function Normalize() As Boolean
    Dim Ret As Boolean
    If m_pIXmlNode <> 0& Then
        If Invoke(m_pIXmlNode, _
                  IXmlNode_Normalize) = S_OK Then
            Ret = True
        End If
    End If
    Normalize = Ret
End Function

' IXmlNodeSelector
Public Function SelectSingleNode(ByVal xpath As String) As XmlNode
    Dim Ret As XmlNode
    If m_pIXmlNodeSelector <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(xpath)
        Dim pIXmlNode As Long
        If Invoke(m_pIXmlNodeSelector, _
                  IXmlNodeSelector_SelectSingleNode, _
                  hString, _
                  VarPtr(pIXmlNode)) = S_OK Then
            If pIXmlNode <> 0& Then
                Set Ret = New XmlNode
                Ret.Ifc = pIXmlNode
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set SelectSingleNode = Ret
End Function

Public Function SelectNodes(ByVal xpath As String) As XmlNodeList
    Dim Ret As XmlNodeList
    If m_pIXmlNodeSelector <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(xpath)
        Dim pIXmlNodeList As Long
        If Invoke(m_pIXmlNodeSelector, _
                  IXmlNodeSelector_SelectNodes, _
                  hString, _
                  VarPtr(pIXmlNodeList)) = S_OK Then
            If pIXmlNodeList <> 0& Then
                Set Ret = New XmlNodeList
                Ret.Ifc = pIXmlNodeList
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set SelectNodes = Ret
End Function

Public Function SelectSingleNodeNS(ByVal xpath As String, _
                                   ByVal namespaces As Inspectable) As XmlNode
    Dim Ret As XmlNode
    If m_pIXmlNodeSelector <> 0& And _
       IsNotNothing(namespaces) Then
        Dim hString As Long
        hString = CreateWindowsString(xpath)
        Dim pIXmlNode As Long
        If Invoke(m_pIXmlNodeSelector, _
                  IXmlNodeSelector_SelectSingleNodeNS, _
                  hString, _
                  namespaces.Ifc, _
                  VarPtr(pIXmlNode)) = S_OK Then
            If pIXmlNode <> 0& Then
                Set Ret = New XmlNode
                Ret.Ifc = pIXmlNode
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set SelectSingleNodeNS = Ret
End Function

Public Function SelectNodesNS(ByVal xpath As String, _
                              ByVal namespaces As Inspectable) As XmlNodeList
    Dim Ret As XmlNodeList
    If m_pIXmlNodeSelector <> 0& And _
       IsNotNothing(namespaces) Then
        Dim hString As Long
        hString = CreateWindowsString(xpath)
        Dim pIXmlNodeList As Long
        If Invoke(m_pIXmlNodeSelector, _
                  IXmlNodeSelector_SelectNodesNS, _
                  hString, _
                  namespaces.Ifc, _
                  VarPtr(pIXmlNodeList)) = S_OK Then
            If pIXmlNodeList <> 0& Then
                Set Ret = New XmlNodeList
                Ret.Ifc = pIXmlNodeList
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set SelectNodesNS = Ret
End Function

' IXmlNodeSerializer
Public Function GetXml() As String
    Dim Ret As String
    If m_pIXmlNodeSerializer <> 0& Then
        Dim value As Long
        If Invoke(m_pIXmlNodeSerializer, _
                  IXmlNodeSerializer_GetXml, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    GetXml = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



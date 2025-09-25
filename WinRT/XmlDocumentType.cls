VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XmlDocumentType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 08/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Data_Xml_Dom_XmlDocumentType As String = "Windows.Data.Xml.Dom.XmlDocumentType"
Private Const IID_IXmlDocumentType As String = "{f7342425-9781-4964-8e94-9b1c6dfc9bc7}"
Private Const IID_IXmlNode As String = "{1c741d59-2122-47d5-a856-83f3d4214875}"
Private Const IID_IXmlNodeSelector As String = "{63dbba8b-d0db-4fe1-b745-f9433afdc25b}"
Private Const IID_IXmlNodeSerializer As String = "{5cc5b382-e6dd-4991-abef-06d8d2e7bd0c}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IXmlDocumentType
    IXmlDocumentType_GetName = 6
    IXmlDocumentType_GetEntities = 7
    IXmlDocumentType_GetNotations = 8
                            
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
Private m_pIXmlDocumentType As Long
Private m_pIXmlNode As Long
Private m_pIXmlNodeSelector As Long
Private m_pIXmlNodeSerializer As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIXmlNode)
    Call ReleaseIfc(m_pIXmlNodeSelector)
    Call ReleaseIfc(m_pIXmlNodeSerializer)
    Call ReleaseIfc(m_pIXmlDocumentType)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIXmlDocumentType
End Property

Public Property Let Ifc(ByVal pIXmlDocumentType As Long)
    Call ReleaseIfc(m_pIXmlNode)
    Call ReleaseIfc(m_pIXmlNodeSelector)
    Call ReleaseIfc(m_pIXmlNodeSerializer)
    Call ReleaseIfc(m_pIXmlDocumentType)
    m_pIXmlDocumentType = pIXmlDocumentType
    Call QueryIfc(m_pIXmlDocumentType, IID_IXmlNode, m_pIXmlNode)
    Call QueryIfc(m_pIXmlDocumentType, IID_IXmlNodeSelector, m_pIXmlNodeSelector)
    Call QueryIfc(m_pIXmlDocumentType, IID_IXmlNodeSerializer, m_pIXmlNodeSerializer)
End Property

' IXmlDocumentType
Public Property Get Name() As String
    Dim Ret As String
    If m_pIXmlDocumentType <> 0& Then
        Dim Value As Long
        If Invoke(m_pIXmlDocumentType, _
                  IXmlDocumentType_GetName, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    Name = Ret
End Property

Public Property Get Entities() As XmlNamedNodeMap
    Dim Ret As XmlNamedNodeMap
    If m_pIXmlDocumentType <> 0& Then
        Dim pIXmlNamedNodeMap As Long
        If Invoke(m_pIXmlDocumentType, _
                  IXmlDocumentType_GetEntities, _
                  VarPtr(pIXmlNamedNodeMap)) = S_OK Then
            If pIXmlNamedNodeMap <> 0& Then
                Set Ret = New XmlNamedNodeMap
                Ret.Ifc = pIXmlNamedNodeMap
            End If
        End If
    End If
    Set Entities = Ret
End Property

Public Property Get Notations() As XmlNamedNodeMap
    Dim Ret As XmlNamedNodeMap
    If m_pIXmlDocumentType <> 0& Then
        Dim pIXmlNamedNodeMap As Long
        If Invoke(m_pIXmlDocumentType, _
                  IXmlDocumentType_GetNotations, _
                  VarPtr(pIXmlNamedNodeMap)) = S_OK Then
            If pIXmlNamedNodeMap <> 0& Then
                Set Ret = New XmlNamedNodeMap
                Ret.Ifc = pIXmlNamedNodeMap
            End If
        End If
    End If
    Set Notations = Ret
End Property

' IXmlNode
Public Property Get NodeValue() As Inspectable
    Dim Ret As Inspectable
    If m_pIXmlNode <> 0& Then
        Dim Value As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_GetNodeValue, _
                  VarPtr(Value)) = S_OK Then
            If Value <> 0& Then
                Set Ret = New Inspectable
                Ret.Ifc = Value
            End If
        End If
    End If
    Set NodeValue = Ret
End Property

Public Property Let NodeValue(ByVal Value As Inspectable)
    If m_pIXmlNode <> 0& And _
       IsNotNothing(Value) Then
        If Invoke(m_pIXmlNode, _
                  IXmlNode_PutNodeValue, _
                  Value.Ifc) = S_OK Then
        End If
    End If
End Property

Public Property Get NodeType() As NodeType
    Dim Ret As NodeType
    If m_pIXmlNode <> 0& Then
        Dim Value As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_GetNodeType, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    NodeType = Ret
End Property

Public Property Get NodeName() As String
    Dim Ret As String
    If m_pIXmlNode <> 0& Then
        Dim Value As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_GetNodeName, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
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
        Dim Value As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_GetNamespaceUri, _
                  VarPtr(Value)) = S_OK Then
            If Value <> 0& Then
                Set Ret = New Inspectable
                Ret.Ifc = Value
            End If
        End If
    End If
    Set namespaceUri = Ret
End Property

Public Property Get localName() As Inspectable
    Dim Ret As Inspectable
    If m_pIXmlNode <> 0& Then
        Dim Value As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_GetLocalName, _
                  VarPtr(Value)) = S_OK Then
            If Value <> 0& Then
                Set Ret = New Inspectable
                Ret.Ifc = Value
            End If
        End If
    End If
    Set localName = Ret
End Property

Public Property Get Prefix() As Inspectable
    Dim Ret As Inspectable
    If m_pIXmlNode <> 0& Then
        Dim Value As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_GetPrefix, _
                  VarPtr(Value)) = S_OK Then
            If Value <> 0& Then
                Set Ret = New Inspectable
                Ret.Ifc = Value
            End If
        End If
    End If
    Set Prefix = Ret
End Property

Public Property Let Prefix(ByVal Value As Inspectable)
    If m_pIXmlNode <> 0& And _
       IsNotNothing(Value) Then
        If Invoke(m_pIXmlNode, _
                  IXmlNode_PutPrefix, _
                  Value.Ifc) = S_OK Then
        End If
    End If
End Property

' IXmlNodeSerializer
Public Property Get InnerText() As String
    Dim Ret As String
    If m_pIXmlNodeSerializer <> 0& Then
        Dim Value As Long
        If Invoke(m_pIXmlNodeSerializer, _
                  IXmlNodeSerializer_GetInnerText, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    InnerText = Ret
End Property

Public Property Let InnerText(ByVal Value As String)
    If m_pIXmlNodeSerializer <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Value)
        If Invoke(m_pIXmlNodeSerializer, _
                  IXmlNodeSerializer_PutInnerText, _
                  hString) = S_OK Then
        End If
        Call DeleteWindowsString(hString)
    End If
End Property

' ----==== Functions ====----
' IXmlNode
Public Function HasChildNodes() As Boolean
    Dim Ret As Boolean
    If m_pIXmlNode <> 0& Then
        Dim Value As Long
        If Invoke(m_pIXmlNode, _
                  IXmlNode_HasChildNodes, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
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
        Dim Value As Long
        If Invoke(m_pIXmlNodeSerializer, _
                  IXmlNodeSerializer_GetXml, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
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


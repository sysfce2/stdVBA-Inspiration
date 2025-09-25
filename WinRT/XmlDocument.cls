VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XmlDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 07/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Data_Xml_Dom_XmlDocument As String = "Windows.Data.Xml.Dom.XmlDocument"
Private Const IID_IXmlNode As String = "{1c741d59-2122-47d5-a856-83f3d4214875}"
Private Const IID_IXmlDocument As String = "{f7f3a506-1e87-42d6-bcfb-b8c809fa5494}"
Private Const IID_IXmlDocumentIO As String = "{6cd0e74e-ee65-4489-9ebf-ca43e87ba637}"
Private Const IID_IXmlDocumentIO2 As String = "{5d034661-7bd8-4ad5-9ebf-81e6347263b1}"
Private Const IID_IXmlNodeSelector As String = "{63dbba8b-d0db-4fe1-b745-f9433afdc25b}"
Private Const IID_IXmlNodeSerializer As String = "{5cc5b382-e6dd-4991-abef-06d8d2e7bd0c}"
Private Const IID_IXmlDocumentStatics As String = "{5543d254-d757-4b79-9539-232b18f50bf1}"

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IXmlDocumentStatics
    IXmlDocumentStatics_LoadFromUriAsync = 6
    IXmlDocumentStatics_LoadFromUriWithSettingsAsync = 7
    IXmlDocumentStatics_LoadFromFileAsync = 8
    IXmlDocumentStatics_LoadFromFileWithSettingsAsync = 9
    
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
    
    ' IXmlDocument
    IXmlDocument_GetDoctype = 6
    IXmlDocument_GetImplementation = 7
    IXmlDocument_GetDocumentElement = 8
    IXmlDocument_CreateElement = 9
    IXmlDocument_CreateDocumentFragment = 10
    IXmlDocument_CreateTextNode = 11
    IXmlDocument_CreateComment = 12
    IXmlDocument_CreateProcessingInstruction = 13
    IXmlDocument_CreateAttribute = 14
    IXmlDocument_CreateEntityReference = 15
    IXmlDocument_GetElementsByTagName = 16
    IXmlDocument_CreateCDataSection = 17
    IXmlDocument_GetDocumentUri = 18
    IXmlDocument_CreateAttributeNS = 19
    IXmlDocument_CreateElementNS = 20
    IXmlDocument_GetElementById = 21
    IXmlDocument_ImportNode = 22
    
    ' IXmlDocumentIO
    IXmlDocumentIO_LoadXml = 6
    IXmlDocumentIO_LoadXmlWithSettings = 7
    IXmlDocumentIO_SaveToFileAsync = 8
                            
    ' IXmlDocumentIO2
    IXmlDocumentIO2_LoadXmlFromBuffer = 6
    IXmlDocumentIO2_LoadXmlFromBufferWithSettings = 7
                            
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
Private m_pIXmlNode As Long
Private m_pIXmlDocument As Long
Private m_pIXmlDocumentIO As Long
Private m_pIXmlDocumentIO2 As Long
Private m_pIXmlNodeSelector As Long
Private m_pIXmlNodeSerializer As Long
Private m_pIXmlDocumentStatics As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Data_Xml_Dom_XmlDocument, _
                            IID_IXmlDocumentStatics, _
                            m_pIXmlDocumentStatics) Then
    End If
    
    Dim pIXmlDocument As Long
    If GetActivateInstance(Windows_Data_Xml_Dom_XmlDocument, _
                           IID_IXmlDocument, _
                           pIXmlDocument) Then
        Me.Ifc = pIXmlDocument
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIXmlNode)
    Call ReleaseIfc(m_pIXmlDocument)
    Call ReleaseIfc(m_pIXmlDocumentIO)
    Call ReleaseIfc(m_pIXmlDocumentIO2)
    Call ReleaseIfc(m_pIXmlNodeSelector)
    Call ReleaseIfc(m_pIXmlNodeSerializer)
    Call ReleaseIfc(m_pIXmlDocumentStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIXmlDocument
End Property

Public Property Let Ifc(ByVal pIXmlDocument As Long)
    Call ReleaseIfc(m_pIXmlNode)
    Call ReleaseIfc(m_pIXmlDocument)
    Call ReleaseIfc(m_pIXmlDocumentIO)
    Call ReleaseIfc(m_pIXmlDocumentIO2)
    Call ReleaseIfc(m_pIXmlNodeSelector)
    Call ReleaseIfc(m_pIXmlNodeSerializer)
    m_pIXmlDocument = pIXmlDocument
    Call QueryIfc(m_pIXmlDocument, IID_IXmlNode, m_pIXmlNode)
    Call QueryIfc(m_pIXmlDocument, IID_IXmlDocumentIO, m_pIXmlDocumentIO)
    Call QueryIfc(m_pIXmlDocument, IID_IXmlDocumentIO2, m_pIXmlDocumentIO2)
    Call QueryIfc(m_pIXmlDocument, IID_IXmlNodeSelector, m_pIXmlNodeSelector)
    Call QueryIfc(m_pIXmlDocument, IID_IXmlNodeSerializer, m_pIXmlNodeSerializer)
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

' IXmlDocument
Public Property Get Doctype() As XmlDocumentType
    Dim Ret As XmlDocumentType
    If m_pIXmlDocument <> 0& Then
        Dim value As Long
        If Invoke(m_pIXmlDocument, _
                  IXmlDocument_GetDoctype, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                Set Ret = New XmlDocumentType
                Ret.Ifc = value
            End If
        End If
    End If
    Set Doctype = Ret
End Property

Public Property Get Implementation() As XmlDomImplementation
    Dim Ret As XmlDomImplementation
    If m_pIXmlDocument <> 0& Then
        Dim value As Long
        If Invoke(m_pIXmlDocument, _
                  IXmlDocument_GetImplementation, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                Set Ret = New XmlDomImplementation
                Ret.Ifc = value
            End If
        End If
    End If
    Set Implementation = Ret
End Property

Public Property Get DocumentElement() As XmlElement
    Dim Ret As XmlDocument
    If m_pIXmlDocument <> 0& Then
        Dim value As Long
        If Invoke(m_pIXmlDocument, _
                  IXmlDocument_GetDocumentElement, _
                  VarPtr(value)) = S_OK Then
            If value <> 0& Then
                Set Ret = New XmlDocument
                Ret.Ifc = value
            End If
        End If
    End If
    Set DocumentElement = Ret
End Property

Public Property Get DocumentUri() As String
    Dim Ret As String
    If m_pIXmlDocument <> 0& Then
        Dim value As Long
        If Invoke(m_pIXmlDocument, _
                  IXmlDocument_GetDocumentUri, _
                  VarPtr(value)) = S_OK Then
            Ret = GetWindowsString(value)
        End If
    End If
    DocumentUri = Ret
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
' IXmlDocumentStatics
Public Function LoadFromUriAsync(ByVal Uri As Uri) As XmlDocument
    Dim Ret As XmlDocument
    If m_pIXmlDocumentStatics <> 0& And _
       IsNotNothing(Uri) Then
        Dim pIXmlDocument As Long
        If Invoke(m_pIXmlDocumentStatics, _
                  IXmlDocumentStatics_LoadFromUriAsync, _
                  Uri.Ifc, _
                  VarPtr(pIXmlDocument)) = S_OK Then
            If pIXmlDocument <> 0& Then
                If Await(pIXmlDocument) Then
                    Set Ret = New XmlDocument
                    Ret.Ifc = pIXmlDocument
                End If
            End If
        End If
    End If
    Set LoadFromUriAsync = Ret
End Function

Public Function LoadFromUriWithSettingsAsync(ByVal Uri As Uri, _
                                             ByVal loadSettings As XmlLoadSettings) As XmlDocument
    Dim Ret As XmlDocument
    If m_pIXmlDocumentStatics <> 0& And _
       IsNotNothing(Uri) And _
       IsNotNothing(loadSettings) Then
        Dim pIXmlDocument As Long
        If Invoke(m_pIXmlDocumentStatics, _
                  IXmlDocumentStatics_LoadFromUriWithSettingsAsync, _
                  Uri.Ifc, _
                  loadSettings.Ifc, _
                  VarPtr(pIXmlDocument)) = S_OK Then
            If pIXmlDocument <> 0& Then
                If Await(pIXmlDocument) Then
                    Set Ret = New XmlDocument
                    Ret.Ifc = pIXmlDocument
                End If
            End If
        End If
    End If
    Set LoadFromUriWithSettingsAsync = Ret
End Function

Public Function LoadFromFileAsync(ByVal file As StorageFile) As XmlDocument
    Dim Ret As XmlDocument
    If m_pIXmlDocumentStatics <> 0& And _
       IsNotNothing(file) Then
        Dim pIXmlDocument As Long
        If Invoke(m_pIXmlDocumentStatics, _
                  IXmlDocumentStatics_LoadFromFileAsync, _
                  file.Ifc, _
                  VarPtr(pIXmlDocument)) = S_OK Then
            If pIXmlDocument <> 0& Then
                If Await(pIXmlDocument) Then
                    Set Ret = New XmlDocument
                    Ret.Ifc = pIXmlDocument
                End If
            End If
        End If
    End If
    Set LoadFromFileAsync = Ret
End Function

Public Function LoadFromFileWithSettingsAsync(ByVal file As StorageFile, _
                                              ByVal loadSettings As XmlLoadSettings) As XmlDocument
    Dim Ret As XmlDocument
    If m_pIXmlDocumentStatics <> 0& And _
       IsNotNothing(file) And _
       IsNotNothing(loadSettings) Then
        Dim pIXmlDocument As Long
        If Invoke(m_pIXmlDocumentStatics, _
                  IXmlDocumentStatics_LoadFromFileWithSettingsAsync, _
                  file.Ifc, _
                  loadSettings.Ifc, _
                  VarPtr(pIXmlDocument)) = S_OK Then
            If pIXmlDocument <> 0& Then
                If Await(pIXmlDocument) Then
                    Set Ret = New XmlDocument
                    Ret.Ifc = pIXmlDocument
                End If
            End If
        End If
    End If
    Set LoadFromFileWithSettingsAsync = Ret
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

' IXmlDocument
Public Function CreateElement(ByVal tagName As String) As XmlElement
    Dim Ret As XmlElement
    If m_pIXmlDocument <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(tagName)
        Dim pIXmlElement As Long
        If Invoke(m_pIXmlDocument, _
                  IXmlDocument_CreateElement, _
                  hString, _
                  VarPtr(pIXmlElement)) = S_OK Then
            If pIXmlElement <> 0& Then
                Set Ret = New XmlElement
                Ret.Ifc = pIXmlElement
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set CreateElement = Ret
End Function

Public Function CreateDocumentFragment() As XmlDocumentFragment
    Dim Ret As XmlDocumentFragment
    If m_pIXmlDocument <> 0& Then
        Dim pIXmlDocumentFragment As Long
        If Invoke(m_pIXmlDocument, _
                  IXmlDocument_CreateDocumentFragment, _
                  VarPtr(pIXmlDocumentFragment)) = S_OK Then
            If pIXmlDocumentFragment <> 0& Then
                Set Ret = New XmlDocumentFragment
                Ret.Ifc = pIXmlDocumentFragment
            End If
        End If
    End If
    Set CreateDocumentFragment = Ret
End Function

Public Function CreateTextNode(ByVal data As String) As XmlText
    Dim Ret As XmlText
    If m_pIXmlDocument <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(data)
        Dim pIXmlText As Long
        If Invoke(m_pIXmlDocument, _
                  IXmlDocument_CreateTextNode, _
                  hString, _
                  VarPtr(pIXmlText)) = S_OK Then
            If pIXmlText <> 0& Then
                Set Ret = New XmlText
                Ret.Ifc = pIXmlText
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set CreateTextNode = Ret
End Function

Public Function CreateComment(ByVal data As String) As XmlComment
    Dim Ret As XmlComment
    If m_pIXmlDocument <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(data)
        Dim pIXmlComment As Long
        If Invoke(m_pIXmlDocument, _
                  IXmlDocument_CreateComment, _
                  hString, _
                  VarPtr(pIXmlComment)) = S_OK Then
            If pIXmlComment <> 0& Then
                Set Ret = New XmlComment
                Ret.Ifc = pIXmlComment
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set CreateComment = Ret
End Function

Public Function CreateProcessingInstruction(ByVal target As String, _
                                            ByVal data As String) As XmlProcessingInstruction
    Dim Ret As XmlProcessingInstruction
    If m_pIXmlDocument <> 0& Then
        Dim hTarget As Long
        Dim hData As Long
        hTarget = CreateWindowsString(target)
        hData = CreateWindowsString(data)
        Dim pIXmlProcessingInstruction As Long
        If Invoke(m_pIXmlDocument, _
                  IXmlDocument_CreateProcessingInstruction, _
                  hTarget, _
                  hData, _
                  VarPtr(pIXmlProcessingInstruction)) = S_OK Then
            If pIXmlProcessingInstruction <> 0& Then
                Set Ret = New XmlProcessingInstruction
                Ret.Ifc = pIXmlProcessingInstruction
            End If
        End If
        Call DeleteWindowsString(hData)
        Call DeleteWindowsString(hTarget)
    End If
    Set CreateProcessingInstruction = Ret
End Function

Public Function CreateAttribute(ByVal Name As String) As XmlAttribute
    Dim Ret As XmlAttribute
    If m_pIXmlDocument <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Name)
        Dim pIXmlAttribute As Long
        If Invoke(m_pIXmlDocument, _
                  IXmlDocument_CreateAttribute, _
                  hString, _
                  VarPtr(pIXmlAttribute)) = S_OK Then
            If pIXmlAttribute <> 0& Then
                Set Ret = New XmlAttribute
                Ret.Ifc = pIXmlAttribute
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set CreateAttribute = Ret
End Function

Public Function CreateEntityReference(ByVal Name As String) As XmlEntityReference
    Dim Ret As XmlEntityReference
    If m_pIXmlDocument <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Name)
        Dim pIXmlEntityReference As Long
        If Invoke(m_pIXmlDocument, _
                  IXmlDocument_CreateEntityReference, _
                  hString, _
                  VarPtr(pIXmlEntityReference)) = S_OK Then
            If pIXmlEntityReference <> 0& Then
                Set Ret = New XmlEntityReference
                Ret.Ifc = pIXmlEntityReference
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set CreateEntityReference = Ret
End Function

Public Function GetElementsByTagName(ByVal tagName As String) As XmlNodeList
    Dim Ret As XmlNodeList
    If m_pIXmlDocument <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(tagName)
        Dim pIXmlNodeList As Long
        If Invoke(m_pIXmlDocument, _
                  IXmlDocument_GetElementsByTagName, _
                  hString, _
                  VarPtr(pIXmlNodeList)) = S_OK Then
            If pIXmlNodeList <> 0& Then
                Set Ret = New XmlNodeList
                Ret.Ifc = pIXmlNodeList
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set GetElementsByTagName = Ret
End Function

Public Function CreateCDataSection(ByVal data As String) As XmlCDataSection
    Dim Ret As XmlCDataSection
    If m_pIXmlDocument <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(data)
        Dim pIXmlCDataSection As Long
        If Invoke(m_pIXmlDocument, _
                  IXmlDocument_CreateCDataSection, _
                  hString, _
                  VarPtr(pIXmlCDataSection)) = S_OK Then
            If pIXmlCDataSection <> 0& Then
                Set Ret = New XmlCDataSection
                Ret.Ifc = pIXmlCDataSection
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set CreateCDataSection = Ret
End Function

Public Function CreateAttributeNS(ByVal namespaceUri As Inspectable, _
                                  ByVal qualifiedName As String) As XmlAttribute
    Dim Ret As XmlAttribute
    If m_pIXmlDocument <> 0& And _
       IsNotNothing(namespaceUri) Then
        Dim hString As Long
        hString = CreateWindowsString(qualifiedName)
        Dim pIXmlAttribute As Long
        If Invoke(m_pIXmlDocument, _
                  IXmlDocument_CreateAttributeNS, _
                  namespaceUri.Ifc, _
                  hString, _
                  VarPtr(pIXmlAttribute)) = S_OK Then
            If pIXmlAttribute <> 0& Then
                Set Ret = New XmlAttribute
                Ret.Ifc = pIXmlAttribute
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set CreateAttributeNS = Ret
End Function

Public Function CreateElementNS(ByVal namespaceUri As Inspectable, _
                                ByVal qualifiedName As String) As XmlElement
    Dim Ret As XmlElement
    If m_pIXmlDocument <> 0& And _
       IsNotNothing(namespaceUri) Then
        Dim hString As Long
        hString = CreateWindowsString(qualifiedName)
        Dim pIXmlElement As Long
        If Invoke(m_pIXmlDocument, _
                  IXmlDocument_CreateElementNS, _
                  namespaceUri.Ifc, _
                  hString, _
                  VarPtr(pIXmlElement)) = S_OK Then
            If pIXmlElement <> 0& Then
                Set Ret = New XmlElement
                Ret.Ifc = pIXmlElement
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set CreateElementNS = Ret
End Function

Public Function GetElementById(ByVal elementId As String) As XmlElement
    Dim Ret As XmlElement
    If m_pIXmlDocument <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(elementId)
        Dim pIXmlElement As Long
        If Invoke(m_pIXmlDocument, _
                  IXmlDocument_GetElementById, _
                  hString, _
                  VarPtr(pIXmlElement)) = S_OK Then
            If pIXmlElement <> 0& Then
                Set Ret = New XmlElement
                Ret.Ifc = pIXmlElement
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set GetElementById = Ret
End Function

Public Function ImportNode(ByVal node As XmlNode, _
                           ByVal deep As Boolean) As XmlNode
    Dim Ret As XmlNode
    If m_pIXmlDocument <> 0& And _
       IsNotNothing(node) Then
        Dim pIXmlNode As Long
        If Invoke(m_pIXmlDocument, _
                  IXmlDocument_ImportNode, _
                  node.Ifc, _
                  deep, _
                  VarPtr(pIXmlNode)) = S_OK Then
            If pIXmlNode <> 0& Then
                Set Ret = New XmlNode
                Ret.Ifc = pIXmlNode
            End If
        End If
    End If
    Set ImportNode = Ret
End Function

' IXmlDocumentIO
Public Function LoadXML(ByVal Xml As String) As Boolean
    Dim Ret As Boolean
    If m_pIXmlDocumentIO <> 0& Then
        Dim hString As Long
        hString = CreateWindowsString(Xml)
        If Invoke(m_pIXmlDocumentIO, _
                  IXmlDocumentIO_LoadXml, _
                  hString) = S_OK Then
            Ret = True
        End If
        Call DeleteWindowsString(hString)
    End If
    LoadXML = Ret
End Function

Public Function LoadXmlWithSettings(ByVal Xml As String, _
                                    ByVal loadSettings As XmlLoadSettings) As Boolean
    Dim Ret As Boolean
    If m_pIXmlDocumentIO <> 0& And _
       IsNotNothing(loadSettings) Then
        Dim hString As Long
        hString = CreateWindowsString(Xml)
        If Invoke(m_pIXmlDocumentIO, _
                  IXmlDocumentIO_LoadXmlWithSettings, _
                  hString, _
                  loadSettings.Ifc) = S_OK Then
            Ret = True
        End If
        Call DeleteWindowsString(hString)
    End If
    LoadXmlWithSettings = Ret
End Function

Public Function SaveToFileAsync(ByVal file As StorageFile) As Boolean
    Dim Ret As Boolean
    If m_pIXmlDocumentIO <> 0& And _
       IsNotNothing(file) Then
        Dim pIAsyncAction As Long
        If Invoke(m_pIXmlDocumentIO, _
                  IXmlDocumentIO_SaveToFileAsync, _
                  file.Ifc, _
                  VarPtr(pIAsyncAction)) = S_OK Then
            If pIAsyncAction <> 0& Then
                If Await(pIAsyncAction, False) Then
                    Ret = True
                End If
            End If
        End If
    End If
    SaveToFileAsync = Ret
End Function
                            
' IXmlDocumentIO2
Public Function LoadXmlFromBuffer(ByVal Buffer As Buffer) As Boolean
    Dim Ret As Boolean
    If m_pIXmlDocumentIO2 <> 0& And _
       IsNotNothing(Buffer) Then
        If Invoke(m_pIXmlDocumentIO2, _
                  IXmlDocumentIO2_LoadXmlFromBuffer, _
                  Buffer.Ifc) = S_OK Then
            Ret = True
        End If
    End If
    LoadXmlFromBuffer = Ret
End Function

Public Function LoadXmlFromBufferWithSettings(ByVal Buffer As Buffer, _
                                              ByVal loadSettings As XmlLoadSettings) As Boolean
    Dim Ret As Boolean
    If m_pIXmlDocumentIO2 <> 0& And _
       IsNotNothing(Buffer) And _
       IsNotNothing(loadSettings) Then
        If Invoke(m_pIXmlDocumentIO2, _
                  IXmlDocumentIO2_LoadXmlFromBufferWithSettings, _
                  Buffer.Ifc, _
                  loadSettings.Ifc) = S_OK Then
            Ret = True
        End If
    End If
    LoadXmlFromBufferWithSettings = Ret
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


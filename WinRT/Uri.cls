VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Uri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Foundation_Uri As String = "Windows.Foundation.Uri"
Private Const IID_IStringable As String = "{96369f54-8eb6-48f0-abce-c1b211e627c3}"
Private Const IID_IUriRuntimeClass As String = "{9e365e57-48b2-4160-956f-c7385120bbfc}"
Private Const IID_IUriEscapeStatics As String = "{c1d432ba-c824-4452-a7fd-512bc3bbe9a1}"
Private Const IID_IUriRuntimeClassFactory As String = "{44a9796f-723e-4fdf-a218-033e75b0c084}"
Private Const IID_IUriRuntimeClassWithAbsoluteCanonicalUri As String = "{758d9661-221c-480f-a339-50656673f46f}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IUriRuntimeClassFactory
    IUriRuntimeClassFactory_CreateUri = 6
    IUriRuntimeClassFactory_CreateWithRelativeUri = 7
    
    ' IUriEscapeStatics
    IUriEscapeStatics_UnescapeComponent = 6
    IUriEscapeStatics_EscapeComponent = 7
    
    ' IUriRuntimeClass
    IUriRuntimeClass_GetAbsoluteUri = 6
    IUriRuntimeClass_GetDisplayUri = 7
    IUriRuntimeClass_GetDomain = 8
    IUriRuntimeClass_GetExtension = 9
    IUriRuntimeClass_GetFragment = 10
    IUriRuntimeClass_GetHost = 11
    IUriRuntimeClass_GetPassword = 12
    IUriRuntimeClass_GetPath = 13
    IUriRuntimeClass_GetQuery = 14
    'IUriRuntimeClass_GetQueryParsed = 15
    IUriRuntimeClass_GetRawUri = 16
    IUriRuntimeClass_GetSchemeName = 17
    IUriRuntimeClass_GetUserName = 18
    IUriRuntimeClass_GetPort = 19
    IUriRuntimeClass_GetSuspicious = 20
    IUriRuntimeClass_Equals = 21
    IUriRuntimeClass_CombineUri = 22
                    
    ' IUriRuntimeClassWithAbsoluteCanonicalUri
    IUriRuntimeClassWithAbsoluteCanonicalUri_GetAbsoluteCanonicalUri = 6
    IUriRuntimeClassWithAbsoluteCanonicalUri_GetDisplayIri = 7
    
    ' IStringable
    IStringable_ToString = 6
                    
End Enum

' ----==== Vars ====----
Private m_pIStringable As Long
Private m_pIUriRuntimeClass As Long
Private m_pIUriEscapeStatics As Long
Private m_pIUriRuntimeClassFactory As Long
Private m_pIUriRuntimeClassWithAbsoluteCanonicalUri As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_Foundation_Uri, _
                            IID_IUriRuntimeClassFactory, _
                            m_pIUriRuntimeClassFactory) Then
        Call QueryIfc(m_pIUriRuntimeClassFactory, IID_IUriEscapeStatics, m_pIUriEscapeStatics)
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIStringable)
    Call ReleaseIfc(m_pIUriRuntimeClass)
    Call ReleaseIfc(m_pIUriRuntimeClassWithAbsoluteCanonicalUri)
    Call ReleaseIfc(m_pIUriEscapeStatics)
    Call ReleaseIfc(m_pIUriRuntimeClassFactory)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIUriRuntimeClass
End Property

Public Property Let Ifc(ByVal pIUriRuntimeClass As Long)
    Call ReleaseIfc(m_pIStringable)
    Call ReleaseIfc(m_pIUriRuntimeClass)
    Call ReleaseIfc(m_pIUriRuntimeClassWithAbsoluteCanonicalUri)
    m_pIUriRuntimeClass = pIUriRuntimeClass
    If m_pIUriRuntimeClass <> 0& Then
        Call QueryIfc(m_pIUriRuntimeClass, _
                      IID_IStringable, _
                      m_pIStringable)
        Call QueryIfc(m_pIUriRuntimeClass, _
                      IID_IUriRuntimeClassWithAbsoluteCanonicalUri, _
                      m_pIUriRuntimeClassWithAbsoluteCanonicalUri)
    End If
End Property

' IUriRuntimeClass
Public Property Get AbsoluteUri() As String
    Dim Ret As String
    If m_pIUriRuntimeClass <> 0& Then
        Dim Value As Long
        If Invoke(m_pIUriRuntimeClass, _
                  IUriRuntimeClass_GetAbsoluteUri, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    AbsoluteUri = Ret
End Property

Public Property Get DisplayUri() As String
    Dim Ret As String
    If m_pIUriRuntimeClass <> 0& Then
        Dim Value As Long
        If Invoke(m_pIUriRuntimeClass, _
                  IUriRuntimeClass_GetDisplayUri, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    DisplayUri = Ret
End Property

Public Property Get Domain() As String
    Dim Ret As String
    If m_pIUriRuntimeClass <> 0& Then
        Dim Value As Long
        If Invoke(m_pIUriRuntimeClass, _
                  IUriRuntimeClass_GetDomain, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    Domain = Ret
End Property

Public Property Get Extension() As String
    Dim Ret As String
    If m_pIUriRuntimeClass <> 0& Then
        Dim Value As Long
        If Invoke(m_pIUriRuntimeClass, _
                  IUriRuntimeClass_GetExtension, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    Extension = Ret
End Property

Public Property Get Fragment() As String
    Dim Ret As String
    If m_pIUriRuntimeClass <> 0& Then
        Dim Value As Long
        If Invoke(m_pIUriRuntimeClass, _
                  IUriRuntimeClass_GetFragment, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    Fragment = Ret
End Property

Public Property Get Host() As String
    Dim Ret As String
    If m_pIUriRuntimeClass <> 0& Then
        Dim Value As Long
        If Invoke(m_pIUriRuntimeClass, _
                  IUriRuntimeClass_GetHost, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    Host = Ret
End Property

Public Property Get password() As String
    Dim Ret As String
    If m_pIUriRuntimeClass <> 0& Then
        Dim Value As Long
        If Invoke(m_pIUriRuntimeClass, _
                  IUriRuntimeClass_GetPassword, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    password = Ret
End Property

Public Property Get Path() As String
    Dim Ret As String
    If m_pIUriRuntimeClass <> 0& Then
        Dim Value As Long
        If Invoke(m_pIUriRuntimeClass, _
                  IUriRuntimeClass_GetPath, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    Path = Ret
End Property

Public Property Get Query() As String
    Dim Ret As String
    If m_pIUriRuntimeClass <> 0& Then
        Dim Value As Long
        If Invoke(m_pIUriRuntimeClass, _
                  IUriRuntimeClass_GetQuery, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    Query = Ret
End Property

' IUriRuntimeClass_GetQueryParsed = 15

Public Property Get RawUri() As String
    Dim Ret As String
    If m_pIUriRuntimeClass <> 0& Then
        Dim Value As Long
        If Invoke(m_pIUriRuntimeClass, _
                  IUriRuntimeClass_GetRawUri, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    RawUri = Ret
End Property

Public Property Get SchemeName() As String
    Dim Ret As String
    If m_pIUriRuntimeClass <> 0& Then
        Dim Value As Long
        If Invoke(m_pIUriRuntimeClass, _
                  IUriRuntimeClass_GetSchemeName, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    SchemeName = Ret
End Property

Public Property Get UserName() As String
    Dim Ret As String
    If m_pIUriRuntimeClass <> 0& Then
        Dim Value As Long
        If Invoke(m_pIUriRuntimeClass, _
                  IUriRuntimeClass_GetUserName, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    UserName = Ret
End Property

Public Property Get Port() As Long
    Dim Ret As Long
    If m_pIUriRuntimeClass <> 0& Then
        Dim Value As Long
        If Invoke(m_pIUriRuntimeClass, _
                  IUriRuntimeClass_GetPort, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Port = Ret
End Property

Public Property Get Suspicious() As Boolean
    Dim Ret As Boolean
    If m_pIUriRuntimeClass <> 0& Then
        Dim Value As Long
        If Invoke(m_pIUriRuntimeClass, _
                  IUriRuntimeClass_GetSuspicious, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    Suspicious = Ret
End Property

' IUriRuntimeClassWithAbsoluteCanonicalUri
Public Property Get AbsoluteCanonicalUri() As String
    Dim Ret As String
    If m_pIUriRuntimeClassWithAbsoluteCanonicalUri <> 0& Then
        Dim Value As Long
        If Invoke(m_pIUriRuntimeClassWithAbsoluteCanonicalUri, _
                  IUriRuntimeClassWithAbsoluteCanonicalUri_GetAbsoluteCanonicalUri, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    AbsoluteCanonicalUri = Ret
End Property

Public Property Get DisplayIri() As String
    Dim Ret As String
    If m_pIUriRuntimeClassWithAbsoluteCanonicalUri <> 0& Then
        Dim Value As Long
        If Invoke(m_pIUriRuntimeClassWithAbsoluteCanonicalUri, _
                  IUriRuntimeClassWithAbsoluteCanonicalUri_GetDisplayIri, _
                  VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
    End If
    DisplayIri = Ret
End Property

' ----==== Functions ====----
' IUriRuntimeClassFactory
Public Function CreateUri(ByVal Uri As String) As Uri
    Dim Ret As Uri
    If m_pIUriRuntimeClassFactory <> 0& Then
        Dim hUri As Long
        Dim pIUriRuntimeClass As Long
        hUri = CreateWindowsString(Uri)
        If Invoke(m_pIUriRuntimeClassFactory, _
                  IUriRuntimeClassFactory_CreateUri, _
                  hUri, VarPtr(pIUriRuntimeClass)) = S_OK Then
            If pIUriRuntimeClass <> 0 Then
                Set Ret = New Uri
                Ret.Ifc = pIUriRuntimeClass
            End If
        End If
        Call DeleteWindowsString(hUri)
    End If
    Set CreateUri = Ret
End Function

Public Function CreateWithRelativeUri(ByVal baseUri As String, _
                                      ByVal relativeUri As String) As Uri
    Dim Ret As Uri
    If m_pIUriRuntimeClassFactory <> 0& Then
        Dim hBaseUri As Long
        Dim hRelativeUri As Long
        Dim pIUriRuntimeClass As Long
        hBaseUri = CreateWindowsString(baseUri)
        hRelativeUri = CreateWindowsString(relativeUri)
        If Invoke(m_pIUriRuntimeClassFactory, _
                  IUriRuntimeClassFactory_CreateWithRelativeUri, _
                  hBaseUri, _
                  hRelativeUri, _
                  VarPtr(pIUriRuntimeClass)) = S_OK Then
            If pIUriRuntimeClass <> 0 Then
                Set Ret = New Uri
                Ret.Ifc = pIUriRuntimeClass
            End If
        End If
        Call DeleteWindowsString(hRelativeUri)
        Call DeleteWindowsString(hBaseUri)
    End If
    Set CreateWithRelativeUri = Ret
End Function

' IUriEscapeStatics
Public Function UnescapeComponent(ByVal toUnescape As String) As String
    Dim Ret As String
    If m_pIUriEscapeStatics <> 0& Then
        Dim Value As Long
        Dim hString As Long
        hString = CreateWindowsString(toUnescape)
        If Invoke(m_pIUriEscapeStatics, _
                  IUriEscapeStatics_UnescapeComponent, _
                  hString, VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
        Call DeleteWindowsString(hString)
    End If
    UnescapeComponent = Ret
End Function

Public Function EscapeComponent(ByVal toEscape As String) As String
    Dim Ret As String
    If m_pIUriEscapeStatics <> 0& Then
        Dim Value As Long
        Dim hString As Long
        hString = CreateWindowsString(toEscape)
        If Invoke(m_pIUriEscapeStatics, _
                  IUriEscapeStatics_EscapeComponent, _
                  hString, VarPtr(Value)) = S_OK Then
            Ret = GetWindowsString(Value)
        End If
        Call DeleteWindowsString(hString)
    End If
    EscapeComponent = Ret
End Function

' IUriRuntimeClass
Public Function Equals(ByVal Uri As Uri) As Boolean
    Dim Ret As Boolean
    If m_pIUriRuntimeClass <> 0& And _
       IsNotNothing(Uri) Then
        Dim Value As Long
        If Invoke(m_pIUriRuntimeClass, _
                  IUriRuntimeClass_Equals, _
                  Uri.Ifc, _
                  VarPtr(Value)) = S_OK Then
            Ret = CBool(Value)
        End If
    End If
    Equals = Ret
End Function

Public Function CombineUri(ByVal relativeUri As String) As Uri
    Dim Ret As Uri
    If m_pIUriRuntimeClass <> 0& Then
        Dim hString As Long
        Dim pIUriRuntimeClass As Long
        hString = CreateWindowsString(relativeUri)
        If Invoke(m_pIUriRuntimeClass, _
                  IUriRuntimeClass_CombineUri, _
                  hString, _
                  VarPtr(pIUriRuntimeClass)) = S_OK Then
            If pIUriRuntimeClass <> 0& Then
                Set Ret = New Uri
                Ret.Ifc = pIUriRuntimeClass
            End If
        End If
        Call DeleteWindowsString(hString)
    End If
    Set CombineUri = Ret
End Function

' IStringable
Public Function ToString() As String
    Dim Ret As String
    If m_pIStringable <> 0& Then
        Dim hString As Long
        If Invoke(m_pIStringable, _
                  IStringable_ToString, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    ToString = Ret
End Function

' ----==== Invoke ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function


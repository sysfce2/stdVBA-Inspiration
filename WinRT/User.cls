VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 08/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_System_User As String = "Windows.System.User"
Private Const IID_IUser As String = "{df9a26c6-e746-4bcd-b5d4-120103c4209b}"
Private Const IID_IUser2 As String = "{98ba5628-a6e3-518e-89d9-d3b2b1991a10}"
Private Const IID_IUserStatics As String = "{155eb23b-242a-45e0-a2e9-3171fc6a7fdd}"
Private Const IID_IUserStatics2 As String = "{74a37e11-2eb5-4487-b0d5-2c6790e013e9}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IUserStatics
    IUserStatics_CreateWatcher = 6
    IUserStatics_FindAllAsync = 7
'   DEPRECATED ("FindAllAsyncByType is deprecated and might not function consistently on all platforms. Instead, use FindAllAsync or GetDefault.")
'    IUserStatics_FindAllAsyncByType = 8
'   DEPRECATED ("FindAllAsyncByTypeAndStatus is deprecated and might not function consistently on all platforms. Instead, use FindAllAsync or GetDefault.")
'    IUserStatics_FindAllAsyncByTypeAndStatus = 9
    IUserStatics_GetFromId = 10
                    
    ' IUserStatics2
    IUserStatics2_GetDefault = 6
    
    ' IUser
    IUser_GetNonRoamableId = 6
    IUser_GetAuthenticationStatus = 7
    IUser_GetType = 8
    IUser_GetPropertyAsync = 9
'    IUser_GetPropertiesAsync = 10
    IUser_GetPictureAsync = 11
                    
    ' IUser2
    IUser2_CheckUserAgeConsentGroupAsync = 6

End Enum

' ----==== Vars ====----
Private m_pIUser As Long
Private m_pIUser2 As Long
Private m_pIUserStatics As Long
Private m_pIUserStatics2 As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivationFactory(Windows_System_User, _
                            IID_IUserStatics, _
                            m_pIUserStatics) Then
        Call QueryIfc(m_pIUserStatics, IID_IUserStatics2, m_pIUserStatics2)
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIUser2)
    Call ReleaseIfc(m_pIUser)
    Call ReleaseIfc(m_pIUserStatics2)
    Call ReleaseIfc(m_pIUserStatics)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIUser
End Property

Public Property Let Ifc(ByVal pIUser As Long)
    Call ReleaseIfc(m_pIUser2)
    Call ReleaseIfc(m_pIUser)
    m_pIUser = pIUser
    Call QueryIfc(m_pIUser, IID_IUser2, m_pIUser2)
End Property

' IUser
Public Property Get NonRoamableId() As String
    Dim Ret As String
    If m_pIUser <> 0& Then
        Dim hString As Long
        If Invoke(m_pIUser, _
                  IUser_GetNonRoamableId, _
                  VarPtr(hString)) = S_OK Then
            Ret = GetWindowsString(hString)
        End If
    End If
    NonRoamableId = Ret
End Property

Public Property Get AuthenticationStatus() As UserAuthenticationStatus
    Dim Ret As UserAuthenticationStatus
    If m_pIUser <> 0& Then
        Dim value As Long
        If Invoke(m_pIUser, _
                  IUser_GetAuthenticationStatus, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    AuthenticationStatus = Ret
End Property

Public Property Get UserType() As UserType
    Dim Ret As UserType
    If m_pIUser <> 0& Then
        Dim value As Long
        If Invoke(m_pIUser, _
                  IUser_GetType, _
                  VarPtr(value)) = S_OK Then
            Ret = value
        End If
    End If
    UserType = Ret
End Property

' ----==== Functions ====----
' IUserStatics
Public Function CreateWatcher() As UserWatcher
    Dim Ret As UserWatcher
    If m_pIUserStatics <> 0& Then
        Dim pIUserWatcher As Long
        If Invoke(m_pIUserStatics, _
                  IUserStatics_CreateWatcher, _
                  VarPtr(pIUserWatcher)) = S_OK Then
            If pIUserWatcher <> 0& Then
                Set Ret = New UserWatcher
                Ret.Ifc = pIUserWatcher
            End If
        End If
    End If
    Set CreateWatcher = Ret
End Function

Public Function FindAllAsync() As ReadOnlyList_1 'ReadOnlyList_User
    Dim Ret As ReadOnlyList_1 'ReadOnlyList_User
    If m_pIUserStatics <> 0& Then
        Dim pIVectorView As Long
        If Invoke(m_pIUserStatics, _
                  IUserStatics_FindAllAsync, _
                  VarPtr(pIVectorView)) = S_OK Then
            If pIVectorView <> 0& Then
                If Await(pIVectorView) Then
                    Set Ret = New_ReadOnlyList_1(Of_User, pIVectorView) 'ReadOnlyList_User
                End If
            End If
        End If
    End If
    Set FindAllAsync = Ret
End Function

Public Function GetFromId(ByVal NonRoamableId As String) As User
    Dim Ret As User
    If m_pIUserStatics <> 0& Then
        Dim pIUser As Long
        Dim hNonRoamableId As Long
        hNonRoamableId = CreateWindowsString(NonRoamableId)
        If Invoke(m_pIUserStatics, _
                  IUserStatics_GetFromId, _
                  hNonRoamableId, _
                  VarPtr(pIUser)) = S_OK Then
            If pIUser <> 0& Then
                Set Ret = New User
                Ret.Ifc = pIUser
            End If
        End If
        Call DeleteWindowsString(hNonRoamableId)
    End If
    Set GetFromId = Ret
End Function

' IUserStatics2
Public Function GetDefault() As User
    Dim Ret As User
    If m_pIUserStatics2 <> 0& Then
        Dim pIUser As Long
        If Invoke(m_pIUserStatics2, _
                  IUserStatics2_GetDefault, _
                  VarPtr(pIUser)) = S_OK Then
            If pIUser <> 0& Then
                Set Ret = New User
                Ret.Ifc = pIUser
            End If
        End If
    End If
    Set GetDefault = Ret
End Function

' IUser
Public Function GetPropertyAsync(ByVal propValue As String) As PropertyValue
    Dim Ret As PropertyValue
    If m_pIUser <> 0& Then
        Dim hPropValue As Long
        Dim pIPropertyValue As Long
        hPropValue = CreateWindowsString(propValue)
        If Invoke(m_pIUser, _
                  IUser_GetPropertyAsync, _
                  hPropValue, _
                  VarPtr(pIPropertyValue)) = S_OK Then
            If pIPropertyValue <> 0& Then
                If Await(pIPropertyValue) Then
                    Set Ret = New PropertyValue
                    Ret.Ifc = pIPropertyValue
                End If
            End If
        End If
        Call DeleteWindowsString(hPropValue)
    End If
    Set GetPropertyAsync = Ret
End Function

'    IUser_GetPropertiesAsync = 10

Public Function GetPictureAsync(ByVal desiredSize As UserPictureSize) As RandomAccessStreamReference
    Dim Ret As RandomAccessStreamReference
    If m_pIUser <> 0& Then
        Dim pIRandomAccessStreamReference As Long
        If Invoke(m_pIUser, _
                  IUser_GetPictureAsync, _
                  desiredSize, _
                  VarPtr(pIRandomAccessStreamReference)) = S_OK Then
            If pIRandomAccessStreamReference <> 0& Then
                If Await(pIRandomAccessStreamReference) Then
                    Set Ret = New RandomAccessStreamReference
                    Ret.Ifc = pIRandomAccessStreamReference
                End If
            End If
        End If
    End If
    Set GetPictureAsync = Ret
End Function
                    
' IUser2
Public Function CheckUserAgeConsentGroupAsync(ByVal consentGroup As UserAgeConsentGroup) As UserAgeConsentResult
    Dim Ret As UserAgeConsentResult
    If m_pIUser2 <> 0& Then
        Dim operation As Long
        If Invoke(m_pIUser2, _
                  IUser2_CheckUserAgeConsentGroupAsync, _
                  consentGroup, _
                  VarPtr(operation)) = S_OK Then
            If operation <> 0& Then
                If Await(operation) Then
                    Ret = operation
                End If
            End If
        End If
    End If
    CheckUserAgeConsentGroupAsync = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



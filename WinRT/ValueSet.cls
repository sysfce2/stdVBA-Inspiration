VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValueSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 02/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Foundation_Collections_ValueSet As String = "Windows.Foundation.Collections.ValueSet"

Private Const IID_IPropertySet As String = "{8a43ed9f-f4e6-4421-acf9-1dab2986820c}"
Private Const IObservableMap_HSTRING_IInspectable As String = "{236aac9d-fb12-5c4d-a41c-9e445fb4d7ec}"
Private Const IMap_HSTRING_IInspectable As String = "{1b0d3570-0877-5ec2-8a2c-3b9539506aca}"
Private Const IIterable_IKeyValuePair_HSTRING_IInspectable As String = "{fe2f3d47-5d47-5499-8374-430c7cda0204}"
Private Const IIterator_IKeyValuePair_HSTRING_IInspectable As String = "{5db5fa32-707c-5849-a06b-91c8eb9d10e8}"
Private Const IKeyValuePair_String_Object As String = "{09335560-6c6b-5a26-9348-97b781132b20}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' IPropertySet -> has no functions
    
    ' IObservableMap_HSTRING_IInspectable
'    IObservableMap_HSTRING_IInspectable_AddMapChanged = 6
'    IObservableMap_HSTRING_IInspectable_RemoveMapChanged = 7
    
    ' IMap_HSTRING_IInspectable
'    IMap_HSTRING_IInspectable_Lookup = 6
    IMap_HSTRING_IInspectable_GetSize = 7
'    IMap_HSTRING_IInspectable_HasKey = 8
'    IMap_HSTRING_IInspectable_GetView = 9
'    IMap_HSTRING_IInspectable_Insert = 10
'    IMap_HSTRING_IInspectable_Remove = 11
'    IMap_HSTRING_IInspectable_Clear = 12
    
    ' IIterable_IKeyValuePair_HSTRING_IInspectable
    IIterable_IKeyValuePair_HSTRING_IInspectable_First = 6
    
    ' IIterator_IKeyValuePair_HSTRING_IInspectable
    IIterator_IKeyValuePair_HSTRING_IInspectable_GetCurrent = 6
    IIterator_IKeyValuePair_HSTRING_IInspectable_GetHasCurrent = 7
    IIterator_IKeyValuePair_HSTRING_IInspectable_MoveNext = 8
'    IIterator_IKeyValuePair_HSTRING_IInspectable_GetMany = 9
    
    ' IKeyValuePair_String_Object
    IKeyValuePair_String_Object_GetKey = 6
    IKeyValuePair_String_Object_GetValue = 7
    
    ' IReference_xxx
    IReference_GetValue = 6
    
End Enum

' ----==== Variablen ====----
Private m_pIPropertySet As Long
Private m_pIMap_HSTRING_IInspectable As Long
Private m_pIObservableMap_HSTRING_IInspectable As Long
Private m_pIIterable_IKeyValuePair_HSTRING_IInspectable As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    If GetActivateInstance(Windows_Foundation_Collections_ValueSet, _
                           IID_IPropertySet, _
                           m_pIPropertySet) Then
        Call QueryIfc(m_pIPropertySet, _
                      IMap_HSTRING_IInspectable, _
                      m_pIMap_HSTRING_IInspectable)
        
        Call QueryIfc(m_pIPropertySet, _
                      IObservableMap_HSTRING_IInspectable, _
                      m_pIObservableMap_HSTRING_IInspectable)
        
        Call QueryIfc(m_pIPropertySet, _
                      IIterable_IKeyValuePair_HSTRING_IInspectable, _
                      m_pIIterable_IKeyValuePair_HSTRING_IInspectable)
    End If
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIPropertySet)
    Call ReleaseIfc(m_pIMap_HSTRING_IInspectable)
    Call ReleaseIfc(m_pIObservableMap_HSTRING_IInspectable)
    Call ReleaseIfc(m_pIIterable_IKeyValuePair_HSTRING_IInspectable)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIPropertySet
End Property

Public Property Let Ifc(ByVal pIPropertySet As Long)
    Call ReleaseIfc(m_pIPropertySet)
    Call ReleaseIfc(m_pIMap_HSTRING_IInspectable)
    Call ReleaseIfc(m_pIObservableMap_HSTRING_IInspectable)
    Call ReleaseIfc(m_pIIterable_IKeyValuePair_HSTRING_IInspectable)
    m_pIPropertySet = pIPropertySet
    If m_pIPropertySet <> 0& Then
        Call QueryIfc(m_pIPropertySet, _
                      IMap_HSTRING_IInspectable, _
                      m_pIMap_HSTRING_IInspectable)
        
        Call QueryIfc(m_pIPropertySet, _
                      IObservableMap_HSTRING_IInspectable, _
                      m_pIObservableMap_HSTRING_IInspectable)
        
        Call QueryIfc(m_pIPropertySet, _
                      IIterable_IKeyValuePair_HSTRING_IInspectable, _
                      m_pIIterable_IKeyValuePair_HSTRING_IInspectable)
    End If
End Property

' IMap_HSTRING_IInspectable
Public Property Get Size() As Long
    Dim Ret As Long
        If m_pIMap_HSTRING_IInspectable <> 0& Then
            Dim value As Long
            If Invoke(m_pIMap_HSTRING_IInspectable, _
                      IMap_HSTRING_IInspectable_GetSize, _
                      VarPtr(value)) = S_OK Then
                Ret = value
            End If
        End If
    Size = Ret
End Property
    
' ----==== Functions ====----
' IMap_HSTRING_IInspectable
'    IMap_HSTRING_IInspectable_Lookup = 6
'    IMap_HSTRING_IInspectable_HasKey = 8
'    IMap_HSTRING_IInspectable_GetView = 9
'    IMap_HSTRING_IInspectable_Insert = 10
'    IMap_HSTRING_IInspectable_Remove = 11
'    IMap_HSTRING_IInspectable_Clear = 12
    
'IIterable_IKeyValuePair_HSTRING_IInspectable
'    IIterable_IKeyValuePair_HSTRING_IInspectable_First = 6

Public Function Test()

    If m_pIIterable_IKeyValuePair_HSTRING_IInspectable <> 0& Then
        Dim pIIterator_IKeyValuePair_HSTRING_IInspectable As Long
        If Invoke(m_pIIterable_IKeyValuePair_HSTRING_IInspectable, _
                  IIterable_IKeyValuePair_HSTRING_IInspectable_First, _
                  VarPtr(pIIterator_IKeyValuePair_HSTRING_IInspectable)) = S_OK Then
            Dim pIKeyValuePair_String_Object As Long
            If Invoke(pIIterator_IKeyValuePair_HSTRING_IInspectable, _
                      IIterator_IKeyValuePair_HSTRING_IInspectable_GetCurrent, _
                      VarPtr(pIKeyValuePair_String_Object)) = S_OK Then
                Dim hString As Long
                If Invoke(pIKeyValuePair_String_Object, _
                          IKeyValuePair_String_Object_GetKey, _
                          VarPtr(hString)) = S_OK Then
                    Debug.Print "Key = " & GetWindowsString(hString)
                End If
                
                Dim pIPropertyValue As Long
                If Invoke(pIKeyValuePair_String_Object, _
                          IKeyValuePair_String_Object_GetValue, _
                          VarPtr(pIPropertyValue)) = S_OK Then
                
                    Dim PropertyValue As New PropertyValue
                    PropertyValue.Ifc = pIPropertyValue
                    
                    Debug.Print "Value: " & PropertyValue.GetString
                
                    Set PropertyValue = Nothing
                
                End If
            
                Call ReleaseIfc(pIKeyValuePair_String_Object)
            End If
            Call ReleaseIfc(pIIterator_IKeyValuePair_HSTRING_IInspectable)
        End If
    
    End If

End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReadOnlyList_StorageFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----

' ----==== Enums ====----
Private Enum vtb_Interfaces

    ' IVectorView_IStorageFile
    IVectorView_IStorageFile_GetAt = 6
    IVectorView_IStorageFile_GetSize = 7
    
End Enum

' ----==== Variablen ====----
Private m_pIVectorView_IStorageFile As Long

' ----==== Class ====----
Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pIVectorView_IStorageFile)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pIVectorView_IStorageFile
End Property

Public Property Let Ifc(ByVal pIVectorView_IStorageFile As Long)
    Call ReleaseIfc(m_pIVectorView_IStorageFile)
    m_pIVectorView_IStorageFile = pIVectorView_IStorageFile
End Property

'IVectorView_IStorageFile
Public Property Get Size() As Long
    Dim Ret As Long
    If m_pIVectorView_IStorageFile <> 0& Then
        Dim Value As Long
        If Invoke(m_pIVectorView_IStorageFile, _
                  IVectorView_IStorageFile_GetSize, _
                  VarPtr(Value)) = S_OK Then
            Ret = Value
        End If
    End If
    Size = Ret
End Property

' ----==== Functions ====----
'IVectorView_IStorageFile
Public Function GetAt(ByVal index As Long) As StorageFile
    Dim Ret As StorageFile
    If m_pIVectorView_IStorageFile <> 0& Then
        Dim pIStorageFile As Long
        If Invoke(m_pIVectorView_IStorageFile, _
                  IVectorView_IStorageFile_GetAt, _
                  index, _
                  VarPtr(pIStorageFile)) = S_OK Then
            If pIStorageFile <> 0& Then
                Set Ret = New StorageFile
                Ret.Ifc = pIStorageFile
            End If
        End If
    End If
    Set GetAt = Ret
End Function

' ----==== Extra Functions ====----
Public Function GetAll() As StorageFile()
    Dim Ret() As StorageFile
    Dim count As Long
    count = Me.Size
    If count > 0& Then
        ReDim Ret(count - 1)
        Dim Item As Long
        For Item = 0 To count - 1
            Set Ret(Item) = Me.GetAt(Item)
        Next
    End If
    GetAll = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



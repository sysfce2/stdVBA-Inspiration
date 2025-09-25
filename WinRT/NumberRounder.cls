VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NumberRounder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Autor: F. Schüler (frank@activevb.de)
' Datum: 05/2023

Option Explicit

' ----==== Const ====----
Private Const Windows_Globalization_NumberFormatting_NumberRounder As String = "Windows.Globalization.NumberFormatting.NumberRounder"
Private Const IID_INumberRounder As String = "{5473c375-38ed-4631-b80c-ef34fc48b7f5}"

' ----==== Enums ====----
Private Enum vtb_Interfaces
    
    ' INumberRounder
    INumberRounder_RoundInt32 = 6
    INumberRounder_RoundUInt32 = 7
    INumberRounder_RoundInt64 = 8
    INumberRounder_RoundUInt64 = 9
    INumberRounder_RoundSingle = 10
    INumberRounder_RoundDouble = 11
                        
End Enum

' ----==== Variablen ====----
Private m_pINumberRounder As Long

' ----==== Class ====----
Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Call ReleaseIfc(m_pINumberRounder)
End Sub

' ----==== Properties ====----
Public Property Get Ifc() As Long
    Ifc = m_pINumberRounder
End Property

Public Property Let Ifc(ByVal pINumberRounder As Long)
    Call ReleaseIfc(m_pINumberRounder)
    m_pINumberRounder = pINumberRounder
End Property

' ----==== Functions ====----
' INumberRounder
Public Function RoundInt32(ByVal value As Long) As Long
    Dim Ret As Long
    If m_pINumberRounder <> 0& Then
        Dim Result As Long
        If Invoke(m_pINumberRounder, _
                  INumberRounder_RoundInt32, _
                  value, _
                  VarPtr(Result)) = S_OK Then
            Ret = Result
        End If
    End If
    RoundInt32 = Ret
End Function

Public Function RoundUInt32(ByVal value As Long) As Long
    Dim Ret As Long
    If m_pINumberRounder <> 0& Then
        Dim Result As Long
        If Invoke(m_pINumberRounder, _
                  INumberRounder_RoundUInt32, _
                  value, _
                  VarPtr(Result)) = S_OK Then
            Ret = Result
        End If
    End If
    RoundUInt32 = Ret
End Function

Public Function RoundInt64(ByVal value As Currency) As Currency
    Dim Ret As Currency
    If m_pINumberRounder <> 0& Then
        Dim Result As Currency
        If Invoke(m_pINumberRounder, _
                  INumberRounder_RoundInt64, _
                  value, _
                  VarPtr(Result)) = S_OK Then
            Ret = Result
        End If
    End If
    RoundInt64 = Ret
End Function

Public Function RoundUInt64(ByVal value As Currency) As Currency
    Dim Ret As Currency
    If m_pINumberRounder <> 0& Then
        Dim Result As Currency
        If Invoke(m_pINumberRounder, _
                  INumberRounder_RoundUInt64, _
                  value, _
                  VarPtr(Result)) = S_OK Then
            Ret = Result
        End If
    End If
    RoundUInt64 = Ret
End Function

Public Function RoundSingle(ByVal value As Single) As Single
    Dim Ret As Single
    If m_pINumberRounder <> 0& Then
        Dim Result As Single
        If Invoke(m_pINumberRounder, _
                  INumberRounder_RoundSingle, _
                  value, _
                  VarPtr(Result)) = S_OK Then
            Ret = Result
        End If
    End If
    RoundSingle = Ret
End Function

Public Function RoundDouble(ByVal value As Double) As Double
    Dim Ret As Double
    If m_pINumberRounder <> 0& Then
        Dim Result As Double
        If Invoke(m_pINumberRounder, _
                  INumberRounder_RoundDouble, _
                  value, _
                  VarPtr(Result)) = S_OK Then
            Ret = Result
        End If
    End If
    RoundDouble = Ret
End Function

' ----==== Invoke Interface ====----
Private Function Invoke(ByVal pInterface As Long, _
                        ByVal vtb As vtb_Interfaces, _
                        ParamArray var()) As Variant
    If pInterface <> 0& Then
        Invoke = OleInvoke(pInterface, vtb, var)
    End If
End Function



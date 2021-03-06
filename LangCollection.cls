VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LangCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"LanguageCountry"
Attribute VB_Ext_KEY = "Member0" ,"LanguageCountry"
Private mCol As Collection

Public Function Add(Name As String, IsoCode As String, Codigo As Integer, Optional sKey As String) As LinguasdoPais
    'criar um novo objeto
    Dim objNewMember As LinguasdoPais
    Set objNewMember = New LinguasdoPais


    'definir as propriedades passadas para o m�todo
    objNewMember.LangName = Name
    objNewMember.LangIsoCode = IsoCode
    objNewMember.CodigoContryInfo = Codigo
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As LinguasdoPais
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'cria a cole��o quando esta classe � criada
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    'destr�i a cole��o quando esta classe � encerrada
    Set mCol = Nothing
End Sub



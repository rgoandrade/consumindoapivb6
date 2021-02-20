VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LinguasdoPais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Classe usada para pegar as informações das linguagens de cada pais
'Definindo os métodos GET,SET e LET

Private v_sIsoCode As String
Private v_sName As String
Private v_iCodigoLang As Integer
Private v_iCodigoContryInfo As Integer

Public Property Let CodigoContryInfo(ByVal vData As Integer)
    v_iCodigoContryInfo = vData
End Property

Public Property Get CodigoContryInfo() As Integer
    CodigoContryInfo = v_iCodigoContryInfo
End Property

Public Property Let CodigoLang(ByVal vData As Integer)
    v_iCodigoLang = vData
End Property

Public Property Get CodigoLang() As Integer
    CodigoLang = v_iCodigoLang
End Property

Public Property Let LangIsoCode(novoValor As String)
 v_sIsoCode = novoValor
End Property

Public Property Get LangIsoCode() As String
  LangIsoCode = v_sIsoCode
End Property

Public Property Let LangName(novoValor As String)
 v_sName = novoValor
End Property

Public Property Get LangName() As String
  LangName = v_sName
End Property


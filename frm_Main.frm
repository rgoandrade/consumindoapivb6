VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frm_Main 
   Caption         =   "Consumindo API - vb6"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13815
   ScaleHeight     =   4440
   ScaleWidth      =   13815
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inetl 
      Left            =   330
      Top             =   6690
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AccessType      =   1
      Protocol        =   4
      URL             =   "http://"
   End
   Begin VB.CommandButton btnSalvarDados 
      Caption         =   "&Salvar Dados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   705
      TabIndex        =   2
      Top             =   3030
      Width           =   2940
   End
   Begin VB.CommandButton btnBaixarDados 
      Caption         =   "&Baixar Dados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   705
      TabIndex        =   1
      Top             =   1605
      Width           =   2940
   End
   Begin VB.TextBox txtDadosRetornados 
      Height          =   1830
      Left            =   3930
      TabIndex        =   0
      Top             =   1800
      Width           =   9705
   End
   Begin VB.Label lblTitulo 
      Caption         =   "Consumindo API no Vb6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   555
      Left            =   4755
      TabIndex        =   3
      Top             =   270
      Width           =   6975
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'usado referências
' e também um componente Inetl
'Visual Basic For Applications
'Visual Basic runtime objects and procedures
'Visual Basic objects and procedures
'OLE Automation
'Microsoft VBScript Regular Expressions 5.5
'Microsoft Visual Basic 6.0 Extensibility
'Microsoft XML, v6.0
'Microsoft ActiveX Data objects 2.8 Library


Dim clang As New LinguasdoPais
Dim contryes As New Collection
Dim lang As New Collection
Dim Pais As New Pais
Dim fields As New Collection

Dim indiceLang As Integer
Dim vcodigo As Integer

Dim enderecoUrl As String
Dim textParse As String
Dim arquivoXml As DOMDocument60
Dim log As New GeraLog

Private Sub btnBaixarDados_Click()

   Dim posicao As Integer
   Dim str1 As String

   On Error GoTo Click_Error

   Screen.MousePointer = vbHourglass

   Set arquivoXml = New DOMDocument60

   enderecoUrl = "http://webservices.oorsprong.org/websamples.countryinfo/CountryInfoService.wso/FullCountryInfoAllCountries"
                  
                  
   'Acesso/uso por meio do componente Inetl de transferênia de arquivos via ftp
   With Inetl
      .AccessType = icDirect
      .Proxy = ""
      .Protocol = icHTTP
      textParse = .OpenURL(enderecoUrl)
   End With

   'le o arquivo xml para e coloca no FOR e pegar aqueles que começam com a letra A
   arquivoXml.loadXML textParse

   For Each obj In arquivoXml.documentElement.childNodes
   
      str1 = obj.childNodes(0).Text
      posicao = InStr(1, str1, "A", 1)
      If posicao = 0 Or posicao > 1 Then
         arquivoXml.documentElement.removeChild obj
      End If

   Next

   txtDadosRetornados.Text = arquivoXml.xml

   Screen.MousePointer = vbDefault
   log.Registrar "Xml Lido com sucesso!"
   On Error GoTo 0
   Exit Sub

Click_Error:
   log.Registrar "Erro " & Err.Description
   Screen.MousePointer = vbDefault
   MsgBox "Error " & Err.Number & " (" & Err.Description & ")"
End Sub

Private Sub btnSalvarDados_Click()

   On Error GoTo Save_Error

   Screen.MousePointer = vbHourglass

   vcodigo = 1

   If arquivoXml Is Nothing Then
      MsgBox "Não existe informações para serem gravadas", vbOKOnly = vbExclamation, "Atenção"
      Screen.MousePointer = vbDefault
      Exit Sub
   End If

   'lendo informação do País
   For Each obj In arquivoXml.documentElement.childNodes
      indiceLang = 0
      Pais.Codigo = vcodigo
      Pais.IsoCode = obj.childNodes(0).Text
      Pais.Name = obj.childNodes(1).Text
      Pais.CapitalCity = obj.childNodes(2).Text
      Pais.PhoneCode = obj.childNodes(3).Text
      Pais.ContinentCode = obj.childNodes(4).Text
      Pais.CurrencyIsoCode = obj.childNodes(5).Text
      Pais.CountryFla = obj.childNodes(6).Text
    
      Dim languageCount As Integer
      languageCount = obj.childNodes(7).childNodes.Length
    
      While indiceLang <= languageCount - 1
         clang.CodigoContryInfo = vcodigo
         clang.LangIsoCode = obj.childNodes(7).childNodes(indiceLang).childNodes(0).Text
         clang.LangName = obj.childNodes(7).childNodes(indiceLang).childNodes(1).Text
         Pais.LangCollection.Add clang.LangName, clang.LangIsoCode, clang.CodigoContryInfo
         indiceLang = indiceLang + 1
      Wend

      contryes.Add Pais, Pais.IsoCode
      vcodigo = vcodigo + 1
      Set Pais = Nothing
      Set clang = Nothing

   Next


   Dim conecta As New ConectaDatabase
   conecta.GravarDados contryes
 
   Screen.MousePointer = vbDefault
   log.Registrar "Dados Gravados "
   On Error GoTo 0
   Exit Sub
   
'estrututra padrão para tratamento de exceção
Save_Error:
   log.Registrar "Erro " & Err.Description
   Screen.MousePointer = vbDefault
   MsgBox "Error " & Err.Number & " (" & Err.Description & ")"
End Sub


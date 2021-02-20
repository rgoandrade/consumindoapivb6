Attribute VB_Name = "AbrirBase"
' le a api no ini
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias _
   "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
   lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As _
   String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'grava a api no ini
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias _
   "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
   lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As _
   Long
   

Public Sub LerIni(sServer As String, sDb As String)
   sServer = ReadINI("conexao", "server", App.Path & "\config_api.ini")
   sDb = ReadINI("conexao", "db", App.Path & "\config_api.ini")
End Sub

Public Function ReadINI(Secao As String, Entrada As String, Arquivo As String)
   Dim Retlen As String
   Dim ret As String
   ret = String$(255, 0)
   Retlen = GetPrivateProfileString(Secao, Entrada, "", ret, Len(ret), Arquivo)
   ret = Left$(ret, Retlen)
   ReadINI = ret
End Function

Public Sub WriteINI(Secao As String, Entrada As String, texto As String, Arquivo As String)
   WritePrivateProfileString Secao, Entrada, texto, Arquivo
End Sub

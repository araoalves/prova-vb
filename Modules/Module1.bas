Attribute VB_Name = "Module1"
'Criação das variaveis globais para utilização em todos os módulos

Option Explicit

Global conexao As New ADODB.Connection
Global consulta As Recordset

Sub Conectar_BD()
   'Abre a conexao com o BD
   
   Set conexao = New ADODB.Connection
   conexao.ConnectionString = "DRIVER={MySQL ODBC 3.51 DRIVER};" _
   & "SERVER=localhost;" _
   & "DATABASE=sistema_ceuma;" _
   & "UID=root;" _
   & "PASSWORD="
   '& "PORT=3306;" _

   conexao.Open
End Sub

'Fecha a conexao e libera as variaveis
Sub Desconectar_BD()
   'conexao.Close
   'Set conexao = Nothing
End Sub


'Private Sub class_initialize()
 '  DoEvents
  ' Set con = New ADODB.Connection
   
 '  With con
 '     .ConnectionString = "Driver=(MySQL ODBC 3.51 Driver);SERVER=localhost;PWD=;UID=root;PORT=3306;DATABASE=;sistema_ceuma"
 '     .CursorLocation = adUseClient = adUseClient
  '    .Open
  ' End With
'End Sub




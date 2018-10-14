VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "listaAlunos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim conexao As ADODB.Connection, consulta As ADODB.Recordset

'Função que retorna uma listagem de alunos da base de dados
Public Function listaAlunos() As ADODB.Recordset
   consulta.Open "SELECT * FROM Alunos ORDER BY cod_aluno", conexao, adOpenKeyset, adLockBatchOptimistic
   
   Set listaAlunos = consulta
End Function

'codigo para aparecer a lista de alunos no form
'Private Sub Form_Load()
'Dim obj As New cDados
'Set MSHFlexGrid1.DataSource = obj.ListaAlunos
'obj.Desconectar_BD
'End Sub

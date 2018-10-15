VERSION 5.00
Begin VB.Form FormPrincipal 
   Caption         =   "Tela Principal"
   ClientHeight    =   8775
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   15240
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   Picture         =   "Principal.frx":0000
   ScaleHeight     =   8775
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Menu menuCadastro 
      Caption         =   "Cadastro"
      Index           =   1
      Begin VB.Menu menu_curso 
         Caption         =   "Curso"
         Index           =   1
      End
      Begin VB.Menu menu_aluno 
         Caption         =   "Aluno"
         Index           =   1
      End
   End
   Begin VB.Menu menu_alterar 
      Caption         =   "Alterar"
      Index           =   1
   End
   Begin VB.Menu menu_listar 
      Caption         =   "Listar Todos"
      Index           =   1
   End
End
Attribute VB_Name = "FormPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub menu_aluno_Click(Index As Integer)
   FormCadAluno.Show
   Unload Me
End Sub

Private Sub menu_curso_Click(Index As Integer)
   Form2.Show
   Unload Me
End Sub

Private Sub nebu_listar_Click(Index As Integer)

End Sub

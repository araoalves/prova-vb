VERSION 5.00
Begin VB.Form FormPrincipal 
   Caption         =   "Tela Principal"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10350
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton btnListarTodos 
      Caption         =   "Listar Todos"
      Height          =   390
      Left            =   4230
      TabIndex        =   2
      Top             =   0
      Width           =   2220
   End
   Begin VB.CommandButton btnAlterar 
      Caption         =   "Alterar"
      Height          =   390
      Left            =   2160
      TabIndex        =   1
      Top             =   0
      Width           =   2085
   End
   Begin VB.ComboBox Cadastro 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "Principal.frx":0000
      Left            =   15
      List            =   "Principal.frx":000A
      TabIndex        =   0
      Text            =   "Cadastro"
      Top             =   15
      Width           =   2160
   End
End
Attribute VB_Name = "FormPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbCadastro_Change()
  If Click = "Aluno" Then
      FormCadAluno.Show
   End If
   
   
   'If cbCadastro = cbCadastro.Text("Aluno") Then
    '  FormPrincipal.Show
     ' Unload Me
   'End If
End Sub


Private Sub btnAlterar_Click()
   
      FormCadAluno.Show
      Unload Me
End Sub


VERSION 5.00
Begin VB.Form FormCadCurso 
   Caption         =   "Formulário de Curso"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSalvar 
      Caption         =   "Salvar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   8925
      TabIndex        =   7
      Top             =   6105
      Width           =   1680
   End
   Begin VB.TextBox txtHorario 
      Height          =   435
      Left            =   885
      TabIndex        =   6
      Top             =   3390
      Width           =   2265
   End
   Begin VB.TextBox txtCod 
      Height          =   465
      Left            =   900
      TabIndex        =   5
      Top             =   2190
      Width           =   2235
   End
   Begin VB.Data date 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   780
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4770
      Width           =   3045
   End
   Begin VB.TextBox txtNome 
      Height          =   540
      Left            =   945
      TabIndex        =   1
      Top             =   1035
      Width           =   8520
   End
   Begin VB.Label lblCod 
      Caption         =   "Código do Curso"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   915
      TabIndex        =   4
      Top             =   1845
      Width           =   1710
   End
   Begin VB.Label lblHorario 
      Caption         =   "Carga horária"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   945
      TabIndex        =   3
      Top             =   2910
      Width           =   1785
   End
   Begin VB.Label LblData 
      Caption         =   "Data de Cadastro"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   795
      TabIndex        =   2
      Top             =   4290
      Width           =   2715
   End
   Begin VB.Label lblNome 
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   975
      TabIndex        =   0
      Top             =   660
      Width           =   1965
   End
   Begin VB.Shape Shape1 
      Height          =   6615
      Left            =   450
      Top             =   420
      Width           =   10815
   End
End
Attribute VB_Name = "FormCadCurso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub btnSalvar_Click()
   FormPrincipal.Show
   Unload Me
End Sub

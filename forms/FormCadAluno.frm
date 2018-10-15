VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormCadAluno 
   Caption         =   "Cadastro de Aluno"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11550
   LinkTopic       =   "Cadastro de alunos"
   ScaleHeight     =   7710
   ScaleWidth      =   11550
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   420
      Left            =   645
      Top             =   6180
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   741
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox TxtBairro 
      Height          =   375
      Left            =   600
      TabIndex        =   16
      Top             =   5445
      Width           =   7005
   End
   Begin VB.TextBox TxtRua 
      Height          =   465
      Left            =   585
      TabIndex        =   15
      Top             =   4320
      Width           =   6930
   End
   Begin VB.TextBox TxtTelefone 
      Height          =   405
      Left            =   3570
      TabIndex        =   14
      Top             =   3075
      Width           =   4110
   End
   Begin VB.TextBox TxtCep 
      Height          =   435
      Left            =   600
      TabIndex        =   13
      Top             =   3150
      Width           =   1530
   End
   Begin VB.TextBox TxtEmail 
      Height          =   420
      Left            =   3585
      TabIndex        =   12
      Top             =   1965
      Width           =   4080
   End
   Begin VB.TextBox TxtCpf 
      Height          =   390
      Left            =   585
      TabIndex        =   11
      Top             =   1935
      Width           =   1455
   End
   Begin VB.TextBox TxtNome 
      Height          =   405
      Left            =   3630
      TabIndex        =   10
      Top             =   930
      Width           =   4260
   End
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
      Left            =   7530
      TabIndex        =   9
      Top             =   6075
      Width           =   1680
   End
   Begin VB.ComboBox cbCursos 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "FormCadAluno.frx":0000
      Left            =   615
      List            =   "FormCadAluno.frx":000D
      TabIndex        =   7
      Text            =   "Cursos"
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblCep 
      Caption         =   "CEP:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   600
      TabIndex        =   8
      Top             =   2730
      Width           =   1230
   End
   Begin VB.Label lblCpf 
      Caption         =   "CPF:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   585
      TabIndex        =   6
      Top             =   1560
      Width           =   1230
   End
   Begin VB.Label lblEmail 
      Caption         =   "E-mail:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   3600
      TabIndex        =   5
      Top             =   1590
      Width           =   1230
   End
   Begin VB.Label Label1 
      Caption         =   "Bairro:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   615
      TabIndex        =   4
      Top             =   5025
      Width           =   1230
   End
   Begin VB.Label LblRua 
      Caption         =   "Rua:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   585
      TabIndex        =   3
      Top             =   3930
      Width           =   1230
   End
   Begin VB.Label lblTelefone 
      Caption         =   "Telefone:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   3615
      TabIndex        =   2
      Top             =   2640
      Width           =   1230
   End
   Begin VB.Label lblNome 
      Caption         =   "Nome:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   3630
      TabIndex        =   1
      Top             =   480
      Width           =   1230
   End
   Begin VB.Label lblCursos 
      Caption         =   "Cursos:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   615
      TabIndex        =   0
      Top             =   480
      Width           =   1230
   End
   Begin VB.Shape Shape1 
      Height          =   6720
      Index           =   7
      Left            =   240
      Top             =   210
      Width           =   9510
   End
End
Attribute VB_Name = "FormCadAluno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSalvar_Click()
   
   Dim comando_SQL As String
   
   Dim nome As String
   Dim cpf As String
   Dim email As String
   Dim cep As String
   Dim rua As String
   Dim bairro As String
   Dim telefone As String
   
   nome = Me.txtNome
   cpf = Me.TxtCpf
   email = Me.TxtEmail
   cep = Me.TxtCep
   rua = Me.TxtRua
   bairro = Me.TxtBairro
   telefone = Me.TxtTelefone
   
   'Adiciona dados a tabela
   Call Conectar_BD
   
   comando_SQL = "INSERT INTO sistema_ceuma.alunos(nome, cpf, email, cep, rua, bairro, telefone) VALUES ('" & nome & "', '" & cpf & "', '" & email & "', '" & cep & "', '" & rua & "', '" & bairro & "', '" & telefone & "')"
   
   conexao.Execute comando_SQL
   
   Call Desconectar_BD
   
   MsgBox "Dados inseridos com sucesso"
      
   FormPrincipal.Show
   Unload Me
End Sub

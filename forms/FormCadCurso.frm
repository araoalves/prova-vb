VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
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
   Begin VB.TextBox codCurso 
      Height          =   510
      Left            =   930
      TabIndex        =   8
      Top             =   2310
      Width           =   8565
   End
   Begin VB.CommandButton btnVoltar 
      Caption         =   "Voltar"
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
      Index           =   1
      Left            =   6735
      TabIndex        =   7
      Top             =   6120
      Width           =   1680
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   435
      Left            =   735
      Top             =   6180
      Visible         =   0   'False
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   767
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
   Begin VB.CommandButton btnSalvar1 
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
      Index           =   0
      Left            =   8925
      TabIndex        =   6
      Top             =   6105
      Width           =   1680
   End
   Begin VB.TextBox txtHorario 
      Height          =   435
      Left            =   885
      TabIndex        =   5
      Top             =   3390
      Width           =   2265
   End
   Begin VB.Data txtData 
      BOFAction       =   1  'BOF
      Caption         =   "Data"
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
Private Sub btnSalvar1_Click(Index As Integer)
   Dim comando_SQL As String
   
   Dim codCurso As Integer
   Dim nome As String
   Dim horario As String
   Dim data As String
   
   codCurso = Me.codCurso
   
   nome = Me.txtNome
   horario = Me.txtHorario
   data = Me.txtData
   'data = Year(data) & "/" & Month(data) & "/" & Day(data) 'Conversão de data para o formato de BD MYSQL ISO-8601
   
    
   
   'Adiciona dados a tabela
   Call Conectar_BD
   
   comando_SQL = "INSERT INTO sistema_ceuma.cursos(cod_curso, carga_horaria, nome, data_cad) VALUES ('" & codCurso & "', '" & horario & "', '" & nome & "', '" & data & "')"
   
   conexao.Execute comando_SQL
   
   Call Desconectar_BD
   
   MsgBox "Dados inseridos com sucesso"
      
   FormPrincipal.Show
   Unload Me
End Sub

Private Sub btnVoltar_Click(Index As Integer)
   FormPrincipal.Show
   Unload Me
End Sub

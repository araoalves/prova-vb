VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14400
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   14400
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   600
      Left            =   8415
      Top             =   5880
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   1058
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
   Begin VB.CommandButton btn_buscar 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   3165
      TabIndex        =   7
      Top             =   1560
      Width           =   1185
   End
   Begin VB.TextBox txt_codigo 
      Height          =   540
      Left            =   315
      TabIndex        =   6
      Top             =   1560
      Width           =   2340
   End
   Begin VB.CommandButton btn_editar 
      Caption         =   "Editar"
      Height          =   600
      Left            =   3600
      TabIndex        =   4
      Top             =   5820
      Width           =   2925
   End
   Begin VB.TextBox txtNome 
      Height          =   540
      Left            =   420
      TabIndex        =   1
      Top             =   4680
      Width           =   8520
   End
   Begin VB.TextBox txtHorario 
      Height          =   435
      Left            =   375
      TabIndex        =   0
      Top             =   5955
      Width           =   2265
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo do curso"
      Height          =   540
      Left            =   330
      TabIndex        =   5
      Top             =   1170
      Width           =   2040
   End
   Begin VB.Label lblNome 
      Caption         =   "Nome do Curso:"
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
      Index           =   0
      Left            =   450
      TabIndex        =   3
      Top             =   4305
      Width           =   1965
   End
   Begin VB.Label lblHorario 
      Caption         =   "Carga Horária:"
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
      Index           =   1
      Left            =   390
      TabIndex        =   2
      Top             =   5490
      Width           =   1965
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub pesquisa()
   Call Conectar_BD
   
   Dim nome As Integer
   Dim comando_Sql As String
   
   Set consulta = New ADODB.Recordset
   
   comando_Sql = "SELECT * FROM sistema_ceuma.cursos where cod_curso= " & txt_codigo & ""
   consulta.Open comando_Sql, conexao, adOpenStatic, adLockReadOnly
   
   On Error Resume Next
   
   Me.txt_codigo = consulta(1)
   Me.txtHorario = consulta(2)
   Me.txtNome = consulta(3)
   
   Call Desconectar_BD
   
   Exit Sub
End Sub

Private Sub btn_buscar_Click()

Call pesquisa

btn_editar.Enabled = True

End Sub

Private Sub btn_editar_Click()
   Call editarCursos
   
End Sub

Private Sub Form_Load()
Me.btn_editar.Enabled = False

End Sub

Private Sub editarCursos()
Call Conectar_BD
   

   Dim comando_Sql As String
   Set consulta = New ADODB.Recordset
   
   comando_Sql = "UPDATE sistema_ceuma.cursos SET carga_horaria = " & txtHorario & ", nome = '" & txtNome & "' WHERE cod_curso = " & txt_codigo & ""
   consulta.Open comando_Sql, conexao, adOpenStatic, adLockReadOnly
   
   MsgBox "dados alterados com sucesso"
   
   Call pesquisa
   
   Call Desconectar_BD
   
   MsgBox "Dados inseridos com sucesso"
   
   Call Desconectar_BD
   Call limpar_campos
End Sub


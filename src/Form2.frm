VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   Caption         =   "Form1"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   675
      Top             =   3360
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   582
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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   2370
      TabIndex        =   2
      Top             =   2130
      Width           =   1425
   End
   Begin VB.TextBox senhaa 
      Height          =   465
      Left            =   3720
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   435
      Width           =   2070
   End
   Begin VB.TextBox nomee 
      Height          =   480
      Left            =   885
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   375
      Width           =   2265
   End
   Begin VB.Label Label2 
      Caption         =   "SENHA"
      Height          =   300
      Left            =   3825
      TabIndex        =   4
      Top             =   1155
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "NOME"
      Height          =   240
      Left            =   915
      TabIndex        =   3
      Top             =   1185
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   
   Dim nome As String
   Dim senha As String
   
   nome = Me.nomee
   
   senha = Me.senhaa
   
   'Adiciona dados a tabela
   Call Conectar_BD
   
   comando_SQL = "INSERT INTO sistema_ceuma.info(fname, lname) VALUES ('" & nome & "',  '" & senha & "')"
   conexao.Execute comando_SQL
   
   Call Desconectar_BD
   
   MsgBox "Dados inseridos com sucesso"

End Sub



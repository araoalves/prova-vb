VERSION 5.00
Begin VB.Form telaLogin 
   Caption         =   "Form1"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   645
      Left            =   2295
      Picture         =   "telaLogin.frx":0000
      ScaleHeight     =   585
      ScaleWidth      =   570
      TabIndex        =   6
      Top             =   840
      Width           =   630
   End
   Begin VB.CommandButton btnSair 
      Caption         =   "Sair"
      Height          =   540
      Left            =   3015
      TabIndex        =   5
      Top             =   3180
      Width           =   1170
   End
   Begin VB.CommandButton btnEntrar 
      Caption         =   "Entrar"
      Height          =   540
      Left            =   1470
      TabIndex        =   4
      Top             =   3195
      Width           =   1260
   End
   Begin VB.TextBox txtSenha 
      Height          =   480
      IMEMode         =   3  'DISABLE
      Left            =   1485
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2430
      Width           =   2685
   End
   Begin VB.TextBox txtLogin 
      Height          =   495
      Left            =   1500
      TabIndex        =   0
      Top             =   1620
      Width           =   2700
   End
   Begin VB.Label lblSenha 
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   735
      TabIndex        =   3
      Top             =   2520
      Width           =   945
   End
   Begin VB.Label lblLogin 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   735
      TabIndex        =   2
      Top             =   1725
      Width           =   1020
   End
End
Attribute VB_Name = "telaLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnEntrar_Click()
   If txtLogin.Text = "User" And txtSenha.Text = "123" Then
      FormPrincipal.Show
      Unload Me
   Else
      MsgBox "Login ou senha incorretos, tente novamente."
   End If
End Sub

Private Sub btnSair_Click()
   LoginSucceded = False
   Me.Hide
   Unload Me
End Sub

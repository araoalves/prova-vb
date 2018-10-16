VERSION 5.00
Begin VB.Form FormLogin 
   Caption         =   "(none)"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6495
   FontTransparent =   0   'False
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Login"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Login.frx":000C
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "Login.frx":5A5AE
   ScaleHeight     =   4830
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   2730
      Picture         =   "Login.frx":B4B50
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   1050
      Width           =   540
   End
   Begin VB.CommandButton BtnSair 
      Caption         =   "Sair"
      Height          =   540
      Left            =   3240
      TabIndex        =   5
      Top             =   3600
      Width           =   1320
   End
   Begin VB.CommandButton BtnEntrar 
      Caption         =   "Entrar"
      Height          =   525
      Left            =   1590
      TabIndex        =   2
      Top             =   3600
      Width           =   1485
   End
   Begin VB.TextBox TxtSenha 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2865
      Width           =   2850
   End
   Begin VB.TextBox TxtLogin 
      Height          =   495
      Left            =   1695
      MaxLength       =   50
      TabIndex        =   0
      Top             =   2025
      Width           =   2865
   End
   Begin VB.Label LblLogin 
      Caption         =   "Login"
      Height          =   450
      Left            =   645
      TabIndex        =   4
      Top             =   2055
      Width           =   840
   End
   Begin VB.Label LblSenha 
      Caption         =   "Senha"
      Height          =   420
      Left            =   615
      TabIndex        =   3
      Top             =   2955
      Width           =   810
   End
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean
Private Sub BtnEntrar_Click()
   If TxtLogin.Text = "User" And TxtSenha.Text = "123" Then
      FormPrincipal.Show
      Unload Me
   Else
      MsgBox "Login ou senha incorretos, tente novamente."
   End If
End Sub

Private Sub BtnSair_Click()

   LoginSucceded = False
   Me.Hide
   Unload Me
   
End Sub




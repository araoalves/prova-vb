VERSION 5.00
Begin VB.Form FormLogin 
   Caption         =   "Login"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5295
   FontTransparent =   0   'False
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Login"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton BtnSair 
      Caption         =   "Sair"
      Height          =   510
      Left            =   3315
      TabIndex        =   5
      Top             =   2295
      Width           =   1230
   End
   Begin VB.CommandButton BtnEntrar 
      Caption         =   "Entrar>>"
      Height          =   525
      Left            =   1605
      TabIndex        =   2
      Top             =   2280
      Width           =   1485
   End
   Begin VB.TextBox TxtSenha 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1590
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1575
      Width           =   2850
   End
   Begin VB.TextBox TxtLogin 
      Height          =   495
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   0
      Top             =   645
      Width           =   2865
   End
   Begin VB.Label LblLogin 
      Caption         =   "Login"
      Height          =   450
      Left            =   780
      TabIndex        =   4
      Top             =   825
      Width           =   840
   End
   Begin VB.Label LblSenha 
      Caption         =   "Senha"
      Height          =   420
      Left            =   795
      TabIndex        =   3
      Top             =   1800
      Width           =   810
   End
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnEntrar_Click()
   If TxtLogin.Text = "User" And TxtSenha.Text = "123" Then
      Form2.Show
      Unload Me
   Else
      MsgBox "Login ou senha incorretos, tente novamente."
   End If
End Sub

Private Sub BtnSair_Click()
   Unload Me
End Sub


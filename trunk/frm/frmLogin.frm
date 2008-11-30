VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "User Login"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4905
   LinkTopic       =   "LoginForm"
   MaxButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtPswd 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Text            =   "passwort"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtUser 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Text            =   "user"
      Top             =   1080
      Width           =   1935
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQuit_Click()
    'Private cmdQuit_Click
    'End Program immediately
    End
End Sub

Private Function testa(ByRef otto As String, ByVal fritz As Integer) As Boolean
    'Private testa As Boolean
    'Unsere Funktion tut nichts
    ' - [IN] ByRef otto As String: Den Dateinamen der Datei
    ' - [IN] ByVal fritz As Integer: fritz ist n toller hecht
    ' - wenn datei existiert dann true
End Function

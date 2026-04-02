VERSION 5.00
Begin VB.Form frmend 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fim!"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7080
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H000000FF&
      Caption         =   "&Jogar novamente"
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
      Left            =   750
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2925
      Width           =   2085
   End
   Begin VB.CommandButton cmdsair 
      BackColor       =   &H000000FF&
      Caption         =   "&Sair"
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
      Left            =   4245
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2940
      Width           =   1125
   End
   Begin VB.Label lblperd 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "você perdeu o jogo!"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   705
      Left            =   1290
      TabIndex        =   3
      Top             =   915
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.Label lblganha 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Parabéns! você completou o jogo!"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1395
      Left            =   1245
      TabIndex        =   0
      Top             =   615
      Visible         =   0   'False
      Width           =   4320
   End
End
Attribute VB_Name = "frmend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdnew_Click()
lblperd.Visible = False
lblganha.Visible = False
'jogar novamente
Beep 800, 250
Unload Me
Load frmcereija
frmcereija.Visible = True
End Sub
Private Sub cmdsair_Click()
'para sair
Beep 800, 250
End
End Sub

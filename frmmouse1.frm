VERSION 5.00
Begin VB.Form frmmouse1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Treinando com o mouse"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6405
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picinfo 
      BackColor       =   &H00C0C0C0&
      Height          =   4425
      Left            =   60
      ScaleHeight     =   4365
      ScaleWidth      =   6345
      TabIndex        =   5
      Top             =   75
      Visible         =   0   'False
      Width           =   6405
      Begin VB.CommandButton cmdclose 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   2595
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3480
         Width           =   900
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Atenção, só há 300 cerejas, se você deixar cair mais de 200, não consiguirá completar seu objetivo e perderá o jogo. Boa sorte!"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   570
         TabIndex        =   8
         Top             =   1665
         Width           =   5325
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmmouse1.frx":0000
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   165
         TabIndex        =   7
         Top             =   75
         Width           =   5325
      End
      Begin VB.Image img 
         Height          =   885
         Left            =   5610
         Picture         =   "frmmouse1.frx":00DC
         Top             =   30
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdinfo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Info"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   5085
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1500
      Width           =   900
   End
   Begin VB.Timer tmr 
      Interval        =   1000
      Left            =   165
      Top             =   150
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00C0C0C0&
      Caption         =   "5"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1395
      Width           =   990
   End
   Begin VB.PictureBox picok 
      BackColor       =   &H00808080&
      Height          =   765
      Left            =   2580
      ScaleHeight     =   705
      ScaleWidth      =   975
      TabIndex        =   2
      Top             =   3390
      Width           =   1035
   End
   Begin VB.Label lblinfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O jogo começara em:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   450
      Index           =   1
      Left            =   1740
      TabIndex        =   3
      Top             =   2790
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.Label lblinfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Este é um jogo de treino com o mouse. para instruções click em info, para começar arraste o botão para o espaço abaixo."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Index           =   0
      Left            =   465
      TabIndex        =   1
      Top             =   195
      Width           =   5295
   End
End
Attribute VB_Name = "frmmouse1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdclose_Click()
Beep 800, 250
'ouculta as instruções
picinfo.Visible = False
End Sub
Private Sub cmdinfo_Click()
'mostra as instruções
picinfo.Visible = True
Beep 800, 250
End Sub
Private Sub cmdinfo_DragDrop(Source As Control, X As Single, Y As Single)
'para um botão não ficar em cima do outro eu "cancelo" o drag and drop
cmdok.Visible = True
End Sub
Private Sub cmdok_Click()
'fecha esse form e abre o outro
If cmdok.Caption <> "OK" Then Exit Sub
Unload Me
Load frmcereija
frmcereija.Visible = True
End Sub

Private Sub cmdok_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'ocultao botaão
cmdok.Visible = False
End Sub
Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
'ajeita o botão no lugar
cmdok.Left = X - (cmdok.Width / 2)
cmdok.Top = Y - (cmdok.Height / 2)
cmdok.Visible = True
End Sub

Private Sub lblinfo_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'ajeita o botão no lugar
cmdok.Left = X - (cmdok.Width / 2)
cmdok.Top = Y - (cmdok.Height / 2)
cmdok.Visible = True
End Sub

Private Sub picok_DragDrop(Source As Control, X As Single, Y As Single)
' faz o botão "encaixar" na picture, e destiva o drag ond drop
cmdok.Top = picok.Top
cmdok.Left = picok.Left
lblinfo(1).Visible = True
cmdok.DragMode = False
cmdok.Visible = True
End Sub
Private Sub tmr_Timer()
'faz a contagem regreciva, para começar o jogoe e desativa o draag and drop
If cmdok.Top = picok.Top And cmdok.Left = picok.Left Then
    cmdok.Caption = cmdok.Caption - 1
    Beep 1000, 250
End If
If cmdok.Caption = 0 Then
    cmdok.Caption = "OK"
    tmr.Enabled = False
End If
End Sub

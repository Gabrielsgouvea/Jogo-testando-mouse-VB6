VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmcereija 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14580
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   9.75
      Charset         =   0
      Weight          =   900
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmmouse.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   6915
   ScaleWidth      =   14580
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdl 
      Left            =   -15
      Top             =   1170
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdmusic 
      BackColor       =   &H000000FF&
      Caption         =   "Música"
      Height          =   375
      Left            =   10935
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   375
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MCI.MMControl mmcmp3 
      Height          =   555
      Left            =   4410
      TabIndex        =   8
      Top             =   1170
      Visible         =   0   'False
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   979
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.PictureBox piccancel 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   8085
      ScaleHeight     =   690
      ScaleWidth      =   2580
      TabIndex        =   6
      Top             =   495
      Visible         =   0   'False
      Width           =   2580
      Begin VB.Label lblcancel 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         BackStyle       =   0  'Transparent
         Caption         =   "Sair"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   -30
         TabIndex        =   7
         Top             =   0
         Width           =   2610
      End
   End
   Begin MCI.MMControl mmc 
      Height          =   465
      Left            =   4410
      TabIndex        =   4
      Top             =   555
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   820
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer tmrpause 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   -30
      Top             =   750
   End
   Begin VB.Timer tmr 
      Interval        =   1
      Left            =   -15
      Top             =   375
   End
   Begin VB.PictureBox picmenu 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   -15
      ScaleHeight     =   390
      ScaleWidth      =   14550
      TabIndex        =   0
      Top             =   -30
      Width           =   14550
      Begin VB.CommandButton cmdmute 
         BackColor       =   &H000000FF&
         Caption         =   "pausar música"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1815
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   15
         Width           =   1500
      End
      Begin VB.Label lblnivel 
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel: 1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6735
         TabIndex        =   5
         Top             =   0
         Width           =   930
      End
      Begin VB.Label lblpontos 
         BackStyle       =   0  'Transparent
         Caption         =   "Pontos: 0"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   12975
         TabIndex        =   2
         Top             =   0
         Width           =   1275
      End
      Begin VB.Label lblperdido 
         BackStyle       =   0  'Transparent
         Caption         =   "Perdidos: 0"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   45
         TabIndex        =   1
         Top             =   0
         Width           =   1770
      End
   End
   Begin VB.Label lblpause 
      BackStyle       =   0  'Transparent
      Caption         =   "Pausa"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   615
      TabIndex        =   3
      Top             =   465
      Visible         =   0   'False
      Width           =   2760
   End
   Begin VB.Image img 
      Height          =   885
      Left            =   3390
      Picture         =   "frmmouse.frx":000C
      Top             =   405
      Width           =   720
   End
End
Attribute VB_Name = "frmcereija"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'variavei utilizadas para criar degrade de fundo
'------------------------------------------------
Dim cor As Single
Dim larg As Integer
Dim alt As Integer
Dim step As Single
'------------------------------------------------
'variaveis usadas para gerar circulos randomicos no pause
'------------------------------------------------
Dim l As Integer, p As Integer, t As Integer
'------------------------------------------------
Dim n As Integer 'usada em loop
Dim score As Integer 'guarda os pontos
Dim perdidos As Integer 'guarda os erros
Dim velocidade As Integer 'velocidade em que a cereija cai
Dim clique As String ' guarda o som de quando acerta a cereija
Dim musica As String ' guarda o nome da música

Private Sub cmdmusic_Click()
cdl.Filter = "Files (*.mp3)|*.mp3"
cdl.DefaultExt = "mp3"
cdl.DialogTitle = "Escolher música"
'cdlsave.Flags = cdlOFNFileMustExist + cdlOFNPathMustExist
cdl.ShowOpen
musica = cdl.InitDir + cdl.FileName
mmcmp3.FileName = musica
mmcmp3.Command = "open"
End Sub

Private Sub cmdmute_Click()
mmcmp3.Command = "pause"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'pausa/despausa o jogo
If KeyAscii = 112 Or KeyAscii = 80 Then
    pause
End If
End Sub
Private Sub Form_Load()
'randomiza o tempo
Randomize Timer

'velocidade inicial
velocidade = 50

'carrega o som do clique
clique = App.Path + "\som.wav"

'carrega musica de fundo
musica = App.Path + "/metal.mp3"
mmcmp3.FileName = musica
mmcmp3.Command = "open"
mmcmp3.Command = "play"

'colore o form
cor = 255
larg = Me.ScaleWidth
alt = Me.ScaleHeight
step = 255 / alt
For n = 0 To alt
    Me.Line (0, n)-(larg, n), RGB(0, 0, cor), BF
    cor = cor - step
Next

'coloca efeito no botão sair
cor = 255
larg = piccancel.ScaleWidth
alt = piccancel.ScaleHeight
step = 255 / alt
For n = 0 To alt
piccancel.Line (0, n)-(larg, n), RGB(cor, cor, cor), BF
cor = cor - step
Next

'randomiza a posição da cereija
img.Left = (Me.ScaleWidth - img.Width) * Rnd
img.Top = 0

'centralizando barra
picmenu.Width = Me.Width
picmenu.Left = 0
picmenu.Top = 0
lblnivel.Top = 0
lblnivel.Left = (Me.ScaleWidth / 2) - (lblnivel.Width / 2)
lblperdido.Top = 0
lblperdido.Left = 10
lblpontos.Top = 0
lblpontos.Left = Me.ScaleWidth - lblpontos.Width
cmdmusic.Top = picmenu.Height
cmdmusic.Left = Me.ScaleWidth - cmdmusic.Width
cmdmute.Left = lblperdido.Width + 5
cmdmute.Top = 0

'coloca efeito na barra de cima
cor = 255
larg = picmenu.ScaleWidth
alt = picmenu.ScaleHeight
step = 255 / alt
For n = 0 To alt
picmenu.Line (0, n)-(larg, n), RGB(cor, cor, cor), BF
cor = cor - step
Next

'centraliza o label pause
lblpause.Left = (Me.ScaleWidth / 2) - (lblpause.Width / 2)
lblpause.Top = (Me.ScaleHeight / 2) - (lblpause.Height / 2)

'centraliza o botão de sair
piccancel.Left = (Me.ScaleWidth / 2) - (piccancel.Width / 2)
piccancel.Top = (Me.ScaleHeight / 2) - (piccancel.Height / 2) + lblpause.Height
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'pausa o jogo caso cclique com o botão direito
If Button = 2 Then
Call pause
End If
End Sub

Private Sub img_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'permite apenas o botão esquerdo
If Button <> 1 Then Exit Sub

'aumenta a pontuação e valta com o objeto para cima, randomizando sua posição
score = score + 1
If score = 100 Then
    Load frmend
    frmend.lblganha.Visible = True
    frmend.Visible = True
    Unload Me
    Set frmcereija = Nothing
End If
'aumenta a velocidade a cada cinco acertos
If score Mod 5 = 0 Then
     velocidade = velocidade + 10
End If
lblpontos.Caption = "Pontos: " + Str(score)
'randomizando onde a "cereija" vai cair
img.Left = (Me.ScaleWidth - img.Width) * Rnd
img.Top = 0

'muda a cor do fundo de acordo com o nivel
Call fundo(score)

'reproduz o som
mmc.Command = "close"
mmc.FileName = clique
mmc.Command = "open"
mmc.Command = "play"
End Sub

Private Sub lblcancel_Click()
Beep 1000, 100
'sai do jogo
If MsgBox("Deseja sair?", vbYesNo, "Sair?") = vbYes Then
    Beep 1000, 100
    End
End If
Beep 1000, 100
End Sub

Private Sub tmr_Timer()
'faz a cereija movimentar, e retornar para cima caso passe a tela
img.Top = img.Top + velocidade
If img.Top > Me.ScaleHeight Then
    Beep 1000, 50
    perdidos = perdidos + 1
    img.Top = 0
    img.Left = (Me.ScaleWidth - img.Width) * Rnd
    lblperdido.Caption = "Perdidos: " + Str(perdidos)
End If
If perdidos = 201 Then
    Load frmend
    frmend.lblperd.Visible = True
    frmend.Visible = True
    Unload Me
    Set frmcereija = Nothing
End If
End Sub
Private Sub tmrpause_Timer()
' randomiza bolas no pausa
t = (200 * Rnd) + 3 ' randomiza variáveis
l = (Me.ScaleWidth * Rnd) + 1
p = (Me.ScaleHeight * Rnd) + 1
Me.DrawWidth = t ' usa uma delas para o tamanho dos círculos
Me.ForeColor = RGB(256 * Rnd, 256 * Rnd, 256 * Rnd) ' randomiza as cores
Me.PSet (l, p) ' usa outras duas para a posição das figuras
End Sub

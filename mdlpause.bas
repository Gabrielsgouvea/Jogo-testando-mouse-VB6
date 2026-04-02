Attribute VB_Name = "mdlpause"
'biblioteca de beep
Rem OBS: inserido no module pora ser publico, ja que não posso declarar biblioteca publica no form
Public Declare Function Beep Lib "kernel32" _
  (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Sub pause()
'instruçõe para pausar o jogo
With frmcereija
    If .tmr.Enabled = True Then
        .tmr.Enabled = False
        .img.Enabled = False
        .img.Visible = False
        .cmdmute.Visible = False
        .lblpause.Visible = True
        .tmrpause.Enabled = True
        .piccancel.Visible = True
        .cmdmusic.Visible = True
        .MousePointer = 0 'padrão
        'pausa a musica de fundo
        .mmcmp3.Command = "pause"
    Else
        .tmr.Enabled = True
        .img.Enabled = True
        .img.Visible = True
        .cmdmute.Visible = True
        .lblpause.Visible = False
        .tmrpause.Enabled = False
        .piccancel.Visible = False
        .cmdmusic.Visible = False
        'toca a música de fundo
        .mmcmp3.Command = "play"
        .MousePointer = 2 'cruz
        'limpa as bolas do form
        .Cls
        Call fundo_pos_pause
    End If
End With
Beep 1000, 50
Beep 1100, 50
Beep 1200, 50
End Sub
Sub fundo(score)
With frmcereija
    'mudançã no fundo
    'azul
'    If score < 15 Then
'        'colore o form
'        cor = 255
'        larg = .ScaleWidth
'        alt = .ScaleHeight
'        step = 255 / alt
'        For n = 0 To alt
'            frmcereija.Line (0, n)-(larg, n), RGB(0, 0, cor), BF
'            cor = cor - step
'        Next
'    End If
    
    'amarelo
    If score = 15 Then
        .lblnivel.Caption = "Nivel: 2"
        'colore o form
        cor = 255
        larg = .ScaleWidth
        alt = .ScaleHeight
        step = 255 / alt
        For n = 0 To alt
            frmcereija.Line (0, n)-(larg, n), RGB(cor, cor, 0), BF
            cor = cor - step
        Next
    End If

    'magento
    If score = 30 Then
        .lblnivel.Caption = "Nivel: 3"
        'colore o form
        cor = 255
        larg = .ScaleWidth
        alt = .ScaleHeight
        step = 255 / alt
        For n = 0 To alt
            frmcereija.Line (0, n)-(larg, n), RGB(cor, 0, cor), BF
            cor = cor - step
        Next
    End If

    'verde
    If score = 50 Then
        .lblnivel.Caption = "Nivel: 4"
        'colore o form
        cor = 255
        larg = .ScaleWidth
        alt = .ScaleHeight
        step = 255 / alt
        For n = 0 To alt
            frmcereija.Line (0, n)-(larg, n), RGB(0, cor, 0), BF
            cor = cor - step
        Next
    End If

    'verde
    If score = 70 Then
        .lblnivel.Caption = "Nivel: 5"
        'colore o form
        cor = 255
        larg = .ScaleWidth
        alt = .ScaleHeight
        step = 255 / alt
        For n = 0 To alt
            frmcereija.Line (0, n)-(larg, n), RGB(cor, 0, 0), BF
            cor = cor - step
        Next
    End If
    
    'preto
    If score = 90 Then
        .lblnivel.Caption = "Nivel: 6"
        'colore o form
        cor = 255
        larg = .ScaleWidth
        alt = .ScaleHeight
        step = 255 / alt
        For n = 0 To alt
            frmcereija.Line (0, n)-(larg, n), RGB(cor, cor, cor), BF
            cor = cor - step
        Next
    End If
End With
End Sub
Sub fundo_pos_pause()
'recoloreo fundo apos o pausa
cor = 255
larg = frmcereija.ScaleWidth
alt = frmcereija.ScaleHeight
step = 255 / alt
For n = 0 To alt
    If Right(frmcereija.lblnivel.Caption, 1) = 1 Then
        frmcereija.Line (0, n)-(larg, n), RGB(0, 0, cor), BF
        cor = cor - step
    End If
    
    If Right(frmcereija.lblnivel.Caption, 1) = 2 Then
        frmcereija.Line (0, n)-(larg, n), RGB(cor, cor, 0), BF
        cor = cor - step
    End If
    
    If Right(frmcereija.lblnivel.Caption, 1) = 3 Then
        frmcereija.Line (0, n)-(larg, n), RGB(cor, 0, cor), BF
        cor = cor - step
    End If
    
    If Right(frmcereija.lblnivel.Caption, 1) = 4 Then
        frmcereija.Line (0, n)-(larg, n), RGB(0, cor, 0), BF
        cor = cor - step
    End If
    
    If Right(frmcereija.lblnivel.Caption, 1) = 5 Then
        frmcereija.Line (0, n)-(larg, n), RGB(cor, 0, 0), BF
        cor = cor - step
    End If
    
    If Right(frmcereija.lblnivel.Caption, 1) = 6 Then
        frmcereija.Line (0, n)-(larg, n), RGB(cor, cor, cor), BF
        cor = cor - step
    End If
Next
End Sub

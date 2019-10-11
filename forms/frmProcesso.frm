VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProcesso 
   Caption         =   "Processo"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "frmProcesso.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Activate()
'Configura a largura do lblProcesso (verde) para 0
      frmProcesso.lblProcesso.Width = 0

      'Chama a sub principal que � o c�digo das ac�es
      Call contar
            
End Sub

Sub contar()

Dim Percentual As Single 'vari�vel que armazena resultado de divis�o
Dim contador As Integer 'conta atual quantidade de la�os feitos
Dim limite As Integer 'apresentando uma vari�vel para armazenar o valor m�ximo

'atribui a quantidade m�xima de c�lulas a serem preenchidas
limite = 3000

'seleciona coluna A para iniciar a contagem
Range("A1").Select

'la�o repete a��o at� vari�vel limite. X � vari�vel in�cio e incrementada
For x = 1 To limite

    'atribui o valor atual de X na c�lula ativa/selecionada
    ActiveCell = x
    'percorre uma linha abaixo e n�o muda de coluna
    ActiveCell.Offset(1, 0).Select

    'conta qual � a quantidade j� realizada de a��es
    contador = contador + 1
    'divide a quantidade feita pelo limite e a fra��o %
    Percentual = contador / limite

    ' Chama atualiza�ao de barra
    AtualizaBarra Percentual

Next x 'repete o la�o se n�o chegou ainda no limite

'fecha a janela (formul�rio) ap�s concluir
frmProcesso.Hide

End Sub
Sub AtualizaBarra(Percentual As Single) 'vari�vel reservada para ser %

    With frmProcesso 'With usa o frmprocesso para as a��es abaixo
    'sem ter que repetir o nome do objeto frmprocesso

        ' Atualiza o T�tulo do Quadro que comporta a barra para %
        .FrameProcesso.Caption = Format(Percentual, "0%")

        ' Atualza o tamanho da Barra (label)
        .lblProcesso.Width = Percentual * _
        (.FrameProcesso.Width - 10)
    End With 'final do uso de frmprocesso diretamente
    
    'Habilita o userform para ser atualizado
    DoEvents
End Sub
Sub Executar()
frmProcesso.Show
End Sub




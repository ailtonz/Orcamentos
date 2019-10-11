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

      'Chama a sub principal que é o código das acões
      Call contar
            
End Sub

Sub contar()

Dim Percentual As Single 'variável que armazena resultado de divisão
Dim contador As Integer 'conta atual quantidade de laços feitos
Dim limite As Integer 'apresentando uma variável para armazenar o valor máximo

'atribui a quantidade máxima de células a serem preenchidas
limite = 3000

'seleciona coluna A para iniciar a contagem
Range("A1").Select

'laço repete ação até variável limite. X é variável início e incrementada
For x = 1 To limite

    'atribui o valor atual de X na cálula ativa/selecionada
    ActiveCell = x
    'percorre uma linha abaixo e não muda de coluna
    ActiveCell.Offset(1, 0).Select

    'conta qual é a quantidade já realizada de ações
    contador = contador + 1
    'divide a quantidade feita pelo limite e a fração %
    Percentual = contador / limite

    ' Chama atualizaçao de barra
    AtualizaBarra Percentual

Next x 'repete o laço se não chegou ainda no limite

'fecha a janela (formulário) após concluir
frmProcesso.Hide

End Sub
Sub AtualizaBarra(Percentual As Single) 'variável reservada para ser %

    With frmProcesso 'With usa o frmprocesso para as ações abaixo
    'sem ter que repetir o nome do objeto frmprocesso

        ' Atualiza o Título do Quadro que comporta a barra para %
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




## Mostrar a guia desenvolvedor no Word

A guia **Desenvolvedor** não é exibida por padrão, mas você pode adicioná-la à faixa de opções.

1. Na guia **Arquivo**, vá para **Opções** > **Personalizar Faixa de Opções**.

2. Em **Personalizar a Faixa de Opções** e em **Guias Principais**, marque a caixa de seleção **Desenvolvedor**.

> :bell: **Importante:** <br>Depois de mostrar a guia, a guia **Desenvolvedor** permanecerá visível, a menos que você desinstale a caixa de seleção ou tenha que reinstalar o `Microsoft Office`.

3. Com a guia **Desenvolvedor** visível, clique nela e clique em `Visual Basic` para exibir o `IDE` do `Microsoft Visual Basic for Applications`.

4. Pesquise na **Caixa de Ferramentas** do projeto por **Project** > **Microsoft Word Objetos**, clique duas vezes em **ThisDocument** para exibir a janela **ThisDocument(Código)**.

5. Na janela **ThisDocument(Código)**, altere a aba **(Geral)** para **Document**, para exibir os procedimentos existentes para este documento. Em **New** altere para o procedimento **Open**, para executar o procedimento uma vez quando o documento for aberto e deixe com o seguinte código-fonte:

```VBA
Option Explicit

Private Sub Document_Open()

    Project.UserForm1.Show 1    ' Exibir como caixa de diálogo modal (restrita).
    
End Sub
```

6. Clique no menu **Inserir** > **Userform**, para inserir o formulário **UserForm1**, pressione F7 ou clique duas vezes no corpo do formulário para exibir a janela **UserForm1(Código)** e deixe com o seguinte código-fonte:

```VBA
Option Explicit

Private Sub UserForm_Initialize()

    ' Escreva seu código-fonte VBA aqui!
    
End Sub
```

> :+1: **Parabéns:** <br>Aqui você conclui nosso artigo, mas precisamos **Salvar** este documento como __*.docm__ que é um tipo de documento especial que contém a programação.

7. Feche o `IDE` do `Visual Basic for Application` e vá para o `Word` do `Microsoft Office`, no menu `Word` vá para **Arquivo** > **Salvar como** > **Procure** o local e em **Tipo** escolha o tipo de __Documento Habilitado para Macro do Word (*.docm)__, fechar tudo e abrir o documento `Word`. Habilitar a execução na primeira vez, se necessário, e o formulário será exibido.

![screenshot](https://github.com/fabasa-pro/vba01userform/blob/main/vba01userform.png)

## Bem-vindo à programação em Visual Basic

Este artigo visa apenas mostrar como utilizar o `Visual Basic` no `Microsoft Word` e requer apenas [conhecimentos básicos](https://learn.microsoft.com/pt-br/office/vba/library-reference/concepts/getting-started-with-vba-in-office) da linguagem.

* No documento inserimos um código na função para chamar o formulário, assim como fazemos para escolher o formulário a ser executado primeiro.
* No formulário inserimos um código na função apenas para iniciar, assim você pode inserir seu código apenas para testes e estudos.

A seguir veremos como [ocultar o Microsoft Office](https://github.com/fabasa-pro/vba02hideapplication), deixando apenas a janela principal do `Visual Basic` visível, para quem precisa programar um `Desktop Application` e não quer ver aquele documento `Word` aberto o tempo todo . Lembrando que podemos manipular todo o documento `Word` para imprimir ou visualizar relatórios, imprimir ou salvar modelos de documentos com total automação, sem a necessidade de abrir e digitar diretamente, apenas utilizando códigos em `Visual Bassic`.

## Licenciado sob a licença MIT

Copyright (C) 2012 - 2024 @Fabasa-Pro. Todos os direitos reservados.

Consulte [LICENSE.TXT](https://github.com/fabasa-pro/vba01userform/blob/main/LICENSE.TXT) na raiz do projeto para obter informações.

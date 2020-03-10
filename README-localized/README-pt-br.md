---
page_type: sample
products:
- office-365
- office-outlook
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 10/19/2017 1:01:24 PM
---
# Suplemento do Outlook: Enviar mensagem acionável

O exemplo demonstra como [ativar um suplemento de uma mensagem acionável](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message) e passar um contexto de inicialização para o suplemento.

## Execução do exemplo

Para executar esse exemplo, você precisará hospedar os arquivos incluídos em um servidor da Web. A escolha do servidor Web é inteiramente por você. O único requisito é que o servidor Web seja protegido por um certificado SSL válido. 

### Atualizar o manifesto

Antes de carregar o suplemento, você deve atualizar o](actionable-add-in.xml)manifesto[para substituir todas as instâncias de `localhost:8080` com a URL base do servidor em que você está hospedando seu suplemento.

### Instalar o suplemento

Siga as instruções em [Realizar sideload dos suplementos do Outlook para teste ](https://docs.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing) para realizar o sideload `actionable-add-in.xml` para instalar o suplemento.

### Use o suplemento

O suplemento adiciona dois botões à faixa de opções durante a leitura de mensagens de email.

- **Enviar ativação do suplemento** \- esse botão permite que você envie uma mensagem de email acionável com botões que chamam suplementos.
- **exibir contexto de inicialização**-esse botão abre um painel de tarefas e exibe o contexto de inicialização atual (se houver).

> **Observação:** O contexto de inicialização só é apresentado quando o painel de tarefas é chamado de uma mensagem acionável. Se você abrir o painel de tarefas usando o botão da faixa de opções, verá uma mensagem informando que é necessário invocar o suplemento a partir de uma mensagem.
>
> ![Uma captura de tela da mensagem apresentada quando você ativa o suplemento manualmente](readme-images/manual-activation.PNG)

1. Clique no botão **enviar ativação de suplemento**. É aberta uma caixa de diálogo para enviar a mensagem: 

    ![Uma captura de tela da caixa de diálogo Enviar mensagem](readme-images/send-message.PNG)
1. Opcional: Modificar o contexto de inicialização.
1. Clique no botão **enviar mensagem**.
1. Quando a mensagem chegar na caixa de entrada, abra-a.

    ![Uma captura de tela da mensagem acionável enviada pelo suplemento](readme-images/actionable-message.PNG)
1. Clique no botão **invocar "contexto de inicialização de exibição**".
1. O painel de tarefas de suplemento abre e exibe o contexto de inicialização.

    ![Captura de tela do painel de tarefas do suplemento](readme-images/activated-taskpane.PNG)

### Testar a instalação sob demanda do suplemento da loja

Você pode usar o segundo botão na mensagem acionável para ver como funciona a instalação sob demanda dos suplementos da loja. Antes de experimentá-lo, se você já tem o [analisador de cabeçalho de mensagem](https://appsource.microsoft.com/en-us/product/office/WA104005406), desinstale-o.

Este projeto adotou o [Código de Conduta de Código Aberto da Microsoft](https://opensource.microsoft.com/codeofconduct/).  Para saber mais, confira [Perguntas frequentes sobre o Código de Conduta](https://opensource.microsoft.com/codeofconduct/faq/) ou contate [opencode@microsoft.com](mailto:opencode@microsoft.com) se tiver outras dúvidas ou comentários.

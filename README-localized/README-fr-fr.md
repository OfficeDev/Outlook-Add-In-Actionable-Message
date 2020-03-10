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
# Complément Outlook : Activation de message actionnable

Cet exemple présente comment [activer un complément à partir d’un message actionnable](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message) et transmettre un contexte d'initialisation vers le complément.

## Exécution de l’exemple

Pour exécuter cet exemple, vous devez héberger les fichiers inclus sur un serveur web. Le choix du serveur web vous appartient. La seule exigence est que le serveur soit protégé par un certificat SSL valide. 

### Mise à jour du manifeste

Avant de charger le complément, vous devez mettre à jour le [manifeste](actionable-add-in.xml) pour remplacer toutes les instances du `localhost:8080` par l’URL de base du serveur sur lequel vous hébergez le complément.

### Installation du complément

Suivez les instructions indiquées dans l’article [Chargement de version test des compléments Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing) pour charger le fichier `actionable-add-in.xml` afin d'installer le complément.

### Utilisation du complément

Le complément ajoute deux boutons au ruban lors de la lecture de courriers électroniques.

- **Envoyer une activation de complément** : ce bouton vous permet de vous envoyer un courrier électronique actionnable à l’aide de boutons appelant des compléments.
- **Afficher le contexte d’initialisation** : ce bouton permet d’ouvrir un volet de tâches et d’afficher le contexte d’initialisation actuel (le cas échéant).

> **Remarque :** Le contexte d’initialisation est uniquement présent lorsque le volet de tâches est appelé à partir d’un message actionnable. Si vous ouvrez le volet de tâches à l’aide du bouton du ruban, un message s’affiche vous informant que vous devez appeler le complément à partir d’un message.
>
> ![Capture d’écran du message présenté lorsque vous activez manuellement le complément](readme-images/manual-activation.PNG)

1. Cliquez sur le bouton **Envoyer une activation de complément**. Une boîte de dialogue s’ouvre pour envoyer le message : 

    ![Capture d’écran de la boîte de dialogue du message envoyé](readme-images/send-message.PNG)
1. (facultatif) : Modifier le contexte d’initialisation.
1. Cliquez sur le bouton **Envoyer un message**.
1. Ouvrez le message lorsqu'il arrive dans votre boîte aux lettres.

    ![Capture d’écran du message actionnable envoyé par le complément](readme-images/actionable-message.PNG)
1. Cliquez sur le bouton **Appeler l'« Affichage du contexte d’initialisation »**.
1. Le volet de tâches du complément s’ouvre et affiche le contexte d’initialisation.

    ![Capture d’écran du volet de tâches du complément](readme-images/activated-taskpane.PNG)

### Essayez l’installation à la demande du complément du store

Vous pouvez utiliser le second bouton dans le message actionnable pour voir le fonctionnement de l’installation à la demande des compléments du store. Avant de faire un essai, veillez à désinstaller l'[Analyseur d’en-têtes de messages](https://appsource.microsoft.com/en-us/product/office/WA104005406) si vous le possédez déjà.

Ce projet a adopté le [Code de conduite Open Source de Microsoft](https://opensource.microsoft.com/codeofconduct/). Pour en savoir plus, reportez-vous à la [FAQ relative au code de conduite](https://opensource.microsoft.com/codeofconduct/faq/) ou contactez [opencode@microsoft.com](mailto:opencode@microsoft.com) pour toute question ou tout commentaire.

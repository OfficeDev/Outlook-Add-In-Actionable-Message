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
# Complemento de Outlook: Activación de mensajes que requieren acción

En este ejemplo se muestra cómo [activar un complemento desde un mensaje que requiere acción](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message) y cómo pasar un contexto de inicialización al complemento.

## Ejecución del ejemplo

Para ejecutar este ejemplo, tendrá que hospedar los archivos incluidos en un servidor web. Puede elegir el servidor web que prefiera. El único requisito que debe cumplir el servidor web es que esté protegido con un certificado SSL válido. 

### Actualizar el manifiesto

Antes de cargar el complemento, debe actualizar el [manifiesto](actionable-add-in.xml) para reemplazar todas las instancias de `localhost:8080` por la dirección URL base del servidor en el que se hospeda el complemento.

### Instalar el complemento

Siga las instrucciones que se indican en [Transferir localmente complementos de Outlook para pruebas](https://docs.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing) para transferir localmente el archivo `actionable-add-in.xml` para instalar el complemento.

### Usar el complemento

El complemento agrega dos botones a la cinta de opciones cuando se leen mensajes de correo electrónico.

- **Enviar activación de complemento**: este botón le permite enviarse un mensaje de correo electrónico con botones que invocan complementos.
- **Ver contexto de inicialización**: este botón abre un panel de tareas y muestra el contexto de inicialización actual (si existe).

> **Nota:** El contexto de inicialización solo aparece cuando se invoca el panel de tareas desde un mensaje que requiere acción. Si abre el panel de tareas mediante el botón de la cinta de opciones, verá un mensaje que le indicará que debe invocar el complemento desde un mensaje.
>
> ![Captura de pantalla del mensaje que se muestra al activar el complemento manualmente](readme-images/manual-activation.PNG)

1. Haga clic en el botón **Enviar activación de complemento**. Se abre un cuadro de diálogo para enviar el mensaje: 

    ![Captura de pantalla del cuadro de diálogo de envío del mensaje](readme-images/send-message.PNG)
1. (Opcional): Modifique el contexto de inicialización.
1. Haga clic en el botón **Enviar mensaje**.
1. Cuando llegue el mensaje a su bandeja de entrada, ábralo.

    ![Captura de pantalla del mensaje que requiere acción enviado por el complemento](readme-images/actionable-message.PNG)
1. Haga clic en el botón **Invocar "Ver contexto de inicialización"**.
1. Se abre el panel de tareas del complemento y se muestra el contexto de inicialización.

    ![Captura de pantalla del panel de tareas abierto](readme-images/activated-taskpane.PNG)

### Pruebe la instalación a petición del complemento de la tienda

Puede usar el segundo botón del mensaje accionable para ver cómo funciona la instalación a petición de complementos de la tienda. Si ya tiene el [Analizador de encabezados de mensaje](https://appsource.microsoft.com/en-us/product/office/WA104005406), desinstálelo antes de probarlo.

Este proyecto ha adoptado el [Código de conducta de código abierto de Microsoft](https://opensource.microsoft.com/codeofconduct/). Para obtener más información, vea [Preguntas frecuentes sobre el código de conducta](https://opensource.microsoft.com/codeofconduct/faq/) o póngase en contacto con [opencode@microsoft.com](mailto:opencode@microsoft.com) si tiene otras preguntas o comentarios.

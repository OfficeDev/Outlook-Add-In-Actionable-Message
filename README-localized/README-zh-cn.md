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
# Outlook 加载项：可操作邮件激活

该示例演示如何[从可操作邮件中激活加载项](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)以及如何将初始化上下文传递给该加载项。

## 运行示例

为了运行此示例，需要将所包含的文件托管在 Web 服务器上。Web 服务器的选择完全取决于你。唯一的要求就是 Web 服务器必须受有效的 SSL 证书保护。 

### 更新清单

在加载加载项之前，必须更新[清单](actionable-add-in.xml)，以将 `localhost:8080` 的所有实例替换为托管该加载项的服务器的基本 URL。

### 安装加载项

按照[旁加载 Outlook 加载项以供测试](https://docs.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing)中的说明，旁加载 `actionable-add-in.xml` 文件以安装该加载项。

### 使用加载项

当你阅读电子邮件时，该加载项会将两个按钮添加到功能区。

- **发送加载项激活** \- 此按钮允许你通过用于调用加载项的按钮向自己发送可操作电子邮件。
- **查看初始化上下文** \- 此按钮会打开一个任务窗格，并显示当前的初始化上下文（如果存在）。

> **注意：**仅当从可操作邮件中调用任务窗格时，才存在初始化上下文。如果使用功能区按钮打开任务窗格，则会显示一条消息，告知你需要从邮件中调用加载项。
>
> ![手动激活加载项时显示的消息的屏幕截图](readme-images/manual-activation.PNG)

1. 单击“**发送加载项激活**”按钮。此时将打开用于发送邮件的对话框： 

    ![“发送邮件”对话框的屏幕截图](readme-images/send-message.PNG)
1. （可选）：修改初始化上下文。
1. 单击“**发送邮件**”按钮。
1. 当邮件到达你的收件箱时，打开它。

    ![由加载项发送的可操作邮件的屏幕截图](readme-images/actionable-message.PNG)
1. 单击“**调用查看初始化上下文**”按钮。
1. 加载项任务窗格将打开并显示初始化上下文。

    ![已打开的任务窗格的屏幕截图](readme-images/activated-taskpane.PNG)

### 试用按需安装的应用商店加载项

你可以使用可操作邮件中的第二个按钮来查看按需安装的应用商店加载项的工作方式。试用之前，如果你已拥有[邮件头分析器](https://appsource.microsoft.com/en-us/product/office/WA104005406)，请确保卸载它。

此项目已采用 [Microsoft 开放源代码行为准则](https://opensource.microsoft.com/codeofconduct/)。有关详细信息，请参阅[行为准则 FAQ](https://opensource.microsoft.com/codeofconduct/faq/)。如有其他任何问题或意见，也可联系 [opencode@microsoft.com](mailto:opencode@microsoft.com)。

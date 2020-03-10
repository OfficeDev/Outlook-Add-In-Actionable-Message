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
# Outlook アドイン:操作可能なメッセージのアクティブ化

このサンプルは、[操作可能なメッセージからアドインをアクティブ化](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)し、初期化コンテキストをアドインに渡す方法を示しています。

## サンプルの実行

このサンプルを実行するには、含まれているファイルを Web サーバーでホストする必要があります。Web サーバーの選択は、完全にユーザー次第です。唯一の要件は、Web サーバーが有効な SSL 証明書で保護されていることです。 

### マニフェストを更新する

アドインを読み込む前に、[マニフェスト](actionable-add-in.xml)を更新して、`localhost:8080` のすべてのインスタンスを、アドインをホストしているサーバーのベース URL に置き換える必要があります。

### アドインをインストールする

「[テスト用に Outlook アドインをサイドロードする](https://docs.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing)」の手順に従って、`actionable-add-in.xml` ファイルをサイドロードしてアドインをインストールします。

### アドインを使用する

メール メッセージを読んでいるとき、アドインはリボンに 2 つのボタンを追加します。

- **アドインのアクティブ化を送信する** \- このボタンを使用すると、アドインを起動するボタンを含む操作可能なメール メッセージを自分自身に送信できます。
- **初期化コンテキストを表示する** \- このボタンは作業ウィンドウを開き、現在の初期化コンテキストを表示します (存在する場合)。

> **注:**初期化コンテキストは、作業ウィンドウが操作可能なメッセージから呼び出された場合にのみ存在します。リボン ボタンを使用して作業ウィンドウを開くと、メッセージからアドインを呼び出す必要があることを知らせるメッセージが表示されます。
>
> ![アドインを手動でアクティブ化したときに表示されるメッセージのスクリーンショット](readme-images/manual-activation.PNG)

1. \[**アドインのアクティブ化を送信する**] ボタンをクリックします。メッセージを送信するダイアログが開きます。 

    ![メッセージ送信ダイアログのスクリーンショット](readme-images/send-message.PNG)
1. (オプション):初期化コンテキストを変更します。
1. \[**メッセージの送信**] ボタンをクリックします。
1. メッセージが受信トレイに届いたら、それを開きます。

    ![アドインによって送信された操作可能なメッセージのスクリーンショット](readme-images/actionable-message.PNG)
1. \[**"初期化コンテキストの表示" を呼び出す**] ボタンをクリックします。
1. アドインの作業ウィンドウが開き、初期化コンテキストが表示されます。

    ![開いている作業ウィンドウのスクリーンショット](readme-images/activated-taskpane.PNG)

### ストア アドインのオンデマンド インストールを試す

操作可能なメッセージの 2 番目のボタンを使用して、ストア アドインのオンデマンド インストールの動作を確認できます。試す前に、すでに[メッセージ ヘッダー アナライザー](https://appsource.microsoft.com/en-us/product/office/WA104005406)をお持ちの場合は、必ずアンインストールしてください。

このプロジェクトでは、[Microsoft Open Source Code of Conduct (Microsoft オープン ソース倫理規定)](https://opensource.microsoft.com/codeofconduct/) が採用されています。詳細については、「[倫理規定の FAQ](https://opensource.microsoft.com/codeofconduct/faq/)」を参照してください。また、その他の質問やコメントがあれば、[opencode@microsoft.com](mailto:opencode@microsoft.com) までお問い合わせください。

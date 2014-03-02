hMailServer-run-vbs-and-relay
=============================

Windows 用多機能メールサーバの [hMailServer](http://www.hmailserver.com/) を使って、
VBScript を実行しつつ本来の送信先に送る方法のメモです。


### 1. hMailServer をインストール

[Download](http://www.hmailserver.com/index.php?page=download) から、最新版をダウンロードして
インストールします。

インストーラは特に何も考えずにデフォルト設定でインストールしてください。

インストーラで管理パスワードの入力を求められますが、入力したパスワードは管理で必要になるので忘れないでください。

### 2. Events フォルダの権限変更

hMailServer は、インストールフォルダの配下にある ``Events`` ディレクトリ (64bit環境のデフォルトインストールの場合、 ``C:\Program Files (x86)\hMailServer\Events``) に VBScript ファイルを使用します。

このディレクトリはそのままだと管理者に昇格しないと書き込みできないため、スクリプトの修正が若干煩わしいので、
以下の手順で権限を修正することをおすすめします。

1. ``C:\Program Files (x86)\hMailServer`` ディレクトリを開く
1. ``Events`` ディレクトリを右クリックする
1. 「プロパティ」を開く
1. 「セキュリティ」タブを開く
1. 「編集」をクリックする
1. 「Users」を選択する
1. 「許可」の「フルコントロール」にチェックを入れて「OK」をクリックする

### 3. hMailServer Administrator を使って設定する

hMailServer Administrator を最初起動すると、以下の画面が表示されます。

![hMailServer Administrator 起動画面](https://raw.github.com/micclly/hMailServer-run-vbs-and-relay/master/images/admin01.png)

管理対象は1個しかありませんので、上の画像のとおり「Automatically connect on startup」にチェックをいれてください。

「Connect」をクリックすると、以下のパスワード入力画面が表示されるので、インストール時に設定した管理パスワードを入力してください。

![hMailServer Administrator パスワード入力画面](https://raw.github.com/micclly/hMailServer-run-vbs-and-relay/master/images/admin02.png)

「OK」をクリックして認証に成功すれば、以下の管理画面が表示されます。

![hMailServer Administrator 管理画面](https://raw.github.com/micclly/hMailServer-run-vbs-and-relay/master/images/admin03.png)

#### リレーサーバの設定

hMailServer で受け取ったメールを、外部の SMTP サーバにリレーするように設定します。

以下の画像のように、「Settings」－「Protocols」－「SMTP」ツリーを開いて、「Delivery of e-mail」タブを開きます。

![hMailServer Administrator 管理画面: Delivery of e-mails タブ](https://raw.github.com/micclly/hMailServer-run-vbs-and-relay/master/images/admin04.png)

「SMTP Relayer」から下を、通常使用している SMTP サーバの設定にしてください。

Yahoo! Japan の場合は、以下のようになります。

* Remote host name
    * ``smtp.mail.yahoo.co.jp``
* Remote TCP/IP Port
    * ``587``
* Server requires authentication
    * チェックON
* User name
    * SMTP ユーザ名
* Password
    * SMTP パスワード



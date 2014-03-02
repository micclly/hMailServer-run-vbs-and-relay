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

設定を変えたら、右下の「Save」ボタンをクリックして保存してください。

#### POP3, IMAP サーバの無効化

POP3, IMAP サーバは使わないので無効化しましょう。

以下の画像のように、「Settings」－「Protocols」ツリーを開いて、「POP3」と「IMAP」のチェックを外して「Save」ボタンをクリックします。

![hMailServer Administrator 管理画面: Protocols](https://raw.github.com/micclly/hMailServer-run-vbs-and-relay/master/images/admin05.png)


#### ロギングの設定

必須ではありませんが、うまくいかないときに原因追究する際、ログがあったほうがよいので、ロギングを有効にしましょう。

以下の画像のように、「Settings」－「Logging」ツリーを開いて設定し、「Save」ボタンをクリックします。

DEBUG ログはそれなりに多くのログが出るため、動作が確認できたら OFF にするとよいでしょう。

![hMailServer Administrator 管理画面: Logging](https://raw.github.com/micclly/hMailServer-run-vbs-and-relay/master/images/admin06.png)


#### SMTP 認証をしないように設定

デフォルトでは、

* ``From: <example@yahoo.co.jp>``
* ``To: <example@gmail.com>``

というように、From と To が自分のドメインではない送信の場合、認証が要求されます。

これを認証されないようにするため、「Settings」－「Advanced」－「IP Ranges」－「My computer」ツリーを開き、「Require SMTP Authentication」の列の一番下にある「External to external e-mail address」のチェックをオフにしてください。

![hMailServer Administrator 管理画面: Disable SMTP auth](https://raw.github.com/micclly/hMailServer-run-vbs-and-relay/master/images/admin07.png)

#### VBScript を配置

[EventHandlers.vbs](https://raw.github.com/micclly/hMailServer-run-vbs-and-relay/master/Events/EventHandlers.vbs) を ``C:\Program Files (x86)\hMailServer\Events\EventHandlers.vbs`` に上書きしてください。

この VBScript には ``ReceiveFromMT4`` というプロシージャを追加定義していて、受信したメールの本文を単純に ``C:\Windows\Temp\mail-from-mt4.txt`` に書き出すようになっています。

コマンドを実行したいような場合は、 ``objShell.Run`` を呼び出すように書き換えてください。

ただし、 GUI が表示されるコマンドは実行できませんので注意してください。

#### VBScript 有効化、VBScript を読み込み

デフォルトでは VBScript の実行が有効になっていないため、「Settings」－「Advanced」ー「Scripts」を開き、「Enabled」のチェックをONにしてください。

![hMailServer Administrator 管理画面: Scripts](https://raw.github.com/micclly/hMailServer-run-vbs-and-relay/master/images/admin08.png)

また、hMailServer は VBScript をメモリにキャッシュするようになっていて、自動では再読み込みされません。

ファイルを修正したら、必ず上記画面で「Reload scripts」と「Check syntax」してください。

VBScript にシンタックスエラーがある場合は、「Check syntax」でエラーダイアログが表示されます。



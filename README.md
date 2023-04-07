# BroadcastEmail

メールの一斉送信を行えるアプリケーションです。

# 実行環境
Python 3.11.2以降
`tcl-tk`パッケージがインストール済であること(`tkinter`ライブラリが利用可能であること)

# 使い方

## CSVファイルを用意する

SMTPサーバ情報、送信元情報、送信先情報をCSV形式で用意します。
CSVファイルには1行目から以下の順にデータを記載してください。

> SMTPサーバアドレス, SMTPサーバのポート番号, 認証用メールアドレス(アカウントのメールアドレス), パスワード
> 送信元メールアドレス
> 会社名1, 顧客名1, 送信先メールアドレス(複数ある場合は**半角カンマ(,)**区切り)
> 会社名2, 顧客名2, 送信先メールアドレス(同上)
> 会社名3, 顧客名3, 送信先メールアドレス(同上)
> ... *以降繰り返し*

## TXTファイルを用意する

件名、添付ファイルのフルパス、メッセージ内容をTXT形式で用意します。
TXTファイルには1行目から以下の順にデータを記載してください。

> 件名
> 添付ファイルの絶対パス(複数ある場合は**半角カンマ(,)**区切り, 拡張子名も記載)
> 内容(複数行記載しても良い)

## 実行環境を構築する

① Homebrewをインストール
② Python3をHomebrew経由でインストール
→ ターミナルで`python3 -V`を実行し、バージョン情報が出力されることを確認
③ `tcl-tk`パッケージをHomebrew経由でインストール
→ ターミナルで以下のコマンドを実行し、tkinterのダイアログが表示されることを確認

```sh
python3 -m tkinter
```

④ ターミナルからPythonプロジェクト(module)の実行

```sh
python3 -m broadcastemail
```
# VBA-EventBackup
# イベント機能活用による自動バックアップ用VBA

- License: The MIT license

- Copyright (c) 2021 YujiFukami

- 開発テスト環境 Excel: Microsoft® Excel® 2019 32bit 

- 開発テスト環境 OS: Windows 10 Pro

その他、実行環境など報告していただくと感謝感激雨霰。

# 使い方

## 「実行サンプル 自動バックアップ.xlsm」の使い方

「実行サンプル 自動バックアップ.xlsm」は保存時のイベントプロシージャ「Workbook_BeforeSave」実行時に自分のブックを自動的にバックアップするプロシージャの実行サンプルである。


## 設定

実行サンプル「実行サンプル 自動バックアップ.xlsm」の中の設定は以下の通り。

### 設定1（使用モジュール）

-  ModEvemtBackup.bas

### 設定2（参照ライブラリ）

-  「Microsoft Scripting Runtime」 FileSystemObjectを使用するため

![参照ライブラリ](https://user-images.githubusercontent.com/73621859/130732662-861cbc29-ef1f-46e9-ac3c-0f53db1ce02c.jpg)

### 設定3 (イベントプロシージャ設定)

　ワークブック保存直前時イベント（Workbook_BeforeSave）のコード内に、バックアップ実行用のプロシージャを実行するように設定する。

![イベント設定](https://user-images.githubusercontent.com/73621859/130732631-f650ea95-185c-40a1-b70c-3cf74000fed0.jpg)

　これにより、ワークブック保存時に自動的にバックアップが生成されるようになる。

![バックアップ状況](https://user-images.githubusercontent.com/73621859/130732651-52e87c28-1167-4678-8e59-4869f2aedbc6.jpg)


## 現在「Dictionary.bas」にて使用できるプロシージャ一覧

-  ワークブック保存時にフォルダに上書きバックアップ
-  ワークブック保存時にフォルダに日付をつけてバックアップ
-  ワークブック保存時に同じフォルダ上に上書きバックアップ
-  ワークブック保存時に同じフォルダ上に日付をつけて上書きバックアップ

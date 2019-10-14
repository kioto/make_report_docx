# JSONファイルからdocxのレポートを生成するスクリプト

JSONファイルを元に、Word文書(docx)を生成する。
Word文書は以下の情報を持つ。

* 表紙ページ
  * ドキュメントタイトル
  * 日付
  * 作成者名
* 概要ページ
  * 概要ページタイトル
  * 本文
* レポートページ
  * 対象サンプル概要
  * 対象サンプルデータ

詳細はJSONファイルを参照。

## 2. 動作環境

以下の環境で開発、動作確認した。

* Python 3.7
* python-docx 0.8.10
* MacOS 10.15で動作確認


## 3. インストール

### 3-1. python-docx

以下の手順でインストールする。

```
pip3 install python-docx
```

# 4. 使用方法

使用方法を以下に示す。

```
python3 MakeDocxReport.py <JSON_FILE>
```

サンプルのJSONファイルから、レポートdocxファイルの生成手順をいかに示す。

``` bash
$ python3 MakeDocxReport.py sample/report.json
$ ls *.docx
report.docx
```

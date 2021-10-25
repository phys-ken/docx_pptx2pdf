# docx_pptx2pdf.py

## できること
* ワードで作ったファイルと、パワポのファイルを、PDFに一括変換します。

## 必要な環境
* windowsでしか動きません。
* wordやpower pointがインストールされている必要があります。
* 以下のライブラリを使用します。

```Python

import shutil
import os
import PyPDF2
import glob
import os
from docx2pdf import convert
import sys
import re
import comtypes.client

```


## 使い方
* docx_pptx2pdf.pyと同じ位置にinputfを作成します。
* その中に、変換したいdocxとpptxを保存します
  * フォルダ構造を持っていても、中まで潜って処理を行ってくれます。

# 個人用メモ
* inputfの有無で処理を分ける(エラー処理)
* tqdmでプログレスバーとして表示する。




---

# ここより下の文章は、oldフォルダ内のコードについての説明です。


## PDFを結合するexeファイルを作成しました！
* pyinstallerで、pdf_controler.pyをexeファイルに変更しました！
  * word2pdfは内部でwordを呼び出しているため、pyinstaller対象外です。
  * その代わりに、word2pdf.xlsmというマクロ付きエクセルファイルをdistのフォルダに入れました。

* 使い方
  * word2pdf.xlsmからwordファイルをpdfに変換
  * pdf_controler.exeで、pdfファイルを結合




### メモ
```
pip install pypdf2
```

だと、パスが通らないので、

```
pip3.9 install pupdf2
```

でインストールする必要あり

--- 
### 旧データ

* できること
  * WordをPDFに出力します。
  * PDFを結合します。

* 使い方
  * ワードが入ったフォルダのパスを指定して、実行するのみ
  * os.mkdirで、勝手にフォルダを作ってくれます。

* 注意点
  * エラー回避はしていません。必ずパスを指定してから実行してください。
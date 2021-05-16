# PDFを扱うノートブック

## exeファイルを作成しました！
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
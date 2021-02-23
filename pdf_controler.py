import PyPDF2
import glob
import os
from docx2pdf import convert

### まずは、ここでパスを指定します。

# 変換する前のワードファイルがあるフォルダ
input_dir = input("ワードファイルが入っているフォルダのパスをフルパスで入力してください。>>")

# 変換したPDFファイルを保存するフォルダ
# デフォルトは input_dir + PDF
output_dir = input_dir + "/PDF"
os.mkdir(output_dir)

# 変換したPDFを結合して、保存するファイル名
output_file =  input_dir + "/PDF/000.marge.pdf"

#このセルでは、input_dirのwordファイルを、全てpdfに変換します。



"""
use library: https://github.com/AlJohri/docx2pdf
"""

def convert_pdf(input_dir="output/", output_dir="output_pdf/"):
    """
    docxファイルの保存されたフォルダを指定して、フォルダ格納データを全てpdfにして、指定フォルダに保存する
    :param input_dir: dir_name, default:output/,  outputフォルダを利用
    :param output_dir: dir_name, default:output_pdf/, output_pdfフォルダを利用
    :return: output_pdfフォルダにoutputフォルダのpdfが全て保存される
    """
    convert(input_dir, output_dir)

    
convert_pdf(input_dir, output_dir)

print("PDF化が完了しました！")

#このセルでは、output_dirのpdfファイルを、全てoutput_fileに変換します。

import PyPDF2
import glob
import os

def merge_pdf_in_dir(dir_path, dst_path):
    l = glob.glob(os.path.join(dir_path, '*.pdf'))
    l.sort()

    merger = PyPDF2.PdfFileMerger()
    for p in l:
        if not PyPDF2.PdfFileReader(p).isEncrypted:
            merger.append(p)

    merger.write(dst_path)
    merger.close()

    
# インプットパスとアウトプットファイル名を指定する。
#https://note.nkmk.me/python-pypdf2-pdf-merge-insert-split/
merge_pdf_in_dir(output_dir, output_file)

print("PDFの結合が完了しました！")
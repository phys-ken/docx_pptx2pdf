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


### まずは、ここでパスを指定します。

# 変換する前のワードファイルがあるフォルダ
input_dir = input("ワードファイルが入っているフォルダのパスをフルパスで入力してください。>>")

# 変換したPDFファイルを保存するフォルダ
# デフォルトは input_dir + PDF
output_dir = input_dir + "/merge"
os.mkdir(output_dir)

# 変換したPDFを結合して、保存するファイル名
output_file =  input_dir + "/merge/000.merge.pdf"

# インプットパスとアウトプットファイル名を指定する。
#https://note.nkmk.me/python-pypdf2-pdf-merge-insert-split/
merge_pdf_in_dir(input_dir, output_file)



print("PDFの結合が完了しました！")
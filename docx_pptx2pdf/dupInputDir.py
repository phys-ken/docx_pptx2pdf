import shutil
import os
import PyPDF2
import glob
import os
from docx2pdf import convert
import sys
import re
import comtypes.client

def yes_no_input():
    while True:
        choice = input("Please respond with 'yes' or 'no' [y/N]: ").lower()
        if choice in ['y', 'ye', 'yes']:
            return True
        elif choice in ['n', 'no']:
            return False

def merge_pdf_in_dir(dir_path, dst_path):
    getl = glob.glob(os.path.join(dir_path, '*.pdf'))
    if len(getl) == 0:
      return 0

    indir , gomi =  os.path.split(getl[0])
    dirs = []
    flNms = []
    
    for p in getl:
      dir , flNm = os.path.split(p)
      dirs.append(dir)
      flNms.append(flNm)

    print("並び替え前のファイル順>>>")
    print(flNms)
    print("sort l")
    tmp1 = []
    for fnstr in flNms:
      if re.search(r'\d+', fnstr) == None:
        os.rename(dir + "\\" + fnstr , dir + "\\000" + fnstr)
        tmp1.append("000" + fnstr)
      else:
        tmp1.append(fnstr)
    tmp2 = []
    tmp2 = sorted(tmp1, key=lambda s: int(re.search(r'\d+', s).group()))
    print("並び替え前のファイル後>>>")
    print(tmp2)

    sortl = []
    for p in tmp2:
      sortl.append(indir + "\\" +p)

    l = sortl

    if len(l) == 0:
      pdfFlag = False
    else:
      pdfFlag = True

    if pdfFlag:
      merger = PyPDF2.PdfFileMerger()
      for p in l:
          if not PyPDF2.PdfFileReader(p).isEncrypted:
              merger.append(p)

      merger.write(dst_path)
      merger.close()


def convert_pdf(input_dir="output/", output_dir="output_pdf/"):
    """
    use library: https://github.com/AlJohri/docx2pdf
    """
    """
    docxファイルの保存されたフォルダを指定して、フォルダ格納データを全てpdfにして、指定フォルダに保存する
    :param input_dir: dir_name, default:output/,  outputフォルダを利用
    :param output_dir: dir_name, default:output_pdf/, output_pdfフォルダを利用
    :return: output_pdfフォルダにoutputフォルダのpdfが全て保存される
    """
    convert(input_dir, output_dir)

def pptx2pdf(input_folder_path , output_folder_path):
  #%% Convert folder paths to Windows format
  input_folder_path = os.path.abspath(input_folder_path)
  output_folder_path = os.path.abspath(output_folder_path)

  #%% Get files in input folder
  input_file_paths = os.listdir(input_folder_path)

  #%% Convert each file
  for input_file_name in input_file_paths:

      # Skip if file does not contain a power point extension
      if not input_file_name.lower().endswith((".ppt", ".pptx")):
          continue
      
      # Create input file path
      input_file_path = os.path.join(input_folder_path, input_file_name)
          
      # Create powerpoint application object
      powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
      
      # Set visibility to minimize
      powerpoint.Visible = 1
      
      # Open the powerpoint slides
      slides = powerpoint.Presentations.Open(input_file_path)
      
      # Get base file name
      file_name = os.path.splitext(input_file_name)[0]
      
      # Create output file path
      output_file_path = os.path.join(output_folder_path, file_name + ".pdf")
      
      # Save as PDF (formatType = 32)
      slides.SaveAs(output_file_path, 32)
      
      # Close the slide deck
      slides.Close()






inputDir = "./inputf"

files = os.listdir(inputDir)

print("すでにあるoutputf内のデータは削除されます。よろしいですか？")
if yes_no_input():
  pass
else:
  print("処理を中断します")
  sys.exit()

shutil.rmtree('./outputf')

for curDir, dirs, files in os.walk(inputDir):
    for dir in dirs:
        outputDir = "outputf\\" + curDir + "\\" + dir
        print(curDir + "\\" + dir + "__の中身を処理中...")

        try : 
          os.makedirs(outputDir, exist_ok=True)
          convert_pdf(curDir + "\\" + dir, outputDir)

          os.makedirs(outputDir + "\\slids", exist_ok=True)
          merge_pdf_in_dir(outputDir, outputDir + "\_Marge.pdf")

          getpptx = glob.glob(os.path.join(curDir + "\\" + dir, '*.pptx'))
          print(getpptx)
          if len(getpptx) >= 1 :
            pptx2pdf(curDir + "\\" + dir,outputDir + "\\slids" )
            merge_pdf_in_dir(outputDir + "\\slids", outputDir + "\\slids" + "\_Marge.pdf")
        except:
          print("nannjakorya?????????????????????????????????????")


print("処理が終了しました。")
from win32com.client import constants, gencache
import os
import glob


# Word转pdf方法,第一个参数代表word文档路径，第二个参数代表pdf文档路径
def Word_to_Pdf(Word_path, Pdf_path):
    word = gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(Word_path, ReadOnly=1)
    # 转换方法
    doc.ExportAsFixedFormat(Pdf_path, constants.wdExportFormatPDF)
    word.Quit()


if __name__ == '__main__':
    Word_files = []
    for file in glob.glob('D:/Destop/办公/办公/data/1word/机器人（SI）21-2/实验4/*'):        # TODO
        # 1.找出所有后缀为doc或者docx的文件
        if file.endswith(('.doc', '.docx')):
            Word_files.append(file)
    print(Word_files)

    # 2.转换word to pdf
    a = 1
    for file in Word_files:
        # os.path.abspath 返回绝对路径
        file_path = os.path.abspath(file)

        pdf_path = 'D:/Destop/办公/办公/data/2pdf' + "/" + file_path.split("\\")[-1] + '.pdf'      # TODO
        pdf_path = os.path.abspath(pdf_path)

        Word_to_Pdf(file_path, pdf_path)
        print(str(a), pdf_path)
        a += 1

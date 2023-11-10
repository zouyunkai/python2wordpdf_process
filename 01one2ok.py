import glob
import PyPDF2
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib import colors
import random
import PyPDF2
import glob
from win32com.client import constants, gencache
import os
import glob
import shutil
import os


from win32com import client
# 转换doc为pdf
def Word_to_Pdf_back(Word_path, Pdf_path):
    word = client.Dispatch("Word.Application")  # 打开word应用程序
    # for file in files:
    doc = word.Documents.Open(Word_path)  # 打开word文件
    doc.SaveAs(Pdf_path, 17)  # 另存为后缀为".pdf"的文件，其中参数17表示为pdf
    doc.Close()  # 关闭原来word文件
    word.Quit()


# Word转pdf方法,第一个参数代表word文档路径，第二个参数代表pdf文档路径
def Word_to_Pdf(Word_path, Pdf_path):
    word = gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(Word_path, ReadOnly=1)
    # 转换方法
    doc.ExportAsFixedFormat(Pdf_path, constants.wdExportFormatPDF)
    doc.Close()
    word.Quit()


# pdf合并
def add_watermark_qianming_pingyu(water_file, page_pdf):
    pdfReader = PyPDF2.PdfFileReader(water_file)      # 读取电子签章的pdf内容
    x = int(float(page_pdf.mediaBox.getWidth()) - float(pdfReader.getPage(0).mediaBox.getWidth())*0.7) - 100
    y = 80
    page_pdf.mergeScaledTranslatedPage(pdfReader.getPage(0), 0.7, x, y)     # TODO, 设置图片位置
    return page_pdf

# 添加签名
# def pdf_add_pdf(sourcepath, outputpath, pic_pdf_path):
#     """
#     :param sourcepath: 原PDF路径
#     :param outputpath: 新PDF路径
#     :param pic_pdf_path: 需要加入原PDF的图片PDF路径
#     :return:
#     """

# 加评语
def add_watermark_pingyu(pdfReader,water_file):    
  # 遍历pdf的每一页,在某页添加图片, TODO:某页
    for page in range(pdfReader.numPages):
        if page in [pdfReader.numPages-1]:
            page_pdf = add_watermark_qianming_pingyu(water_file, pdfReader.getPage(page))
        else:
            page_pdf = pdfReader.getPage(page)
    return page_pdf
# pdf all
def add_watermark(   page_pdf, file_stu_score):
    """
    :param water_file: 水印
    :param page_pdf: 原pdf的某页
    :param file_stu_score: 分数
    :return: 写入分数，并加评语
    """
    # 最后一页添加签名

    # # 加评语
    # waterReader = PyPDF2.PdfFileReader(water_file)  # 读取电子签章的pdf内容
    x = int(float(page_pdf.mediaBox.getWidth()) ) - 30
    # page_pdf.mergeScaledTranslatedPage(waterReader.getPage(0), 0.3,
    #        x,
    #        int(page_pdf.mediaBox.getHeight())-int(int(waterReader.getPage(0).mediaBox.getHeight())*0.3)-150 )  # TODO, 设置图片位置

    # 写分数, 1.使用reportlab先生成分数pdf，2.读取分数pdf，合并入page_pdf
    # 1.
    inch = 72.0
    score_pdf_file = "data/pic/score/score.pdf"  # 设置输出的PDF文件名
    c = canvas.Canvas(score_pdf_file, pagesize=(0.32 * inch, 0.5 * inch))
    c.setFillColor(colors.red)
    c.setFont("Times-Roman", 18)
    c.drawString(1, 1, str(file_stu_score))
    c.save()
    # 2.
    scoreReader = PyPDF2.PdfFileReader(score_pdf_file)
    page_pdf.mergeScaledTranslatedPage(scoreReader.getPage(0), 1,
                                       x-int(scoreReader.getPage(0).mediaBox.getWidth())-50,
                                       int(page_pdf.mediaBox.getHeight()) - int(
                                           int(scoreReader.getPage(0).mediaBox.getHeight())) - 30)
    return page_pdf

# 添加分数，电子签，评语
def pdf_add_pdf_all(sourcepath, outputpath, pic_pdf_path, file_stu_score):
    """
    :param sourcepath: 原PDF路径
    :param outputpath: 新PDF路径
    :param pic_pdf_path: 需要加入原PDF的图片PDF路径
    :param file_stu_score: 分数
    :return:
    """
    pdfWriter = PyPDF2.PdfFileWriter()
    pdfReader = PyPDF2.PdfFileReader(sourcepath)

    # 加分数，遍历pdf的每一页,在某页添加图片 
    for page in range(pdfReader.numPages):
        if page in [0]:
            page_pdf = add_watermark(   pdfReader.getPage(page), file_stu_score)
        else:
            page_pdf = pdfReader.getPage(page)
        pdfWriter.addPage(page_pdf)
    with open(sourcepath, 'wb') as target_file:
        pdfWriter.write(target_file)
        # print(sourcepath)
    # 加评语和签名
    pdfWriter = PyPDF2.PdfFileWriter()
    pdfReader = PyPDF2.PdfFileReader(sourcepath)
    for page in range(pdfReader.numPages):
        if page in [pdfReader.numPages-1]:
            page_pdf = add_watermark_qianming_pingyu(pic_pdf_path, pdfReader.getPage(page))
        else:
            page_pdf = pdfReader.getPage(page)
        pdfWriter.addPage(page_pdf)
    
    with open(sourcepath, 'wb') as target_file:
        pdfWriter.write(target_file)
    shutil.copy(sourcepath,outputpath)
    print('ok '+sourcepath)



def process(score_col_index,excel_path,sheet_name,temp_path_shiyan):
    # 参数修改汇总
    # score_col_index=3表示读取第四列，score_col_index=0表示读取整个表格的第1列
    # 默认第一行的"序号"	学号	姓名	实验1.... 不读取，随机查一个同学的实验成绩对不对就行
    num=score_col_index+1
    num=str(num)
    df = pd.read_excel(excel_path, sheet_name )      # 读excel成绩表,sheet_name=0表示读取第一个工作簿，=1表示读取第二个工作簿
    files_to_process=  'D:/Destop/办公/办公/data/1word'+'/'+temp_path_shiyan+num+'/*'  
    files_to_process=os.path.abspath(files_to_process)
    
    # 抽出来的实验存放的位置
    output_ans_path =temp_path_shiyan+num+'\\'  # 最终的结果存放整个结果，按照办公\data\4pdf_sign_score\班级\实验4的目录格式存放的
    # temppath=' '
    # basedir 目录,修改成办公/data下面的,最后要加个反斜杠和后面的拼起来
    basedir = 'D:/Destop/办公/办公/data/'
    dirpath_list = [
        'D:/Destop/办公/办公/data/2pdf',
        # 'D:/Destop/办公/办公/data/3pdf_sign',
    ]

    # 先清空data下除了1word和4pdf_sgin_score的其他目录

    

    for dirpath in dirpath_list:
        shutil.rmtree(dirpath)
        os.mkdir(dirpath)

    ###########################################################################
    # 第一步，将word转成pdf，或者使用powershell来做
    # Word_files = []
    # for file in glob.glob(files_to_process):        # TODO
    #     # 1.找出所有后缀为doc或者docx的文件
    #     if file.endswith(('.doc', '.docx')):
    #         Word_files.append(file)
    # # print(Word_files)
    # if os.path.exists(basedir+'3pdf_sign'+temppath ) :
    #     shutil.rmtree(basedir+'3pdf_sign'+temppath )
    #     os.mkdir(basedir+'3pdf_sign'+temppath )
    # else: 
    #     os.mkdir(basedir+'3pdf_sign'+temppath )
    # # 2.转换word to pdf
    # # a = 1
    # for file in Word_files:
    #     # os.path.abspath 返回绝对路径
    #     file_path = os.path.abspath(file)

    #     pdf_path = basedir+'3pdf_sign'+temppath + "/" + (file_path.split("\\")[-1]).split('.')[-2] + '.pdf'      # TODO
    #     pdf_path = os.path.abspath(pdf_path)

    #     Word_to_Pdf(file_path, pdf_path)
        # print(str(a), pdf_path)
        # a += 1

    # 第二步，打签名
    # pdf_files = []
    # for file in glob.glob(basedir+'2pdf/*'):  # TODO
    #     if file.endswith('.pdf'):
    #         pdf_files.append(file)
    # print(pdf_files)

    # for file in pdf_files:
    #     sourcepath = file  # 原始pdf位置
    #     outputpath = basedir+"3pdf_sign/" + file.split("\\")[-1]  # TODO, 新生成pdf位置
    #     pic_pdf_path = basedir+"pic/score/sign.pdf"                    # TODO, sign的pdf位置
    #     pdf_add_pdf(sourcepath, outputpath, pic_pdf_path)


    # 第三步，打分数
    pdf_files = []
    # for file in glob.glob(basedir+'3pdf_sign'+temppath+'/*'):  # TODO
    for file in glob.glob(files_to_process):     
        if file.endswith('.pdf'):
            pdf_files.append(file)
    # print(pdf_files)

    # 读excel
    # df = pd.read_excel('D:/Destop/办公/办公/data/excel/2122-1+大学计算机基础+赵晶+机器人21-1、机器人（SI）21-12、电竞21-1/2122-1实验+大作业成绩-机器人.xlsx')   # TODO, 成绩表位置
    # print(int(df.iloc[1][6]))   # 第一列是学号，第4列是分数（从0开始）
    file_stu_score =80
    for file in pdf_files:
        # 获取当前文件的分数
        file_stu_id = file.split("\\")[-1].split(" ")[0]    # 当前文件的学号
        for i in range(1, len(df)):
            if df.iloc[i][1] is not None and float(file_stu_id) == float(df.iloc[i][1]):
                # score_col_index = 4
                if df.iloc[i][score_col_index] == '未批' or df.iloc[i][score_col_index] == '未交' or df.iloc[i][score_col_index] == '':
                    file_stu_score = 0
                else:
                    file_stu_score = int(df.iloc[i][score_col_index])
                file_stu_name = df.iloc[i][2]
                break

        sourcepath = file  # 原始pdf位置
        # outputpath = "data/4pdf_sign_score/" + str(file_stu_id) + "_" + str(file_stu_name) + ".pdf"  # TODO, 新生成pdf位置
        # outputpath = basedir+"3pdf_sign/" + file.split("\\")[-1]  # TODO, 新生成pdf位置
        outputpath = basedir+"4pdf_sign_score/"+output_ans_path   # TODO, 新生成pdf位置

        # 随机选择每一类等级中的某一个评分pdf
        if file_stu_score >= 93:    # 优秀
            pic_pdf_path = basedir+"pic/score/A1.pdf"               # TODO, score的pdf位置
        elif file_stu_score >= 90:  # 良好
            pic_pdf_path = basedir+"pic/score/A2.pdf"
        elif file_stu_score >= 87:  # 良好
            pic_pdf_path = basedir+"pic/score/A3.pdf" 
        elif file_stu_score >= 84:  # 良好
            pic_pdf_path = basedir+"pic/score/B1.pdf" 
        elif file_stu_score >= 80:  # 良好
            pic_pdf_path = basedir+"pic/score/B2.pdf" 
        elif file_stu_score >= 75:  # 良好
            pic_pdf_path = basedir+"pic/score/C1.pdf" 
        else:
            pic_pdf_path = basedir+"pic/score/C2.pdf"  
        # 将反斜杠转换成斜杠形式的地址
        outputpath = os.path.abspath(outputpath)
        if not os.path.exists(outputpath ) :
            os . makedirs(outputpath)
        pdf_add_pdf_all(sourcepath, outputpath, pic_pdf_path, file_stu_score)


if __name__ == '__main__':
    excel_path=r'D:\Destop\办公\办公\data\1word\2122-1实验+大作业成绩-机器人.xlsx'
    sheet_name=3 # 读取excel成绩表中的第几个工作簿，0表示第一个
    temp_path_shiyan='电竞21-1\实验' #生成结果的地址前缀和读取数据的前缀
    process(3,excel_path,sheet_name,temp_path_shiyan)
    process(4,excel_path,sheet_name,temp_path_shiyan)
    process(5,excel_path,sheet_name,temp_path_shiyan)
    process(6,excel_path,sheet_name,temp_path_shiyan)

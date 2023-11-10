import glob
import PyPDF2
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib import colors
import random


# pdf合并
def add_watermark(water_file, page_pdf, file_stu_score):
    """
    :param water_file: 水印
    :param page_pdf: 原pdf的某页
    :param file_stu_score: 分数
    :return: 写入分数，并加评语
    """
    # 加评语
    waterReader = PyPDF2.PdfFileReader(water_file)  # 读取电子签章的pdf内容
    x = int(float(page_pdf.mediaBox.getWidth()) - float(waterReader.getPage(0).mediaBox.getWidth())*0.3) - 10
    page_pdf.mergeScaledTranslatedPage(waterReader.getPage(0), 0.3,
           x,
           int(page_pdf.mediaBox.getHeight())-int(int(waterReader.getPage(0).mediaBox.getHeight())*0.3)-150 )  # TODO, 设置图片位置

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
                                       x-int(scoreReader.getPage(0).mediaBox.getWidth())-10,
                                       int(page_pdf.mediaBox.getHeight()) - int(
                                           int(scoreReader.getPage(0).mediaBox.getHeight())) - 132)
    return page_pdf


def pdf_add_pdf(sourcepath, outputpath, pic_pdf_path, file_stu_score):
    """
    :param sourcepath: 原PDF路径
    :param outputpath: 新PDF路径
    :param pic_pdf_path: 需要加入原PDF的图片PDF路径
    :param file_stu_score: 分数
    :return:
    """
    pdfWriter = PyPDF2.PdfFileWriter()
    pdfReader = PyPDF2.PdfFileReader(sourcepath)

    # 遍历pdf的每一页,在某页添加图片, TODO:某页
    for page in range(pdfReader.numPages):
        if page in [0]:
            page_pdf = add_watermark(pic_pdf_path, pdfReader.getPage(page), file_stu_score)
        else:
            page_pdf = pdfReader.getPage(page)
        pdfWriter.addPage(page_pdf)

    with open(outputpath, 'wb') as target_file:
        pdfWriter.write(target_file)
        print(outputpath)


if __name__ == '__main__':
    # 多个文件的转换
    pdf_files = []
    for file in glob.glob('D:/Destop/办公/办公/data/3pdf_sign/*'):  # TODO
        if file.endswith('.pdf'):
            pdf_files.append(file)
    print(pdf_files)

    # 读excel
    df = pd.read_excel('D:/Destop/办公/办公/data/excel/2122-1+大学计算机基础+赵晶+机器人21-1、机器人（SI）21-12、电竞21-1/2122-1实验+大作业成绩-机器人.xlsx')   # TODO, 成绩表位置
    print(int(df.iloc[1][6]))   # 第一列是学号，第4列是分数（从0开始）

    for file in pdf_files:
        # 获取当前文件的分数
        file_stu_id = file.split("\\")[-1].split(" ")[0]    # 当前文件的学号
        for i in range(1, len(df)):
            if df.iloc[i][1] is not None and float(file_stu_id) == float(df.iloc[i][1]):
                score_col_index = 4
                if df.iloc[i][score_col_index] == '未批' or df.iloc[i][score_col_index] == '未交':
                    file_stu_score = 80
                else:
                    file_stu_score = int(df.iloc[i][score_col_index])
                file_stu_name = df.iloc[i][2]
                break

        sourcepath = file  # 原始pdf位置
        # outputpath = "data/4pdf_sign_score/" + str(file_stu_id) + "_" + str(file_stu_name) + ".pdf"  # TODO, 新生成pdf位置
        outputpath = "D:/Destop/办公/办公/data/4pdf_sign_score/" + file.split("\\")[-1]  # TODO, 新生成pdf位置

        # 随机选择每一类等级中的某一个评分pdf
        if file_stu_score >= 85:    # 优秀
            pic_pdf_path = "D:/Destop/办公/办公/data/pic/score/A" + str(random.randint(1, 3)) + ".pdf"               # TODO, score的pdf位置
        elif file_stu_score >= 75:  # 良好
            pic_pdf_path = "D:/Destop/办公/办公/data/pic/score/B" + str(random.randint(1, 6)) + ".pdf"
        else:
            pic_pdf_path = "D:/Destop/办公/办公/data/pic/score/C" + str(random.randint(1, 3)) + ".pdf"

        pdf_add_pdf(sourcepath, outputpath, pic_pdf_path, file_stu_score)

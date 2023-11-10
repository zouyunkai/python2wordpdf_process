import PyPDF2
import glob


# pdf合并
def add_watermark(water_file, page_pdf):
    pdfReader = PyPDF2.PdfFileReader(water_file)      # 读取电子签章的pdf内容
    x = int(float(page_pdf.mediaBox.getWidth()) - float(pdfReader.getPage(0).mediaBox.getWidth())*0.7) - 100
    y = 80
    page_pdf.mergeScaledTranslatedPage(pdfReader.getPage(0), 0.7, x, y)     # TODO, 设置图片位置
    return page_pdf


def pdf_add_pdf(sourcepath, outputpath, pic_pdf_path):
    """
    :param sourcepath: 原PDF路径
    :param outputpath: 新PDF路径
    :param pic_pdf_path: 需要加入原PDF的图片PDF路径
    :return:
    """
    pdfWriter = PyPDF2.PdfFileWriter()
    pdfReader = PyPDF2.PdfFileReader(sourcepath)

    # 遍历pdf的每一页,在某页添加图片, TODO:某页
    for page in range(pdfReader.numPages):
        if page in [pdfReader.numPages-1]:
            page_pdf = add_watermark(pic_pdf_path, pdfReader.getPage(page))
        else:
            page_pdf = pdfReader.getPage(page)
        pdfWriter.addPage(page_pdf)

    with open(outputpath, 'wb') as target_file:
        pdfWriter.write(target_file)
        print(outputpath)


if __name__ == '__main__':
    # 多个文件的转换
    pdf_files = []
    for file in glob.glob('C:/Users/yunkai/Desktop/办公/办公/data/2pdf/*'):  # TODO
        if file.endswith('.pdf'):
            pdf_files.append(file)
    print(pdf_files)

    for file in pdf_files:
        sourcepath = file  # 原始pdf位置
        outputpath = "C:/Users/yunkai/Desktop/办公/办公/data/3pdf_sign/" + file.split("\\")[-1]  # TODO, 新生成pdf位置
        pic_pdf_path = "C:/Users/yunkai/Desktop/办公/办公/data/pic/score/sign.pdf"                    # TODO, sign的pdf位置
        pdf_add_pdf(sourcepath, outputpath, pic_pdf_path)


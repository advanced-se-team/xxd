import os.path
from time import sleep

import convertapi
from docx import Document
from loguru import logger
from docx.oxml import OxmlElement

def extract_text(dox_dir: str):
    # 打开Word文档
    doc = Document(dox_dir)
    navigationBra = 1
    authorContent = 0
    abstract = 0
    paperDict = {}  # 保存的内容为'contextType' = 'content'
    pre_paragraph = doc.paragraphs[0]
    # 遍历文档中的段落并打印文本内容
    for paragraph in doc.paragraphs:
        tmp = paragraph.text.replace(" ", "").replace("\n", "")
        if not tmp:
            continue
        # 获取段落的样式
        style = paragraph.style.name
        for run in paragraph.runs:
            # 获取文本的具体属性
            font_name = run.font.name  # 字体名称
            font_size = run.font.size  # 字体大小
            bold = run.font.bold  # 是否加粗
            italic = run.font.italic  # 是否斜体
            underline = run.font.underline  # 是否有下划线
            # 获取段落的对齐方式
            alignment = paragraph.alignment
            # 根据对齐方式输出文字对齐方式
            # if alignment == 0:
            #     #print("左对齐")
            # elif alignment == 1:
            #     #print("居中对齐")
            # elif alignment == 2:
            #     #print("右对齐")
            # elif alignment == 3:
            #     #print("两端对齐")
            # else:
            # print("未知对齐方式")
            break
        if style == 'Title':
            contentType = 'Title'
            content = paragraph.text
            navigationBra = 0
            authorContent = 1
            paperDict.update({contentType: content})

        elif navigationBra == 1:
            contentType = 'Navigation Bra'
            content = paragraph.text
            navigationBra = 0
            paperDict.update({contentType: content})

        elif alignment == 1 and authorContent == 1:
            contentType = 'Authors Information'
            content = paragraph.text
            abstract = 1
            if not contentType in paperDict:
                paperDict[contentType] = paragraph.text
            else:
                paperDict[contentType] += '\n' + paragraph.text

        elif alignment == 3 and abstract == 1:
            authorContent = 0
            contentType = 'Abstract'
            content = paragraph.text
            if not contentType in paperDict:
                paperDict[contentType] = paragraph.text
            else:
                paperDict[contentType] += '\n' + paragraph.text

        else:
            abstract = 0
            contentType = 'Main Body'
            if not contentType in paperDict:
                paperDict[contentType] = paragraph.text
            else:
                paperDict[contentType] += '\n' + paragraph.text


        pre_paragraph = paragraph
    return paperDict

def pdf_to_docx(PDFDir):
    convertapi.api_secret = 'NiINCWmQ2PYMdDRU'
    convertapi.convert('docx', {
        'File': PDFDir
    }, from_format='pdf').save_files('./')

if __name__ == '__main__':

    PDFDir = 'Kingsbury and Alvaro - 2020 - Elle inferring isolation anomalies from experimen(1).pdf'
    DOCXDir = PDFDir.replace(".pdf", ".docx")

    # 将pdf转为word，如果该word已经存在则不执行
    if not (os.path.exists(PDFDir.replace(".pdf", ".docx"))):
        logger.info('Convertion starts...')
        pdf_to_docx(PDFDir)
    else:
        logger.info('file has been converted')

    # 调用输出解析内容的方法
    mydict = extract_text(DOCXDir)
    for key, value in mydict.items():
        logger.info(f'解析内容:{key}')
        sleep(1)
        print(value)
        sleep(1)

    # 打印Word文档里的每段文字属性
    doc = Document(DOCXDir)

    # 遍历文档中的段落并打印文本内容
    for paragraph in doc.paragraphs:
        tmp = paragraph.text.replace(" ", "").replace("\n", "")
        if not tmp:
            continue
        print(paragraph.text)
        # 获取段落的样式
        style = paragraph.style.name
        print(f"段落样式：{style}")
        for run in paragraph.runs:
            # 获取文本的具体属性
            text = run.text  # 文本内容
            font_name = run.font.name  # 字体名称
            font_size = run.font.size  # 字体大小
            bold = run.font.bold  # 是否加粗
            italic = run.font.italic  # 是否斜体
            underline = run.font.underline  # 是否有下划线
            # 获取段落的对齐方式
            alignment = paragraph.alignment

            # 根据对齐方式输出文字对齐方式
            if alignment == 0:
                print("左对齐")
            elif alignment == 1:
                print("居中对齐")
            elif alignment == 2:
                print("右对齐")
            elif alignment == 3:
                print("两端对齐")
            else:
                print("未知对齐方式")
            # 打印文本的具体属性
            print(f"文本内容：{text}")
            print(f"字体名称：{font_name}")
            print(f"字体大小：{font_size}")
            print(f"是否加粗：{bold}")
            print(f"是否斜体：{italic}")
            print(f"是否有下划线：{underline}")
            print()
            break
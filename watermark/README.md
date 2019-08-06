## python 批量给 word 或 pdf 文件添加水印

### 1.依赖库
```
# win32com，用于调用 word 等应用程序
pip install pypiwin32

# reportlab，用于生成 pdf 水印文件
pip install reportlab

# PyPDF4，用于合成 pdf 文件
pip install pypdf4
```

### 2.修改pypdf4 的 utils.py，以支持中英文合成
代码路径: \PyPDF4\utils.py

将其中的 r = s.encode('latin-1') ，改为如下
```
    try:
        r = s.encode('latin-1')            
        if len(s) < 2:
            bc[s] = r
        return r
    except Exception as e:
        r = s.encode('utf-8')
        if len(s) < 2:
            bc[s] = r
        return r
```

### 3.word 转换为 pdf 格式
通过 win32com.client 调用 Word 应用程序，打开 word 文件并另存为 pdf 格式。
```
def word2pdf(word_file):
    doc_ext = os.path.splitext(os.path.basename(word_file))[1].lower()
    pdf_file = word_file.replace(doc_ext, '.pdf')

    try:
        word = client.DispatchEx("Word.Application")

        if os.path.exists(pdf_file):
            os.remove(pdf_file)
        worddoc = word.Documents.Open(word_file, ReadOnly=1)
        worddoc.SaveAs(pdf_file, FileFormat=17)  # 保存文件为 pdf 格式
        worddoc.Close()
        word.Quit()
        return pdf_file
    except Exception as e:
        print(e)
        return None
```

### 4.创建水印文件
水印的位置与角度需要手动去调，以满足需求。还可以设置字体，字体的大小、颜色、透明度等参数。
```
def create_watermark(content, angle, font_size=36):
    """
    创建 PDF 水印模板
    """

    wm_file = 'watermark.pdf'
    pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))  # 加载字体
    c = canvas.Canvas(wm_file)  # 使用reportlab来创建一个PDF文件来作为一个水印文件
    
    # 配置水印文件
    c.setFont('Arial', font_size)
    c.saveState()
    c.rotate(angle)
    c.setFillAlpha(0.2)

    content_lst = content.split(' ')
    y = 0
    for c_item in content_lst:
        c.drawString(200, 100 - y, c_item)
        c.drawString(450, 350 - y, c_item)
        c.drawString(750, 250 - y, c_item)
        c.drawString(550, 0 - y, c_item)
        y += font_size + 4

    c.restoreState()
    c.save()

    return wm_file
```

### 5.合并水印文件与资料文件
通过 PdfFileReader 读取 pdf 文件，并使用 PdfFileWriter 新建一个空的 pdf，merge 水印文件，并保存。原始的 pypdf4 不支持中文的编码，需要修改 utils.py 文件（见 2.）。
```
def add_watermark(pdf_file, wm_file):
    wm_obj = PdfFileReader(wm_file)
    wm_page = wm_obj.getPage(0)

    out_file = os.path.join(os.path.dirname(pdf_file), 'wm_%s' % os.path.basename(pdf_file))
    pdf_reader = PdfFileReader(pdf_file)
    pdf_writer = PdfFileWriter()

    for page_num in range(pdf_reader.getNumPages()):
        page = pdf_reader.getPage(page_num)
        page.mergePage(wm_page)
        pdf_writer.addPage(page)

    with open(out_file, 'wb') as out:
        pdf_writer.write(out)
```

### 6.使用说明
命令
```
python add_watermark.py input_file 
                        --watermark DANPE
                        --angle 45
                        --font_file arial.ttf                        
                        --font_size 36
                        --color black
                        --alpha 0.2

# 参数说明 input_file 输入单一文件或文件夹路径，目前支持 word， excel， powerpoint 的新旧 6 种格式与 pdf 格式
           --watermark 水印文本，通过符号 ‘|’ 换行
           --angle 水印文本方向
           --font_file 可自定义字体文件，若无输入或字体文件不存在，则使用默认的字体
           --font_size 字体大小
           --color 水印颜色，可选常见的颜色，如 [black, red, blue, green, yellow, white, gold, purple, pink, orange] 等
           --alpha 字体透明度
           --only_pdf 只转换文本为 pdf，不添加水印
           --no_date 水印不加入日期
# 输出
    若输入为单一文件，会新建一个 with-watermark 目录，将添加水印的文件放置到该目录下；
    若输入为文件夹，则会遍历目录，将所有符合格式的文件添加水印，并新建一个 文件夹名+-with-watermark 的目录，存放结果。
```

### 7.常见错误
- 转换 ppt 文件时，出现错误 “The Python instance can not be converted to a COM object”
  
  在保存成 pdf 文件时，需要输入参数 PrintRange
  ```
  office_file.ExportAsFixedFormat(pdf_file, 32, PrintRange=None)
  ```
  
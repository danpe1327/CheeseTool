## python 批量给 word 或 pdf 文件添加水印

项目地址：https://github.com/danpe1327/CheeseTool/tree/master/watermark

### 1.依赖库
```
pip install -r requirements.txt
```
### 2.部分python库说明

- win32com，用于调用 word 等应用程序
- reportlab，用于生成 pdf 水印文件

### 3.安装pdf工具包
```
# PyPDF4，用于合成 pdf 文件
git clone https://github.com/danpe1327/PyPDF4.git pypdf4
cd pypdf4
python setup.py install --record files.txt
```

### 4.使用说明
命令
```
python add_watermark.py input_file 
                        --watermark DANPE
                        --angle 45
                        --font_file arial.ttf                        
                        --font_size 36
                        --color black
                        --alpha 0.2

# 参数说明 
            input_file 输入单一文件或文件夹路径，目前支持 word， excel， powerpoint 的新旧 6 种格式与 pdf 格式
           --watermark 水印文本，通过符号 ‘|’ 换行
           --angle 水印文本方向
           --font_file 可自定义字体文件，若无输入或字体文件不存在，则使用默认的字体
           --font_size 字体大小
           --color 水印颜色，可选常见的颜色，如 [black, red, blue, green, yellow, white, gold, purple, pink, orange] 等
           --alpha 字体透明度
           --only_pdf 只转换文本为 pdf，不添加水印
           --no_date 水印不加入日期
# 输出
    若输入为单一文件，会新建一个 wm-files 目录，将添加水印的文件放置到该目录下；
    若输入为文件夹，则会遍历目录，将所有符合格式的文件添加水印，并新建一个 文件夹名+"-wm-files" 的目录，存放结果。
```

### 5.常见错误
- 转换 ppt 文件时，出现错误 “The Python instance can not be converted to a COM object”
  
  在保存成 pdf 文件时，需要输入参数 PrintRange
  ```
  office_file.ExportAsFixedFormat(pdf_file, 32, PrintRange=None)
  ```
- 为中文文档添加水印报错 “'latin-1' codec can't encode characters in position 8-12: ordinal not in range(256)”
  修改pypdf4 的 utils.py，以支持中英文合成。
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
### 6.pdf 权限
- -1 permit everything
- -4096 deny anything
- -4092 only print

```
0000 0000 0001 unknown
0000 0000 0010 unknown
0000 0000 0011 unknown
0000 0000 0100 打印
0000 0000 1000 更改文档、文档组合、填写表单域、签名、创建模板页面
0000 0001 0000 内容复制、复制内容用于辅助工具
0000 0010 0000 注释、填写表单域、签名
0000 0100 0000 unknown
0000 1000 0000 unknown
0001 0000 0000 填写表单域、签名、创建模板页面
0010 0000 0000 复制内容用于辅助工具
0100 0000 0000 文档组合
1000 0000 0000 unknown
```
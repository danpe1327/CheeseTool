import os
import time
import argparse
import shutil
from tqdm import tqdm
from win32com import client
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from pypdf import PdfFileReader, PdfFileWriter
import pythoncom
import uuid

TRY_TIMES = 3
DEFAULT_FONT_SIZE_SCALE = 0.045
OFFICE_PDF_EXT = ['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.pdf']
ORGIN_LIST = [
    (0.7, 0.7),
    (0.3, 0.7),
    (0.3, 0.3),
    (0.7, 0.3),
]


class PdfConvert(object):

    def run_convert(self, in_file, save_dir):
        file_ext = os.path.splitext(os.path.basename(in_file))[1]
        pdf_file = in_file.replace(file_ext, '.pdf')
        pdf_file = os.path.join(save_dir, os.path.basename(pdf_file))

        out_file = pdf_file
        file_ext = file_ext.lower()
        if not os.path.exists(pdf_file) or file_ext in ['.xls', '.xlsx']:
            if file_ext in ['.doc', '.docx']:
                out_file = self.word2pdf(in_file, pdf_file)
            elif file_ext in ['.ppt', '.pptx']:
                out_file = self.ppt2pdf(in_file, pdf_file)
            elif file_ext in ['.xls', '.xlsx']:
                out_file = self.excel2pdf(in_file, pdf_file)
            else:
                return None

        # return a list of pdf files
        if isinstance(out_file, str):
            pdf_list = [out_file]
        else:
            pdf_list = out_file

        return pdf_list

    def word2pdf(self, in_file, pdf_file):
        office_app = None
        try:
            pythoncom.CoInitialize()
            office_app = client.DispatchEx("Word.Application")
            office_app.DisplayAlerts = False
            if os.path.exists(pdf_file):
                os.remove(pdf_file)

            office_file = office_app.Documents.Open(in_file, Visible=False, ReadOnly=1)
            office_file.ExportAsFixedFormat(pdf_file, 17)
            office_file.Close()
        except Exception as e:
            print('failed to convert word %s, %s' % (in_file, e))
            pdf_file = None
        finally:
            if office_app is not None:
                office_app.DisplayAlerts = True
                office_app.Quit()
            pythoncom.CoUninitialize()
            return pdf_file

    def excel2pdf(self, in_file, pdf_file):
        office_app = None
        out_list = []
        try:
            pythoncom.CoInitialize()
            office_app = client.DispatchEx("Excel.Application")
            office_app.DisplayAlerts = False
            # office_app.Application.DisplayAlerts = False

            office_file = office_app.Workbooks.Open(in_file, ReadOnly=1)
            sheet_num = office_file.Sheets.Count

            # save every sheet that is not empty
            for i in range(1, sheet_num + 1):
                sheet_name = office_file.Sheets(i).Name
                xls_sheet = office_file.Worksheets(sheet_name)

                if str(xls_sheet.UsedRange) == 'None':  # filter the empty sheet
                    continue

                tmp_file = pdf_file.replace('.pdf', '_%s.pdf' % sheet_name)

                if os.path.exists(tmp_file):
                    os.remove(tmp_file)
                out_list.append(tmp_file)
                xls_sheet.ExportAsFixedFormat(0, tmp_file)

            office_file.Close()
        except Exception as e:
            print('failed to convert excel %s, %s' % (in_file, e))
            if out_list and len(out_list) > 0:
                for f in out_list:
                    os.remove(f)
            out_list = None
        finally:
            if office_app is not None:
                office_app.DisplayAlerts = True
                office_app.Quit()

            pythoncom.CoUninitialize()
            return out_list

    def ppt2pdf(self, in_file, pdf_file):
        office_app = None
        try:
            pythoncom.CoInitialize()
            office_app = client.DispatchEx("Powerpoint.Application")
            office_app.DisplayAlerts = False
            if os.path.exists(pdf_file):
                os.remove(pdf_file)

            office_file = office_app.Presentations.Open(in_file, WithWindow=False, ReadOnly=1)
            office_file.ExportAsFixedFormat(pdf_file, 32, PrintRange=None)
            office_file.Close()
        except Exception as e:
            print('failed to convert ppt %s, %s' % (in_file, e))
            pdf_file = None
        finally:
            if office_app is not None:
                office_app.DisplayAlerts = True
                office_app.Quit()

            pythoncom.CoUninitialize()
            return pdf_file


def create_watermark(content, out_dir, angle, pagesize=None, font_file=None, font_size=None, color='black', alpha=0.2):
    """
    create PDF watermark file
    """
    if not isinstance(pagesize, float):
        pagesize = (float(pagesize[0]), float(pagesize[1]))

    uuid_str = uuid.uuid4().hex
    wm_file = os.path.join(out_dir, 'watermark_%s.pdf' % uuid_str)
    if font_file is None or not os.path.exists(font_file):
        available_fonts = pdfmetrics.getRegisteredFontNames()
        font_name = available_fonts[0]
    else:
        font_name = os.path.splitext(os.path.basename(font_file))[0]
        pdfmetrics.registerFont(TTFont(font_name, font_file))  # register custom font

    c = canvas.Canvas(wm_file, pagesize=pagesize)  # create an empty pdf file

    # setting pdf parameters
    w, h = pagesize
    if font_size is None:
        font_size = max(w, h) * DEFAULT_FONT_SIZE_SCALE
    c.setFont(font_name, font_size)
    c.setFillColor(eval('colors.%s' % color))
    c.setFillAlpha(alpha)
    c.saveState()

    content_list = content.split('|')
    # create 4 watermarks in page
    for i, orgin in enumerate(ORGIN_LIST):
        c.restoreState()
        c.saveState()
        c.translate(orgin[0] * w, orgin[1] * h)
        c.rotate(angle)
        y = 0
        for c_item in content_list:
            c.drawCentredString(0, 0 - y, c_item)
            y += font_size

    c.save()

    return wm_file


def merge_watermark(pdf_file, save_dir, owner_pwd, p_value, wm_attrs):
    out_file = os.path.join(save_dir, os.path.basename(pdf_file))

    try:
        pdf_reader = PdfFileReader(pdf_file)
    except Exception as e:
        print('try to repair %s' % pdf_file)
        import fitz
        pdf_doc = fitz.open(pdf_file)
        repair_pdf_file = pdf_file.replace('.pdf', '_repaired.pdf')
        pdf_doc.save(repair_pdf_file)
        pdf_doc.close()
        shutil.move(repair_pdf_file, pdf_file)
        pdf_reader = PdfFileReader(pdf_file)

    if pdf_reader.isEncrypted:
        pdf_reader.decrypt('')
    pdf_writer = PdfFileWriter(out_file)

    first_page = pdf_reader.getPage(0)
    page_width = first_page.mediaBox.getWidth()
    page_height = first_page.mediaBox.getHeight()

    wm_attrs.update({'pagesize': (page_width, page_height)})
    wm_file = create_watermark(**wm_attrs)  # for Portrait

    wm_obj = PdfFileReader(wm_file)
    wm_page = wm_obj.getPage(0)

    for page_num in range(pdf_reader.numPages):
        current_page = pdf_reader.getPage(page_num)
        # width = current_page.mediaBox.getWidth()
        # height = current_page.mediaBox.getHeight()
        ## merge the watermark file which is suitable
        # wm_page = wm_page_v if height >= width else wm_page_h
        current_page.mergePage(wm_page)
        pdf_writer.addPage(current_page)

    if owner_pwd.lower() not in ['-1', 'no', 'none', 'null']:
        pdf_writer.encrypt('', ownerPwd=owner_pwd, P=p_value)
        with open(os.path.join(wm_attrs['out_dir'], '..', 'log'), 'a', encoding='utf-8') as f_log:
            f_log.write('%s %s %s\n' % (time.strftime('%Y-%m-%d %H:%M:%S'), os.path.relpath(out_file), owner_pwd))

    pdf_writer.write()
    pdf_writer.close()

    if os.path.exists(wm_file):
        os.remove(wm_file)


def listFiles(dir, out_list, types, recursion=False):
    files = os.listdir(dir)
    for name in files:
        fullname = os.path.join(dir, name)
        if os.path.isdir(fullname):
            if recursion:
                listFiles(fullname, out_list, types, recursion)
        else:
            _, ext = os.path.splitext(name)
            if ext != '' and ext.lower() in types:
                out_list.append(fullname)
    return out_list


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('input_file', type=str, help='input docx or pdf file or directory')
    parser.add_argument('--output_dir', type=str, help='output directory', default='./watermark_output')
    # watermark params
    parser.add_argument('--watermark', type=str, help='Wrap through |', default='DANPE')
    parser.add_argument('--angle', type=int, help='', default=45)
    parser.add_argument('--font_file', type=str, help='', default='arial.ttf')
    parser.add_argument('--font_size', type=int, help='None for autoset', default=None)
    parser.add_argument('--color', type=str, help='', default='black')
    parser.add_argument('--alpha', type=float, help='', default=0.2)
    parser.add_argument('--only_pdf', action='store_true', help='', default=False)
    parser.add_argument('--no_date', action='store_true', help='the watermark with no date information', default=False)
    # encrypt params
    parser.add_argument('--pwd', type=str, help='owner password', default='123456')
    parser.add_argument(
        '--p',
        type=int,
        help='permission value, default(-4092) permit print only, -1 permit everything, -4096 deny anything',
        default=-4092)
    args = parser.parse_args()

    return args


if __name__ == '__main__':
    args = parse_args()
    input_file = os.path.abspath(args.input_file)
    output_dir = os.path.abspath(args.output_dir)
    owner_pwd = args.pwd
    if owner_pwd.lower() == 'random':
        owner_pwd = uuid.uuid4().hex

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    wm_content = args.watermark
    if not args.no_date:
        date_str = time.strftime('%Y.%m.%d')
        wm_content += '|' + date_str

    input_file_list = []
    if os.path.isdir(input_file):
        listFiles(input_file, input_file_list, OFFICE_PDF_EXT, True)
        watermark_dir = os.path.join(output_dir, '%s-wm-files' % os.path.basename(input_file))
        pdf_dir = os.path.join(output_dir, '%s-pdf-files' % os.path.basename(input_file))
    else:
        input_file_ext = os.path.splitext(os.path.basename(input_file))[1].lower()
        assert input_file_ext in OFFICE_PDF_EXT, 'Do not support %s file' % input_file_ext
        input_file_list = [input_file]
        watermark_dir = os.path.join(output_dir, 'wm-files')
        pdf_dir = os.path.join(output_dir, 'pdf-files')
    wm_attrs = {
        'content': wm_content,
        'out_dir': output_dir,
        'angle': args.angle,
        'pagesize': None,
        'font_file': args.font_file,
        'font_size': args.font_size,
        'color': args.color,
        'alpha': args.alpha,
    }

    pdf_convert = PdfConvert()
    pdf_list = []
    failure_list = []
    for src_file in tqdm(input_file_list):
        src_file = os.path.normpath(src_file)
        if os.path.basename(src_file).startswith(tuple(('wm_', '~'))) or 'wm-files' in src_file:
            print('illegal file %s' % src_file)
            continue

        watermark_save_dir = watermark_dir
        pdf_save_dir = pdf_dir
        if os.path.isdir(input_file):
            sub_dir = os.path.dirname(src_file)
            watermark_save_dir = os.path.join(watermark_dir, sub_dir.split(input_file)[1][1:])
            pdf_save_dir = os.path.join(pdf_dir, sub_dir.split(input_file)[1][1:])

        if not os.path.exists(watermark_save_dir):
            os.makedirs(watermark_save_dir)

        if not os.path.exists(pdf_save_dir):
            os.makedirs(pdf_save_dir)

        # print('add watermark for %s' % src_file)

        file_ext = os.path.splitext(os.path.basename(src_file))[1].lower()
        if file_ext == '.pdf':
            pdf_file = os.path.join(pdf_save_dir, os.path.basename(src_file))
            shutil.copy(src_file, pdf_file)
            pdf_list = [pdf_file]
        else:
            # if convert failed, try again
            left_try_times = TRY_TIMES
            while left_try_times > 0:
                try:
                    pdf_list = pdf_convert.run_convert(src_file, pdf_save_dir)
                    if pdf_list is not None:
                        print('Try to convert and result success!', left_try_times)
                        break
                finally:
                    left_try_times -= 1

        if not args.only_pdf and pdf_list is not None:
            try:
                for pdf_item in pdf_list:
                    merge_watermark(pdf_item, watermark_save_dir, owner_pwd, args.p,
                                    wm_attrs)  # add watermark, overwrite the pdf file

            except Exception as e:
                print('failed to add watermark %s' % src_file, e)
                failure_list.append(src_file)
                continue
    for i, failure_file in enumerate(failure_list):
        print(i, failure_file)
    # if not args.only_pdf:
    #     shutil.rmtree(pdf_dir)

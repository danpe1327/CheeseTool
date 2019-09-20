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


class PdfConvert(object):

    def run_convert(self, in_file, save_dir):
        file_ext = os.path.splitext(os.path.basename(in_file))[1].lower()
        pdf_file = in_file.replace(file_ext, '.pdf')
        pdf_file = os.path.join(save_dir, os.path.basename(pdf_file))

        out_file = pdf_file
        if not os.path.exists(pdf_file):
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
            pdf_lst = [out_file]
        else:
            pdf_lst = out_file

        return pdf_lst

    def word2pdf(self, in_file, pdf_file):
        office_app = None
        try:
            pythoncom.CoInitialize()
            office_app = client.DispatchEx("Word.Application")

            if os.path.exists(pdf_file):
                os.remove(pdf_file)

            office_file = office_app.Documents.Open(in_file, Visible=False, ReadOnly=1)
            office_file.ExportAsFixedFormat(pdf_file, 17)
            office_file.Close()
        except Exception as e:
            print('failed to convert word %s' % in_file, e)
            pdf_file = None
        finally:
            if office_app is not None:
                office_app.Quit()
            pythoncom.CoUninitialize()
            return pdf_file

    def excel2pdf(self, in_file, pdf_file):
        office_app = None
        pdf_lst = list()
        try:
            pythoncom.CoInitialize()
            office_app = client.DispatchEx("Excel.Application")

            office_file = office_app.Workbooks.Open(in_file, ReadOnly=1)
            sheet_num = office_file.Sheets.Count

            # save every sheet that is not empty
            for i in range(1, sheet_num + 1):
                sheet_name = office_file.Sheets(i).Name
                xls_sheet = office_file.Worksheets(sheet_name)

                if str(xls_sheet.UsedRange) == 'None':  # filter the empty sheet
                    continue

                pdf_file = pdf_file.replace('.pdf', '_%s.pdf' % sheet_name)

                if os.path.exists(pdf_file):
                    os.remove(pdf_file)
                pdf_lst.append(pdf_file)
                xls_sheet.ExportAsFixedFormat(0, pdf_file)

            office_file.Close()
        except Exception as e:
            print('failed to convert excel %s' % in_file, e)
            if pdf_lst and len(pdf_lst) > 0:
                for f in pdf_lst:
                    os.remove(f)
            pdf_lst = None
        finally:
            if office_app is not None:
                office_app.Quit()
            pythoncom.CoUninitialize()
            return pdf_lst

    def ppt2pdf(self, in_file, pdf_file):
        office_app = None
        try:
            pythoncom.CoInitialize()
            office_app = client.DispatchEx("Powerpoint.Application")

            if os.path.exists(pdf_file):
                os.remove(pdf_file)

            office_file = office_app.Presentations.Open(in_file, WithWindow=False, ReadOnly=1)
            office_file.ExportAsFixedFormat(pdf_file, 32, PrintRange=None)
            office_file.Close()
        except Exception as e:
            print('failed to convert ppt %s' % in_file, e)
            pdf_file = None
        finally:
            if office_app is not None:
                office_app.Quit()

            pythoncom.CoUninitialize()
            return pdf_file


def create_watermark(content,
                     out_dir,
                     angle,
                     pagesize=None,
                     direction='v',
                     font_file=None,
                     font_size=36,
                     color='black',
                     alpha=0.2):
    """
    create PDF watermark file
    """
    import uuid
    if not isinstance(pagesize, float):
        pagesize = (float(pagesize[0]), float(pagesize[1]))

    uuid_str = uuid.uuid4().hex
    wm_file = os.path.join(out_dir, 'watermark_%s_%s.pdf' % (uuid_str, direction))
    if font_file is None or not os.path.exists(font_file):
        available_fonts = pdfmetrics.getRegisteredFontNames()
        font_name = available_fonts[0]
    else:
        font_name = os.path.splitext(os.path.basename(font_file))[0]
        pdfmetrics.registerFont(TTFont(font_name, font_file))  # register custom font

    c = canvas.Canvas(wm_file, pagesize=pagesize)  # create an empty pdf file

    # setting pdf parameters
    c.setFont(font_name, font_size)
    c.saveState()
    c.rotate(angle)
    c.setFillColor(eval('colors.%s' % color))
    c.setFillAlpha(alpha)

    content_lst = content.split('|')
    y = 0
    w, h = pagesize
    if angle == 45:
        if direction == 'v':
            for c_item in content_lst:
                c.drawString(0.3 * w, 0.1 * h - y, c_item)
                c.drawString(0.8 * w, 0.4 * h - y, c_item)
                c.drawString(1.3 * w, 0.25 * h - y, c_item)
                c.drawString(w, 0 - y, c_item)
                y += font_size + 4
        else:
            for c_item in content_lst:
                c.drawString(0.4 * w, 0.2 * h - y, c_item)
                c.drawString(0.3 * w, -0.1 * h - y, c_item)
                c.drawString(0.5 * w, -0.6 * h - y, c_item)
                c.drawString(0.6 * w, -0.3 * h - y, c_item)
                y += font_size + 4
    else:  # angle=0
        for c_item in content_lst:
            c.drawString(0.1 * w, 0.3 * h - y, c_item)
            c.drawString(0.6 * w, 0.3 * h - y, c_item)
            c.drawString(0.1 * w, 0.7 * h - y, c_item)
            c.drawString(0.6 * w, 0.7 * h - y, c_item)
            y += font_size + 4

    c.restoreState()
    c.save()

    return wm_file


def merge_watermark(pdf_file, save_dir, wm_attrs):
    out_file = os.path.join(save_dir, os.path.basename(pdf_file))

    pdf_reader = PdfFileReader(pdf_file)
    pdf_writer = PdfFileWriter(out_file)

    first_page = pdf_reader.getPage(0)
    page_width = first_page.mediaBox.getWidth()
    page_height = first_page.mediaBox.getHeight()

    wm_attrs.update({'pagesize': (page_width, page_height)})
    wm_file_v = create_watermark(**wm_attrs)  # for Portrait

    wm_attrs.update({'direction': 'h'})
    wm_file_h = create_watermark(**wm_attrs)  # for Landscape

    wm_obj_v = PdfFileReader(wm_file_v)
    wm_page_v = wm_obj_v.getPage(0)
    wm_obj_h = PdfFileReader(wm_file_h)
    wm_page_h = wm_obj_h.getPage(0)

    for page_num in range(pdf_reader.numPages):
        current_page = pdf_reader.getPage(page_num)
        width = current_page.mediaBox.getWidth()
        height = current_page.mediaBox.getHeight()

        # merge the watermark file which is suitable
        wm_page = wm_page_v if height >= width else wm_page_h
        current_page.mergePage(wm_page)
        pdf_writer.addPage(current_page)

    pdf_writer.write()
    pdf_writer.close()

    if os.path.exists(wm_file_v):
        os.remove(wm_file_v)
    if os.path.exists(wm_file_h):
        os.remove(wm_file_h)


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
    parser.add_argument('--output_dir', type=str, help='output directory', default='.\watermark_output')
    parser.add_argument('--watermark', type=str, help='Wrap through |', default='DANPE')
    parser.add_argument('--angle', type=int, help='', default=45)
    parser.add_argument('--font_file', type=str, help='', default='arial.ttf')
    parser.add_argument('--font_size', type=int, help='', default=36)
    parser.add_argument('--color', type=str, help='', default='black')
    parser.add_argument('--alpha', type=float, help='', default=0.2)
    parser.add_argument('--only_pdf', action='store_true', help='', default=False)
    parser.add_argument('--no_date', action='store_true', help='the watermark with no date information', default=False)

    args = parser.parse_args()

    return args


if __name__ == '__main__':
    OFFICE_PDF_EXT = ['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.pdf']
    args = parse_args()
    input_file = os.path.abspath(args.input_file)    
    output_dir = os.path.abspath(args.output_dir)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    wm_content = args.watermark
    if not args.no_date:
        date_str = time.strftime('%Y.%m.%d')
        wm_content += '|' + date_str

    word_file_lst = list()
    if os.path.isdir(input_file):
        listFiles(input_file, word_file_lst, OFFICE_PDF_EXT, True)
        watermark_dir = os.path.join(output_dir, '%s-with-watermark' % os.path.basename(input_file))
        pdf_dir = os.path.join(output_dir, '%s-pdf-files' % os.path.basename(input_file))
    else:
        word_file_lst = [input_file]
        watermark_dir = os.path.join(output_dir, 'with-watermark')
        pdf_dir = os.path.join(output_dir, 'pdf-files')
    wm_attrs = {
        'content': wm_content,
        'out_dir': output_dir,
        'angle': args.angle,
        'pagesize': None,
        'direction': 'v',
        'font_file': args.font_file,
        'font_size': args.font_size,
        'color': args.color,
        'alpha': args.alpha,
    }

    pdf_convert = PdfConvert()
    for src_file in tqdm(word_file_lst):
        src_file = os.path.normpath(src_file)
        if os.path.basename(src_file).startswith(tuple(('wm_', '~'))) or 'with-watermark' in src_file:
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

        print('add watermark for %s' % src_file)

        file_ext = os.path.splitext(os.path.basename(src_file))[1].lower()
        if file_ext == '.pdf':
            pdf_file = os.path.join(pdf_save_dir, os.path.basename(src_file))
            shutil.copy(src_file, pdf_file)
            pdf_lst = [pdf_file]
        else:
            pdf_lst = pdf_convert.run_convert(src_file, pdf_save_dir)

        if not args.only_pdf and pdf_lst is not None:
            try:
                for pdf_item in pdf_lst:
                    merge_watermark(pdf_item, watermark_save_dir, wm_attrs)  # add watermark, overwrite the pdf file

            except Exception as e:
                print('failed to add watermark %s' % src_file, e)
                continue

    # if not args.only_pdf:
    #     shutil.rmtree(pdf_dir)

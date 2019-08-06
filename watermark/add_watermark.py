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
from PyPDF4 import PdfFileReader, PdfFileWriter


class PdfConvert(object):

    def run_convert(self, in_file, save_dir):
        file_ext = os.path.splitext(os.path.basename(in_file))[1].lower()
        pdf_file = in_file.replace(file_ext, '.pdf')
        pdf_file = os.path.join(save_dir, os.path.basename(pdf_file))

        if file_ext in ['.doc', '.docx']:
            out_file = self.word2pdf(src_file, pdf_file)
        elif file_ext in ['.ppt', '.pptx']:
            out_file = self.ppt2pdf(src_file, pdf_file)
        elif file_ext in ['.xls', '.xlsx']:
            out_file = self.excel2pdf(src_file, pdf_file)
        else:
            return None

        # return a list of pdf files
        if isinstance(out_file, str):
            pdf_lst = [out_file]
        else:
            pdf_lst = out_file

        return pdf_lst

    def word2pdf(self, in_file, pdf_file):
        try:
            office_app = client.DispatchEx("Word.Application")

            if os.path.exists(pdf_file):
                os.remove(pdf_file)

            office_file = office_app.Documents.Open(in_file, Visible=False, ReadOnly=1)
            office_file.ExportAsFixedFormat(pdf_file, 17)
            office_file.Close()
            office_app.Quit()
            return pdf_file
        except Exception as e:
            print('failed to convert word %s' % src_file, e)
            office_app.Quit()
            return None

    def excel2pdf(self, in_file, pdf_file):
        try:
            pdf_lst = list()
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

                xls_sheet.ExportAsFixedFormat(0, pdf_file)
                pdf_lst.append(pdf_file)

            office_file.Close()
            office_app.Quit()
            return pdf_lst
        except Exception as e:
            print('failed to convert excel %s' % src_file, e)
            office_app.Quit()
            return None

    def ppt2pdf(self, in_file, pdf_file):
        try:
            office_app = client.DispatchEx("Powerpoint.Application")

            if os.path.exists(pdf_file):
                os.remove(pdf_file)

            office_file = office_app.Presentations.Open(in_file, WithWindow=False, ReadOnly=1)
            office_file.ExportAsFixedFormat(pdf_file, 32, PrintRange=None)
            office_file.Close()
            office_app.Quit()
            return pdf_file
        except Exception as e:
            print('failed to convert ppt %s' % src_file, e)
            office_app.Quit()
            return None


def create_watermark(content, angle, direction='v', font_file=None, font_size=36, color='black', alpha=0.2):
    """
    create PDF watermark file
    """

    wm_file = 'watermark_%s.pdf' % direction
    if font_file is None or not os.path.exists(font_file):
        available_fonts = pdfmetrics.getRegisteredFontNames()
        font_name = available_fonts[0]
    else:
        font_name = os.path.splitext(os.path.basename(font_file))[0]
        pdfmetrics.registerFont(TTFont(font_name, font_file))  # register custom font

    c = canvas.Canvas(wm_file)  # create an empty pdf file

    # setting pdf parameters
    c.setFont(font_name, font_size)
    c.saveState()
    c.rotate(angle)
    c.setFillColor(eval('colors.%s' % color))
    c.setFillAlpha(alpha)

    content_lst = content.split('|')
    y = 0
    if direction == 'v':
        for c_item in content_lst:
            c.drawString(200, 100 - y, c_item)
            c.drawString(450, 350 - y, c_item)
            c.drawString(750, 250 - y, c_item)
            c.drawString(550, 0 - y, c_item)
            y += font_size + 4
    else:
        for c_item in content_lst:
            c.drawString(200, 20 - y, c_item)
            c.drawString(500, 60 - y, c_item)
            c.drawString(750, -200 - y, c_item)
            c.drawString(400, -300 - y, c_item)
            y += font_size + 4
    c.restoreState()
    c.save()

    return wm_file


def add_watermark(pdf_file, save_dir, wm_file_v, wm_file_h):
    out_file = os.path.join(save_dir, os.path.basename(pdf_file))

    pdf_reader = PdfFileReader(pdf_file)
    pdf_writer = PdfFileWriter()

    wm_obj_v = PdfFileReader(wm_file_v)
    wm_page_v = wm_obj_v.getPage(0)
    wm_obj_h = PdfFileReader(wm_file_h)
    wm_page_h = wm_obj_h.getPage(0)

    for page_num in range(pdf_reader.getNumPages()):
        current_page = pdf_reader.getPage(page_num)
        width = current_page.mediaBox.getWidth()
        height = current_page.mediaBox.getHeight()

        # merge the watermark file which is suitable
        wm_page = wm_page_v if height >= width else wm_page_h
        current_page.mergePage(wm_page)
        pdf_writer.addPage(current_page)

    with open(out_file, 'wb') as out:
        pdf_writer.write(out)


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
    parser.add_argument('--watermark', type=str, help='Wrap through |', default='DANPE|')
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

    wm_content = args.watermark
    if not args.no_date:
        date_str = time.strftime('%Y.%m.%d')
        wm_content += date_str

    word_file_lst = list()
    if os.path.isdir(input_file):
        listFiles(input_file, word_file_lst, OFFICE_PDF_EXT, True)
        out_dir = os.path.join(input_file, '..', '%s-with-watermark' % os.path.basename(input_file))
    else:
        word_file_lst = [input_file]
        out_dir = os.path.join(os.getcwd(), 'with-watermark')

    wm_file_v = create_watermark(wm_content, args.angle, 'v', args.font_file, args.font_size, args.color,
                                 args.alpha)  # for Portrait
    wm_file_h = create_watermark(wm_content, args.angle, 'h', args.font_file, args.font_size, args.color,
                                 args.alpha)  # for Landscape
    pdf_convert = PdfConvert()
    for src_file in tqdm(word_file_lst):
        src_file = os.path.normpath(src_file)
        if os.path.basename(src_file).startswith(tuple(('wm_', '~'))) or 'with-watermark' in src_file:
            print('illegal file %s' % src_file)
            continue

        save_dir = out_dir
        if os.path.isdir(input_file):
            sub_dir = os.path.dirname(src_file)
            save_dir = os.path.join(out_dir, sub_dir.split(input_file)[1][1:])

        if not os.path.exists(save_dir):
            os.makedirs(save_dir)

        # print('add watermark for %s' % src_file)

        file_ext = os.path.splitext(os.path.basename(src_file))[1].lower()
        if file_ext == '.pdf':
            pdf_file = os.path.join(save_dir, os.path.basename(src_file))
            shutil.copy(src_file, pdf_file)
            pdf_lst = [pdf_file]
        else:
            pdf_lst = pdf_convert.run_convert(src_file, save_dir)

        if not args.only_pdf and pdf_lst is not None:
            try:
                for pdf_item in pdf_lst:
                    add_watermark(pdf_item, save_dir, wm_file_v, wm_file_h)  # add watermark, overwrite the pdf file

            except Exception as e:
                print('failed to add watermark %s' % src_file, e)
                for pdf_item in pdf_lst:
                    os.remove(pdf_item)
                continue

    os.remove(wm_file_v)
    os.remove(wm_file_h)

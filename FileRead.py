import os
import chardet
import docx2txt
from win32com.client import Dispatch
import io
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.converter import XMLConverter, HTMLConverter, TextConverter
from pdfminer.layout import LAParams


class FileRead:

    def __init__(self, filepath, img_dir=None):
        # path is exist?
        if not os.path.exists(filepath):
            print('File {} does not exist.'.format(filepath))
            exit(1)
        # is a directory?
        if os.path.isdir(filepath):
            print('{} is a directory.'.format(filepath))
            exit(1)
        # is a temp file?
        _, name = os.path.split(filepath)
        if name.startswith('~$'):
            print('{} is a temp file.'.format(filepath))
            exit(1)
        # is a supported file format
        _, ext = os.path.splitext(filepath)
        if ext not in ['.txt','.doc','.docx','.pdf']:
            print('The file format is not supported')
            exit(1)
        self.filepath = filepath
        self._name = name
        self._ext = ext
        self.img_dir = img_dir
        # self.newpath = filepath
        self._info = []
        self._text = ''

    def txt2text(self):
        with open(self.filepath, 'rb') as f_:
            try:
                text = f_.read()
                # get the coding info of the file. e.g. ansi,utf-8,gbk...
                coding_info = chardet.detect(text)
                text = text.decode(encoding=coding_info['encoding'], errors='ignore')
            except:
                print('Read text failed. The file may be empty.')
                text = ''
        self._text = text
        return text

    def docx2text(self):
        text = docx2txt.process(self.filepath,self.img_dir)
        self._text = text
        return text

    def doc2docx(self):
        word = Dispatch('Word.Application')
        word.Visible = 0
        word.DisplayAlerts = 0
        doc = word.Documents.Open(self.filepath)
        newpath = os.path.splitext(self.filepath)[0] + '.docx'
        doc.SaveAs(newpath, 12, False, "", True, "", False, False, False, False)
        doc.Close()
        word.Quit()
        # os.remove(self.filepath)
        newpath = self.filepath
        return newpath

    def doc2text(self):
        newpath = self.doc2docx()
        text = docx2txt.process(self.filepath,self.img_dir)
        os.remove(newpath)
        self._text = text
        return text

    def pdf2text(self):
        fp = open(self.filepath, 'rb')
        rsrcmgr = PDFResourceManager()
        retstr = io.StringIO()
        codec = 'utf-8'
        laparams = LAParams()
        device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        text = ''
        for page in PDFPage.get_pages(fp):
            interpreter.process_page(page)
            text = text + retstr.getvalue()
        fp.close()
        self._text = text
        return text

    def readtext(self):
        text = ''
        if self._ext == '.txt':
            text = self.txt2text()
        elif self._ext == '.docx':
            text = self.docx2text()
        elif self._ext == '.doc':
            text = self.doc2text()
        elif self._ext == '.pdf':
            text = self.pdf2text()
        else:
            pass
        self._text = text
        return text


    def formatTime(self,longtime):
        '''格式化时间的函数'''
        import time
        return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(longtime))

    def formatByte(self,number):
        '''格式化文件大小的函数'''
        for (scale,lable) in [(1024 * 1024 * 1024, "GB"), (1024 * 1024, "MB"), (1024, "KB")]:
            if number >= scale:
                return "%.2f %s" % (number * 1.0 / scale, lable)

            elif number == 1:
                return "1字节"

            else:  # 小于1字节
                byte = "%.2f" % (number or 0)
        return (byte[:-3]) if byte.endswith(".00") else byte + "B"

    def getinfo(self):
        self._info.append(self.filepath)
        self._info.append(self._name)
        self._info.append(self._ext)
        fileinfo = os.stat(self.filepath)
        self._info.append(self.formatByte(fileinfo.st_size))
        self._info.append(self.formatTime(fileinfo.st_atime))
        self._info.append(self.formatTime(fileinfo.st_mtime))
        return self._info


if __name__ == '__main__':
    filepath = r'E:\gui-config.txt'
    fr = FileRead(filepath)
    print(fr.readtext())
    print(fr.getinfo())





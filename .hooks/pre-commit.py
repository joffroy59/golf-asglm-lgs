import os
import shutil
from oletools.olevba3 import VBA_Parser


EXCEL_FILE_EXTENSIONS = ('xlsb', 'xls', 'xlsm', 'xla', 'xlt', 'xlam',)
KEEP_NAME = False  # Set this to True if you would like to keep "Attribute VB_Name"


def parse(workbook_path):
    vba_path = './vba/' + workbook_path + '.vba'
    vba_parser = VBA_Parser(workbook_path)
    vba_modules = vba_parser.extract_all_macros() if vba_parser.detect_vba_macros() else []

    new_file = open("git.log", "w")
    new_file.write('workbook_path:'+workbook_path)
    new_file.write('\n')
    new_file.write('vba_path:'+vba_path)
    new_file.write('\n')

    for _, _, filename, content in vba_modules:
        lines = []
        if '\r\n' in content:
            lines = content.split('\r\n')
        else:
            lines = content.split('\n')
        if lines:
            content = []
            for line in lines:
                if line.startswith('Attribute') and 'VB_' in line:
                    if 'VB_Name' in line and KEEP_NAME:
                        content.append(line)
                else:
                    content.append(line)
            if content and content[-1] == '':
                content.pop(len(content)-1)
                non_empty_lines_of_code = len([c for c in content if c])
                if non_empty_lines_of_code > 0:
                    if not os.path.exists(os.path.join(vba_path)):
                        os.makedirs(vba_path)
                    with open(os.path.join(vba_path, filename), 'w', encoding='utf-8') as f:
                        f.write('\n'.join(content))


if __name__ == '__main__':
    toParse = []
    for root, dirs, files in os.walk('.'):
        for f in dirs:
            if f.endswith('.vba'):
                shutil.rmtree(os.path.join(root, f))

        for f in files:
            if f.endswith(EXCEL_FILE_EXTENSIONS):
                toParse.append(os.path.join(root, f))

    for f in toParse:
        parse(f)
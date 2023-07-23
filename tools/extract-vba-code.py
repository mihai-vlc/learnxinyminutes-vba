import os
import shutil
from oletools.olevba3 import VBA_Parser


EXCEL_FILE_EXTENSIONS = ('xlsb', 'xls', 'xlsm', 'xla', 'xlt', 'xlam',)


def parse(workbook_path):
    vba_path = workbook_path + '.vba'
    vba_parser = VBA_Parser(workbook_path)
    vba_modules = vba_parser.extract_all_macros() if vba_parser.detect_vba_macros() else []

    for _, _, filename, content in vba_modules:
        if not os.path.exists(os.path.join(vba_path)):
            os.makedirs(vba_path)

        content = content.replace('\r\n', '\n').replace('\r', '\n')

        with open(os.path.join(vba_path, filename), 'w', encoding='utf-8') as f:
            f.write(content)
            print(f"extract-vba-code.py: written {filename}")


if __name__ == '__main__':
    for root, dirs, files in os.walk('.'):
        for f in dirs:
            if f.endswith('.vba'):
                shutil.rmtree(os.path.join(root, f))

        for f in files:
            if f.endswith(EXCEL_FILE_EXTENSIONS) and not f.startswith("~$"):
                parse(os.path.join(root, f))
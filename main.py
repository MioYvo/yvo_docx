# __author__ = "Mio"
# __email__: "liurusi.101@gmail.com"
# created: 5/19/21 11:34 PM
from pathlib import Path

from docx import Document
from htmldocx import HtmlToDocx
import httpx
from lxml import etree


def print_dir_file():
    print('-' * 30)
    for i in Path().glob('output_docx/*'):
        print(i)
    print('-' * 30)


def get_data(url: str):
    document = Document()
    new_parser = HtmlToDocx()
    res = httpx.get(url)
    content = res.content.decode()
    content = content.replace("XiXisys.com 免费提供，仅供参考。", " ")
    content = content.replace(" 如有疑问，请联系 sds@xixisys.com 咨询。", " ")
    new_parser.add_html_to_document(content, document)
    name = etree.HTML(res.content).xpath('/html/body/article/section[1]/div[1]/span[2]')[0].text
    document.save(f'output_docx/{name}.docx')
    return name


if __name__ == '__main__':
    new_added = None
    while True:
        print_dir_file()
        if new_added:
            print(f'{new_added} added')
        try:
            url = input('输入 xixisys-api 地址: ')
        except KeyboardInterrupt:
            print('Bye baby...')
            break
        except Exception:
            break
        new_added = get_data(url)

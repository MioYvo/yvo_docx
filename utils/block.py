# __author__ = "Mio"
# __email__: "liurusi.101@gmail.com"
# created: 6/4/21 10:05 PM
from docx.table import _Cell, _Row, Table
from docx.text.paragraph import Paragraph
from docx.document import Document as T_Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl


def iter_block_items(parent):
    """
    Generate a reference to each paragraph and table child within *parent*,
    in document order. Each returned value is an instance of either Table or
    Paragraph. *parent* would most commonly be a reference to a main
    Document object, but also works for a _Cell object, which itself can
    contain paragraphs and tables.
    """
    if isinstance(parent, T_Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    elif isinstance(parent, _Row):
        parent_elm = parent._tr
    else:
        raise ValueError("something's not right")
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

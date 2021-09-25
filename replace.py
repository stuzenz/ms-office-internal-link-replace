from docx import Document

from docx.opc.constants import RELATIONSHIP_TYPE as RT

document = Document('/home/stuart/Development/2021/09-Sep/docx-replace/CoreMod Programme Hibernation Report test.docx')

rels = document.part.rels

for rel in rels:
   if rels[rel].reltype == RT.HYPERLINK:
      print("\n Original link id -", rel, "with detected URL: ", rels[rel]._target)
      print("\n Original link id -", rel, "with detected URL: ", rels
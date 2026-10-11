"""Create a three-page DOCX with SVG and navigation fields, using standard Open XML."""
import io
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from PIL import Image

folder = Path(sys.argv[1])
folder.mkdir(parents=True, exist_ok=True)
svg = '<svg xmlns="http://www.w3.org/2000/svg" width="420" height="200"><rect x="1" y="1" width="418" height="198" fill="white" stroke="black"/><path d="M30 170L390 170M30 170L30 20" fill="none" stroke="black" stroke-width="2"/><path d="M30 165C130 160 160 100 220 85S330 35 380 30" fill="none" stroke="#146be8" stroke-width="3"/><circle cx="220" cy="85" r="7" fill="#e83232"/></svg>'

def run(text):
    return f'<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'

def paragraph(text, style=None):
    properties = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ''
    return f'<w:p>{properties}{run(text)}</w:p>'

def field(code, result):
    return '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' + f'<w:r><w:instrText xml:space="preserve"> {escape(code)} </w:instrText></w:r>' + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' + run(result) + '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

page_break = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
drawing = '''<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="4000500" cy="1905000"/><wp:docPr id="1" name="SVG Figure"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="figure.svg"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rPng"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="rSvg"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000500" cy="1905000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'''
body = paragraph('Contents') + '<w:p>' + field('TOC \\o "1-2" \\u', 'Contents will update') + '</w:p>' + page_break
body += paragraph('Chapter One', 'Heading1') + paragraph('Figure details', 'Heading2')
body += '<w:p><w:bookmarkStart w:id="10" w:name="FigureTarget"/>' + run('Figure 1: SVG vector curve') + '<w:bookmarkEnd w:id="10"/></w:p>' + drawing + page_break
body += paragraph('Chapter Two', 'Heading1')
body += '<w:p>' + run('Reference: ') + field('REF FigureTarget', 'Figure 1: SVG vector curve') + '</w:p>'
body += '<w:p>' + run('Figure page: ') + field('PAGEREF FigureTarget', '2') + '</w:p>'
body += '<w:p><w:hyperlink r:id="rLink">' + run('External link') + '</w:hyperlink></w:p>'
body += '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>'
document = '<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"><w:body>' + body + '</w:body></w:document>'
styles = '''<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style></w:styles>'''
content_types = '''<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="svg" ContentType="image/svg+xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'''
root_rels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
relations = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rPng" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/fallback.png"/><Relationship Id="rSvg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/figure.svg"/><Relationship Id="rLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/></Relationships>'
png = io.BytesIO()
Image.new('RGB', (1, 1), 'white').save(png, 'PNG')
path = folder / 'SVG-navigation.DOCX'
with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as archive:
    for name, content in {'[Content_Types].xml': content_types, '_rels/.rels': root_rels,
                          'word/document.xml': document, 'word/styles.xml': styles,
                          'word/_rels/document.xml.rels': relations, 'word/media/figure.svg': svg,
                          'word/media/fallback.png': png.getvalue()}.items():
        archive.writestr(name, content)
print(path)

"""Check native and converted regression PDFs. Requires pypdf and pdfplumber."""
import json
import sys
from pathlib import Path

import pdfplumber
from pypdf import PdfReader
from pypdf.generic import ArrayObject


def inspect(path):
    reader = PdfReader(path)
    pages = {page.indirect_reference.idnum: index + 1 for index, page in enumerate(reader.pages)}
    outlines = []

    def collect(items):
        for item in items:
            if isinstance(item, list):
                collect(item)
            else:
                outlines.append((item.title, reader.get_destination_page_number(item) + 1))

    collect(reader.outline)
    links = []
    for index, page in enumerate(reader.pages):
        for reference in page.get('/Annots', []):
            annotation = reference.get_object()
            if annotation.get('/Subtype') != '/Link':
                continue
            action = annotation.get('/A', {}).get_object() if annotation.get('/A') else {}
            if action.get('/S') == '/URI':
                links.append({'from': index + 1, 'uri': str(action['/URI'])})
                continue
            destination = annotation.get('/Dest', action.get('/D'))
            if isinstance(destination, ArrayObject):
                target = pages.get(getattr(destination[0], 'idnum', None))
            else:
                named = reader.named_destinations.get(str(destination))
                target = reader.get_destination_page_number(named) + 1 if named else None
            assert target is not None, f'{path.name}: unresolved link {destination}'
            links.append({'from': index + 1, 'to': target})

    with pdfplumber.open(path) as doc:
        images = sum(len(page.images) for page in doc.pages)
        vectors = sum(len(page.curves) + len(page.lines) + len(page.rects) for page in doc.pages)
    result = {'file': path.name, 'pages': len(reader.pages), 'outlines': outlines,
              'links': links, 'images': images, 'vectors': vectors}
    print(json.dumps(result, ensure_ascii=True, indent=2))
    assert len(reader.pages) == 3, f'{path.name}: expected 3 pages'
    assert ('Chapter One', 2) in outlines, f'{path.name}: Chapter One bookmark target'
    assert ('Chapter Two', 3) in outlines, f'{path.name}: Chapter Two bookmark target'
    assert sum(link.get('from') == 1 and link.get('to') == 2 for link in links) >= 2, 'TOC links to chapter and subheading'
    assert any(link.get('from') == 1 and link.get('to') == 3 for link in links), 'TOC link to Chapter Two'
    assert sum(link.get('from') == 3 and link.get('to') == 2 for link in links) >= 2, 'REF and PAGEREF targets'
    assert any(link.get('uri', '').rstrip('/') == 'https://example.com' for link in links), 'External hyperlink preservation'
    assert images == 0 and vectors >= 4, f'{path.name}: SVG must remain vector'
    return result


folder = Path(sys.argv[1])
reports = [inspect(folder / name) for name in ['word-native.pdf', 'clawpdf-converted.pdf']]
assert reports[0]['outlines'] == reports[1]['outlines'], 'Converted bookmark structure must match'
assert reports[0]['links'] == reports[1]['links'], 'Converted navigation targets must match'
print('PASS: SVG vectors, chapter bookmarks, clickable TOC, REF/PAGEREF and external links')

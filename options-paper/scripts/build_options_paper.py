#!/usr/bin/env python3
"""
Build a branded options paper (.docx) from a JSON content file.

Usage:
    python3 build_options_paper.py --content content.json --template your-template.docx --output paper.docx
"""

import argparse, json, zipfile, io, os, sys

# ── XML infrastructure ──

NS = (
    'xmlns:o="urn:schemas-microsoft-com:office:office" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:w10="urn:schemas-microsoft-com:office:word" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
    'mc:Ignorable="w14 wp14 w15"'
)

SECT_PR = '''<w:sectPr>
  <w:headerReference w:type="even" r:id="rId12"/>
  <w:headerReference w:type="default" r:id="rId13"/>
  <w:headerReference w:type="first" r:id="rId14"/>
  <w:footerReference w:type="even" r:id="rId15"/>
  <w:footerReference w:type="default" r:id="rId16"/>
  <w:footerReference w:type="first" r:id="rId17"/>
  <w:type w:val="nextPage"/>
  <w:pgSz w:w="11906" w:h="16838"/>
  <w:pgMar w:left="1440" w:right="1440" w:gutter="0" w:header="708"
           w:top="1440" w:footer="708" w:bottom="1440"/>
  <w:pgNumType w:fmt="decimal"/>
  <w:formProt w:val="false"/>
  <w:titlePg/>
  <w:textDirection w:val="lrTb"/>
  <w:docGrid w:type="default" w:linePitch="360" w:charSpace="0"/>
</w:sectPr>'''


def esc(t):
    """XML-escape text content."""
    return t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


# ── Element builders ──

def title(text):
    return f'<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'

def subtitle(text):
    return f'<w:p><w:pPr><w:pStyle w:val="Subtitle"/></w:pPr><w:r><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'

def heading1(text):
    return f'<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'

def heading2(text):
    return f'<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'

def heading3(text):
    return f'<w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'

def para(text):
    return f'<w:p><w:r><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'

def bold_para(text):
    return f'<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'

def italic_para(text):
    return f'<w:p><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'

def bullet(text):
    """Plain bullet point."""
    return (f'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/>'
            f'<w:numId w:val="14"/></w:numPr></w:pPr>'
            f'<w:r><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>')

def bullet_bold_prefix(prefix, rest):
    """Bullet with a bold opening label."""
    return (f'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/>'
            f'<w:numId w:val="14"/></w:numPr></w:pPr>'
            f'<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{esc(prefix)}</w:t></w:r>'
            f'<w:r><w:t xml:space="preserve">{esc(rest)}</w:t></w:r></w:p>')

def empty():
    return '<w:p/>'


# ── Table builder ──

def table_row(cells, bold=False):
    row = '<w:tr>'
    for c in cells:
        rpr = '<w:rPr><w:b/></w:rPr>' if bold else ''
        row += (f'<w:tc><w:tcPr><w:tcBorders>'
                f'<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                f'<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                f'<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                f'</w:tcBorders></w:tcPr>'
                f'<w:p><w:r>{rpr}<w:t xml:space="preserve">{esc(c)}</w:t></w:r></w:p></w:tc>')
    row += '</w:tr>'
    return row

def build_table(headers, rows):
    """Build a complete table element. headers is a list of strings, rows is list of list of strings."""
    t = '<w:tbl><w:tblPr><w:tblW w:w="9016" w:type="dxa"/><w:tblBorders>'
    for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        t += f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    t += '</w:tblBorders></w:tblPr>'
    t += table_row(headers, bold=True)
    for row in rows:
        t += table_row(row, bold=False)
    t += '</w:tbl>'
    return t


# ── Bullet parsing ──

def parse_bullet(text):
    """Parse a bullet string. If it contains a bold prefix (text ending with '. ' before the
    rest), split it. Otherwise return as a plain bullet.

    Bullets from strengths/risks typically look like:
      "Bold label. Rest of the explanation..."
    """
    # Check if this looks like "Bold label. Rest..." pattern
    dot_pos = text.find('. ')
    if dot_pos > 0 and dot_pos < 60:  # reasonable label length
        prefix = text[:dot_pos + 2]  # includes ". "
        rest = text[dot_pos + 2:]
        return bullet_bold_prefix(prefix, rest)
    return bullet(text)


# ── Content assembly ──

def build_content(data):
    """Convert the JSON content structure into a list of XML paragraph elements."""
    parts = []

    # Title block
    parts.append(title(data['title']))
    if data.get('subtitle'):
        parts.append(subtitle(data['subtitle']))
    parts.append(empty())

    # Context
    parts.append(heading1('Context'))
    for p in data.get('context', []):
        parts.append(para(p))
        parts.append(empty())

    # Problem / tension / opportunity
    problem_heading = data.get('problem_heading', 'The problem')
    parts.append(heading1(problem_heading))
    for p in data.get('problem_paragraphs', []):
        parts.append(para(p))
        parts.append(empty())
    for b in data.get('problem_bullets', []):
        if isinstance(b, dict) and 'bold' in b:
            parts.append(bullet_bold_prefix(b['bold'], b['text']))
        else:
            parts.append(bullet(b if isinstance(b, str) else str(b)))
    if data.get('problem_bullets'):
        parts.append(empty())
    if data.get('problem_closing'):
        parts.append(para(data['problem_closing']))
        parts.append(empty())

    # Options
    for opt in data.get('options', []):
        star = ' \u2605' if opt.get('star') else ''
        parts.append(heading1(f"Option {opt['number']}: {opt['name']}{star}"))
        parts.append(empty())

        if opt.get('star'):
            parts.append(italic_para('The \u201Cmore out there\u201D option.'))
            parts.append(empty())

        parts.append(heading2('Core idea'))
        parts.append(para(opt['core_idea']))
        parts.append(empty())

        parts.append(heading2('How it works'))
        how = opt.get('how_it_works', [])
        if isinstance(how, str):
            how = [how]
        for p in how:
            parts.append(para(p))
            parts.append(empty())

        parts.append(heading2('Strengths'))
        for s in opt.get('strengths', []):
            parts.append(parse_bullet(s))
        parts.append(empty())

        parts.append(heading2('Risks'))
        for r in opt.get('risks', []):
            parts.append(parse_bullet(r))
        parts.append(empty())

        parts.append(heading2('Best for'))
        parts.append(para(opt['best_for']))
        parts.append(empty())

    # Comparison table
    comp = data.get('comparison', {})
    if comp:
        parts.append(heading1('Comparison'))
        parts.append(empty())
        parts.append(build_table(
            comp.get('headers', ['Option', 'Strength', 'Main risk', 'Operational cost', 'Best pairing']),
            comp.get('rows', [])
        ))
        parts.append(empty())

    # Suggested approach
    if data.get('suggested_approach_intro') or data.get('suggested_approach_items'):
        parts.append(heading1('Suggested approach'))
        if data.get('suggested_approach_intro'):
            parts.append(para(data['suggested_approach_intro']))
            parts.append(empty())
        for item in data.get('suggested_approach_items', []):
            if isinstance(item, dict) and 'bold' in item:
                parts.append(bullet_bold_prefix(item['bold'], item['text']))
            else:
                parts.append(bullet(item if isinstance(item, str) else str(item)))
        parts.append(empty())

    # Next steps
    if data.get('next_steps'):
        parts.append(heading1('Next steps'))
        for step in data['next_steps']:
            parts.append(bullet(step))
        parts.append(empty())

    return parts


# ── DOCX assembly ──

def build_docx(content_parts, template_path, output_path):
    """Assemble the XML and inject into the template .docx."""
    body_xml = '\n'.join(content_parts)

    doc_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document {NS}>
<w:body>
{body_xml}
{SECT_PR}
</w:body>
</w:document>'''

    buf = io.BytesIO()
    with zipfile.ZipFile(template_path, 'r') as src, \
         zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            if item.filename == 'word/document.xml':
                dst.writestr(item, doc_xml.encode('utf-8'))
            else:
                dst.writestr(item, src.read(item.filename))

    os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)
    with open(output_path, 'wb') as f:
        f.write(buf.getvalue())

    return os.path.getsize(output_path)


def main():
    parser = argparse.ArgumentParser(description='Build a branded options paper (.docx)')
    parser.add_argument('--content', required=True, help='Path to JSON content file')
    parser.add_argument('--template', required=True, help='Path to a Word template (.docx) with your styles/branding')
    parser.add_argument('--output', required=True, help='Output .docx path')
    args = parser.parse_args()

    with open(args.content, 'r') as f:
        data = json.load(f)

    parts = build_content(data)
    size = build_docx(parts, args.template, args.output)
    print(f'Written to {args.output} ({size:,} bytes)')


if __name__ == '__main__':
    main()

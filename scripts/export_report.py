import sys
import os
import re
import markdown
import requests
from io import BytesIO
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def verify_bibliography_dois(md_text):
    """Scans all DOIs in the markdown text and queries the Crossref API to verify metadata."""
    doi_pattern = re.compile(r'https?://doi\.org/(10\.\d{4,9}/[-._;()/:A-Za-z0-9]+)')
    dois = list(set(doi_pattern.findall(md_text)))

    print("\n--- Running Automated Crossref DOI Verification ---")
    verified_count = 0
    for doi in dois:
        clean_doi = doi.rstrip('.)]')
        url = f"https://api.crossref.org/works/{clean_doi}"
        try:
            r = requests.get(url, headers={'User-Agent': 'ResearchVerifier/1.0 (mailto:test@example.com)'}, timeout=6)
            if r.status_code == 200:
                item = r.json().get('message', {})
                title = item.get('title', ['Unknown Title'])[0]
                authors = ', '.join([a.get('family', '') for a in item.get('author', [])[:3]])
                print(f"  [VERIFIED] DOI: 10.{clean_doi.split('10.',1)[-1]} -> \"{title[:45]}...\" by {authors}")
                verified_count += 1
            else:
                print(f"  [WARNING] DOI not found or blocked: {clean_doi} (HTTP {r.status_code})")
        except Exception as e:
            print(f"  [NOTICE] Verification check timed out for: {clean_doi} ({e})")

    print(f"Verified {verified_count}/{len(dois)} DOIs against official Crossref metadata.\n")

def add_hyperlink(paragraph, url, text, color="0366D6", underline=True):
    """Adds a native clickable hyperlink or internal bookmark to a python-docx paragraph."""
    try:
        if url.startswith('#'):
            anchor = url.lstrip('#')
            hyperlink = parse_xml(f'<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:anchor="{anchor}"/>')
            new_run = parse_xml(f'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:color w:val="{color}"/><w:u w:val="{"single" if underline else "none"}"/></w:rPr><w:t>{text}</w:t></w:r>')
            hyperlink.append(new_run)
            paragraph._p.append(hyperlink)
            return

        part = paragraph.part
        r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
        hyperlink = parse_xml(f'<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" r:id="{r_id}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>')
        new_run = parse_xml(f'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:color w:val="{color}"/><w:u w:val="{"single" if underline else "none"}"/></w:rPr><w:t>{text}</w:t></w:r>')
        hyperlink.append(new_run)
        paragraph._p.append(hyperlink)
    except Exception:
        run = paragraph.add_run(text)
        run.font.color.rgb = RGBColor(3, 102, 214)
        run.font.underline = True


def inline_local_svgs(html_content, base_dir):
    """Replaces <img src="images/*.svg" ... /> with raw inline <svg>...</svg> for 100% standalone HTML."""
    def svg_replacer(match):
        img_tag = match.group(0)
        src_match = re.search(r'src=["\']([^"\']+\.svg)["\']', img_tag)
        if not src_match:
            return img_tag
        rel_path = src_match.group(1)
        full_path = os.path.normpath(os.path.join(base_dir, rel_path))
        if os.path.exists(full_path):
            try:
                with open(full_path, 'r', encoding='utf-8') as f:
                    svg_code = f.read()
                svg_code = re.sub(r'<\?xml[^>]*\?>', '', svg_code)
                return f'<div class="svg-figure-container" style="text-align:center; margin: 1.5em 0;">{svg_code}</div>'
            except Exception as e:
                print(f"Warning: Failed to inline SVG {full_path}: {e}")
        return img_tag

    return re.sub(r'<img[^>]+src=["\'][^"\']+\.svg["\'][^>]*>', svg_replacer, html_content)

def protect_latex_math(md_text):
    """
    Extracts $$...$$ and $...$ blocks and replaces them with safe unique tokens
    so Python's markdown parser NEVER mangles underscores (_) into <em> tags.
    """
    math_tokens = {}
    counter = 0

    # 1. Protect display math $$ ... $$
    def display_replacer(match):
        nonlocal counter
        token = f"ZZZDISPLAYMATHBLOCK{counter}ZZZ"
        math_content = match.group(0)
        # Defensive: escape any unescaped '#' so MathJax doesn't throw macro parameter errors
        math_content = re.sub(r'(?<!\\)#', r'\\#', math_content)
        math_tokens[token] = math_content
        counter += 1
        return token

    # Match $$...$$ across multiple lines
    text = re.sub(r'\$\$[\s\S]*?\$\$', display_replacer, md_text)

    # 2. Protect inline math $ ... $
    def inline_replacer(match):
        nonlocal counter
        token = f"ZZZINLINEMATHBLOCK{counter}ZZZ"
        math_content = match.group(0)
        math_content = re.sub(r'(?<!\\)#', r'\\#', math_content)
        math_tokens[token] = math_content
        counter += 1
        return token


    # Match $...$ on a single line where content has no $ and is not empty
    text = re.sub(r'\$(?!\s)[^$\n]+?(?<!\s)\$', inline_replacer, text)

    return text, math_tokens

def restore_latex_math(html_text, math_tokens):
    """Restores the exact, untouched raw LaTeX math expressions into the HTML."""
    for token, original_math in math_tokens.items():
        html_text = html_text.replace(token, original_math)
    return html_text

def create_html(md_text, output_path, base_dir):
    # 1. Protect all LaTeX math expressions from markdown parser underscore/italic mangling
    safe_md, math_tokens = protect_latex_math(md_text)

    # 2. Convert Markdown to HTML
    raw_html_body = markdown.markdown(
        safe_md, 
        extensions=['fenced_code', 'tables', 'attr_list', 'sane_lists']
    )

    # 3. Restore the pure, pristine LaTeX math expressions
    pure_html_body = restore_latex_math(raw_html_body, math_tokens)
    
    # 4. Inline local SVGs for 100% standalone reliability
    html_body = inline_local_svgs(pure_html_body, base_dir)

    # 5. Format In-Text Citations with interactive badges
    html_body = re.sub(r'<a href="#ref-([^"]+)">', r'<a href="#ref-\1" class="citation-ref" title="Jump to Reference \1">', html_body)

    doc_title_match = re.search(r'^#\s+(.+)$', md_text, re.MULTILINE)
    doc_title = doc_title_match.group(1).strip() if doc_title_match else "Plasmonic Hotspot Catalyst Loading Literature Survey"

    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{doc_title}</title>

    <!-- 1. MathJax v3 Configuration (MUST precede the script loader) -->
    <script>
      window.MathJax = {{
        tex: {{
          inlineMath: [['$', '$'], ['\\\\(', '\\\\)']],
          displayMath: [['$$', '$$'], ['\\\\[', '\\\\]']],
          processEscapes: true
        }},
        svg: {{
          fontCache: 'global'
        }}
      }};
    </script>
    <script id="MathJax-script" async src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-svg.js"></script>

    <style>
        :root {{
            --primary: #1e40af;
            --text-main: #0f172a;
            --text-muted: #475569;
            --bg-page: #f8fafc;
            --bg-card: #ffffff;
            --border: #e2e8f0;
            --link: #0284c7;
        }}

        body {{
            font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif;
            line-height: 1.7;
            background-color: var(--bg-page);
            color: var(--text-main);
            padding: 2.5em 1.5em;
            margin: 0;
        }}

        .article-container {{
            max-width: 960px;
            margin: auto;
            background: var(--bg-card);
            padding: 3.5em 4em;
            border-radius: 12px;
            box-shadow: 0 4px 20px -2px rgba(0, 0, 0, 0.06), 0 2px 6px -1px rgba(0, 0, 0, 0.04);
            border: 1px solid var(--border);
        }}

        h1 {{
            color: #0f172a;
            font-size: 2.3em;
            line-height: 1.25;
            margin-top: 0;
            border-bottom: 2px solid #0284c7;
            padding-bottom: 0.35em;
        }}

        h2 {{
            color: #1e3a8a;
            font-size: 1.6em;
            margin-top: 1.8em;
            margin-bottom: 0.6em;
            border-bottom: 1px solid var(--border);
            padding-bottom: 0.25em;
        }}

        h3 {{
            color: #1e40af;
            font-size: 1.25em;
            margin-top: 1.4em;
        }}

        p, li {{
            font-size: 16px;
            color: #334155;
        }}

        /* Table Styling */
        table {{
            border-collapse: collapse;
            width: 100%;
            margin: 1.8em 0;
            font-size: 14.5px;
            background: #ffffff;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            border-radius: 8px;
            overflow: hidden;
        }}

        th {{
            background-color: #f1f5f9;
            color: #0f172a;
            font-weight: 600;
            text-align: left;
            padding: 12px 14px;
            border: 1px solid var(--border);
        }}

        td {{
            padding: 12px 14px;
            border: 1px solid var(--border);
            color: #334155;
            vertical-align: top;
        }}

        tr:nth-child(even) {{
            background-color: #f8fafc;
        }}

        /* Code & Pre Blocks */
        pre {{
            background: #0f172a;
            color: #f8fafc;
            padding: 1.2em;
            border-radius: 8px;
            overflow-x: auto;
            font-size: 13.5px;
            line-height: 1.5;
        }}

        code {{
            font-family: 'Consolas', 'Courier New', monospace;
            background: #f1f5f9;
            color: #0f172a;
            padding: 2px 6px;
            border-radius: 4px;
            font-size: 0.9em;
        }}

        pre code {{
            background: transparent;
            color: inherit;
            padding: 0;
        }}

        hr {{
            border: none;
            border-top: 1px solid var(--border);
            margin: 2.5em 0;
        }}

        /* SVGs & Figures */
        .svg-figure-container svg {{
            max-width: 100%;
            height: auto;
            display: inline-block;
            filter: drop-shadow(0 4px 8px rgba(0,0,0,0.04));
        }}

        img {{
            max-width: 100%;
            height: auto;
            border-radius: 8px;
        }}

        /* Hyperlink & Instant Tooltip Styling */
        a {{
            color: var(--link);
            text-decoration: none;
            transition: color 0.15s ease;
        }}

        a:hover {{
            color: #0369a1;
            text-decoration: underline;
        }}

        .tooltip-term {{
            border-bottom: 1.5px dotted #0284c7;
            font-weight: 500;
            color: #0369a1;
            cursor: help;
        }}

        .tooltip-term:hover {{
            background-color: #e0f2fe;
            border-radius: 2px;
        }}

        /* In-Text Citation Badges */
        a.citation-ref {{
            display: inline-block;
            font-size: 0.82em;
            font-weight: 600;
            color: #1d4ed8;
            background: #eff6ff;
            border: 1px solid #bfdbfe;
            border-radius: 4px;
            padding: 0px 5px;
            margin: 0 1px;
            text-decoration: none;
            vertical-align: baseline;
            transition: all 0.15s ease;
        }}

        a.citation-ref:hover {{
            background: #1d4ed8;
            color: #ffffff;
            border-color: #1e40af;
            text-decoration: none;
            transform: translateY(-1px);
            box-shadow: 0 2px 4px rgba(29, 78, 216, 0.2);
        }}

        .reference-item {{
            padding: 8px 12px;
            border-radius: 6px;
            margin-bottom: 8px;
            transition: background-color 0.4s ease, border-left 0.4s ease;
        }}

        .reference-item:target {{
            background-color: #fef9c3 !important;
            border-left: 4px solid #ca8a04 !important;
        }}

        html {{
            scroll-behavior: smooth;
        }}

        /* Floating Tooltip Box (Zero Lag) */
        #custom-floating-tooltip {{
            position: fixed;
            display: none;
            max-width: 330px;
            background: #0f172a;
            color: #f8fafc;
            padding: 9px 14px;
            border-radius: 6px;
            font-size: 13px;
            line-height: 1.45;
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.4), 0 8px 10px -6px rgba(0, 0, 0, 0.2);
            z-index: 999999;
            pointer-events: none;
            border: 1px solid #334155;
            opacity: 0;
            transition: opacity 0.1s ease;
        }}

        #custom-floating-tooltip .tooltip-header {{
            color: #38bdf8;
            font-weight: 600;
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            margin-bottom: 3px;
        }}
    </style>
</head>
<body>
    <div id="custom-floating-tooltip">
        <div class="tooltip-header">Technical Definition</div>
        <div class="tooltip-body"></div>
    </div>

    <div class="article-container">
        {html_body}
    </div>

    <!-- Instant Floating Hover Tooltip Engine -->
    <script>
      (function() {{
        const tooltip = document.getElementById('custom-floating-tooltip');
        const tooltipBody = tooltip.querySelector('.tooltip-body');

        function positionTooltip(e) {{
          const offset = 14;
          let x = e.clientX + offset;
          let y = e.clientY + offset;

          const tooltipWidth = 330;
          const tooltipHeight = 110;
          if (x + tooltipWidth > window.innerWidth) {{
            x = e.clientX - tooltipWidth - 6;
          }}
          if (y + tooltipHeight > window.innerHeight) {{
            y = e.clientY - tooltipHeight - 6;
          }}

          tooltip.style.left = x + 'px';
          tooltip.style.top = y + 'px';
        }}

        document.querySelectorAll('a[title]').forEach(function(link) {{
          const def = link.getAttribute('title');
          if (!def || def.trim() === '') return;

          link.classList.add('tooltip-term');
          link.setAttribute('data-definition', def);
          link.removeAttribute('title'); // Strip native title to eliminate 2-second browser lag

          link.addEventListener('mouseenter', function(e) {{
            tooltipBody.textContent = def;
            tooltip.style.display = 'block';
            tooltip.style.opacity = '1';
            positionTooltip(e);
          }});

          link.addEventListener('mousemove', positionTooltip);

          link.addEventListener('mouseleave', function() {{
            tooltip.style.display = 'none';
            tooltip.style.opacity = '0';
          }});
        }});
      }})();
    </script>
</body>
</html>
"""
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    print(f"Successfully generated HTML with Protected MathJax, Tooltips, and Inline SVGs: {output_path}")

def create_docx(md_text, output_path, base_dir):
    doc = Document()

    # Page Margins: 1 inch all around
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    lines = md_text.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # Skip raw HTML div or image wrappers
        if stripped.startswith('<div') or stripped.startswith('</div>') or stripped.startswith('</p>'):
            i += 1
            continue

        # Check for Markdown Headings
        if stripped.startswith('# '):
            p = doc.add_heading(stripped[2:].strip(), level=1)
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(6)
            i += 1
            continue
        elif stripped.startswith('## '):
            p = doc.add_heading(stripped[3:].strip(), level=2)
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(4)
            i += 1
            continue
        elif stripped.startswith('### '):
            p = doc.add_heading(stripped[4:].strip(), level=3)
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(3)
            i += 1
            continue
        elif stripped.startswith('#### '):
            p = doc.add_heading(stripped[5:].strip(), level=4)
            i += 1
            continue

        # Check for Horizontal Rules
        if stripped in ['---', '***', '___']:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(6)
            p_border = parse_xml(r'<w:pBdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:bottom w:val="single" w:sz="6" w:space="1" w:color="CBD5E1"/></w:pBdr>')
            p._p.get_or_add_pPr().append(p_border)
            i += 1
            continue

        # Check for Tables (Markdown Pipe Syntax)
        if stripped.startswith('|') and '|' in stripped[1:]:
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                table_lines.append(lines[i].strip())
                i += 1

            if len(table_lines) >= 3:
                headers = [c.strip() for c in table_lines[0].split('|')[1:-1]]
                data_rows = []
                for row_line in table_lines[2:]:
                    cells = [c.strip() for c in row_line.split('|')[1:-1]]
                    data_rows.append(cells)

                table = doc.add_table(rows=len(data_rows) + 1, cols=len(headers))
                table.style = 'Table Grid'

                # Style header row
                hdr_cells = table.rows[0].cells
                for col_idx, h_text in enumerate(headers):
                    hdr_cells[col_idx].text = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', h_text)
                    shd = parse_xml(r'<w:shd {} w:fill="F1F5F9"/>'.format(nsdecls('w')))
                    hdr_cells[col_idx]._tc.get_or_add_tcPr().append(shd)

                # Populate data rows
                for row_idx, row_data in enumerate(data_rows):
                    row_cells = table.rows[row_idx + 1].cells
                    for col_idx, cell_text in enumerate(row_data):
                        if col_idx < len(row_cells):
                            clean_cell = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', cell_text)
                            row_cells[col_idx].text = clean_cell

                doc.add_paragraph()
            continue

        # Check for SVG or Image tags
        img_match = re.search(r'!\[([^\]]*)\]\(([^)]+)\)', stripped) or re.search(r'<img[^>]+src=["\']([^"\']+)["\']', stripped)
        if img_match:
            img_src = img_match.group(2) if img_match.group(0).startswith('!') else img_match.group(1)
            full_img_path = os.path.normpath(os.path.join(base_dir, img_src))
            if os.path.exists(full_img_path):
                picture_path = full_img_path
                if full_img_path.lower().endswith('.svg'):
                    png_path = os.path.splitext(full_img_path)[0] + '.png'
                    if not os.path.exists(png_path):
                        edge_path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
                        if os.path.exists(edge_path):
                            try:
                                import subprocess
                                file_url = f"file:///{os.path.abspath(full_img_path).replace(os.sep, '/')}"
                                subprocess.run([edge_path, "--headless=new", "--disable-gpu", "--window-size=1150,560", f"--screenshot={png_path}", file_url], check=True, capture_output=True)
                            except Exception as e:
                                print(f"Warning: Failed to rasterize SVG: {e}")
                    if os.path.exists(png_path):
                        picture_path = png_path

                try:
                    p = doc.add_paragraph()
                    p.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
                    p.add_run().add_picture(picture_path, width=Inches(6.2))
                except Exception as e:
                    print(f"Warning: Failed to add picture {picture_path}: {e}")
                    doc.add_paragraph(f"[Embedded Figure: {os.path.basename(img_src)}]")
            i += 1
            continue

        # Check for Bullet Lists
        if stripped.startswith('* ') or stripped.startswith('- '):
            p = doc.add_paragraph(style='List Bullet')
            content = stripped[2:].strip()
            add_formatted_runs_with_hyperlinks(p, content)
            i += 1
            continue

        # Check for Numbered Lists
        num_match = re.match(r'^\d+\.\s+(.*)$', stripped)
        if num_match:
            p = doc.add_paragraph(style='List Number')
            content = num_match.group(1)
            add_formatted_runs_with_hyperlinks(p, content)
            i += 1
            continue

        # Check for Math Blocks ($$ ... $$)
        if stripped.startswith('$$') and stripped.endswith('$$') and len(stripped) > 4:
            p = doc.add_paragraph()
            p.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(stripped[2:-2].strip())
            run.italic = True
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(15, 23, 42)
            i += 1
            continue

        # Standard Paragraph
        if stripped != '':
            p = doc.add_paragraph()
            add_formatted_runs_with_hyperlinks(p, stripped)

        i += 1

    doc.save(output_path)
    print(f"Successfully generated DOCX with native hyperlinks and tables: {output_path}")

def add_formatted_runs_with_hyperlinks(paragraph, text):
    """Parses inline bold, italics, math, and hyperlinks, injecting native OpenXML runs."""
    link_pattern = re.compile(r'\[([^\]]+)\]\(([^ ")]+)(?:\s+"([^"]*)")?\)')

    last_idx = 0
    for match in link_pattern.finditer(text):
        start, end = match.span()
        if start > last_idx:
            segment = text[last_idx:start]
            add_inline_runs(paragraph, segment)

        link_text = match.group(1)
        url = match.group(2)
        add_hyperlink(paragraph, url, link_text)
        last_idx = end

    if last_idx < len(text):
        add_inline_runs(paragraph, text[last_idx:])

def add_inline_runs(paragraph, text):
    """Splits bold (**text**) and code (`code`) segments into styled runs."""
    parts = re.split(r'(\*\*[^*]+\*\*)', text)
    for part in parts:
        if part.startswith('**') and part.endswith('**'):
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        else:
            paragraph.add_run(part)

def main():
    if len(sys.argv) < 3:
        print("Usage: python export_report.py <input_md_file> <output_dir>")
        sys.exit(1)

    input_md = os.path.abspath(sys.argv[1].strip('\"\''))
    output_dir = os.path.abspath(sys.argv[2].strip('\"\''))
    base_dir = os.path.dirname(input_md)

    with open(input_md, 'r', encoding='utf-8') as f:
        md_text = f.read()

    # 1. Run automated Crossref DOI verification
    verify_bibliography_dois(md_text)

    # 2. Generate formats
    basename = os.path.splitext(os.path.basename(input_md))[0]
    html_path = os.path.join(output_dir, f"{basename}.html")
    docx_path = os.path.join(output_dir, f"{basename}.docx")

    create_html(md_text, html_path, base_dir)
    create_docx(md_text, docx_path, base_dir)

    # 3. GitHub Pages deployment pipeline
    index_html_path = os.path.join(output_dir, "index.html")
    if os.path.normpath(html_path) != os.path.normpath(index_html_path):
        import shutil
        shutil.copyfile(html_path, index_html_path)
        print(f"Successfully generated GitHub Pages entry point: {index_html_path}")

    nojekyll_path = os.path.join(output_dir, ".nojekyll")
    if not os.path.exists(nojekyll_path):
        with open(nojekyll_path, 'w', encoding='utf-8') as f:
            f.write('')
        print(f"Successfully created GitHub Pages Jekyll bypass: {nojekyll_path}")


if __name__ == "__main__":
    main()

import json
import os
import pypandoc
import sys
import re
import argparse
import tempfile

def convert_book(metadata_path, output_dir, target_chapter=None):
    if not os.path.exists(metadata_path):
        print(f"Error: Metadata file not found at {metadata_path}")
        return

    with open(metadata_path, 'r', encoding='utf-8') as f:
        book_data = json.load(f)

    book_title = book_data.get("book_title", "Book")
    chapters = book_data.get("chapters", [])
    
    if not chapters:
        print("No chapters found in metadata.")
        return

    # Filter chapters if target_chapter is specified
    if target_chapter is not None:
        chapters = [c for c in chapters if c['number'] == target_chapter]
        if not chapters:
            print(f"Error: Chapter {target_chapter} not found in metadata.")
            return

    # Base directory for latex files
    base_latex_dir = "input/latex_files"

    for chapter in chapters:
        chapter_num = chapter['number']
        chapter_title = chapter['title']
        print(f"Processing Chapter {chapter_num}: {chapter_title}")
        
        # Expected chapter directory
        chapter_dir = os.path.join(base_latex_dir, f"Chapter_{chapter_num}")
        
        if not os.path.exists(chapter_dir):
            print(f"Error: Chapter directory not found: {chapter_dir}")
            sys.exit(1)

        # List all files in the chapter directory once
        try:
            dir_files = os.listdir(chapter_dir)
        except OSError as e:
            print(f"Error accessing directory {chapter_dir}: {e}")
            sys.exit(1)

        chapter_files = []
        for section in chapter.get('sections', []):
            section_num = section.get('number') # e.g. "1.1"
            
            found_file = None
            pattern = re.compile(rf"section[\s._]*{re.escape(str(section_num))}(\D|$)", re.IGNORECASE)
            
            for fname in dir_files:
                if pattern.search(fname):
                    found_file = os.path.join(chapter_dir, fname)
                    break 
            
            if not found_file:
                print(f"ERROR: Missing file for Section {section_num} in {chapter_dir}")
                print(f"       Expected file containing 'Section {section_num}'")
                sys.exit(1)
                
            print(f"  Found: {os.path.basename(found_file)}")
            chapter_files.append(found_file)
            
        if not chapter_files:
            print(f"  No valid files found for Chapter {chapter_num}")
            continue

        # Pre-process files
        temp_files = []
        cleaned_chapter_files = []
        
        citation_pattern = re.compile(r'\\(cite|citep|citet|ref)\{[^}]+\}|\[cite:[^\]]+\]|\[cite_start\]')
        graphics_pattern = re.compile(r'\\includegraphics(?:\[(.*?)\])?\{(.*?)\}')

        # Publisher Style Patterns
        regex_eg = re.compile(r'\be\.g\.', re.IGNORECASE)
        regex_vs = re.compile(r'\bvs\.', re.IGNORECASE)
        regex_title_colon = re.compile(r'\\(section|subsection|subsubsection|paragraph)\{([^}]+):\s*\}')
        regex_caption_period = re.compile(r'(\\caption\{((?:[^{}]|\{[^{}]*\})*))\.\s*\}')
        regex_fig_ref_explicit = re.compile(r'\b(Figure|Table)\s+(\d+\.\d+)')
        regex_fig_ref_latex = re.compile(r'\b(Figure|Table)(~|\s+)(\\ref\{[^}]+\})')

        available_images = {}
        # heuristic: match on "figure_d_d"
        prefix_map = {}
        prefix_pattern = re.compile(r'^(figure_\d+_\d+)')

        for f in dir_files:
            name, ext = os.path.splitext(f)
            lower_name = name.lower()
            available_images[lower_name] = f
            
            pmatch = prefix_pattern.match(lower_name)
            if pmatch:
                prefix = pmatch.group(1)
                prefix_map[prefix] = f

        def find_target_image(ref_path):
            ref_basename = os.path.basename(ref_path)
            ref_name_no_ext, _ = os.path.splitext(ref_basename)
            lower_ref_name = ref_name_no_ext.lower()
            
            target = available_images.get(lower_ref_name)
            if not target:
                pmatch = prefix_pattern.match(lower_ref_name)
                if pmatch and pmatch.group(1) in prefix_map:
                    target = prefix_map[pmatch.group(1)]
            return target

        figure_block_pattern = re.compile(
            r'(\\begin\{figure\}(?:.|\n)*?\\end\{figure\})(\s*%\s*Image Prompt:[^\n]*)?', 
            re.IGNORECASE | re.MULTILINE
        )

        def process_figure_block(match):
            full_block = match.group(1)
            prompt_comment = match.group(2) or ""
            
            g_match = graphics_pattern.search(full_block)
            if not g_match:
                return match.group(0) 
            
            options = g_match.group(1)
            ref_path = g_match.group(2)
            
            target_file = find_target_image(ref_path)
            
            if target_file:
                new_tag = f'\\includegraphics[{options}]{{{target_file}}}' if options else f'\\includegraphics{{{target_file}}}'
                s, e = g_match.span()
                new_block = full_block[:s] + new_tag + full_block[e:]
                return new_block + prompt_comment
            else:
                combined = full_block + prompt_comment
                commented = "\n".join([f"% {line}" for line in combined.split('\n')])
                return f"\n% [MISSING IMAGE: {ref_path}] - Block Corrected\n{commented}\n"

        def resolve_inline_image(match):
            options = match.group(1)
            ref_path = match.group(2)
            target = find_target_image(ref_path)
            if target:
                 return f'\\includegraphics[{options}]{{{target}}}' if options else f'\\includegraphics{{{target}}}'
            else:
                 return match.group(0)

        for file_path in chapter_files:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                cleaned_content = citation_pattern.sub('', content)
                cleaned_content = figure_block_pattern.sub(process_figure_block, cleaned_content)
                cleaned_content = graphics_pattern.sub(resolve_inline_image, cleaned_content)

                # Publisher Style Replacements
                cleaned_content = regex_eg.sub("for example", cleaned_content)
                cleaned_content = regex_vs.sub("versus", cleaned_content)
                cleaned_content = regex_title_colon.sub(r'\\\1{\2}', cleaned_content)
                cleaned_content = regex_caption_period.sub(r'\1}', cleaned_content)
                
                # Italicize Figure/Table references
                # Explicit: Figure 1.1 -> \textit{Figure 1.1}
                cleaned_content = regex_fig_ref_explicit.sub(lambda m: f"\\textit{{{m.group(1)} {m.group(2)}}}", cleaned_content)
                # Latex Ref: Figure~\ref{...} -> \textit{Figure~\ref{...}}
                cleaned_content = regex_fig_ref_latex.sub(lambda m: f"\\textit{{{m.group(1)}{m.group(2)}{m.group(3)}}}", cleaned_content)

                cleaned_content = cleaned_content.replace("``", '"').replace("''", '"')
                cleaned_content = cleaned_content.replace("—", "-")
                cleaned_content = cleaned_content.replace("**", "")

                fd, temp_path = tempfile.mkstemp(suffix='.tex', text=True)
                with os.fdopen(fd, 'w', encoding='utf-8') as tf:
                    tf.write(cleaned_content)
                
                cleaned_chapter_files.append(temp_path)
                temp_files.append(temp_path)
                
                print(f"    Processed {os.path.basename(file_path)}: {len(cleaned_content)} chars")
                
            except Exception as e:
                print(f"  Error processing file {file_path}: {e}")
                for t in temp_files: 
                    try: os.remove(t) 
                    except: pass
                continue

        # Custom Title Page for Publisher Style
        # Right aligned, specific text structure to easily style in post-processing
        title_content = (
            f"\\begin{{flushright}}\n"
            f"CHAPTER {chapter_num}\n"
            f"\\par\n" 
            f"\\vspace{{0.5cm}}\n"
            f"{chapter_title}\n"
            f"\\end{{flushright}}\n"
            f"\\thispagestyle{{empty}}\n" 
            f"\\newpage\n"
            f"\\tableofcontents\n"
            f"\\newpage\n"
            f"\\chapter{{{chapter_title}}}\n" 
        )

        fd_title, title_path = tempfile.mkstemp(suffix='.tex', text=True)
        with os.fdopen(fd_title, 'w', encoding='utf-8') as tf:
            tf.write(title_content)
        temp_files.append(title_path)
        
        final_files = [title_path] + cleaned_chapter_files
        
        sanitized_title = "".join(c for c in chapter_title if c.isalnum() or c in (' ', '_', '-')).strip()
        sanitized_title = sanitized_title.replace(" ", "_")
        
        output_filename = f"C{chapter_num:02d}_{sanitized_title}.docx"
        output_path = os.path.join(output_dir, output_filename)
        
        abs_chapter_dir = os.path.abspath(chapter_dir)
        resource_path = f"{abs_chapter_dir};{os.path.join(abs_chapter_dir, 'images')}"
        
        extra_args = [
            f'--resource-path={resource_path}',
            '--top-level-division=chapter'
        ]
        
        print(f"  Combining {len(final_files)} files into {output_path}...")
        
        debug_filename = output_filename.replace('.docx', '.tex')
        debug_tex_path = os.path.join(output_dir, debug_filename)
        with open(debug_tex_path, 'w', encoding='utf-8') as debug_f:
            for fpath in final_files:
                with open(fpath, 'r', encoding='utf-8') as part_f:
                    debug_f.write(part_f.read() + "\n")
        
        try:
            pypandoc.convert_file(
                debug_tex_path,
                'docx',
                format='latex',
                outputfile=output_path,
                extra_args=extra_args
            )
            print(f"  Successfully created {output_path}")
        except Exception as e:
            print(f"  Error: {e}")
            sys.exit(1)
        finally:
            for t in temp_files:
                try: os.remove(t)
                except: pass
        
        return output_path

def post_process_docx(docx_path):
    """
    Applies publisher styles to the generated DOCX file.
    """
    try:
        from docx import Document
        from docx.shared import Pt, RGBColor, Inches, Mm
        from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    except ImportError:
        print("  Warning: python-docx not installed. Skipping post-processing style application.")
        return

    print(f"  Applying Publisher Styles to {docx_path}...")
    doc = Document(docx_path)

    # 1. Page Layout (A4, 0.79" margins)
    # A4 size: 210mm x 297mm
    section = doc.sections[0]
    section.page_height = Mm(297)
    section.page_width = Mm(210)
    
    margin_size = Inches(0.79)
    section.top_margin = margin_size
    section.bottom_margin = margin_size
    section.left_margin = margin_size
    section.right_margin = margin_size

    # 1.5 Fix Numbering Consistency (Force Tab suffix for all levels)
    try:
        if hasattr(doc.part, 'numbering_part') and doc.part.numbering_part:
            numbering_element = doc.part.numbering_part.element
            # Iterate through all abstractNum definitions and their levels
            # We want to enforce <w:suff w:val="tab"/>
            from docx.oxml.ns import qn
            from docx.oxml import OxmlElement
            
            # Find all <w:lvl> elements (levels in abstract numbering)
            for lvl in numbering_element.xpath('.//w:lvl'):
                # Find or create <w:suff>
                suff = lvl.find(qn('w:suff'))
                if suff is None:
                    suff = OxmlElement('w:suff')
                    lvl.append(suff)
                suff.set(qn('w:val'), 'space') # Changed from 'tab' to 'space' per user request
            print("  Enforced Space suffix for numbering levels.")
    except Exception as e:
        print(f"  Warning: Could not patch numbering XML: {e}")

    # 1.6 Add Page Numbers to Footer
    # We need to add a footer to the section and insert a PAGE field.
    # Simple page number centering.
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # We need to insert the page number field using Oxml
    # <w:fldSimple w:instr=" PAGE "/>
    fldSimple = OxmlElement('w:fldSimple')
    fldSimple.set(qn('w:instr'), ' PAGE ')
    # Retrieve the run to put it in, or just append to paragraph
    run = footer_para.add_run()
    run.font.name = "Lora"
    run.font.size = Pt(10)
    # Append the field to the paragraph element, wrapped in a run? 
    # Actually fldSimple contains runs. 
    # Standard way:
    run._element.append(fldSimple)

    def update_style(style_name, font_name="Lora", font_size=11, bold=False, italic=False, 
                     space_before=0, space_after=10, align=None, color_rgb=RGBColor(0,0,0)):
        try:
            style = doc.styles[style_name]
            # Font properties exist on both Character and Paragraph styles (mostly)
            if hasattr(style, 'font'):
                style.font.name = font_name
                # Force rFonts to ensure Theme fallback (like Aptos) doesn't win
                try:
                    rPr = style.element.get_or_add_rPr()
                    # Creating w:rFonts element
                    rFonts = rPr.find(qn('w:rFonts'))
                    if rFonts is None:
                        rFonts = OxmlElement('w:rFonts')
                        rPr.append(rFonts)
                    # Set all types to the target font
                    rFonts.set(qn('w:ascii'), font_name)
                    rFonts.set(qn('w:hAnsi'), font_name)
                    rFonts.set(qn('w:eastAsia'), font_name) # Added
                    rFonts.set(qn('w:cs'), font_name)
                    rFonts.set(qn('w:asciiTheme'), '') # Clear theme bindings
                    rFonts.set(qn('w:hAnsiTheme'), '')
                    rFonts.set(qn('w:eastAsiaTheme'), '')
                    rFonts.set(qn('w:cstheme'), '')
                except Exception as xml_e:
                    print(f"    Warning setting rFonts for {style_name}: {xml_e}")

                style.font.size = Pt(font_size)
                style.font.bold = bold
                style.font.italic = italic
                style.font.color.rgb = color_rgb
            
            # Paragraph formatting only tests on Paragraph styles
            if hasattr(style, 'paragraph_format'):
                style.paragraph_format.space_before = Pt(space_before)
                style.paragraph_format.space_after = Pt(space_after)
                style.paragraph_format.line_spacing = 1.15
                if align is not None:
                    style.paragraph_format.alignment = align
        except KeyError:
            pass

    # Update Styles
    # Normal: Lora, 11pt, Justified
    update_style('Normal', font_name="Lora", font_size=11, space_after=10, align=WD_PARAGRAPH_ALIGNMENT.JUSTIFY)
    
    # Heading 1 (General use in doc, maybe for Chapter Title repeats): Lora 24 Bold
    update_style('Heading 1', font_name="Lora", font_size=24, bold=True, space_before=24, space_after=24)
    
    # Heading 2 (Section): Lora 20 Bold
    # Increased spacing per user request: "start of new section add space before... end of each section add space after"
    # Interpreted as increased margins around the header.
    update_style('Heading 2', font_name="Lora", font_size=20, bold=True, space_before=42, space_after=18)
    
    # Heading 3 (Subsection): Lora 18 Bold (User request: 18pt)
    update_style('Heading 3', font_name="Lora", font_size=18, bold=True, space_before=24, space_after=12, align=WD_PARAGRAPH_ALIGNMENT.JUSTIFY)

    # Heading 4 (Sub-subsection): Lora 16 Bold (User request: 16pt)
    update_style('Heading 4', font_name="Lora", font_size=16, bold=True, space_before=12, space_after=12, align=WD_PARAGRAPH_ALIGNMENT.JUSTIFY)

    # Heading 5: Lora 14 Bold (User request: 14pt)
    update_style('Heading 5', font_name="Lora", font_size=14, bold=True, space_before=12, space_after=12, align=WD_PARAGRAPH_ALIGNMENT.JUSTIFY)

    # Code Blocks (Source Code): Left Aligned
    # Pandoc usually maps verbatims to "Source Code"
    update_style('Source Code', font_name="Consolas", font_size=10, align=WD_PARAGRAPH_ALIGNMENT.LEFT)
    update_style('VerbatimChar', font_name="Consolas", font_size=10, align=WD_PARAGRAPH_ALIGNMENT.LEFT)

    # Captions: Lora 9, Italic, Center
    update_style('Caption', font_name="Lora", font_size=9, italic=True, align=WD_PARAGRAPH_ALIGNMENT.CENTER)
    
    # Also Figure/Table-specific captions often use specific styles or direct formatting.
    # We can try to catch 'Figure Caption' or 'Table Caption' if they exist.
    update_style('Figure Caption', font_name="Lora", font_size=9, italic=True, align=WD_PARAGRAPH_ALIGNMENT.CENTER)
    update_style('Table Caption', font_name="Lora", font_size=9, italic=True, align=WD_PARAGRAPH_ALIGNMENT.CENTER)

    # 2. Custom Title Page Formatting ("CHAPTER X" and Title)
    # R54, G95, B145 Color (User correction: "Red=54, Green=95,Blue=145")
    accent_color = RGBColor(54, 95, 145)

    # We also need to remove the DUPLICATE chapter title that Pandoc generates.
    # The structure is: Custom Title Page (Heading 1 or Normal) -> ... -> \chapter{Title} (Heading 1)
    # Let's clean up the duplicate first.
    # Attempt to find "Heading 1" paragraphs that match the Chapter Title text (assumed known or guessed).
    # Since we don't have the title string here, we look for "Heading 1" paragraphs that appear AFTER the custom title.
    # Our custom title uses explicit formatting (Text 'CHAPTER X' and 'Title'). 
    # The duplicate usually immediately follows the custom title page content.
    
    # Let's perform the formatting of the CUSTOM title page first.
    found_chapter_num = False
    paragraphs_to_remove = []
    
    for p in doc.paragraphs[:20]:
        txt = p.text.strip()
        if not txt: continue
        
        # Check for our custom "CHAPTER X"
        if re.match(r"^CHAPTER \d+$", txt, re.IGNORECASE) and not found_chapter_num:
            # Format Chapter Num
            p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
            for run in p.runs:
                run.font.name = "Lora"
                run.font.size = Pt(35)
                run.font.bold = False
                run.font.color.rgb = accent_color
            found_chapter_num = True
            continue
        
        if found_chapter_num:
            # The next non-empty paragraph IS our custom title.
            # Format Title: Lora, 40, Bold, Right
            p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
            for run in p.runs:
                run.font.name = "Lora"
                run.font.size = Pt(40)
                run.font.bold = True
                run.font.color.rgb = accent_color
            found_chapter_num = False # Done formatting custom title
            
            # NOW: We expect a Duplicate Title to appear shortly after.
            # It will be a Heading 1 with the same text (roughly) or "1 Title".
            # We will mark the NEXT Heading 1 for removal if it seems redundant.
            # Let's look ahead.
            continue
            
    # Iterate again to removing duplicate Heading 1s
    # Strategy: Remove the FIRST Heading 1 found in the document? 
    # Pandoc `\chapter` creates the first structural Heading 1.
    # Our custom text was likely "Right Aligned" but might have "Normal" style or similar.
    # If we remove the FIRST Heading 1, we remove the structural node.
    # Better: Change the "Structure" Heading 1 to be Hidden? Or just remove text?
    # User asked to "do not need this duplicate line".
    # We will remove the paragraph.
    
    first_h1_removed = False
    for p in doc.paragraphs:
        if p.style.name == 'Heading 1' and not first_h1_removed:
            # Check if this is likely the duplicate
            # It usually contains the Chapter Title.
            # We'll just delete it.
            # deleting a paragraph in python-docx: get parent and remove child
            p._element.getparent().remove(p._element)
            print("  Removed duplicate Heading 1 paragraph.")
            first_h1_removed = True
            break

    # Pre-pass: Identify the LAST "Conclusion" subsection to promote
    target_conclusion_index = -1
    for i, p in enumerate(doc.paragraphs):
        if p.style.name == 'Heading 3':
            # Match "1.2.3 Conclusion" or "Conclusion" with spaces/tabs
            if re.search(r'^\s*[\d\.]+\s+Conclusion\s*$', p.text, re.IGNORECASE) or p.text.strip().lower() == "conclusion":
                target_conclusion_index = i

    # 3. Center Align Images and handle captions if manual
    # Also iterate to apply Aggressive Font/Spacing Fixes for Headings
    for i, paragraph in enumerate(doc.paragraphs):
        # Check for images (drawings)
        if 'w:drawing' in paragraph._element.xml or 'w:pict' in paragraph._element.xml:
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            
        # Check if style name contains "Caption"
        if "Caption" in paragraph.style.name:
             paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
             for run in paragraph.runs:
                 run.font.name = "Lora"
                 run.font.size = Pt(9)
                 run.font.italic = True
        
        # Special Case: Promote "Conclusion" subsection to Section (Unnumbered)
        # Only for the LAST instance found in the chapter
        if i == target_conclusion_index:
            print(f"  Promoting Conclusion subsection: {paragraph.text.strip()}")
            paragraph.style = doc.styles['Heading 2']
            paragraph.text = "Conclusion"
            try:
                if paragraph._p.pPr is not None and paragraph._p.pPr.numPr is not None:
                     paragraph._p.pPr.remove(paragraph._p.pPr.numPr)
            except Exception:
                pass
        
        # Aggressive Heading Fixes (Spacing and Font)
        if paragraph.style.name in ['Heading 2', 'Heading 3', 'Heading 1', 'Heading 4']:
            # 1. Fix Tab Spacing in Text (Manually numbered by Pandoc?)
            if '\t' in paragraph.text:
                # Replace tab with space
                # We do this by modifying p.text, which consolidates runs. 
                # This is okay because we will re-apply font formatting immediately after.
                paragraph.text = paragraph.text.replace('\t', ' ')
                print(f"  Fixed spacing (Tab->Space) in: {paragraph.text[:30]}...")
            
            # 2. Force Font on Runs (Aptos Override)
            for run in paragraph.runs:
                run.font.name = "Lora"
                # Re-apply bold?
                run.font.bold = True
                
                if paragraph.style.name == 'Heading 1':
                    run.font.size = Pt(24)
                elif paragraph.style.name == 'Heading 2':
                    run.font.size = Pt(20)
                elif paragraph.style.name == 'Heading 3':
                    run.font.size = Pt(18)
                elif paragraph.style.name == 'Heading 4':
                    run.font.size = Pt(16)
                elif paragraph.style.name == 'Heading 5':
                    run.font.size = Pt(14)
                    
                # Explicitly clear Theme fonts in XML if possible, 
                # but run.font.name usually writes w:rFonts w:ascii="Lora" etc.
                if hasattr(run.element.rPr, 'rFonts'):
                     rFonts = run.element.rPr.rFonts
                     if rFonts is not None:
                         rFonts.set(qn('w:asciiTheme'), '')
                         rFonts.set(qn('w:hAnsiTheme'), '')
                         rFonts.set(qn('w:eastAsiaTheme'), '')
                         rFonts.set(qn('w:cstheme'), '')

    try:
        doc.save(docx_path)
        print("  Publisher Styles applied successfully.")
    except Exception as e:
        print(f"  Error saving styled DOCX: {e}")

def main():
    parser = argparse.ArgumentParser(description="Convert LaTeX to Publisher Style Docx.")
    parser.add_argument("--chapter", type=int, help="Specific chapter number to convert.")
    args = parser.parse_args()

    metadata_file = "input/metadata.json"
    output_dir = "output"
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    output_path = convert_book(metadata_file, output_dir, args.chapter)
    
    if output_path:
        post_process_docx(output_path)

if __name__ == "__main__":
    main()

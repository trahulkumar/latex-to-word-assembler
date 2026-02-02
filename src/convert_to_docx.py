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
            
            # Find a file that contains "section {section_num}" case-insensitive
            # Matches identifiers like "Section 1.1", "section 1.1", "Section_1.1"
            # We treat spaces and underscores flexibly
            found_file = None
            
            # specific pattern: "section" followed by space/underscore/dot then the number
            # simplistic check: filename must contain "section" and the number
            # But "Section 1.10" contains "1.1". We need to be careful. 
            # Let's use regex for robust matching.
            # Pattern: section[ ._]*1\.1(\D|$) to ensure 1.1 isn't 1.10
            
            pattern = re.compile(rf"section[\s._]*{re.escape(str(section_num))}(\D|$)", re.IGNORECASE)
            
            for fname in dir_files:
                if pattern.search(fname):
                    found_file = os.path.join(chapter_dir, fname)
                    break 
            
            if not found_file:
                print(f"ERROR: Missing file for Section {section_num} in {chapter_dir}")
                print(f"       Expected file containing 'Section {section_num}'")
                sys.exit(1) # Strict error as requested
                
            print(f"  Found: {os.path.basename(found_file)}")
            chapter_files.append(found_file)
            
        if not chapter_files:
            print(f"  No valid files found for Chapter {chapter_num}")
            continue

        # Pre-process files to remove citations/references and fix image paths
        temp_files = []
        cleaned_chapter_files = []
        
        # Regex to remove \cite{...}, \citep{...}, \citet{...}, \ref{...}
        # Also remove [cite: ...] and [cite_start] found in input txt files
        citation_pattern = re.compile(r'\\(cite|citep|citet|ref)\{[^}]+\}|\[cite:[^\]]+\]|\[cite_start\]')
        # Regex for \includegraphics[...]{...} or \includegraphics{...}
        # Captures: 1=options (optional), 2=path
        graphics_pattern = re.compile(r'\\includegraphics(?:\[(.*?)\])?\{(.*?)\}')

        # Create a map of basename (no ext) -> actual filename for all files in chapter dir
        # This allows fuzzy matching of images (e.g. asking for .png but having .jpg)
        available_images = {}
        # Also create a prefix map for fuzzy matching (e.g. "figure_1_4" -> "figure_1_4_spiral.jpg")
        # heuristic: match on "figure_d_d"
        prefix_map = {}
        prefix_pattern = re.compile(r'^(figure_\d+_\d+)')

        for f in dir_files:
            name, ext = os.path.splitext(f)
            lower_name = name.lower()
            available_images[lower_name] = f
            
            # Populate prefix map
            pmatch = prefix_pattern.match(lower_name)
            if pmatch:
                prefix = pmatch.group(1)
                # Only map if unique or overwrite (simple approach: last wins or first wins)
                # Given the structure, unique per section is expected.
                prefix_map[prefix] = f

        # Define helper to find image
        def find_target_image(ref_path):
            ref_basename = os.path.basename(ref_path)
            ref_name_no_ext, _ = os.path.splitext(ref_basename)
            lower_ref_name = ref_name_no_ext.lower()
            
            # 1. Exact match
            target = available_images.get(lower_ref_name)
            # 2. Fuzzy match
            if not target:
                pmatch = prefix_pattern.match(lower_ref_name)
                if pmatch and pmatch.group(1) in prefix_map:
                    target = prefix_map[pmatch.group(1)]
            return target

        # Regex for figure block + optional prompt comment
        # Matches: \begin{figure} ... \end{figure} + optional (% Image Prompt: ...)
        figure_block_pattern = re.compile(
            r'(\\begin\{figure\}(?:.|\n)*?\\end\{figure\})(\s*%\s*Image Prompt:[^\n]*)?', 
            re.IGNORECASE | re.MULTILINE
        )

        def process_figure_block(match):
            full_block = match.group(1)
            prompt_comment = match.group(2) or ""
            
            # Find includegraphics inside this block
            g_match = graphics_pattern.search(full_block)
            if not g_match:
                return match.group(0) # No image, leave alone
            
            options = g_match.group(1)
            ref_path = g_match.group(2)
            
            target_file = find_target_image(ref_path)
            
            if target_file:
                # Found: Update the path inside the block
                new_tag = f'\\includegraphics[{options}]{{{target_file}}}' if options else f'\\includegraphics{{{target_file}}}'
                s, e = g_match.span()
                new_block = full_block[:s] + new_tag + full_block[e:]
                return new_block + prompt_comment
            else:
                # Missing: Comment out the ENTIRE block and prompt
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
                
                # Remove citations
                cleaned_content = citation_pattern.sub('', content)
                
                # Apply figure block processing first
                cleaned_content = figure_block_pattern.sub(process_figure_block, cleaned_content)

                cleaned_content = graphics_pattern.sub(resolve_inline_image, cleaned_content)
                
                # Sanitize LaTeX quotes `` and '' to simple "
                cleaned_content = cleaned_content.replace("``", '"').replace("''", '"')
                
                # Replace em dashes — with hyphens -
                cleaned_content = cleaned_content.replace("—", "-")
                
                # Remove LLM-style bold markers **
                cleaned_content = cleaned_content.replace("**", "")

                # Create a localized temp file
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

        # Create a formatted Title Page and Table of Contents
        # We manually structure this to ensure page breaks and correct order
        # We use \setcounter{chapter}{X-1} then \chapter so it numbers correctly as X? 
        # Actually Pandoc default is 1. If we want Chapter 5, we might need adjustments.
        # But for now, user inputs chapter number.
        
        # Note: \chapter will output "Chapter X" in standard styling. 
        # User wants large fonts. We provide a custom title page FIRST.
        
        title_content = (
            f"\\begin{{center}}\n"
            f"\\Huge \\textbf{{Chapter {chapter_num}}}\n"
            f"\\vspace{{1cm}}\n"
            f"\\Huge \\textbf{{{chapter_title}}}\n"
            f"\\end{{center}}\n"
            f"\\thispagestyle{{empty}}\n" # No page number on title
            f"\\newpage\n"
            f"\\tableofcontents\n"
            f"\\newpage\n"
            # We still need the structural \chapter for section numbering (1.1, etc.)
            # But we might hide it or accept it repeats the title.
            # To avoid "0.1", we ensure this comes before sections.
            f"\\chapter{{{chapter_title}}}\n" 
        )

        fd_title, title_path = tempfile.mkstemp(suffix='.tex', text=True)
        with os.fdopen(fd_title, 'w', encoding='utf-8') as tf:
            tf.write(title_content)
        temp_files.append(title_path)
        
        # Prepend title file to cleaned files
        final_files = [title_path] + cleaned_chapter_files
        
        # Sanitize title for filename
        sanitized_title = "".join(c for c in chapter_title if c.isalnum() or c in (' ', '_', '-')).strip()
        sanitized_title = sanitized_title.replace(" ", "_")
        
        # Output filename format: C01_Chapter_Name.docx
        output_filename = f"C{chapter_num:02d}_{sanitized_title}.docx"
        output_path = os.path.join(output_dir, output_filename)
        
        # Resource path using absolute paths
        abs_chapter_dir = os.path.abspath(chapter_dir)
        resource_path = f"{abs_chapter_dir};{os.path.join(abs_chapter_dir, 'images')}"
        
        # Extra args: 
        # - Removed '--toc' (added manually via \tableofcontents)
        # - Added --top-level-division=chapter (structural)
        # - Added --number-sections
        extra_args = [
            '--number-sections', 
            f'--resource-path={resource_path}',
            '--top-level-division=chapter'
        ]
        
        print(f"  Combining {len(final_files)} files (1 title + {len(cleaned_chapter_files)} sections) into {output_path}...")
        
        # DEBUG: Dump combined content to file
        debug_filename = output_filename.replace('.docx', '.tex')
        debug_tex_path = os.path.join(output_dir, debug_filename)
        with open(debug_tex_path, 'w', encoding='utf-8') as debug_f:
            for fpath in final_files:
                with open(fpath, 'r', encoding='utf-8') as part_f:
                    debug_f.write(part_f.read() + "\n")
        print(f"  [DEBUG] Saved combined LaTeX to {debug_tex_path}")

        try:
            pypandoc.convert_file(
                debug_tex_path,
                'docx',
                format='latex',
                outputfile=output_path,
                extra_args=extra_args
            )
            print(f"  Successfully created {output_path}")
        except RuntimeError as e:
            print(f"  Pandoc Error: {e}")
            sys.exit(1)
        except Exception as e:
            print(f"  Error: {e}")
            sys.exit(1)
        finally:
            # Cleanup temp files
            for t in temp_files:
                try:
                    os.remove(t)
                except OSError:
                    pass
        
        return output_path

def post_process_docx(docx_path):
    """
    Applies consistent styling to the generated DOCX file using python-docx.
    """
    try:
        from docx import Document
        from docx.shared import Pt, RGBColor
        from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    except ImportError:
        print("  Warning: python-docx not installed. Skipping post-processing style application.")
        return

    print(f"  Applying styles to {docx_path}...")
    doc = Document(docx_path)

    def update_style(style_last_name, font_name="Arial", font_size=11, bold=False, space_before=0, space_after=0):
        # style_last_name e.g. 'Heading 1'
        try:
            style = doc.styles[style_last_name]
            style.font.name = font_name
            style.font.size = Pt(font_size)
            style.font.bold = bold
            style.font.color.rgb = RGBColor(0, 0, 0) # Black
            style.paragraph_format.space_before = Pt(space_before)
            style.paragraph_format.space_after = Pt(space_after)
        except KeyError:
            pass

    # Heading 1: Chapter Title
    update_style('Heading 1', font_size=24, bold=True, space_before=0, space_after=24)
    
    # Heading 2: Section (1.1, 1.2)
    update_style('Heading 2', font_size=18, bold=True, space_before=24, space_after=12)
    
    # Heading 3: Subsection (1.1.1)
    update_style('Heading 3', font_size=14, bold=True, space_before=18, space_after=6)

    # Normal text
    update_style('Normal', font_size=11, space_after=6)
    
    # Iterate through paragraphs to ensure direct formatting doesn't override style
    # functionality limited here, but main work is done via styles.
    
    # Also Center-Align images
    # Images in docx are InlineShapes. We iterate paragraphs and check for drawing elements.
    for paragraph in doc.paragraphs:
        # Check for images (drawings or pictures)
        # Namespace map might be needed if strictly parsing XML, but searching for tag name in XML string is easier or using xpath
        if 'w:drawing' in paragraph._element.xml or 'w:pict' in paragraph._element.xml:
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    try:
        doc.save(docx_path)
        print("  Styles applied successfully.")
    except Exception as e:
        print(f"  Error saving styled DOCX: {e}")

def main():
    parser = argparse.ArgumentParser(description="Convert LaTeX chapters to Docx.")
    parser.add_argument("--chapter", type=int, help="Specific chapter number to convert (e.g. 1)")
    args = parser.parse_args()

    metadata_file = "input/metadata.json"
    output_dir = "output"
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    output_path = convert_book(metadata_file, output_dir, args.chapter)
    
    # convert_book needs to return the output_path for us to post-process it
    if output_path:
        post_process_docx(output_path)

if __name__ == "__main__":
    main()

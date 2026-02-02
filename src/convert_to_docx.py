import json
import os
import pypandoc
import sys
import re
import argparse

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

        # Sanitize title for filename
        sanitized_title = "".join(c for c in chapter_title if c.isalnum() or c in (' ', '_', '-')).strip()
        sanitized_title = sanitized_title.replace(" ", "_")
        
        # Output filename format: C01_Chapter_Name.docx
        output_filename = f"C{chapter_num:02d}_{sanitized_title}.docx"
        output_path = os.path.join(output_dir, output_filename)
        
        extra_args = ['--number-sections']
        if target_chapter is None:
            extra_args.append('--toc') 
        
        print(f"  Combining {len(chapter_files)} files into {output_path}...")
        
        try:
            # We explicitly specify format='latex' so pandoc treats .txt files as latex
            pypandoc.convert_file(
                chapter_files,
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

def main():
    parser = argparse.ArgumentParser(description="Convert LaTeX chapters to Docx.")
    parser.add_argument("--chapter", type=int, help="Specific chapter number to convert (e.g. 1)")
    args = parser.parse_args()

    metadata_file = "input/metadata.json"
    output_dir = "output"
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    convert_book(metadata_file, output_dir, args.chapter)

if __name__ == "__main__":
    main()

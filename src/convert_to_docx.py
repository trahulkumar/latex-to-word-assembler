import json
import os
import pypandoc
import sys

def convert_book(metadata_path, output_dir):
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

    all_files_ordered = []
    
    print(f"Starting conversion for '{book_title}'...")
    
    # Process each chapter
    for chapter in chapters:
        chapter_num = chapter['number']
        chapter_title = chapter['title']
        print(f"Processing Chapter {chapter_num}: {chapter_title}")
        
        chapter_files = []
        for section in chapter.get('sections', []):
            file_path = section.get('file_path')
            
            # Check if file exists
            if not os.path.exists(file_path):
                print(f"  WARNING: File not found: {file_path}")
                # Try to find it case-insensitively or with slight variations if needed?
                # For now, strictly erroring or skipping is safer to avoid garbage output.
                # However, let's just skip and warn.
                continue
                
            chapter_files.append(file_path)
            all_files_ordered.append(file_path)
            
        if not chapter_files:
            print(f"  No valid files found for Chapter {chapter_num}")
            continue

        # Option: Convert individual chapters
        # "All these files need to be combined to create a single word docx which will be by chapter"
        # Use pandoc to combine per chapter first if we want chapter-level docs
        # But user wants a single word docx.
        
        # Let's create a chapter header in the combined doc?
        # Pandoc combines files by concatenating. To add "Chapter X" headers, we might need a dynamically generated tex file or header.
        # But simplest MVP is just combining the content.
        # If the partial files don't have \chapter{}, it will just be text. 
        # Usually Section files have \section{}. 
        # We might want to insert a title block.
        
        pass

    if not all_files_ordered:
        print("No files to convert.")
        return

    # Combine all files into one docx
    output_filename = f"{book_title.replace(' ', '_')}.docx"
    output_path = os.path.join(output_dir, output_filename)
    
    # Extra args for pandoc
    # --toc for table of contents could be nice, or reference-doc for styling
    extra_args = ['--toc'] 
    
    print(f"Combining {len(all_files_ordered)} files into {output_path}...")
    
    try:
        output = pypandoc.convert_file(
            all_files_ordered,
            'docx',
            outputfile=output_path,
            extra_args=extra_args
        )
        print(f"Successfully created {output_path}")
    except RuntimeError as e:
        print(f"Pandoc Error: {e}")
    except Exception as e:
        print(f"Error: {e}")

def main():
    metadata_file = "input/metadata.json"
    output_dir = "output"
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    convert_book(metadata_file, output_dir)

if __name__ == "__main__":
    main()

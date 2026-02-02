import os
import json
import re

def parse_metadata(file_path):
    """
    Parses the metadata text file (Markdown format) and returns a structured dictionary.
    
    Expected format:
    **Book Title:** Title
    ### Chapter 1: Chapter Name
    * **1.1 Section Name**
    """
    book_data = {
        "book_title": "Untitled Book",
        "chapters": []
    }
    
    current_chapter = None
    
    if not os.path.exists(file_path):
        print(f"Error: Metadata file not found at {file_path}")
        return None

    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        # Parse Book Title
        if "**Book Title:**" in line:
            book_data["book_title"] = line.split("**Book Title:**", 1)[1].strip()
            
        # Parse Chapter
        # Matches "### Chapter 1: From Notebooks to Systems"
        elif line.startswith("### Chapter"):
            match = re.match(r"### Chapter\s*(\d+)[:\s]*(.*)", line, re.IGNORECASE)
            if match:
                chapter_num = int(match.group(1))
                chapter_title = match.group(2).strip()
                
                current_chapter = {
                    "number": chapter_num,
                    "title": chapter_title,
                    "sections": []
                }
                book_data["chapters"].append(current_chapter)
            else:
                print(f"Warning: Could not parse chapter line: {line}")
                
        # Parse Section
        # Matches "* **1.1 The Identity Crisis...**"
        elif line.startswith("* **"):
            if current_chapter is None:
                # Some lines might look like sections but appear before chapters or in other contexts
                continue
                
            # Regex to capture number and title inside the bold markers
            match = re.match(r"\*\s*\*\*([\d\.]+)\s+(.*?)\*\*", line)
            if match:
                section_num_str = match.group(1)
                section_title = match.group(2).strip()
                
                clean_section_num = section_num_str.rstrip('.')
                
                # File path construction
                file_path = f"input/latex_files/Chapter_{current_chapter['number']}/Section_{clean_section_num}.tex"
                
                section_data = {
                    "number": clean_section_num,
                    "title": section_title,
                    "file_path": file_path
                }
                current_chapter["sections"].append(section_data)
            else:
                 # It might be a list item not meant as a section, or different format
                 pass
                
    return book_data

def main():
    input_file = "input/Master Production Manuscript.txt"
    output_file = "input/metadata.json"
    
    print(f"Reading metadata from {input_file}...")
    book_structure = parse_metadata(input_file)
    
    if book_structure:
        # Ensure output directory exists
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(book_structure, f, indent=4)
        print(f"Successfully generated {output_file}")
        
        # summary
        print(f"Book: {book_structure['book_title']}")
        print(f"Chapters: {len(book_structure['chapters'])}")
        for ch in book_structure['chapters']:
            print(f"  Chapter {ch['number']}: {len(ch['sections'])} sections")

if __name__ == "__main__":
    main()

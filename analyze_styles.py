from docx import Document
import os

def analyze_docx(docx_path):
    if not os.path.exists(docx_path):
        print(f"File not found: {docx_path}")
        return

    doc = Document(docx_path)
    
    # 1. Page Margins (Section properties)
    print("--- Page Margins ---")
    for section in doc.sections:
        # sizes are in Twips (1/1440 inch). 72 pts per inch.
        # Print in inches for clarity
        top = section.top_margin.inches if section.top_margin else "Default"
        bottom = section.bottom_margin.inches if section.bottom_margin else "Default"
        left = section.left_margin.inches if section.left_margin else "Default"
        right = section.right_margin.inches if section.right_margin else "Default"
        print(f"Top: {top}, Bottom: {bottom}, Left: {left}, Right: {right}")
    
    # 2. Styles
    print("\n--- Style Analysis ---")
    # We'll look at styles actually used in paragraphs to get a better sense,
    # plus the style definitions themselves.
    
    unique_styles = set()
    for p in doc.paragraphs:
        unique_styles.add(p.style.name)
        
    print(f"Styles used in document: {unique_styles}")
    
    styles_of_interest = ['Normal', 'Heading 1', 'Heading 2', 'Heading 3', 'Title', 'Subtitle']
    # Add any styles found in the doc that look like headings
    for s in unique_styles:
        if 'Heading' in s and s not in styles_of_interest:
            styles_of_interest.append(s)

    for style_name in styles_of_interest:
        if style_name not in doc.styles:
            continue
            
        style = doc.styles[style_name]
        print(f"\nStyle: {style_name}")
        
        # Font
        font = style.font
        print(f"  Font Name: {font.name}")
        print(f"  Font Size: {font.size.pt if font.size else 'Default'}")
        print(f"  Bold: {font.bold}")
        print(f"  Italic: {font.italic}")
        
        # Paragraph Format
        pf = style.paragraph_format
        print(f"  Space Before: {pf.space_before.pt if pf.space_before else 'Default'}")
        print(f"  Space After: {pf.space_after.pt if pf.space_after else 'Default'}")
        print(f"  Line Spacing: {pf.line_spacing}") 
        print(f"  Alignment: {pf.alignment}")

if __name__ == "__main__":
    analyze_docx(r"d:\AI\Github\agents\Latex-to-docx beautify\input\Sample Chapter-updated.docx")

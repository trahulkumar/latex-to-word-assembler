from docx import Document
from docx.oxml.ns import qn
import zipfile
import re

def inspect_xml(docx_path):
    print(f"Inspecting {docx_path}...")
    doc = Document(docx_path)
    
    # 1. Find a Heading 3 paragraph (Subsection)
    h3_para = None
    for p in doc.paragraphs:
        if p.style.name == 'Heading 3':
            h3_para = p
            break
            
    if not h3_para:
        print("No Heading 3 found.")
        return

    print(f"Found Heading 3: '{h3_para.text}'")
    print("Paragraph XML snippet related to numbering:")
    # Extract numPr
    numPr = h3_para._element.pPr.numPr
    if numPr is not None:
        numId_elem = numPr.find(qn('w:numId'))
        ilvl_elem = numPr.find(qn('w:ilvl'))
        numId = numId_elem.get(qn('w:val')) if numId_elem is not None else "None"
        ilvl = ilvl_elem.get(qn('w:val')) if ilvl_elem is not None else "0"
        print(f"  numId: {numId}")
        print(f"  ilvl: {ilvl}")
        
        # Now let's look at the numbering part directly via zipfile to see raw XML
        # because python-docx abstraction can be opaque
        with zipfile.ZipFile(docx_path, 'r') as z:
            xml_content = z.read('word/numbering.xml').decode('utf-8')
            
        print("\nSearching numbering.xml for configuration...")
        # Find the abstractNumId for this numId
        # <w:num w:numId="X"><w:abstractNumId w:val="Y"/></w:num>
        num_pattern = re.compile(rf'<w:num w:numId="{numId}">.*?<w:abstractNumId w:val="(\d+)"/>', re.DOTALL)
        match = num_pattern.search(xml_content)
        if match:
            abstractNumId = match.group(1)
            print(f"  Maps to abstractNumId: {abstractNumId}")
            
            # Now find the level definition in abstractNum
            # <w:abstractNum w:abstractNumId="Y"> ... <w:lvl w:ilvl="Z"> ... </w:lvl>
            # We want to see if <w:suff w:val="space"/> exists effectively
            
            # Simple regex search for the level in the abstract num
            # This is rough parsing, but sufficient for debug
            abs_pattern = re.compile(rf'<w:abstractNum w:abstractNumId="{abstractNumId}">.*?</w:abstractNum>', re.DOTALL)
            abs_match = abs_pattern.search(xml_content)
            if abs_match:
                abs_block = abs_match.group(0)
                # Find the specific level
                lvl_pattern = re.compile(rf'<w:lvl w:ilvl="{ilvl}">.*?</w:lvl>', re.DOTALL)
                lvl_match = lvl_pattern.search(abs_block)
                if lvl_match:
                    print(f"\n  XML for AbstractNum {abstractNumId} Level {ilvl}:")
                    print(lvl_match.group(0))
                else:
                    print(f"  Could not find Level {ilvl} in AbstractNum {abstractNumId}")
            else:
                print(f"  Could not find definition for AbstractNum {abstractNumId}")
        else:
            print(f"  Could not find w:num definition for numId {numId}")
            
    else:
        print("  No numPr found in paragraph (Manual numbering?)")

    print(f"  Paragraph Text (repr): {repr(h3_para.text)}")
    print("  Runs Inspection:")
    for i, run in enumerate(h3_para.runs):
        print(f"    Run {i}: '{run.text}'")
        font_name = run.font.name
        element_font = "None"
        if run.element.rPr and run.element.rPr.rFonts:
             element_font = run.element.rPr.rFonts.get(qn('w:ascii'))
        print(f"      Font: {font_name} (XML: {element_font})")

if __name__ == "__main__":
    inspect_xml(r"d:\AI\Github\agents\Latex-to-docx beautify\output\C01_From_Notebooks_to_Systems.docx")

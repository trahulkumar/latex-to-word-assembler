import re
import argparse
import os
import sys

def count_words(text):
    """Counts words in a string after stripping most LaTeX commands."""
    # Remove LaTeX comments
    text = re.sub(r'%.*$', '', text, flags=re.MULTILINE)
    
    # Remove core commands but keep their content for some
    # Keep content of: \textbf{...}, \textit{...}, \underline{...}, etc.
    # We'll use a simple regex to remove commands while keeping braced content
    # This is a heuristic and might need refinement
    
    # Remove commands like \cite{...}, \ref{...}, \label{...} completely
    text = re.sub(r'\\(?:cite|ref|label|cite[pt]|url|href)\{[^}]*\}', '', text)
    
    # Keep content of \section{...}, \subsection{...} etc. as they are often counted in the section but we handle them separately
    # Strip common formatting commands but keep content
    text = re.sub(r'\\(?:textbf|textit|underline|emph)\{([^}]*)\}', r'\1', text)
    
    # Remove other commands \command or \command{args}
    # Remove \begin{env}, \end{env}
    text = re.sub(r'\\(?:begin|end|noindent|vfill|hfill|newline|break|pagebreak|clearpage|centering|caption|bibliographystyle|bibliography)\b', '', text)
    
    # Remove commands with arguments that aren't text-heavy
    text = re.sub(r'\\[a-zA-Z]+\{[^}]*\}', '', text)
    
    # Remove commands without arguments
    text = re.sub(r'\\[a-zA-Z]+', '', text)
    
    # Remove math mode content $...$ or $$...$$
    text = re.sub(r'\$\$?.*?\$\$?', ' ', text, flags=re.DOTALL)
    
    # Remove environments like equation, figure, table (heuristic)
    # This is tricky with regex, but we'll try to strip common block environments
    text = re.sub(r'\\begin\{(?:equation|figure|table|align|gather|tabular|enumerate|itemize)\}.*?\\end\{[^}]*\}', ' ', text, flags=re.DOTALL)

    # Count words - specifically words with letters
    words = re.findall(r'\b[a-zA-Z]{2,}\b', text)
    return len(words)

def process_latex(filepath):
    if not os.path.exists(filepath):
        print(f"\n[!] ERROR: File not found at: {filepath}")
        print("[!] Stopping execution.")
        sys.exit(1)

    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # Extract Title
    title_match = re.search(r'\\title\{([^}]*)\}', content)
    paper_title = title_match.group(1).strip() if title_match else "Unknown Title"
    
    print("=" * 73)
    print(f"PAPER TITLE: {paper_title}")
    print("=" * 73)

    # Split into sections
    # Regex to find \section, \subsection, etc. including numeric depth
    # We'll support \section, \subsection, \subsubsection
    section_pattern = re.compile(r'\\(section|subsection|subsubsection)\*?\{([^}]*)\}', re.IGNORECASE)
    
    matches = list(section_pattern.finditer(content))
    
    sections = []
    if not matches:
        sections.append(("Full Document", content))
    else:
        # Text before the first section (Abstract, Title etc.)
        intro_text = content[:matches[0].start()]
        if intro_text.strip():
            sections.append(("Front Matter (Title, Abstract, etc.)", intro_text))
            
        for i in range(len(matches)):
            start = matches[i].end()
            end = matches[i+1].start() if i + 1 < len(matches) else len(content)
            level = matches[i].group(1).capitalize()
            title = matches[i].group(2)
            sections.append((f"{level}: {title}", content[start:end]))

    print(f"{'Section':<60} | {'Word Count':>10}")
    print("-" * 73)
    
    total_words = 0
    for title, text in sections:
        count = count_words(text)
        total_words += count
        # Truncate title if too long but keep it readable
        display_title = title if len(title) <= 60 else title[:57] + "..."
        print(f"{display_title:<60} | {count:>10}")
        
    print("-" * 73)
    print(f"{'TOTAL (Body Text only)':<60} | {total_words:>10}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Count words in a LaTeX file by section.")
    parser.get_default("file")
    parser.add_argument("file", help="Path to the .tex file", nargs='?', default=r"D:\AI\Github\agents\Latex-to-docx beautify\input\paper_journal_latex\manuscript.tex")
    
    args = parser.parse_args()
    process_latex(args.file)

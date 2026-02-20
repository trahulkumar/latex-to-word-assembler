"""sync_listings.py — Inject production code listings into LaTeX files.

Reads the listing registry (docs/listings/chapter_XX.md) from the scripts
repo, extracts code between listing markers in the Python source files,
and replaces matching lstlisting blocks in the LaTeX section files.

Features:
- **Snippet Mode**: Filters code based on the 'Snippet' column in the registry.
  - `core` (default): Strips docstrings and runtime prints; keeps the essence.
  - `full`: Keeps everything (including docstrings).
  - `sig+N`: Keeps signature + N lines.
  - `L5-L10`: Keeps specific line range relative to the marker block.
- **Reference Callouts**: Injects a link to the full GitHub source below each block.

Usage
-----
::

    uv run src/sync_listings.py \
        --scripts-repo "D:\path\to\mastering-predictive-analytics-with-python" \
        --chapter 3 \
        --dry-run

    uv run src/sync_listings.py \
        --scripts-repo "D:\path\to\mastering-predictive-analytics-with-python" \
        --chapter 3

After syncing, run the DOCX conversion as normal:

    uv run src/convert_to_pub_docx.py --chapter 3
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
import textwrap
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path


# ═══════════════════════════════════════════════════════════════════════════
#  Data Structures
# ═══════════════════════════════════════════════════════════════════════════


@dataclass
class ListingEntry:
    """One row from the chapter_XX.md registry table."""

    number: str          # e.g. "3.1"
    section: str         # e.g. "§3.1.1"
    script: str          # e.g. "01_memory_spike.py"
    functions: list[str] # e.g. ["main"]
    title: str           # e.g. "The 5×–10× RAM Rule"
    snippet_mode: str    # e.g. "core", "full", "L5-L10"
    code: str = ""       # extracted source code (populated later)
    script_rel: str = "" # relative path within src/ (populated later)


@dataclass
class LaTeXBlock:
    """A single \\begin{lstlisting}...\\end{lstlisting} block."""

    start_line: int      # 0-indexed line where \begin{lstlisting} is
    end_line: int        # 0-indexed line where \end{lstlisting} is
    caption: str         # parsed caption text (may be empty)
    language: str        # parsed language (default "Python")
    body_lines: list[str] = field(default_factory=list)


# ═══════════════════════════════════════════════════════════════════════════
#  1. Parse the Listing Registry
# ═══════════════════════════════════════════════════════════════════════════

# Regex for markdown link: [`filename`](../../path/to/file)
_LINK_RE = re.compile(r'\[`?([^]`]+)`?\]\([^)]+\)')


def parse_registry(registry_path: Path, chapter: int) -> list[ListingEntry]:
    """Parse docs/listings/chapter_XX.md into ListingEntry objects."""
    text = registry_path.read_text(encoding="utf-8")
    entries: list[ListingEntry] = []

    # Find all table rows (skip header + separator rows)
    for line in text.splitlines():
        line = line.strip()
        if not line.startswith("|"):
            continue
        cells = [c.strip() for c in line.split("|")]
        # Filter out empty cells from leading/trailing pipes
        cells = [c for c in cells if c]
        if len(cells) < 5:
            continue
        # Skip header + separator
        listing_num = cells[0].strip()
        if listing_num.startswith("Listing") or listing_num.startswith("-"):
            continue
        # Must start with the chapter number
        if not listing_num.startswith(str(chapter)):
            continue

        section = cells[1].strip()
        # Extract script filename from markdown link or plain text
        script_cell = cells[2].strip()
        link_match = _LINK_RE.search(script_cell)
        script_name = link_match.group(1) if link_match else script_cell

        functions_raw = cells[3].strip()
        functions = [f.strip().strip("`") for f in functions_raw.split(",")]

        title = cells[4].strip()

        # Parse Snippet column if present (6th column), default to "core"
        snippet_mode = "core"
        if len(cells) >= 6:
            val = cells[5].strip().lower()
            if val:
                snippet_mode = val

        entries.append(ListingEntry(
            number=listing_num,
            section=section,
            script=script_name,
            functions=functions,
            title=title,
            snippet_mode=snippet_mode,
        ))

    if not entries:
        print(f"ERROR: No listing entries found in {registry_path}")
        sys.exit(1)

    print(f"  Parsed {len(entries)} listings from registry")
    return entries


# ═══════════════════════════════════════════════════════════════════════════
#  2. Extract Code from Scripts
# ═══════════════════════════════════════════════════════════════════════════

# Matches: # === Listing 3.1: The 5x-10x RAM Rule === #
_MARKER_START_RE = re.compile(
    r'^# === Listing (\d+\.\d+):\s*(.+?)\s*=== #\s*$'
)
# Matches: # === End Listing 3.1 === #
_MARKER_END_RE = re.compile(
    r'^# === End Listing (\d+\.\d+)\s*=== #\s*$'
)


def extract_listings_from_script(
    script_path: Path,
    chapter: int,
) -> dict[str, str]:
    """Extract code between listing markers in a single Python file.

    Returns a dict mapping listing number (e.g. "3.1") to code string.
    The marker lines themselves are excluded, and the code is dedented.
    """
    lines = script_path.read_text(encoding="utf-8").splitlines()
    results: dict[str, str] = {}

    current_listing: str | None = None
    current_lines: list[str] = []

    for line in lines:
        if current_listing is None:
            m = _MARKER_START_RE.match(line)
            if m and m.group(1).startswith(str(chapter)):
                current_listing = m.group(1)
                current_lines = []
        else:
            m = _MARKER_END_RE.match(line)
            if m and m.group(1) == current_listing:
                # Dedent the extracted code
                code = textwrap.dedent("\n".join(current_lines))
                # Strip leading/trailing blank lines
                code = code.strip("\n")
                results[current_listing] = code
                current_listing = None
                current_lines = []
            else:
                current_lines.append(line)

    return results


def _strip_docstring(code: str) -> str:
    """Remove the leading docstring from a function/class body."""
    # This is a heuristic parser. It assumes the docstring starts
    # inside the first block.
    lines = code.splitlines()
    if not lines:
        return code

    # Check if we are inside a function/class
    first_line = lines[0].strip()
    if not (first_line.startswith("def ") or first_line.startswith("class ") or first_line.startswith("@")):
        # Not a standard function definition, dangerous to auto-strip
        return code

    # Find start of docstring
    doc_start_idx = -1
    quote_char = None
    for i, line in enumerate(lines):
        stripped = line.strip()
        if '"""' in stripped:
            doc_start_idx = i
            quote_char = '"""'
            break
        if "'''" in stripped:
            doc_start_idx = i
            quote_char = "'''"
            break
        # Stop at first code line if it's not a docstring
        if stripped and not stripped.startswith("#") and not (
            stripped.startswith("def ") or stripped.startswith("class ") or
            stripped.startswith("@") or stripped.startswith("async ")
        ):
            # We hit body code before a docstring
            return code

    if doc_start_idx == -1:
        return code

    # Find end of docstring
    # Handle single-line docstrings: """doc"""
    start_line = lines[doc_start_idx]
    if start_line.count(quote_char) >= 2:
        # Check if it ends on the same line
        # Simple check: does the line end with quotes?
        if start_line.strip().endswith(quote_char):
             # It's a one-liner, remove it
             return "\n".join(lines[:doc_start_idx] + lines[doc_start_idx+1:])

    # Multi-line docstring: search for closing quotes
    doc_end_idx = -1
    for i in range(doc_start_idx + 1, len(lines)):
        if quote_char in lines[i]:
            doc_end_idx = i
            break

    if doc_end_idx != -1:
        # Remove the range [doc_start_idx, doc_end_idx]
        return "\n".join(lines[:doc_start_idx] + lines[doc_end_idx+1:])

    return code


def _filter_prints(code: str) -> str:
    """Remove lines that are purely print() statements or their continuations."""
    lines = code.splitlines()
    filtered: list[str] = []
    skip = False
    
    # Simple state machine to catch multi-line print statements
    open_parens = 0
    
    for line in lines:
        stripped = line.strip()
        
        # Check if it looks like a print statement start
        if not skip and stripped.startswith("print("):
            open_parens += line.count("(") - line.count(")")
            if open_parens > 0:
                skip = True
            continue # Skip this line
            
        if skip:
            open_parens += line.count("(") - line.count(")")
            if open_parens <= 0:
                skip = False
                open_parens = 0
            continue # Skip this line
            
        filtered.append(line)
        
    return "\n".join(filtered)


def process_snippet(code: str, mode: str) -> str:
    """Apply snippet filtering rules."""
    mode = mode.lower().strip()
    if mode == "full":
        return code

    lines = code.splitlines()

    if mode.startswith("sig+"):
        # Keep signature + N lines
        try:
            n = int(mode.split("+")[1])
            # Heuristic: Find the end of the definition (colon)
            def_end_idx = 0
            for i, line in enumerate(lines):
                if line.strip().endswith(":"):
                    def_end_idx = i
                    break
            return "\n".join(lines[:def_end_idx + 1 + n]) + "\n    # ..."
        except (IndexError, ValueError):
            pass

    if mode.startswith("l") and "-" in mode:
        # "L5-L10" -> lines 5 to 10 (1-based, relative to marker block)
        try:
            parts = mode.replace("l", "").split("-")
            start = int(parts[0]) - 1
            end = int(parts[1])
            subset = lines[start:end]
            if start > 0:
                subset.insert(0, "# ...")
            if end < len(lines):
                subset.append("# ...")
            return "\n".join(subset)
        except (IndexError, ValueError):
            pass

    if mode == "core":
        # 1. Strip docstrings
        cleaned = _strip_docstring(code)
        # 2. Filter runtime prints (too noisy for book)
        cleaned = _filter_prints(cleaned)
        # 3. Clean up extra newlines
        cleaned = re.sub(r'\n{3,}', '\n\n', cleaned)
        return cleaned.strip()

    return code


def extract_all_listings(
    scripts_dir: Path,
    entries: list[ListingEntry],
    chapter: int,
) -> None:
    """Populate the `code` field on each ListingEntry."""
    # Group entries by script to avoid reading the same file multiple times
    script_entries: dict[str, list[ListingEntry]] = {}
    for entry in entries:
        script_entries.setdefault(entry.script, []).append(entry)

    for script_name, grouped_entries in script_entries.items():
        script_path = scripts_dir / script_name
        if not script_path.exists():
            print(f"  WARNING: Script not found: {script_path}")
            continue

        codes = extract_listings_from_script(script_path, chapter)
        
        # Determine the relative path for display
        rel_parts = script_path.parts
        try:
            src_idx = list(rel_parts).index("src")
            rel_path = "/".join(rel_parts[src_idx + 1:])
        except ValueError:
            rel_path = script_name

        for entry in grouped_entries:
            if entry.number in codes:
                raw_code = codes[entry.number]
                # Apply snippet filtering
                entry.code = process_snippet(raw_code, entry.snippet_mode)
                entry.script_rel = rel_path
                
                print(f"    Listing {entry.number} ({entry.snippet_mode}): "
                      f"kept {len(entry.code.splitlines())} lines "
                      f"(was {len(raw_code.splitlines())})")
            else:
                print(f"  WARNING: Listing {entry.number} marker not found "
                      f"in {script_name}")


# ═══════════════════════════════════════════════════════════════════════════
#  3. Parse LaTeX lstlisting Blocks
# ═══════════════════════════════════════════════════════════════════════════

_LSTLISTING_BEGIN_RE = re.compile(r'\\begin\{lstlisting\}(?:\[(.*?)\])?')
_LSTLISTING_END_RE = re.compile(r'\\end\{lstlisting\}')
_CAPTION_RE = re.compile(r'caption\s*=\s*\{(.+?)\}')
_LANGUAGE_RE = re.compile(r'language\s*=\s*(\w+)')


def parse_lstlisting_blocks(lines: list[str]) -> list[LaTeXBlock]:
    """Find all lstlisting blocks in a list of LaTeX lines."""
    blocks: list[LaTeXBlock] = []
    i = 0
    while i < len(lines):
        m = _LSTLISTING_BEGIN_RE.search(lines[i])
        if m:
            options = m.group(1) or ""
            cap_m = _CAPTION_RE.search(options)
            lang_m = _LANGUAGE_RE.search(options)
            caption = cap_m.group(1) if cap_m else ""
            language = lang_m.group(1) if lang_m else "Python"

            start_line = i
            body: list[str] = []
            i += 1
            while i < len(lines):
                if _LSTLISTING_END_RE.search(lines[i]):
                    blocks.append(LaTeXBlock(
                        start_line=start_line,
                        end_line=i,
                        caption=caption,
                        language=language,
                        body_lines=body,
                    ))
                    break
                body.append(lines[i])
                i += 1
        i += 1
    return blocks


# ═══════════════════════════════════════════════════════════════════════════
#  4. Match Listings to LaTeX Blocks
# ═══════════════════════════════════════════════════════════════════════════


def _caption_similarity(registry_title: str, latex_caption: str) -> float:
    """Compute similarity between a registry title and a LaTeX caption."""
    a = registry_title.lower().strip()
    b = latex_caption.lower().strip()
    return SequenceMatcher(None, a, b).ratio()


def _section_to_latex_section(section_str: str) -> str:
    """Convert registry section like '§3.1.1' to LaTeX section number '3.1'."""
    section_str = section_str.lstrip("§").strip()
    parts = section_str.split(".")
    if len(parts) >= 2:
        return f"{parts[0]}.{parts[1]}"
    return section_str


def _is_big_block(block: LaTeXBlock, min_lines: int = 4) -> bool:
    """Determine if a block is a 'big' code block worth replacing."""
    if block.caption:
        return True
    return len(block.body_lines) >= min_lines


def match_listings_to_blocks(
    entries: list[ListingEntry],
    blocks: list[LaTeXBlock],
    latex_section_num: str,
) -> list[tuple[ListingEntry, LaTeXBlock]]:
    """Match listing entries to LaTeX blocks within a single section file."""
    section_entries = [
        e for e in entries
        if _section_to_latex_section(e.section) == latex_section_num
        and e.code
    ]

    if not section_entries:
        return []

    big_blocks = [b for b in blocks if _is_big_block(b)]
    matches: list[tuple[ListingEntry, LaTeXBlock]] = []
    used_blocks: set[int] = set()

    # Pass 1: Explicit Caption Match (Listing X.X)
    # Allows multiple blocks to match one listing (duplicates)
    for idx, block in enumerate(big_blocks):
        if not block.caption:
            continue
        
        # Try to find matching entry by number
        for entry in section_entries:
            # Check if caption starts with "Listing 3.1" followed by non-digit
            pattern = rf"^Listing {re.escape(entry.number)}(\D|$)"
            if re.match(pattern, block.caption):
                matches.append((entry, block))
                used_blocks.add(idx)
                # Don't break here! Another entry won't match, but we want to know
                # we've greedily claimed this block.
                break 

    # Pass 2: Fuzzy Similarity Match for remaining unmatched entries/blocks
    # Only for entries that haven't been matched yet??
    # Actually, if we have duplicates, maybe we found one but not the other?
    # But usually duplicates have explicit captions.
    # Let's focus on unmatched blocks.
    
    unmatched_entries = [
        e for e in section_entries
        if not any(m[0].number == e.number for m in matches)
    ]
    
    for entry in unmatched_entries:
        best_score = 0.0
        best_block_idx: int | None = None
        for idx, block in enumerate(big_blocks):
            if idx in used_blocks or not block.caption:
                continue
            score = _caption_similarity(entry.title, block.caption)
            if score > best_score:
                best_score = score
                best_block_idx = idx

        if best_block_idx is not None and best_score > 0.35:
            matches.append((entry, big_blocks[best_block_idx]))
            used_blocks.add(best_block_idx)

    # Pass 3: Sequential matching for remaining unmatched entries
    finals_unmatched = [
        e for e in section_entries
        if not any(m[0].number == e.number for m in matches)
    ]
    remaining_blocks = [
        (idx, b) for idx, b in enumerate(big_blocks)
        if idx not in used_blocks
    ]

    for entry, (idx, block) in zip(finals_unmatched, remaining_blocks):
        matches.append((entry, block))
        used_blocks.add(idx)

    return matches


# ═══════════════════════════════════════════════════════════════════════════
#  5. Inject Listings into LaTeX
# ═══════════════════════════════════════════════════════════════════════════


def _build_listing_caption(entry: ListingEntry) -> str:
    """Build the caption string for a listing."""
    func_str = ", ".join(entry.functions)
    # Note: We remove path/function from caption because we add the callout now
    return f"Listing {entry.number} — {entry.title}"


def _build_lstlisting_line(
    entry: ListingEntry,
    language: str = "Python",
) -> str:
    """Build the \\begin{lstlisting}[...] line with updated caption."""
    caption = _build_listing_caption(entry)
    return f"\\begin{{lstlisting}}[language={language}, caption={{{caption}}}]"


def _build_reference_callout(entry: ListingEntry) -> str:
    """Create a LaTeX callout pointing to the companion repo."""
    # Escaping for LaTeX: _ -> \_
    safe_path = entry.script_rel.replace("_", "\\_")
    safe_func = ", ".join(entry.functions).replace("_", "\\_")
    
    return (
        f"\\vspace{{2pt}}\n"
        f"\\noindent\\small\\textit{{%\n"
        f"  \\textbf{{Full Listing {entry.number}}} —\n"
        f"  \\texttt{{{safe_path}::{safe_func}}}\n"
        f"  — companion repository on GitHub.}}\n"
        f"\\normalsize"
    )


def inject_listings(
    lines: list[str],
    matches: list[tuple[ListingEntry, LaTeXBlock]],
) -> list[str]:
    """Replace lstlisting block contents/captions and append reference callout.

    Returns a new list of lines with the replacements applied.
    Works from bottom to top to preserve line indices.
    """
    result = list(lines)  # shallow copy

    # Sort matches by block start_line descending (process from bottom)
    sorted_matches = sorted(
        matches, key=lambda m: m[1].start_line, reverse=True
    )

    for entry, block in sorted_matches:
        # Build replacement parts
        new_header = _build_lstlisting_line(entry, block.language)
        code_lines = entry.code.split("\n")
        callout = _build_reference_callout(entry)
        
        # Determine the lines to replace
        # block.end_line points to \end{lstlisting}
        
        # New block structure:
        # \begin{lstlisting}...
        # ... code ...
        # \end{lstlisting}
        # \vspace... callout ...
        
        # 1. Detect existing callouts immediately following the block
        # Look ahead from block.end_line + 1
        # A callout typically starts with \vspace{2pt} and ends with \normalsize
        # We need to be robust: consume chunks that look like callouts
        
        lines_to_remove = 0
        idx = block.end_line + 1
        
        while idx < len(result):
            # Check for start of callout
            # It usually starts with \vspace{2pt}
            if result[idx].strip() == r"\vspace{2pt}":
                # Check next few lines for "Full Listing" or "companion repository"
                # to confirm it's OUR callout and not random vertical space
                is_our_callout = False
                scan_len = min(6, len(result) - idx)
                chunk = "".join(result[idx : idx + scan_len])
                if "companion repository" in chunk or "Full Listing" in chunk:
                    is_our_callout = True
                
                if is_our_callout:
                    # Consume up to \normalsize
                    found_end = False
                    for offset in range(scan_len + 2): # slightly larger window?
                         if idx + offset < len(result) and result[idx + offset].strip() == r"\normalsize":
                             # Found end of this callout block
                             # Consume this range [idx, idx + offset]
                             # lines_to_remove += (offset + 1)
                             # idx += (offset + 1)
                             # Continue while loop to find NEXT callout if stacked
                             
                             # Actually, simpler: calculate the end index of this callout
                             callout_end = idx + offset
                             lines_to_remove += (callout_end - idx + 1)
                             idx = callout_end + 1
                             found_end = True
                             break
                    
                    if not found_end:
                         # Couldn't find \normalsize, maybe partial? Stop consuming to be safe.
                         break
                else:
                    # Just a random vspace, stop.
                    break
            else:
                # Not a vspace, so not a callout start
                break

        # Calculate replacement list
        replacement = (
            [new_header] + 
            code_lines + 
            [result[block.end_line]] + 
            callout.split("\n")
        )

        # Replace the range [start_line, end_line + lines_to_remove] inclusive
        # If we found callouts, extensions extend the replacement range
        replace_end = block.end_line + lines_to_remove
        
        result[block.start_line : replace_end + 1] = replacement

    return result


# ═══════════════════════════════════════════════════════════════════════════
#  6. Orchestrator
# ═══════════════════════════════════════════════════════════════════════════


def _find_latex_section_files(
    latex_chapter_dir: Path,
    chapter: int,
) -> dict[str, Path]:
    """Map section number (e.g. '3.1') to its LaTeX file path."""
    mapping: dict[str, Path] = {}
    for f in sorted(latex_chapter_dir.iterdir()):
        if not f.is_file():
            continue
        name_lower = f.name.lower()
        m = re.search(
            rf'chapter\s*{chapter}\s*section\s*({chapter}\.\d+)',
            name_lower,
        )
        if m:
            section_num = m.group(1)
            mapping[section_num] = f
    return mapping


def sync_chapter(
    scripts_repo: Path,
    latex_dir: Path,
    chapter: int,
    output_dir: Path | None,
    dry_run: bool = False,
) -> None:
    """Main orchestrator: sync all listings for a chapter."""
    print(f"\n{'='*70}")
    print(f"  Syncing Listings for Chapter {chapter}")
    print(f"{'='*70}\n")

    # ── Locate the registry ─────────────────────────────────────────
    chapter_str = f"{chapter:02d}"
    registry_path = scripts_repo / "docs" / "listings" / f"chapter_{chapter_str}.md"
    if not registry_path.exists():
        print(f"ERROR: Registry not found: {registry_path}")
        sys.exit(1)

    # ── Locate the scripts directory ────────────────────────────────
    src_dir = scripts_repo / "src"
    chapter_dirs = [
        d for d in src_dir.iterdir()
        if d.is_dir() and d.name.startswith(f"chapter_{chapter_str}")
    ]
    if not chapter_dirs:
        print(f"ERROR: No chapter directory found under {src_dir}")
        sys.exit(1)
    scripts_chapter_dir = chapter_dirs[0]
    print(f"  Scripts dir:  {scripts_chapter_dir}")

    # ── Locate LaTeX section files ──────────────────────────────────
    latex_chapter_dir = latex_dir / f"Chapter_{chapter}"
    if not latex_chapter_dir.exists():
        print(f"ERROR: LaTeX chapter dir not found: {latex_chapter_dir}")
        sys.exit(1)
    latex_files = _find_latex_section_files(latex_chapter_dir, chapter)
    print(f"  LaTeX files:  {len(latex_files)} section files found")

    # ── Parse registry ──────────────────────────────────────────────
    entries = parse_registry(registry_path, chapter)

    # ── Extract code from scripts ───────────────────────────────────
    print(f"\n  Extracting code from scripts...")
    extract_all_listings(scripts_chapter_dir, entries, chapter)

    extracted = sum(1 for e in entries if e.code)
    print(f"\n  Extracted: {extracted}/{len(entries)} listings\n")

    # ── Determine output directory ──────────────────────────────────
    if output_dir is None:
        output_dir = latex_chapter_dir  # overwrite in-place
    else:
        output_dir = output_dir / f"Chapter_{chapter}"

    if not dry_run and output_dir != latex_chapter_dir:
        if output_dir.exists():
            shutil.rmtree(output_dir)
        shutil.copytree(latex_chapter_dir, output_dir)
        print(f"  Output dir:   {output_dir}")

    # ── Process each section file ───────────────────────────────────
    total_matches = 0
    for section_num, latex_path in sorted(latex_files.items()):
        lines = latex_path.read_text(encoding="utf-8").splitlines()
        blocks = parse_lstlisting_blocks(lines)

        if not blocks:
            continue

        matches = match_listings_to_blocks(entries, blocks, section_num)

        if not matches:
            continue

        total_matches += len(matches)

        print(f"\n  Section {section_num} ({latex_path.name}):")
        for entry, block in matches:
            caption_preview = block.caption[:50] if block.caption else "(no caption)"
            print(f"    ✓ Listing {entry.number} → block at line {block.start_line + 1} "
                  f'"{caption_preview}"')
            if dry_run:
                # Show a snippet of the new code
                code_preview = entry.code[:100].replace("\n", " ↵ ")
                print(f"      Code: {code_preview}...")
                new_cap = _build_listing_caption(entry)
                print(f"      Caption: {new_cap}")
                print(f"      Ref: Full Listing {entry.number} — ...")

        if not dry_run:
            # Apply replacements
            new_lines = inject_listings(lines, matches)

            # Write to output location
            out_path = output_dir / latex_path.name
            out_path.write_text("\n".join(new_lines), encoding="utf-8")
            print(f"    → Wrote {out_path.name}")

    print(f"\n{'─'*70}")
    print(f"  Total: {total_matches} listings matched and "
          f"{'previewed (dry-run)' if dry_run else 'injected'}")
    print(f"{'─'*70}\n")


# ═══════════════════════════════════════════════════════════════════════════
#  CLI
# ═══════════════════════════════════════════════════════════════════════════


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Sync production code listings into LaTeX files "
                    "before DOCX conversion.",
    )
    parser.add_argument(
        "--scripts-repo",
        type=Path,
        required=True,
        help="Path to the local clone of the scripts repo (read-only).",
    )
    parser.add_argument(
        "--chapter",
        type=int,
        required=True,
        help="Chapter number to sync (e.g. 3).",
    )
    parser.add_argument(
        "--latex-dir",
        type=Path,
        default=Path("input/latex_files"),
        help="LaTeX input directory (default: input/latex_files).",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help="Output directory for synced LaTeX files. "
             "If omitted, modifies the original files in-place.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print matches and previews without writing files.",
    )
    args = parser.parse_args()

    # Validate inputs
    if not args.scripts_repo.exists():
        print(f"ERROR: Scripts repo not found: {args.scripts_repo}")
        sys.exit(1)
    if not args.latex_dir.exists():
        print(f"ERROR: LaTeX directory not found: {args.latex_dir}")
        sys.exit(1)

    sync_chapter(
        scripts_repo=args.scripts_repo,
        latex_dir=args.latex_dir,
        chapter=args.chapter,
        output_dir=args.output_dir,
        dry_run=args.dry_run,
    )


if __name__ == "__main__":
    main()

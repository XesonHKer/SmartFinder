#!/usr/bin/env python3
"""
SmartFinder v9.0.3 — Post-compression font encoding fixer v2.

Simpler approach: instead of injecting ToUnicode CMap via qpdf,
we directly add /Encoding /WinAnsiEncoding to font objects that lack it.
WinAnsiEncoding is a standard PDF encoding that maps bytes 0-255 
to the correct WinAnsi character set. This is safer than ToUnicode
because the PDF viewer already knows WinAnsiEncoding.

For TrueType fonts, /Encoding /WinAnsiEncoding is perfectly valid.
The original MacRomanEncoding was already converted by Ghostscript
to WinAnsi in the font program itself; we just need to declare it.
"""
import re
import os
import sys
import tempfile
import shutil
import subprocess


def fix_encoding(pdf_path):
    """
    Scan the PDF for TrueType fonts missing /Encoding and add /WinAnsiEncoding.
    Uses qpdf --qdf to decompress, edits in text, then recompresses.
    """
    tmpdir = tempfile.mkdtemp()
    try:
        qdf_path = os.path.join(tmpdir, "decompressed.qdf")
        
        # Step 1: Decompress with qpdf
        r = subprocess.run(
            ["qpdf", "--qdf", "--object-streams=disable", pdf_path, qdf_path],
            capture_output=True, timeout=120
        )
        if r.returncode != 0 and r.returncode != 3:
            print(f"qpdf decompress failed (code {r.returncode}): {r.stderr.decode()}")
            return False
        
        # Step 2: Read and fix
        with open(qdf_path, 'rb') as f:
            qdf_data = f.read()
        
        qdf_text = qdf_data.decode('latin-1')
        orig_text = qdf_text
        
        # Find font objects missing Encoding
        # QDF format keeps object structure readable
        # Pattern: 41 0 obj << ... /Subtype /TrueType ... >> endobj
        font_pattern = re.compile(
            r'(\d+ \d+ obj)(.*?)(>>\s*endobj)',
            re.DOTALL
        )
        
        fixed_count = 0
        font_fixes = []
        
        for m in font_pattern.finditer(qdf_text):
            obj_header = m.group(1)
            obj_body = m.group(2)
            obj_tail = m.group(3)
            full_obj = m.group(0)
            
            # Check if it's a font with Widths (indicates it's used)
            if '/Widths' not in obj_body:
                continue
            if '/Type /Font' not in obj_body:
                continue
            if '/Subtype /TrueType' not in obj_body:
                continue
            
            # Skip fonts that already have Encoding or ToUnicode
            if '/Encoding' in obj_body or '/ToUnicode' in obj_body:
                continue
            
            # Extract font name
            fn_m = re.search(r'/BaseFont\s+/([^\s/]+)', obj_body)
            font_name = fn_m.group(1) if fn_m else '?'
            
            # Add /Encoding /WinAnsiEncoding before the closing >>
            # Insert right before ">>" in the object
            new_body = obj_body.rstrip() + '\n  /Encoding /WinAnsiEncoding\n  '
            new_full = obj_header + new_body + obj_tail
            
            font_fixes.append((m.start(), m.end(), new_full, font_name))
            fixed_count += 1
        
        if not font_fixes:
            print("No fonts missing encoding found.")
            return True
        
        # Apply fixes in reverse order
        for start, end, new_full, font_name in reversed(font_fixes):
            qdf_text = qdf_text[:start] + new_full + qdf_text[end:]
            print(f"  Fixed: {font_name}")
        
        if qdf_text == orig_text:
            print("No changes made (text unchanged after processing).")
            return True
        
        # Step 3: Write fixed QDF
        fixed_qdf = os.path.join(tmpdir, "fixed.qdf")
        with open(fixed_qdf, 'w', encoding='latin-1') as f:
            f.write(qdf_text)
        
        # Step 4: Recompress — use qpdf without --linearize to preserve object structure
        tmp_output = os.path.join(tmpdir, "output.pdf")
        r = subprocess.run(
            ["qpdf", "--compress-streams=y", "--recompress-flate", 
             "--object-streams=generate", fixed_qdf, tmp_output],
            capture_output=True, timeout=120
        )
        if r.returncode != 0 and r.returncode != 3:
            print(f"qpdf recompress failed (code {r.returncode}): {r.stderr.decode()[:500]}")
            return False
        
        # Step 5: Verify output exists
        if not os.path.exists(tmp_output) or os.path.getsize(tmp_output) == 0:
            print("qpdf produced empty output")
            return False
        
        # Step 6: Copy back
        shutil.move(tmp_output, pdf_path)
        print(f"\nFixed {fixed_count} font(s) in {pdf_path}")
        return True
    
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 fix_font_encoding.py <compressed_pdf>")
        sys.exit(1)
    
    path = sys.argv[1]
    print(f"Fixing font encoding in: {path}")
    if not os.path.exists(path):
        print(f"File not found: {path}")
        sys.exit(1)
    if fix_encoding(path):
        print("Done!")
    else:
        print("Failed!")
        sys.exit(1)
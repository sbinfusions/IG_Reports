#!/usr/bin/env python3
"""Debug hyperlink extraction from DOCX"""

from docx import Document
import re

doc = Document('Content_Schedule_INPUT/SB_IG_February_Schedule_with_captions.docx')
table = doc.tables[1]  # Posts schedule table

# Print headers
headers = [cell.text.strip() for cell in table.rows[0].cells]
print('Headers:', headers)

# Find photo links column
photo_idx = next((i for i, h in enumerate(headers) if 'photo' in h.lower()), None)
title_idx = next((i for i, h in enumerate(headers) if 'title' in h.lower()), None)
print(f'Title idx: {title_idx}, Photo idx: {photo_idx}')
print()

# Get document relationships
doc_rels = doc.part.rels

# Check first 5 data rows
for row_idx, row in enumerate(table.rows[1:6], 1):
    title = row.cells[title_idx].text.strip()[:30] if title_idx else ''
    photo_text = row.cells[photo_idx].text.strip()[:80] if photo_idx else ''
    cell = row.cells[photo_idx] if photo_idx else None
    
    print(f'=== Row {row_idx}: {title} ===')
    print(f'Cell text: {photo_text}')
    
    if cell:
        xml = cell._tc.xml
        
        # Look for hyperlink refs
        hyperlink_rids = re.findall(r'<w:hyperlink[^>]*r:id="([^"]+)"', xml)
        print(f'Hyperlink rIds: {hyperlink_rids}')
        
        # Resolve rIds to URLs
        for rid in hyperlink_rids[:5]:
            if rid in doc_rels:
                rel = doc_rels[rid]
                if hasattr(rel, 'target_ref'):
                    print(f'  {rid} -> {rel.target_ref[:100]}')
        
        # Also check for any http URLs directly in XML
        direct_urls = re.findall(r'(https?://[^\s<>"]+)', xml)
        if direct_urls:
            print(f'Direct URLs in XML: {direct_urls[:3]}')
    
    print()

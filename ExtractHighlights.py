from docx import Document
from docx.enum.text import WD_COLOR_INDEX  # For color checks if needed

def extract_highlights(docx_path):
    doc = Document(docx_path)
    highlights = []
    current_highlight = ""
    
    for para in doc.paragraphs:
        for run in para.runs:
            if run.font.highlight_color:  # Any highlight
                current_highlight += run.text
            elif current_highlight:
                # Optional: trim and add if it meets criteria (e.g., length)
                highlights.append(current_highlight.strip())
                current_highlight = ""
    
    if current_highlight:
        highlights.append(current_highlight.strip())
    
    return highlights

# Example
highlights = extract_highlights("your_document.docx")
print(highlights)

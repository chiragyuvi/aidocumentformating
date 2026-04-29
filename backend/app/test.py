from docx import Document

def classify_style(style_name):
    style_name = style_name.lower()
    
    if "heading" in style_name:
        return "heading"
    elif "title" in style_name:
        return "title"
    elif "subtitle" in style_name:
        return "subtitle"
    elif "list" in style_name:
        return "list"
    elif "normal" in style_name:
        return "paragraph"
    else:
        return "other"
def extract_docx_with_metadata(file_path):
    doc = Document(file_path)

    results = []

    for para_index, para in enumerate(doc.paragraphs):
        for run_index, run in enumerate(para.runs):
            text = run.text.strip()

            if text:  # skip empty text
                metadata = {
                    # "text": text,
                    # "style": para.style.name,
                    "type": classify_style(para.style.name),
                    # "paragraph_index": para_index,
                    # "run_index": run_index,
                    # "font_name": run.font.name,
                    # "font_size": run.font.size.pt if run.font.size else None,
                    # "bold": run.bold,
                    # "italic": run.italic,
                    # "underline": run.underline,
                }
                results.append(metadata)

    return results


# 🔹 Example usage
file_path = "unformatdoc.docx"
data = extract_docx_with_metadata(file_path)

for item in data:
    print(item)

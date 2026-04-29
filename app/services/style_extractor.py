from docx import Document


class StyleExtractor:

    def extract_styles(self, file_bytes):
        doc = Document(file_bytes)

        styles_map = {}

        for para in doc.paragraphs:
            text = para.text.strip()

            if not text:
                continue

            style_name = para.style.name

            run = para.runs[0] if para.runs else None

            styles_map[style_name] = {
                "font": run.font.name if run else None,
                "size": run.font.size.pt if run and run.font.size else None,
                "bold": run.bold if run else False,
                "italic": run.italic if run else False,
                "align": str(para.alignment)
            }

        return styles_map
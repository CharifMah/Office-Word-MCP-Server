"""Page setup tools - margins, section configuration."""

import os
from docx import Document
from docx.shared import Cm

from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension


async def set_section_page_margins(filename: str, top: float = 2.5, bottom: float = 2.5,
                                    left: float = 2.5, right: float = 2.5) -> str:
    """Set page margins for the document (in cm)."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}"
    try:
        doc = Document(filename)
        for section in doc.sections:
            section.top_margin = Cm(top)
            section.bottom_margin = Cm(bottom)
            section.left_margin = Cm(left)
            section.right_margin = Cm(right)
        doc.save(filename)
        return f"Page margins set: top={top}cm, bottom={bottom}cm, left={left}cm, right={right}cm"
    except Exception as e:
        return f"Failed to set margins: {str(e)}"
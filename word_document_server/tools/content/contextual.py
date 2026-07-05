"""Contextual insertion tools - insert content relative to existing text or paragraphs."""

import os
from typing import List, Optional
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension
from word_document_server.utils.document_utils import insert_header_near_text, insert_numbered_list_near_text, insert_line_or_paragraph_near_text, replace_paragraph_block_below_header, replace_block_between_manual_anchors


async def insert_header_near_text_tool(filename: str, target_text: str = None, header_title: str = "",
                                        position: str = 'after', header_style: str = 'Heading 1',
                                        target_paragraph_index: int = None) -> str:
    """Insert a header near specific text or paragraph index."""
    return insert_header_near_text(filename, target_text, header_title, position, header_style, target_paragraph_index)


async def insert_numbered_list_near_text_tool(filename: str, target_text: str = None, list_items: list = None,
                                               position: str = 'after', target_paragraph_index: int = None,
                                               bullet_type: str = 'bullet') -> str:
    """Insert a numbered or bulleted list near specific text or paragraph index."""
    return insert_numbered_list_near_text(filename, target_text, list_items, position, target_paragraph_index, bullet_type)


async def insert_line_or_paragraph_near_text_tool(filename: str, target_text: str = None, line_text: str = "",
                                                   position: str = 'after', line_style: str = None,
                                                   target_paragraph_index: int = None) -> str:
    """Insert a line or paragraph near specific text or paragraph index."""
    return insert_line_or_paragraph_near_text(filename, target_text, line_text, position, line_style, target_paragraph_index)


async def replace_paragraph_block_below_header_tool(filename: str, header_text: str, new_paragraphs: list,
                                                      detect_block_end_fn=None) -> str:
    """Replace the block of paragraphs below a header, avoiding modifying TOC."""
    return replace_paragraph_block_below_header(filename, header_text, new_paragraphs, detect_block_end_fn)


async def replace_block_between_manual_anchors_tool(filename: str, start_anchor_text: str, new_paragraphs: list,
                                                      end_anchor_text: str = None, match_fn=None,
                                                      new_paragraph_style: str = None) -> str:
    """Replace all content between start_anchor_text and end_anchor_text."""
    return replace_block_between_manual_anchors(filename, start_anchor_text, new_paragraphs, end_anchor_text, match_fn, new_paragraph_style)
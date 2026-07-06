"""
MCP tool implementations for the Word Document Server.

This package contains the MCP tool implementations that expose functionality
to clients through the Model Context Protocol.
"""

# Document tools
from word_document_server.tools.document_tools import (
    create_document, get_document_info, get_document_text,
    get_document_outline, list_available_documents,
    copy_document, merge_documents
)

# Content tools
from word_document_server.tools.content import (
    add_heading, add_paragraph, add_table, add_picture,
    add_page_break, add_bullet_list, delete_paragraph,
    search_and_replace, insert_paragraph_at_index, move_section,
    delete_table, insert_table_at_position, get_paragraph_text,
    format_table_all_cells, delete_paragraphs_range, set_paragraph_spacing,
    set_section_page_margins, set_table_borders, set_cell_margins_all,
    replace_table_data, set_header_footer, add_page_number_to_footer,
    remove_all_indentation_from_table,
    insert_header_near_text_tool, insert_numbered_list_near_text_tool,
    insert_line_or_paragraph_near_text_tool,
    replace_paragraph_block_below_header_tool,
    replace_block_between_manual_anchors_tool,
)

# Format tools
from word_document_server.tools.format_tools import (
    format_text, create_custom_style, format_table
)

# Protection tools
from word_document_server.tools.protection_tools import (
    protect_document, add_restricted_editing,
    add_digital_signature, verify_document
)

# Footnote tools
from word_document_server.tools.footnote_tools import (
    add_footnote_to_document, add_endnote_to_document,
    convert_footnotes_to_endnotes_in_document, customize_footnote_style
)

# Comment tools
from word_document_server.tools.comment_tools import (
    get_all_comments, get_comments_by_author, get_comments_for_paragraph
)

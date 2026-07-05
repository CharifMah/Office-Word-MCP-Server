"""Content tools package - organized by category."""

from word_document_server.tools.content.creation import (
    add_heading, add_paragraph, add_table, add_picture, add_page_break, add_bullet_list
)
from word_document_server.tools.content.modification import (
    delete_paragraph, search_and_replace, insert_paragraph_at_index, move_section,
    delete_paragraphs_range, set_paragraph_spacing, set_header_footer, add_page_number_to_footer
)
from word_document_server.tools.content.contextual import (
    insert_header_near_text_tool, insert_numbered_list_near_text_tool,
    insert_line_or_paragraph_near_text_tool, replace_paragraph_block_below_header_tool,
    replace_block_between_manual_anchors_tool
)
from word_document_server.tools.content.table_tools import (
    delete_table, insert_table_at_position, format_table_all_cells, replace_table_data,
    set_table_borders, set_cell_margins_all, remove_all_indentation_from_table
)
from word_document_server.tools.content.page_setup import set_section_page_margins
from word_document_server.tools.content.read import get_paragraph_text
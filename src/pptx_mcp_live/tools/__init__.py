"""Tools module for PowerPoint MCP Live."""

from .inspection import (
    get_slide_info_sync,
    inspect_presentation_sync,
    list_open_presentations_sync,
    list_slide_layouts_sync,
)
from .readers import (
    get_comments_sync,
    read_shape_text_sync,
    read_slide_notes_sync,
    read_slide_text_sync,
)
from .writers import (
    add_slide_sync,
    add_text_box_sync,
    delete_slide_sync,
    duplicate_slide_sync,
    reorder_slide_sync,
    set_shape_text_sync,
    set_slide_notes_sync,
)
from .formatters import (
    format_shape_sync,
    format_text_sync,
    set_slide_background_sync,
)
from .layout import (
    add_shape_sync,
    align_shapes_sync,
    arrange_shape_sync,
    delete_shape_sync,
    distribute_shapes_sync,
    group_shapes_sync,
    move_shape_sync,
    resize_shape_sync,
    rotate_shape_sync,
    set_slide_layout_sync,
)
from .media import (
    add_chart_sync,
    add_image_sync,
    add_table_sync,
)
from .export import (
    capture_slide_sync,
    export_pdf_sync,
    export_slide_image_sync,
)
from .comments import (
    add_comment_sync,
    delete_comment_sync,
    get_all_comments_sync,
    reply_to_comment_sync,
)
from .advanced import (
    close_presentation_sync,
    find_replace_sync,
    save_presentation_sync,
)
from .advanced import (
    find_replace_sync,
    save_presentation_sync,
)

__all__ = [
    # Inspection
    "list_open_presentations_sync",
    "inspect_presentation_sync",
    "get_slide_info_sync",
    # Readers
    "read_slide_text_sync",
    "read_slide_notes_sync",
    "read_shape_text_sync",
    "get_comments_sync",
    # Writers
    "add_slide_sync",
    "delete_slide_sync",
    "duplicate_slide_sync",
    "reorder_slide_sync",
    "set_shape_text_sync",
    "set_slide_notes_sync",
    "add_text_box_sync",
    # Formatters
    "format_text_sync",
    "format_shape_sync",
    "set_slide_background_sync",
    # Layout
    "move_shape_sync",
    "resize_shape_sync",
    "rotate_shape_sync",
    "arrange_shape_sync",
    "align_shapes_sync",
    "distribute_shapes_sync",
    "group_shapes_sync",
    # Media
    "add_image_sync",
    "add_table_sync",
    "add_chart_sync",
    # Export
    "export_slide_image_sync",
    "capture_slide_sync",
    "export_pdf_sync",
    # Comments
    "add_comment_sync",
    "delete_comment_sync",
    # Advanced
    "find_replace_sync",
    "save_presentation_sync",
]

"""PowerPoint MCP Live Server — COM-based live editing for PowerPoint.

43 tools for real-time PowerPoint automation on Windows.
"""

import asyncio
import logging
from typing import Any, Dict, List, Optional

from mcp.server.fastmcp import FastMCP

from .core.errors import ToolError
from .tools import (
    # inspection
    get_slide_info_sync,
    inspect_presentation_sync,
    list_open_presentations_sync,
    list_slide_layouts_sync,
    # readers
    get_comments_sync,
    read_shape_text_sync,
    read_slide_notes_sync,
    read_slide_text_sync,
    # writers
    add_slide_sync,
    add_text_box_sync,
    delete_slide_sync,
    duplicate_slide_sync,
    reorder_slide_sync,
    set_shape_text_sync,
    set_slide_notes_sync,
    # formatters
    format_shape_sync,
    format_text_sync,
    set_slide_background_sync,
    # layout
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
    # media
    add_chart_sync,
    add_image_sync,
    add_table_sync,
    # export
    capture_slide_sync,
    export_pdf_sync,
    export_slide_image_sync,
    # comments
    add_comment_sync,
    delete_comment_sync,
    get_all_comments_sync,
    reply_to_comment_sync,
    # advanced
    close_presentation_sync,
    find_replace_sync,
    save_presentation_sync,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler()],
)
logger = logging.getLogger(__name__)

mcp = FastMCP(
    name="PowerPoint MCP Live",
    instructions="""
    PowerPoint MCP Live Server
    ==========================
    COM-based live editing for Microsoft PowerPoint on Windows.
    
    43 tools for:
    - Inspecting presentations, slides, and available layouts
    - Reading/writing text, notes, shapes
    - Adding/deleting/reordering slides and shapes
    - Formatting text (whole shape or per-paragraph) and shapes
    - Position, size, alignment, rotation of shapes (including placeholders)
    - Adding images, tables, charts, and auto shapes (arrows, rectangles, etc.)
    - Exporting slides as images/PDF
    - Visual feedback via capture_slide (base64 PNG)
    - Comments: add, delete, reply, get all across slides
    - Find/replace across slides
    - Close presentations programmatically
    
    Prerequisites:
    - Windows OS (required for COM automation)
    - Microsoft PowerPoint must be running
    - At least one presentation must be open
    """,
)


# ============================================================================
# Inspection Tools
# ============================================================================

@mcp.tool()
async def list_open_presentations() -> dict:
    """List all open PowerPoint presentations with slide counts."""
    try:
        return await asyncio.to_thread(list_open_presentations_sync)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def inspect_presentation(presentation_name: Optional[str] = None) -> dict:
    """Get metadata: slide count, dimensions, layouts, slide summaries.
    
    Args:
        presentation_name: Filename (e.g. "slides.pptx"). None = active.
    """
    try:
        return await asyncio.to_thread(inspect_presentation_sync, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def get_slide_info(slide_index: int, presentation_name: Optional[str] = None) -> dict:
    """Get all shapes on a slide with type, text, position, size, and font details.
    
    Args:
        slide_index: 1-based slide number.
    """
    try:
        return await asyncio.to_thread(get_slide_info_sync, slide_index, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def list_slide_layouts(presentation_name: Optional[str] = None) -> dict:
    """List all available slide layouts with names, indices, and placeholder counts.
    Use this to find the right layout_index for add_slide or set_slide_layout.
    """
    try:
        return await asyncio.to_thread(list_slide_layouts_sync, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e


# ============================================================================
# Reading Tools
# ============================================================================

@mcp.tool()
async def read_slide_text(slide_index: int, presentation_name: Optional[str] = None) -> dict:
    """Extract all text from all shapes on a slide."""
    try:
        return await asyncio.to_thread(read_slide_text_sync, slide_index, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def read_slide_notes(slide_index: int, presentation_name: Optional[str] = None) -> dict:
    """Read speaker notes from a slide."""
    try:
        return await asyncio.to_thread(read_slide_notes_sync, slide_index, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def read_shape_text(slide_index: int, shape_ref: str, presentation_name: Optional[str] = None) -> dict:
    """Read text from a specific shape by name or index.
    
    Args:
        shape_ref: Shape name (e.g. "Title 1") or index as string (e.g. "1").
    """
    try:
        ref = int(shape_ref) if shape_ref.isdigit() else shape_ref
        return await asyncio.to_thread(read_shape_text_sync, slide_index, ref, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def get_comments(slide_index: int, presentation_name: Optional[str] = None) -> dict:
    """Get all comments from a slide with author, text, position."""
    try:
        return await asyncio.to_thread(get_comments_sync, slide_index, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def get_all_comments(presentation_name: Optional[str] = None) -> dict:
    """Get all comments from ALL slides in one call. Returns comments grouped by slide."""
    try:
        return await asyncio.to_thread(get_all_comments_sync, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e


# ============================================================================
# Writing Tools
# ============================================================================

@mcp.tool()
async def add_slide(layout_index: int = 1, insert_at: Optional[int] = None, presentation_name: Optional[str] = None) -> dict:
    """Add a new slide. Use inspect_presentation to see available layouts.
    
    Args:
        layout_index: 1-based layout index from the slide master.
        insert_at: Position (1-based). None = end.
    """
    try:
        return await asyncio.to_thread(add_slide_sync, layout_index, insert_at, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def delete_slide(slide_index: int, presentation_name: Optional[str] = None) -> dict:
    """Delete a slide by index."""
    try:
        return await asyncio.to_thread(delete_slide_sync, slide_index, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def duplicate_slide(slide_index: int, presentation_name: Optional[str] = None) -> dict:
    """Duplicate a slide."""
    try:
        return await asyncio.to_thread(duplicate_slide_sync, slide_index, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def reorder_slide(slide_index: int, new_position: int, presentation_name: Optional[str] = None) -> dict:
    """Move a slide to a new position."""
    try:
        return await asyncio.to_thread(reorder_slide_sync, slide_index, new_position, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def set_shape_text(slide_index: int, shape_ref: str, text: str, presentation_name: Optional[str] = None) -> dict:
    """Write text to a shape.
    
    Args:
        shape_ref: Shape name or index as string.
        text: New text content.
    """
    try:
        ref = int(shape_ref) if shape_ref.isdigit() else shape_ref
        return await asyncio.to_thread(set_shape_text_sync, slide_index, ref, text, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def set_slide_notes(slide_index: int, text: str, presentation_name: Optional[str] = None) -> dict:
    """Set speaker notes for a slide."""
    try:
        return await asyncio.to_thread(set_slide_notes_sync, slide_index, text, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def add_text_box(slide_index: int, text: str, left_inches: float = 1.0, top_inches: float = 1.0, width_inches: float = 4.0, height_inches: float = 1.0, presentation_name: Optional[str] = None) -> dict:
    """Add a text box to a slide with position and size in inches."""
    try:
        return await asyncio.to_thread(add_text_box_sync, slide_index, text, left_inches, top_inches, width_inches, height_inches, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e


# ============================================================================
# Formatting Tools
# ============================================================================

@mcp.tool()
async def format_text(slide_index: int, shape_ref: str, bold: Optional[bool] = None, italic: Optional[bool] = None, underline: Optional[bool] = None, font_name: Optional[str] = None, font_size: Optional[float] = None, font_color: Optional[str] = None, alignment: Optional[str] = None, paragraph_index: Optional[int] = None, presentation_name: Optional[str] = None) -> dict:
    """Format text in a shape: bold, italic, font, size, color, alignment.
    
    Args:
        font_color: "#RRGGBB" hex color.
        alignment: "left", "center", "right", "justify".
        paragraph_index: 1-based paragraph number. None = format ALL text in the shape.
    """
    try:
        ref = int(shape_ref) if shape_ref.isdigit() else shape_ref
        return await asyncio.to_thread(format_text_sync, slide_index, ref, bold, italic, underline, font_name, font_size, font_color, alignment, paragraph_index, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def format_shape(slide_index: int, shape_ref: str, fill_color: Optional[str] = None, line_color: Optional[str] = None, line_width: Optional[float] = None, transparency: Optional[float] = None, no_fill: Optional[bool] = None, no_line: Optional[bool] = None, presentation_name: Optional[str] = None) -> dict:
    """Format a shape: fill color, border, transparency.
    
    Args:
        fill_color/line_color: "#RRGGBB" hex.
        transparency: 0.0 (opaque) to 1.0 (fully transparent).
    """
    try:
        ref = int(shape_ref) if shape_ref.isdigit() else shape_ref
        return await asyncio.to_thread(format_shape_sync, slide_index, ref, fill_color, line_color, line_width, transparency, no_fill, no_line, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def set_slide_background(slide_index: int, color: str, presentation_name: Optional[str] = None) -> dict:
    """Set slide background color. Args: color = "#RRGGBB"."""
    try:
        return await asyncio.to_thread(set_slide_background_sync, slide_index, color, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e


# ============================================================================
# Position & Layout Tools
# ============================================================================

@mcp.tool()
async def move_shape(slide_index: int, shape_ref: str, left_inches: Optional[float] = None, top_inches: Optional[float] = None, presentation_name: Optional[str] = None) -> dict:
    """Move a shape to a position (inches from top-left)."""
    try:
        ref = int(shape_ref) if shape_ref.isdigit() else shape_ref
        return await asyncio.to_thread(move_shape_sync, slide_index, ref, left_inches, top_inches, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def resize_shape(slide_index: int, shape_ref: str, width_inches: Optional[float] = None, height_inches: Optional[float] = None, presentation_name: Optional[str] = None) -> dict:
    """Resize a shape (inches)."""
    try:
        ref = int(shape_ref) if shape_ref.isdigit() else shape_ref
        return await asyncio.to_thread(resize_shape_sync, slide_index, ref, width_inches, height_inches, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def rotate_shape(slide_index: int, shape_ref: str, degrees: float, presentation_name: Optional[str] = None) -> dict:
    """Rotate a shape by degrees (0-360)."""
    try:
        ref = int(shape_ref) if shape_ref.isdigit() else shape_ref
        return await asyncio.to_thread(rotate_shape_sync, slide_index, ref, degrees, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def arrange_shape(slide_index: int, shape_ref: str, action: str, presentation_name: Optional[str] = None) -> dict:
    """Change z-order: bring_to_front, send_to_back, bring_forward, send_backward."""
    try:
        ref = int(shape_ref) if shape_ref.isdigit() else shape_ref
        return await asyncio.to_thread(arrange_shape_sync, slide_index, ref, action, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def align_shapes(slide_index: int, shape_names: List[str], alignment: str, presentation_name: Optional[str] = None) -> dict:
    """Align multiple shapes: left, center, right, top, middle, bottom."""
    try:
        return await asyncio.to_thread(align_shapes_sync, slide_index, shape_names, alignment, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def distribute_shapes(slide_index: int, shape_names: List[str], direction: str, presentation_name: Optional[str] = None) -> dict:
    """Distribute shapes evenly: horizontal or vertical."""
    try:
        return await asyncio.to_thread(distribute_shapes_sync, slide_index, shape_names, direction, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def group_shapes(slide_index: int, shape_names: List[str], ungroup: bool = False, presentation_name: Optional[str] = None) -> dict:
    """Group or ungroup shapes. To ungroup, pass a single group shape name with ungroup=True."""
    try:
        return await asyncio.to_thread(group_shapes_sync, slide_index, shape_names, ungroup, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def delete_shape(slide_index: int, shape_ref: str, presentation_name: Optional[str] = None) -> dict:
    """Delete a shape from a slide by name or index.
    
    Args:
        shape_ref: Shape name (e.g. "TextBox 3") or index as string (e.g. "3").
    """
    try:
        ref = int(shape_ref) if shape_ref.isdigit() else shape_ref
        return await asyncio.to_thread(delete_shape_sync, slide_index, ref, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def add_shape(slide_index: int, shape_type: str, left_inches: float = 1.0, top_inches: float = 1.0, width_inches: float = 2.0, height_inches: float = 1.0, text: Optional[str] = None, fill_color: Optional[str] = None, line_color: Optional[str] = None, presentation_name: Optional[str] = None) -> dict:
    """Add an auto shape (rectangle, arrow, oval, etc.) to a slide.
    
    Args:
        shape_type: rectangle, rounded_rectangle, oval, triangle, diamond,
            right_arrow, left_arrow, up_arrow, down_arrow, star_5, heart,
            cross, chevron, hexagon, pentagon, and more.
        text: Optional text inside the shape.
        fill_color: Fill color as "#RRGGBB".
        line_color: Border color as "#RRGGBB".
    """
    try:
        return await asyncio.to_thread(add_shape_sync, slide_index, shape_type, left_inches, top_inches, width_inches, height_inches, text, fill_color, line_color, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def set_slide_layout(slide_index: int, layout_index: int, presentation_name: Optional[str] = None) -> dict:
    """Change the layout of an existing slide. Use list_slide_layouts to see available layouts.
    
    Args:
        slide_index: 1-based slide number.
        layout_index: 1-based layout index.
    """
    try:
        return await asyncio.to_thread(set_slide_layout_sync, slide_index, layout_index, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e


# ============================================================================
# Media Tools
# ============================================================================

@mcp.tool()
async def add_image(slide_index: int, image_path: str, left_inches: float = 1.0, top_inches: float = 1.0, width_inches: Optional[float] = None, height_inches: Optional[float] = None, presentation_name: Optional[str] = None) -> dict:
    """Add an image to a slide. Supports PNG, JPG, BMP. Position/size in inches."""
    try:
        return await asyncio.to_thread(add_image_sync, slide_index, image_path, left_inches, top_inches, width_inches, height_inches, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def add_table(slide_index: int, rows: int, cols: int, data: Optional[List[List[str]]] = None, left_inches: float = 1.0, top_inches: float = 2.0, width_inches: float = 8.0, height_inches: float = 3.0, presentation_name: Optional[str] = None) -> dict:
    """Add a table with optional data. First row = header."""
    try:
        return await asyncio.to_thread(add_table_sync, slide_index, rows, cols, data, left_inches, top_inches, width_inches, height_inches, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def add_chart(slide_index: int, chart_type: str = "column", data: Optional[List[List]] = None, left_inches: float = 1.0, top_inches: float = 1.5, width_inches: float = 8.0, height_inches: float = 5.0, title: Optional[str] = None, presentation_name: Optional[str] = None) -> dict:
    """Add a chart. Types: column, bar, line, pie, area, scatter. Data: first row = categories, subsequent rows = [series_name, val1, val2, ...]."""
    try:
        return await asyncio.to_thread(add_chart_sync, slide_index, chart_type, data, left_inches, top_inches, width_inches, height_inches, title, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e


# ============================================================================
# Export & Visual Feedback Tools
# ============================================================================

@mcp.tool()
async def export_slide_image(slide_index: int, output_path: Optional[str] = None, width: int = 1920, height: int = 1080, presentation_name: Optional[str] = None) -> dict:
    """Export a slide as PNG image file."""
    try:
        return await asyncio.to_thread(export_slide_image_sync, slide_index, output_path, width, height, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def capture_slide(slide_index: int, width: int = 1280, height: int = 720, presentation_name: Optional[str] = None) -> dict:
    """Capture a slide as base64 PNG for visual feedback. Use to verify edits."""
    try:
        return await asyncio.to_thread(capture_slide_sync, slide_index, width, height, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def export_pdf(output_path: str, slide_index: Optional[int] = None, presentation_name: Optional[str] = None) -> dict:
    """Export presentation to PDF. slide_index = specific slide only."""
    try:
        return await asyncio.to_thread(export_pdf_sync, output_path, slide_index, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e


# ============================================================================
# Comment Tools
# ============================================================================

@mcp.tool()
async def add_comment(slide_index: int, text: str, author: str = "", author_initials: str = "", presentation_name: Optional[str] = None) -> dict:
    """Add a comment to a slide."""
    try:
        return await asyncio.to_thread(add_comment_sync, slide_index, text, author, author_initials, 0, 0, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def delete_comment(slide_index: int, comment_index: int, presentation_name: Optional[str] = None) -> dict:
    """Delete a comment by index."""
    try:
        return await asyncio.to_thread(delete_comment_sync, slide_index, comment_index, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def reply_to_comment(slide_index: int, comment_index: int, text: str, author: str = "", author_initials: str = "", presentation_name: Optional[str] = None) -> dict:
    """Reply to an existing comment on a slide.
    
    Args:
        slide_index: 1-based slide number.
        comment_index: 1-based index of the comment to reply to.
        text: Reply text.
    """
    try:
        return await asyncio.to_thread(reply_to_comment_sync, slide_index, comment_index, text, author, author_initials, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e


# ============================================================================
# Advanced Tools
# ============================================================================

@mcp.tool()
async def find_replace(find_text: str, replace_text: str, slide_index: Optional[int] = None, preview_only: bool = False, presentation_name: Optional[str] = None) -> dict:
    """Find and replace text across slides. preview_only=True to count without replacing."""
    try:
        return await asyncio.to_thread(find_replace_sync, find_text, replace_text, slide_index, preview_only, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def save_presentation(presentation_name: Optional[str] = None) -> dict:
    """Save the active presentation."""
    try:
        return await asyncio.to_thread(save_presentation_sync, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e

@mcp.tool()
async def close_presentation(save: bool = True, presentation_name: Optional[str] = None) -> dict:
    """Close a presentation. Optionally save before closing.
    
    Args:
        save: Whether to save before closing. Default True.
        presentation_name: Name of presentation to close. None = active.
    """
    try:
        return await asyncio.to_thread(close_presentation_sync, save, presentation_name)
    except Exception as e:
        if isinstance(e, ToolError): raise
        raise ToolError(f"Failed: {e}") from e


# ============================================================================
# Server Entry Point
# ============================================================================

def create_server() -> FastMCP:
    """Create and return the MCP server instance."""
    return mcp


def run_server():
    """Run the MCP server."""
    mcp.run(transport="stdio")

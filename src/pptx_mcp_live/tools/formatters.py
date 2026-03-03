"""Formatting tools: text format, shape format, slide background."""

from typing import Any, Dict, Optional

from ..core.connection import get_powerpoint, get_presentation, get_slide, get_shape
from ..core.errors import ToolError


def format_text_sync(
    slide_index: int,
    shape_ref,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    underline: Optional[bool] = None,
    font_name: Optional[str] = None,
    font_size: Optional[float] = None,
    font_color: Optional[str] = None,
    alignment: Optional[str] = None,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Format text in a shape.

    Args:
        font_color: Hex color like "#FF0000" for red.
        alignment: "left", "center", "right", "justify".
    """
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)
    shape = get_shape(slide, shape_ref)

    if not shape.HasTextFrame:
        raise ToolError(f"Shape '{shape_ref}' does not contain text.")

    tr = shape.TextFrame.TextRange
    changes = []

    if bold is not None:
        tr.Font.Bold = bold
        changes.append(f"bold={bold}")
    if italic is not None:
        tr.Font.Italic = italic
        changes.append(f"italic={italic}")
    if underline is not None:
        tr.Font.Underline = underline
        changes.append(f"underline={underline}")
    if font_name is not None:
        tr.Font.Name = font_name
        changes.append(f"font={font_name}")
    if font_size is not None:
        tr.Font.Size = font_size
        changes.append(f"size={font_size}")
    if font_color is not None:
        rgb = _hex_to_rgb_int(font_color)
        tr.Font.Color.RGB = rgb
        changes.append(f"color={font_color}")
    if alignment is not None:
        align_map = {
            "left": 1,      # ppAlignLeft
            "center": 2,    # ppAlignCenter
            "right": 3,     # ppAlignRight
            "justify": 4,   # ppAlignJustify
        }
        val = align_map.get(alignment.lower())
        if val is None:
            raise ToolError(f"Invalid alignment '{alignment}'. Use: left, center, right, justify.")
        for i in range(1, tr.Paragraphs().Count + 1):
            tr.Paragraphs(i).ParagraphFormat.Alignment = val
        changes.append(f"alignment={alignment}")

    return {
        "success": True,
        "presentation": pres.Name,
        "slide_index": slide_index,
        "shape_name": shape.Name,
        "changes": changes,
        "message": f"Formatted text on '{shape.Name}': {', '.join(changes)}",
    }


def format_shape_sync(
    slide_index: int,
    shape_ref,
    fill_color: Optional[str] = None,
    line_color: Optional[str] = None,
    line_width: Optional[float] = None,
    transparency: Optional[float] = None,
    no_fill: Optional[bool] = None,
    no_line: Optional[bool] = None,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Format a shape's fill and border.

    Args:
        fill_color: Hex color like "#4472C4".
        line_color: Border color as hex.
        line_width: Border width in points.
        transparency: Fill transparency 0.0-1.0.
    """
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)
    shape = get_shape(slide, shape_ref)

    changes = []

    if no_fill:
        shape.Fill.Background()
        changes.append("no_fill")
    elif fill_color is not None:
        shape.Fill.Solid()
        shape.Fill.ForeColor.RGB = _hex_to_rgb_int(fill_color)
        changes.append(f"fill={fill_color}")

    if transparency is not None:
        shape.Fill.Transparency = transparency
        changes.append(f"transparency={transparency}")

    if no_line:
        shape.Line.Visible = False
        changes.append("no_line")
    elif line_color is not None:
        shape.Line.Visible = True
        shape.Line.ForeColor.RGB = _hex_to_rgb_int(line_color)
        changes.append(f"line_color={line_color}")

    if line_width is not None:
        shape.Line.Visible = True
        shape.Line.Weight = line_width
        changes.append(f"line_width={line_width}pt")

    return {
        "success": True,
        "presentation": pres.Name,
        "slide_index": slide_index,
        "shape_name": shape.Name,
        "changes": changes,
        "message": f"Formatted shape '{shape.Name}': {', '.join(changes)}",
    }


def set_slide_background_sync(
    slide_index: int,
    color: Optional[str] = None,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Set slide background color.

    Args:
        color: Hex color like "#FFFFFF" for white.
    """
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)

    if color is not None:
        slide.FollowMasterBackground = False
        slide.Background.Fill.Solid()
        slide.Background.Fill.ForeColor.RGB = _hex_to_rgb_int(color)

    return {
        "success": True,
        "presentation": pres.Name,
        "slide_index": slide_index,
        "color": color,
        "message": f"Background set to {color} on slide {slide_index}.",
    }


def _hex_to_rgb_int(hex_color: str) -> int:
    """Convert #RRGGBB to COM RGB integer (BGR format)."""
    hex_color = hex_color.lstrip("#")
    if len(hex_color) != 6:
        raise ToolError(f"Invalid color '{hex_color}'. Use #RRGGBB format.")
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return r + (g << 8) + (b << 16)

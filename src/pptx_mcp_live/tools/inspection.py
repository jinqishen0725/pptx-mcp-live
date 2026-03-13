"""Inspection tools: list presentations, inspect, get slide info."""

from typing import Any, Dict, List, Optional

from ..core.connection import get_powerpoint, get_presentation, get_slide, points_to_inches
from ..core.errors import ToolError


def list_open_presentations_sync() -> Dict[str, Any]:
    """List all open presentations with slide counts."""
    app = get_powerpoint()
    presentations = []
    for i in range(1, app.Presentations.Count + 1):
        pres = app.Presentations(i)
        presentations.append({
            "name": pres.Name,
            "path": pres.FullName,
            "slide_count": pres.Slides.Count,
            "read_only": bool(pres.ReadOnly),
        })
    return {
        "success": True,
        "presentations": presentations,
        "count": len(presentations),
    }


def inspect_presentation_sync(
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Get detailed metadata about a presentation."""
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)

    # Slide dimensions
    width_in = points_to_inches(pres.PageSetup.SlideWidth)
    height_in = points_to_inches(pres.PageSetup.SlideHeight)

    # Slide summaries
    slides = []
    for i in range(1, pres.Slides.Count + 1):
        slide = pres.Slides(i)
        slides.append({
            "index": i,
            "layout": slide.Layout,
            "shape_count": slide.Shapes.Count,
            "has_notes": _has_notes(slide),
            "hidden": bool(slide.SlideShowTransition.Hidden) if hasattr(slide.SlideShowTransition, 'Hidden') else False,
        })

    # Slide masters
    masters = []
    for i in range(1, pres.SlideMaster.CustomLayouts.Count + 1):
        layout = pres.SlideMaster.CustomLayouts(i)
        masters.append({"index": i, "name": layout.Name})

    return {
        "success": True,
        "name": pres.Name,
        "path": pres.FullName,
        "slide_count": pres.Slides.Count,
        "dimensions": {
            "width_inches": round(width_in, 2),
            "height_inches": round(height_in, 2),
        },
        "slides": slides,
        "layouts": masters,
    }


def get_slide_info_sync(
    slide_index: int,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Get detailed info about a specific slide including all shapes with font details."""
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)

    shapes = []
    for i in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes(i)
        shape_info = {
            "index": i,
            "name": shape.Name,
            "type": _shape_type_name(shape.Type),
            "left_inches": round(points_to_inches(shape.Left), 2),
            "top_inches": round(points_to_inches(shape.Top), 2),
            "width_inches": round(points_to_inches(shape.Width), 2),
            "height_inches": round(points_to_inches(shape.Height), 2),
        }

        # Text content and font details
        if shape.HasTextFrame:
            try:
                shape_info["text"] = shape.TextFrame.TextRange.Text
                # Get font details from the first run
                tr = shape.TextFrame.TextRange
                font_info = {}
                try:
                    font_info["name"] = tr.Font.Name
                except Exception:
                    pass
                try:
                    font_info["size"] = tr.Font.Size
                except Exception:
                    pass
                try:
                    font_info["bold"] = bool(tr.Font.Bold)
                except Exception:
                    pass
                try:
                    font_info["italic"] = bool(tr.Font.Italic)
                except Exception:
                    pass
                try:
                    rgb_int = tr.Font.Color.RGB
                    r = rgb_int & 0xFF
                    g = (rgb_int >> 8) & 0xFF
                    b = (rgb_int >> 16) & 0xFF
                    font_info["color"] = f"#{r:02X}{g:02X}{b:02X}"
                except Exception:
                    pass
                if font_info:
                    shape_info["font"] = font_info
            except Exception:
                shape_info["text"] = None

        # Table info
        if shape.HasTable:
            shape_info["table"] = {
                "rows": shape.Table.Rows.Count,
                "columns": shape.Table.Columns.Count,
            }

        shapes.append(shape_info)

    # Notes
    notes_text = _get_notes_text(slide)

    return {
        "success": True,
        "presentation": pres.Name,
        "slide_index": slide_index,
        "layout": slide.Layout,
        "shape_count": slide.Shapes.Count,
        "shapes": shapes,
        "notes": notes_text,
    }


def list_slide_layouts_sync(
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """List all available slide layouts with names and indices.

    Returns layout names and indices that can be used with add_slide or set_slide_layout.
    """
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)

    layouts = []
    for i in range(1, pres.SlideMaster.CustomLayouts.Count + 1):
        layout = pres.SlideMaster.CustomLayouts(i)
        # Count placeholders
        placeholder_count = 0
        try:
            placeholder_count = layout.Placeholders.Count
        except Exception:
            pass
        layouts.append({
            "index": i,
            "name": layout.Name,
            "placeholder_count": placeholder_count,
        })

    return {
        "success": True,
        "presentation": pres.Name,
        "layouts": layouts,
        "count": len(layouts),
    }


def _has_notes(slide) -> bool:
    """Check if a slide has speaker notes."""
    try:
        notes = slide.NotesPage.Shapes(2).TextFrame.TextRange.Text
        return bool(notes and notes.strip())
    except Exception:
        return False


def _get_notes_text(slide) -> Optional[str]:
    """Get speaker notes text from a slide."""
    try:
        text = slide.NotesPage.Shapes(2).TextFrame.TextRange.Text
        return text if text and text.strip() else None
    except Exception:
        return None


def _shape_type_name(type_id: int) -> str:
    """Convert shape type ID to readable name."""
    types = {
        1: "AutoShape",
        2: "Callout",
        3: "Chart",
        4: "Comment",
        5: "FreeForm",
        6: "Group",
        7: "EmbeddedOLEObject",
        8: "FormControl",
        9: "Line",
        10: "LinkedOLEObject",
        11: "LinkedPicture",
        12: "OLEControlObject",
        13: "Picture",
        14: "Placeholder",
        15: "TextEffect",
        16: "MediaObject",
        17: "TextBox",
        18: "ScriptAnchor",
        19: "Table",
        20: "Canvas",
        21: "Diagram",
        22: "Ink",
        23: "InkComment",
        24: "SmartArt",
        25: "ContentApp",
        26: "WebVideo",
        27: "3DModel",
        28: "LinkedGraphic",
        29: "Graphic",
    }
    return types.get(type_id, f"Unknown({type_id})")

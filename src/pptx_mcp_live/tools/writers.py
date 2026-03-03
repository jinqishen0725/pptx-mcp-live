"""Writing tools: add/delete/duplicate slides, set text, notes, textbox."""

from typing import Any, Dict, Optional

from ..core.connection import (
    get_powerpoint, get_presentation, get_slide, get_shape, inches_to_points
)
from ..core.errors import ToolError


def add_slide_sync(
    layout_index: int = 1,
    insert_at: Optional[int] = None,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Add a new slide with specified layout.

    Args:
        layout_index: 1-based layout index from the slide master.
        insert_at: Position to insert (1-based). None = end.
        presentation_name: Name of presentation. None = active.
    """
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)

    if insert_at is None:
        insert_at = pres.Slides.Count + 1

    try:
        layout = pres.SlideMaster.CustomLayouts(layout_index)
    except Exception:
        raise ToolError(
            f"Layout index {layout_index} not found. Use inspect_presentation to see available layouts."
        )

    slide = pres.Slides.AddSlide(insert_at, layout)

    return {
        "success": True,
        "presentation": pres.Name,
        "slide_index": slide.SlideIndex,
        "layout": layout.Name,
        "total_slides": pres.Slides.Count,
        "message": f"Slide added at position {slide.SlideIndex} with layout '{layout.Name}'",
    }


def delete_slide_sync(
    slide_index: int,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Delete a slide by index."""
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)

    slide.Delete()

    return {
        "success": True,
        "presentation": pres.Name,
        "deleted_index": slide_index,
        "remaining_slides": pres.Slides.Count,
        "message": f"Slide {slide_index} deleted. {pres.Slides.Count} slides remaining.",
    }


def duplicate_slide_sync(
    slide_index: int,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Duplicate a slide."""
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)

    new_slide = slide.Duplicate()
    new_index = new_slide(1).SlideIndex  # Duplicate returns a SlideRange

    return {
        "success": True,
        "presentation": pres.Name,
        "source_index": slide_index,
        "new_index": new_index,
        "total_slides": pres.Slides.Count,
        "message": f"Slide {slide_index} duplicated to position {new_index}.",
    }


def reorder_slide_sync(
    slide_index: int,
    new_position: int,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Move a slide to a new position."""
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)

    if new_position < 1 or new_position > pres.Slides.Count:
        raise ToolError(f"New position {new_position} out of range (1-{pres.Slides.Count}).")

    slide.MoveTo(new_position)

    return {
        "success": True,
        "presentation": pres.Name,
        "from_index": slide_index,
        "to_index": new_position,
        "message": f"Slide moved from position {slide_index} to {new_position}.",
    }


def set_shape_text_sync(
    slide_index: int,
    shape_ref,
    text: str,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Write text to a shape on a slide."""
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)
    shape = get_shape(slide, shape_ref)

    if not shape.HasTextFrame:
        raise ToolError(f"Shape '{shape_ref}' does not support text.")

    shape.TextFrame.TextRange.Text = text

    return {
        "success": True,
        "presentation": pres.Name,
        "slide_index": slide_index,
        "shape_name": shape.Name,
        "text": text,
        "message": f"Text set on shape '{shape.Name}' on slide {slide_index}.",
    }


def set_slide_notes_sync(
    slide_index: int,
    text: str,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Set speaker notes for a slide."""
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)

    try:
        slide.NotesPage.Shapes(2).TextFrame.TextRange.Text = text
    except Exception as e:
        raise ToolError(f"Failed to set notes: {e}") from e

    return {
        "success": True,
        "presentation": pres.Name,
        "slide_index": slide_index,
        "notes": text,
        "message": f"Speaker notes set on slide {slide_index}.",
    }


def add_text_box_sync(
    slide_index: int,
    text: str,
    left_inches: float = 1.0,
    top_inches: float = 1.0,
    width_inches: float = 4.0,
    height_inches: float = 1.0,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Add a text box to a slide."""
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)

    # msoTextOrientationHorizontal = 1
    shape = slide.Shapes.AddTextbox(
        1,  # Orientation: horizontal
        inches_to_points(left_inches),
        inches_to_points(top_inches),
        inches_to_points(width_inches),
        inches_to_points(height_inches),
    )
    shape.TextFrame.TextRange.Text = text

    return {
        "success": True,
        "presentation": pres.Name,
        "slide_index": slide_index,
        "shape_name": shape.Name,
        "text": text,
        "position": {
            "left_inches": left_inches,
            "top_inches": top_inches,
            "width_inches": width_inches,
            "height_inches": height_inches,
        },
        "message": f"Text box added to slide {slide_index}.",
    }

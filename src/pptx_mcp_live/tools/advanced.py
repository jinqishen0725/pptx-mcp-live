"""Advanced tools: find/replace, save."""

from typing import Any, Dict, List, Optional

from ..core.connection import get_powerpoint, get_presentation, get_slide
from ..core.errors import ToolError


def find_replace_sync(
    find_text: str,
    replace_text: str,
    slide_index: Optional[int] = None,
    preview_only: bool = False,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Find and replace text across slides.

    Args:
        find_text: Text to find.
        replace_text: Text to replace with.
        slide_index: Specific slide only. None = all slides.
        preview_only: If True, count matches without replacing.
    """
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)

    if slide_index:
        slides = [get_slide(pres, slide_index)]
    else:
        slides = [pres.Slides(i) for i in range(1, pres.Slides.Count + 1)]

    total_matches = 0
    total_replacements = 0
    results = []

    for slide in slides:
        slide_matches = 0
        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            if shape.HasTextFrame:
                try:
                    text = shape.TextFrame.TextRange.Text
                    count = text.count(find_text)
                    if count > 0:
                        slide_matches += count
                        if not preview_only:
                            shape.TextFrame.TextRange.Text = text.replace(find_text, replace_text)
                            total_replacements += count
                except Exception:
                    pass

        if slide_matches > 0:
            total_matches += slide_matches
            results.append({
                "slide_index": slide.SlideIndex,
                "matches": slide_matches,
            })

    return {
        "success": True,
        "presentation": pres.Name,
        "find_text": find_text,
        "replace_text": replace_text,
        "preview_only": preview_only,
        "total_matches": total_matches,
        "total_replacements": total_replacements,
        "slides": results,
        "message": f"{'Found' if preview_only else 'Replaced'} {total_matches} matches across {len(results)} slides.",
    }


def save_presentation_sync(
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Save the active presentation."""
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)

    pres.Save()

    return {
        "success": True,
        "presentation": pres.Name,
        "path": pres.FullName,
        "message": f"Presentation '{pres.Name}' saved.",
    }

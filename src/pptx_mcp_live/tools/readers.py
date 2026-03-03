"""Reading tools: read text, notes, shapes, comments."""

from typing import Any, Dict, List, Optional

from ..core.connection import get_powerpoint, get_presentation, get_slide, get_shape
from ..core.errors import ToolError


def read_slide_text_sync(
    slide_index: int,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Extract all text from all shapes on a slide."""
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)

    texts = []
    for i in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes(i)
        if shape.HasTextFrame:
            try:
                text = shape.TextFrame.TextRange.Text
                if text and text.strip():
                    texts.append({
                        "shape_name": shape.Name,
                        "shape_index": i,
                        "text": text,
                    })
            except Exception:
                pass

    return {
        "success": True,
        "presentation": pres.Name,
        "slide_index": slide_index,
        "texts": texts,
        "shape_count": len(texts),
    }


def read_slide_notes_sync(
    slide_index: int,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Read speaker notes from a slide."""
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)

    notes = None
    try:
        notes = slide.NotesPage.Shapes(2).TextFrame.TextRange.Text
    except Exception:
        pass

    return {
        "success": True,
        "presentation": pres.Name,
        "slide_index": slide_index,
        "notes": notes,
    }


def read_shape_text_sync(
    slide_index: int,
    shape_ref,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Read text from a specific shape by name or index."""
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)
    shape = get_shape(slide, shape_ref)

    if not shape.HasTextFrame:
        raise ToolError(f"Shape '{shape_ref}' does not contain text.")

    text = shape.TextFrame.TextRange.Text

    return {
        "success": True,
        "presentation": pres.Name,
        "slide_index": slide_index,
        "shape_name": shape.Name,
        "text": text,
    }


def get_comments_sync(
    slide_index: int,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Get all comments from a slide."""
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)

    comments = []
    try:
        for i in range(1, slide.Comments.Count + 1):
            c = slide.Comments(i)
            comment_data = {
                "index": i,
                "author": c.Author,
                "author_initials": c.AuthorInitials,
                "text": c.Text,
                "date": str(c.DateTime),
                "left": c.Left,
                "top": c.Top,
            }
            # Check for replies (modern comments)
            try:
                if hasattr(c, 'Replies') and c.Replies.Count > 0:
                    replies = []
                    for j in range(1, c.Replies.Count + 1):
                        r = c.Replies(j)
                        replies.append({
                            "index": j,
                            "author": r.Author,
                            "text": r.Text,
                            "date": str(r.DateTime),
                        })
                    comment_data["replies"] = replies
                    comment_data["reply_count"] = c.Replies.Count
            except Exception:
                pass

            comments.append(comment_data)
    except Exception:
        pass

    return {
        "success": True,
        "presentation": pres.Name,
        "slide_index": slide_index,
        "comments": comments,
        "count": len(comments),
    }

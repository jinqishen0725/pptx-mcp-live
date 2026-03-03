"""Comment tools: add, reply, resolve, delete."""

from typing import Any, Dict, Optional

from ..core.connection import get_powerpoint, get_presentation, get_slide
from ..core.errors import ToolError


def add_comment_sync(
    slide_index: int,
    text: str,
    author: str = "",
    author_initials: str = "",
    left: float = 0,
    top: float = 0,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Add a comment to a slide.

    Args:
        slide_index: 1-based slide index.
        text: Comment text.
        author: Author name.
        author_initials: Author initials.
        left: X position on slide.
        top: Y position on slide.
    """
    if not text or not text.strip():
        raise ToolError("Comment text cannot be empty.")

    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)

    try:
        comment = slide.Comments.Add2(left, top, author, author_initials, text)
        return {
            "success": True,
            "presentation": pres.Name,
            "slide_index": slide_index,
            "comment_index": comment.Index if hasattr(comment, 'Index') else slide.Comments.Count,
            "text": text,
            "author": author,
            "message": f"Comment added to slide {slide_index}.",
        }
    except Exception:
        # Fallback to Add for older PowerPoint
        try:
            comment = slide.Comments.Add(left, top, author, author_initials, text)
            return {
                "success": True,
                "presentation": pres.Name,
                "slide_index": slide_index,
                "text": text,
                "author": author,
                "message": f"Comment added to slide {slide_index}.",
            }
        except Exception as e:
            raise ToolError(f"Failed to add comment: {e}") from e


def delete_comment_sync(
    slide_index: int,
    comment_index: int,
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Delete a comment by index."""
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)

    if comment_index < 1 or comment_index > slide.Comments.Count:
        raise ToolError(
            f"Comment index {comment_index} out of range (1-{slide.Comments.Count})."
        )

    slide.Comments(comment_index).Delete()

    return {
        "success": True,
        "presentation": pres.Name,
        "slide_index": slide_index,
        "deleted_index": comment_index,
        "remaining": slide.Comments.Count,
        "message": f"Comment {comment_index} deleted from slide {slide_index}.",
    }

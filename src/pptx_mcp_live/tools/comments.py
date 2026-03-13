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


def get_all_comments_sync(
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Get all comments from all slides in the presentation.

    Returns comments grouped by slide, avoiding the need to loop slide by slide.
    """
    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)

    all_comments = []
    total = 0

    for si in range(1, pres.Slides.Count + 1):
        slide = pres.Slides(si)
        try:
            if slide.Comments.Count == 0:
                continue
        except Exception:
            continue

        slide_comments = []
        for ci in range(1, slide.Comments.Count + 1):
            c = slide.Comments(ci)
            comment_data = {
                "index": ci,
                "author": c.Author,
                "author_initials": c.AuthorInitials,
                "text": c.Text,
                "date": str(c.DateTime),
                "left": c.Left,
                "top": c.Top,
            }
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
            except Exception:
                pass

            slide_comments.append(comment_data)

        if slide_comments:
            all_comments.append({
                "slide_index": si,
                "comments": slide_comments,
                "count": len(slide_comments),
            })
            total += len(slide_comments)

    return {
        "success": True,
        "presentation": pres.Name,
        "slides_with_comments": len(all_comments),
        "total_comments": total,
        "slides": all_comments,
    }


def reply_to_comment_sync(
    slide_index: int,
    comment_index: int,
    text: str,
    author: str = "",
    author_initials: str = "",
    presentation_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Reply to an existing comment on a slide.

    Args:
        slide_index: 1-based slide index.
        comment_index: 1-based index of the comment to reply to.
        text: Reply text.
        author: Author name for the reply.
        author_initials: Author initials for the reply.
    """
    if not text or not text.strip():
        raise ToolError("Reply text cannot be empty.")

    app = get_powerpoint()
    pres = get_presentation(app, presentation_name)
    slide = get_slide(pres, slide_index)

    if comment_index < 1 or comment_index > slide.Comments.Count:
        raise ToolError(
            f"Comment index {comment_index} out of range (1-{slide.Comments.Count})."
        )

    comment = slide.Comments(comment_index)

    try:
        # Modern comments support Replies
        reply = comment.AddReply(text)
        return {
            "success": True,
            "presentation": pres.Name,
            "slide_index": slide_index,
            "parent_comment_index": comment_index,
            "text": text,
            "message": f"Reply added to comment {comment_index} on slide {slide_index}.",
        }
    except Exception:
        # Fallback: add a new comment near the original with [Reply] prefix
        try:
            left = comment.Left + 10
            top = comment.Top + 10
            reply_text = f"[Reply to {comment.Author}] {text}"
            try:
                slide.Comments.Add2(left, top, author, author_initials, reply_text)
            except Exception:
                slide.Comments.Add(left, top, author, author_initials, reply_text)
            return {
                "success": True,
                "presentation": pres.Name,
                "slide_index": slide_index,
                "parent_comment_index": comment_index,
                "text": reply_text,
                "fallback": True,
                "message": f"Reply added as new comment near comment {comment_index} (threaded replies not supported in this PowerPoint version).",
            }
        except Exception as e:
            raise ToolError(f"Failed to reply to comment: {e}") from e
